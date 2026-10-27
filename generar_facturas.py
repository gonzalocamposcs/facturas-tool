# -*- coding: utf-8 -*-
"""
Genera una factura PDF bilingüe (español con el inglés debajo, igual que la
factura LCA del CRM: resources/views/pdf/factura.blade.php) por cada fila del
Excel que tenga relleno el número de factura.

Uso:
    py generar_facturas.py                  -> lee ./excel.xlsx
    py generar_facturas.py ruta\\libro.xlsx  -> lee ese Excel

Salida: carpeta ./facturas/  ->  un PDF por factura.
"""

import os
import re
import sys
from datetime import datetime
from xml.sax.saxutils import escape

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, black
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

# ---------------------------------------------------------------------------
# Configuracion / rutas
# ---------------------------------------------------------------------------
BASE = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE, "excel.xlsx")
LOGO = os.path.join(BASE, "logo.jpeg")
FIRMA = os.path.join(BASE, "firma.jpeg")
OUT_DIR = os.path.join(BASE, "facturas")

PAGE_W, PAGE_H = A4
LEFT = 70          # margen izquierdo
RIGHT = PAGE_W - 70  # margen derecho (x del borde derecho del contenido)
GREEN = HexColor("#2e8b3d")
BLUE = HexColor("#1155cc")
GRIS = HexColor("#555555")

# Datos fijos de la empresa (iguales en todas las facturas)
EMPRESA = [
    ("Campus Education Spain, SLU", "Helvetica-Bold", 9),
    ("NIF: B27847243", "Helvetica", 6),
    ("Domingo Bueno 27, Porriño", "Helvetica", 6),
    ("36400 (Pontevedra) España", "Helvetica", 6),
    ("(+34) 650 319 130 , (+34) 627 671 051", "Helvetica", 6),
    ("www.campusspain.es", "Helvetica", 6),
]

# Datos bancarios: (etiqueta ES, etiqueta EN, valor). Mismas etiquetas que el CRM.
BANCO = [
    ("Banco:", "Bank", "Banco Santander (BSCH)"),
    ("Dirección:", "Address", "Policarpo Sanz 44, Vigo, 36202 (Pontevedra) España"),
    ("Cuenta bancaria/IBAN:", "Bank account/IBAN",
     "0049 5470 50 2616149512 / ES1900495470502616149512"),
    ("Código SWIFT:", "Swift Code", "BSCHESMM"),
    ("Beneficiario:", "Beneficiary", "Campus Education Spain, S.L.U. NIF: B27847243"),
]

# Estilos: el español a 10 pt y el inglés debajo, más pequeño y en gris
# (la clase .en de la plantilla del CRM).
ES = ParagraphStyle("es", fontName="Helvetica", fontSize=10, leading=12.5,
                    textColor=black)
EN = ParagraphStyle("en", fontName="Helvetica", fontSize=7, leading=8.5,
                    textColor=GRIS)


# ---------------------------------------------------------------------------
# Traducción de conceptos (port de Factura::conceptoBilingue() del CRM)
# ---------------------------------------------------------------------------
# Frases en inglés => español. Se aplican en este orden (las largas antes que
# sus trozos). Copia de Factura::FRASES_CONCEPTO: si se añade una allí, aquí.
FRASES_CONCEPTO = [
    ("at the at the ", "at the "),
    ("Reservation for the ", "Reserva del "),
    ("Reservation for ", "Reserva del "),
    ("Español LCA: Language, Culture and Adaptation to University course",
     "curso de Español LCA: Lengua, Cultura y Adaptación a la Universidad"),
    ("(aff./ Rey Juan Carlos University of Madrid)",
     "(adscrito a la Universidad Rey Juan Carlos de Madrid)"),
    ("(affiliated with Rey Juan Carlos University of Madrid)",
     "(adscrito a la Universidad Rey Juan Carlos de Madrid)"),
    ("IUNIT affiliated to Rey Juan Carlos University Madrid",
     "IUNIT (adscrito a la Universidad Rey Juan Carlos de Madrid)"),
    ("IUNIT affiliated to Rey Juan Carlos Madrid",
     "IUNIT (adscrito a la Universidad Rey Juan Carlos de Madrid)"),
    ("or at the University of Vigo", "o en la Universidad de Vigo"),
    ("or University of Vigo", "o en la Universidad de Vigo"),
    ("at the University of Vigo", "en la Universidad de Vigo"),
    ("or at IUNIT", "o en el IUNIT"),
    ("or IUNIT", "o en el IUNIT"),
    ("at the IUNIT", "en el IUNIT"),
    ("at IUNIT", "en el IUNIT"),
    (" until ", " hasta el "),
    ("Guardianship Notary deed", "Escritura notarial de tutela"),
    ("Notarial deed", "Escritura notarial"),
    ("Notarized affidavit", "Declaración jurada notarial"),
    ("Translations", "Traducciones"),
    ("Translation", "Traducción"),
    ("Extra month Swisscare", "Mes adicional de Swisscare"),
    ("Spanish level assessment", "Prueba de nivel de español"),
    ("Accommodation deposit", "Depósito de alojamiento"),
    ("Other Refund", "Otro reembolso"),
]

MESES_CONCEPTO = {
    "January": "enero", "February": "febrero", "March": "marzo",
    "April": "abril", "May": "mayo", "June": "junio",
    "July": "julio", "August": "agosto", "September": "septiembre",
    "October": "octubre", "November": "noviembre", "December": "diciembre",
}
_MESES = "|".join(MESES_CONCEPTO)
_RESTO_INGLES = re.compile(
    r"\b(the|at|for|and|of|or|with|to|is|until|course|University|intake|"
    r"installment|Reservation|Language|Culture|Adaptation|price|total|"
    r"remaining|fees|" + _MESES + r")\b", re.I)


def concepto_bilingue(texto):
    """Concepto como (español, inglés). Igual que en el CRM, pero conserva los
    saltos de línea del Excel. Si tras traducir queda alguna palabra inglesa,
    no arriesga una traducción a medias y devuelve (texto, None)."""
    lineas = [re.sub(r"[ \t ]+", " ", l).strip()
              for l in str(texto or "").splitlines()]
    texto = "\n".join(l for l in lineas if l)
    if texto == "":
        return "", None

    es = texto
    for ingles, espanol in FRASES_CONCEPTO:
        es = es.replace(ingles, espanol)
    # «2nd installment:» / «1st installment» → «2.º plazo:» / «1.er plazo».
    # (El CRM solo reconoce la forma con dos puntos; en el Excel van sin ellos.)
    es = re.sub(r"\b(\d+)(?:st|nd|rd|th) installment\b",
                lambda m: m[1] + (".er" if m[1] in ("1", "3") else ".º") + " plazo", es)
    # «September 2026 intake» → «convocatoria de septiembre de 2026»
    es = re.sub(r"\b(" + _MESES + r") (\d{4}) intake\b",
                lambda m: f"convocatoria de {MESES_CONCEPTO[m[1]]} de {m[2]}", es)
    # «October 1st, 2025» → «1 de octubre de 2025»
    es = re.sub(r"\b(" + _MESES + r") (\d{1,2})(?:st|nd|rd|th)?,? (\d{4})\b",
                lambda m: f"{m[2]} de {MESES_CONCEPTO[m[1]]} de {m[3]}", es)

    if es == texto or _RESTO_INGLES.search(es):
        return texto, None
    return es[:1].upper() + es[1:], texto


# ---------------------------------------------------------------------------
# Utilidades de formato
# ---------------------------------------------------------------------------
def fmt_importe(valor):
    """7900 -> '7.900,00 €'  (formato español)."""
    if valor is None:
        valor = 0
    s = f"{float(valor):,.2f}"          # '7,900.00'
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")  # '7.900,00'
    return f"{s} €"


def fmt_fecha(valor, con_anio_completo=False):
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y" if con_anio_completo else "%d/%m/%y")
    return str(valor) if valor is not None else ""


def parse_alumno(texto):
    """'Khaled Chaouch (U356071)' -> ('Khaled Chaouch', 'U356071').
    También 'Hadia - AF6497843' -> ('Hadia', 'AF6497843')."""
    if not texto:
        return "", ""
    texto = str(texto).strip()
    m = re.search(r"\(([^)]*)\)\s*$", texto)
    if m:
        nombre = texto[: m.start()].strip()
        pasaporte = m.group(1).strip()
        return nombre, pasaporte
    m = re.search(r"\s+-\s+(\S+)\s*$", texto)
    if m:
        return texto[: m.start()].strip(), m.group(1).strip()
    return texto, ""


def cap_nacionalidad(valor):
    if not valor:
        return ""
    return str(valor).strip().capitalize()


# ---------------------------------------------------------------------------
# Dibujo de la factura
# ---------------------------------------------------------------------------
def T(top):
    """Convierte coordenada 'desde arriba' en coordenada PDF."""
    return PAGE_H - top


def parrafo(c, html, estilo, x, top, ancho, alineacion=TA_LEFT):
    """Dibuja un párrafo con su borde superior en 'top'; devuelve dónde acaba."""
    if alineacion != estilo.alignment:
        estilo = ParagraphStyle(estilo.name + str(alineacion), parent=estilo,
                                alignment=alineacion)
    p = Paragraph(html, estilo)
    _, h = p.wrap(ancho, PAGE_H)
    p.drawOn(c, x, T(top + h))
    return top + h


def bilingue(c, html_es, texto_en, x, top, ancho, alineacion=TA_LEFT):
    """Español y, debajo, el inglés pequeño. Devuelve dónde acaba."""
    top = parrafo(c, html_es, ES, x, top, ancho, alineacion)
    if texto_en:
        top = parrafo(c, escape(texto_en).replace("\n", "<br/>"), EN, x, top + 1,
                      ancho, alineacion)
    return top


def etiqueta(es, valor):
    """'<b>Factura:</b> CS2026-001'."""
    return f"<b>{escape(es)}</b> {escape(valor)}"


def generar_factura(c, datos):
    """Dibuja una pagina de factura en el canvas y hace showPage()."""
    # --- Logo (arriba a la izquierda) ---
    logo_w = 165
    logo_h = logo_w * 353 / 1600
    c.drawImage(ImageReader(LOGO), LEFT, T(40 + logo_h), width=logo_w,
                height=logo_h, mask="auto", preserveAspectRatio=True)

    # --- Bloque empresa (arriba a la derecha, alineado a la derecha) ---
    y = 40
    for texto, font, size in EMPRESA:
        c.setFont(font, size)
        if texto == "www.campusspain.es":
            c.setFillColor(BLUE)
            c.drawRightString(RIGHT, T(y + size), texto)
            tw = c.stringWidth(texto, font, size)
            c.linkURL("https://www.campusspain.es",
                      (RIGHT - tw, T(y + size) - 1, RIGHT, T(y + size) + size),
                      relative=0)
            c.setFillColor(black)
        else:
            c.drawRightString(RIGHT, T(y + size), texto)
        y += size + 2.5

    # --- Factura / Fecha ---
    bilingue(c, etiqueta("Factura:", datos["factura"]), "Invoice", LEFT, 125, 220)
    top = bilingue(c, etiqueta("Fecha:", datos["fecha"]), "Date", LEFT + 230, 125, 200)

    # --- Datos del estudiante (indentado) ---
    sx = LEFT + 130
    top += 22
    for es, en, valor in [("Estudiante:", "Student", datos["alumno"]),
                          ("Pasaporte:", "Passport", datos["pasaporte"]),
                          ("Nacionalidad:", "Nationality", datos["nacionalidad"])]:
        top = bilingue(c, etiqueta(es, valor), en, sx, top, RIGHT - sx) + 5

    # --- Tabla de conceptos ---
    concept_right = RIGHT - 95    # ancho de la columna de concepto
    importe_x = concept_right     # columna del importe: de aquí a RIGHT
    importe_w = RIGHT - importe_x
    c.setLineWidth(0.8)

    # Cabecera
    cab_top = top + 20
    h1 = bilingue(c, "<b>Conceptos</b>", "Concepts", LEFT, cab_top,
                  concept_right - LEFT, TA_CENTER)
    h2 = bilingue(c, "<b>Importe (€)</b>", "Amount (€)", importe_x, cab_top,
                  importe_w, TA_CENTER)
    linea = max(h1, h2) + 6
    c.line(LEFT, T(linea), RIGHT, T(linea))

    # Fila de concepto
    fila_top = linea + 7
    fin = bilingue(c, escape(datos["concepto_es"]).replace("\n", "<br/>"),
                   datos["concepto_en"], LEFT, fila_top, concept_right - LEFT - 15)
    parrafo(c, escape(datos["importe"]), ES, importe_x, fila_top, importe_w, TA_CENTER)
    linea = fin + 7
    c.line(LEFT, T(linea), RIGHT, T(linea))

    # --- Importe total ---
    total_top = linea + 8
    bilingue(c, "<b>Importe total:</b>", "Total amount", LEFT,
             total_top, concept_right - LEFT - 10, TA_RIGHT)
    top = parrafo(c, f"<b>{escape(datos['importe'])}</b>", ES, importe_x,
                  total_top, importe_w, TA_CENTER)
    top = total_top + ES.leading + EN.leading + 1

    # --- Nota IVA ---
    nota_es = ParagraphStyle("nota_es", parent=ES, fontSize=6.5, leading=8)
    nota_en = ParagraphStyle("nota_en", parent=EN, fontSize=6.5, leading=8)
    top = parrafo(c, "* Los servicios de esta factura no están sujetos al IVA español "
                  "según los artículos 20.Uno.9 y 69.Dos de la Ley 37/1992 del Impuesto "
                  "sobre el Valor Añadido", nota_es, LEFT - 20, top + 14,
                  RIGHT - LEFT + 40, TA_CENTER)
    top = parrafo(c, "* The services of this invoice are not taxed by Spanish VAT "
                  "according to articles 20.Uno.9 and 69.Dos of Spanish Law 37/1992 on "
                  "Value Added Tax", nota_en, LEFT - 20, top + 1, RIGHT - LEFT + 40,
                  TA_CENTER)

    # --- Texto de pago ---
    enlace = ('<a href="https://campusspain.flywire.com/">'
              '<font color="#1155cc"><u>https://campusspain.flywire.com/</u></font></a>')
    top = bilingue(c, "Los importes mencionados se abonarán a través del siguiente "
                   "enlace, seleccionando la opción de pago que más le convenga: " + enlace,
                   "The amounts mentioned are to be paid through the following link, "
                   "selecting the payment option that suits you best: "
                   "https://campusspain.flywire.com/",
                   LEFT, top + 14, RIGHT - LEFT, TA_JUSTIFY)
    top = bilingue(c, "Las cantidades mencionadas también se pueden pagar en la "
                   "siguiente cuenta bancaria:",
                   "The amounts mentioned can also be paid to the following bank account:",
                   LEFT, top + 10, RIGHT - LEFT)

    # --- Datos bancarios ---
    bx = LEFT + 60
    top += 12
    banco = BANCO + [("Concepto:", "Concept", datos["concepto_transferencia"])]
    for es, en, valor in banco:
        top = bilingue(c, etiqueta(es, valor), en, bx, top, RIGHT - bx) + 4

    # --- Firma (abajo izquierda) ---
    firma_w = 230
    firma_h = firma_w * 132 / 479
    firma_top = top + 12
    c.drawImage(ImageReader(FIRMA), LEFT + 30, T(firma_top + firma_h),
                width=firma_w, height=firma_h, mask="auto",
                preserveAspectRatio=True)

    # --- Sello PAGADO / PAID (abajo derecha) ---
    box_w, box_h = 100, 50
    box_x = RIGHT - box_w
    box_y_top = firma_top + 5
    c.setStrokeColor(GREEN)
    c.setLineWidth(1.5)
    c.roundRect(box_x, T(box_y_top + box_h), box_w, box_h, 4, stroke=1, fill=0)
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(box_x + box_w / 2, T(box_y_top + 26), "PAGADO")
    c.setFont("Helvetica", 11)
    c.drawCentredString(box_x + box_w / 2, T(box_y_top + 41), "PAID")
    c.setFont("Helvetica", 9)
    c.drawCentredString(box_x + box_w / 2, T(box_y_top + box_h + 14),
                        datos["fecha_completa"])
    c.setFillColor(black)
    c.setStrokeColor(black)

    c.showPage()


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def main():
    excel = sys.argv[1] if len(sys.argv) > 1 else EXCEL
    wb = openpyxl.load_workbook(excel, data_only=True)
    ws = wb.active

    # Localizar columnas por cabecera (fila 1)
    headers = {}
    for col, cell in enumerate(ws[1], start=1):
        if cell.value:
            headers[str(cell.value).strip().lower()] = col

    def col(*nombres):
        for n in nombres:
            if n in headers:
                return headers[n]
        return None

    c_factura = col("factura", "nº de factura", "n° de factura", "no de factura",
                    "número de factura", "numero de factura", "invoice")
    c_fecha = col("fecha", "date")
    c_alumno = col("alumno", "nombre/razón social", "nombre/razon social",
                   "nombre", "student", "estudiante")
    c_nac = col("nacionalidad", "nationality")
    c_concepto = col("concepto", "concept", "concepts")
    c_importe = col("€", "importe", "cantidad", "amount")

    if not c_factura:
        raise SystemExit("No se encontro la columna 'Factura' en el Excel.")

    os.makedirs(OUT_DIR, exist_ok=True)

    generadas = 0
    sin_traducir = []
    for row in ws.iter_rows(min_row=2):
        factura = row[c_factura - 1].value
        if factura is None or str(factura).strip() == "":
            continue  # solo filas con 'Factura' rellena

        nombre, pasaporte = parse_alumno(row[c_alumno - 1].value if c_alumno else "")
        concepto_es, concepto_en = concepto_bilingue(
            row[c_concepto - 1].value if c_concepto else "")
        datos = {
            "factura": str(factura).strip(),
            "fecha": fmt_fecha(row[c_fecha - 1].value if c_fecha else None),
            "fecha_completa": fmt_fecha(row[c_fecha - 1].value if c_fecha else None,
                                        con_anio_completo=True),
            "alumno": nombre,
            "pasaporte": pasaporte,
            "nacionalidad": cap_nacionalidad(row[c_nac - 1].value if c_nac else ""),
            "concepto_es": concepto_es,
            "concepto_en": concepto_en,
            "concepto_transferencia": f"{nombre} - {pasaporte}",
            "importe": fmt_importe(row[c_importe - 1].value if c_importe else 0),
        }
        if concepto_es and not concepto_en:
            sin_traducir.append(datos["factura"])

        nombre_archivo = re.sub(r"[^\w\-]", "_", datos["factura"]) + ".pdf"
        ruta = os.path.join(OUT_DIR, nombre_archivo)
        try:
            c = canvas.Canvas(ruta, pagesize=A4)
            c.setTitle(f"Factura / Invoice {datos['factura']}")
            generar_factura(c, datos)
            c.save()
        except PermissionError:
            print(f"  [OMITIDA] {nombre_archivo} (archivo abierto/bloqueado, "
                  "cierralo y vuelve a ejecutar)")
            continue
        generadas += 1
        print(f"  [OK] {nombre_archivo}")

    print(f"\nHecho. {generadas} facturas generadas en: {OUT_DIR}")
    if sin_traducir:
        print(f"{len(sin_traducir)} con el concepto solo en inglés (no se reconoció "
              f"la frase): {', '.join(sin_traducir)}")


if __name__ == "__main__":
    main()
