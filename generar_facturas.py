# -*- coding: utf-8 -*-
"""
Genera una factura PDF (con la estructura de pdf_ej.pdf) por cada fila del
Excel que tenga relleno el campo 'Factura'.

Uso:
    py generar_facturas.py

Salida: carpeta ./facturas/  ->  un PDF por factura.
"""

import os
import re
from datetime import datetime

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black
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

# Datos fijos de la empresa (iguales en todas las facturas)
EMPRESA = [
    ("Campus Education Spain, SLU", "Helvetica-Bold", 9),
    ("NIF: B27847243", "Helvetica", 6),
    ("Domingo Bueno 27, Porriño", "Helvetica", 6),
    ("36400 (Pontevedra) España", "Helvetica", 6),
    ("(+34) 650 319 130 , (+34) 627 671 051", "Helvetica", 6),
    ("www.campusspain.es", "Helvetica", 6),
]


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
    """'Khaled Chaouch (U356071)' -> ('Khaled Chaouch', 'U356071')."""
    if not texto:
        return "", ""
    texto = str(texto).strip()
    m = re.search(r"\(([^)]*)\)\s*$", texto)
    if m:
        nombre = texto[: m.start()].strip()
        pasaporte = m.group(1).strip()
        return nombre, pasaporte
    return texto, ""


def cap_nacionalidad(valor):
    if not valor:
        return ""
    return str(valor).strip().capitalize()


# ---------------------------------------------------------------------------
# Dibujo de la factura
# ---------------------------------------------------------------------------
def label_value(c, x, y, label, value, size=10, label_font="Helvetica-Bold",
                value_font="Helvetica"):
    """Dibuja 'label: value' donde label va en negrita."""
    c.setFont(label_font, size)
    c.setFillColor(black)
    c.drawString(x, y, label)
    w = c.stringWidth(label, label_font, size)
    c.setFont(value_font, size)
    c.drawString(x + w + 3, y, value)


def generar_factura(c, datos):
    """Dibuja una pagina de factura en el canvas y hace showPage()."""
    def T(top):
        """Convierte coordenada 'desde arriba' en coordenada PDF."""
        return PAGE_H - top

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

    # --- Invoice / Date ---
    label_value(c, LEFT, T(135), "Invoice:", datos["factura"], size=10)
    label_value(c, LEFT + 230, T(135), "Date:", datos["fecha"], size=10)

    # --- Datos del estudiante (indentado) ---
    sx = LEFT + 130
    label_value(c, sx, T(180), "Student:", datos["alumno"], size=10)
    label_value(c, sx, T(195), "Passport:", datos["pasaporte"], size=10)
    label_value(c, sx, T(210), "Nationality:", datos["nacionalidad"], size=10)

    # --- Tabla de conceptos ---
    tabla_top = 280
    col_amount_x = RIGHT          # importe alineado a la derecha
    concept_right = RIGHT - 95    # ancho de la columna de concepto

    # Cabecera
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString((LEFT + concept_right) / 2, T(tabla_top), "Concepts")
    c.drawRightString(col_amount_x, T(tabla_top - 5), "Amount")
    c.drawRightString(col_amount_x, T(tabla_top + 7), "(€)")

    # Linea bajo la cabecera
    c.setLineWidth(0.8)
    c.line(LEFT, T(tabla_top + 18), RIGHT, T(tabla_top + 18))

    # Fila de concepto (texto con ajuste de linea)
    estilo = ParagraphStyle("concepto", fontName="Helvetica", fontSize=10,
                            leading=13, textColor=black)
    p = Paragraph(datos["concepto"], estilo)
    avail_w = concept_right - LEFT
    pw, ph = p.wrap(avail_w, 200)
    fila_top = tabla_top + 30
    p.drawOn(c, LEFT, T(fila_top) - ph + 11)
    c.setFont("Helvetica", 10)
    c.drawRightString(col_amount_x, T(fila_top), datos["importe"])

    # Linea bajo la fila
    linea_y = fila_top + max(ph, 13) + 4
    c.line(LEFT, T(linea_y), RIGHT, T(linea_y))

    # --- Total ---
    total_y = linea_y + 28
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(concept_right + 20, T(total_y), "Total amount:")
    c.drawRightString(col_amount_x, T(total_y), datos["importe"])

    # --- Nota IVA ---
    nota_y = total_y + 40
    c.setFont("Helvetica", 7)
    c.drawCentredString(
        PAGE_W / 2, T(nota_y),
        "* The services of this invoice are not taxed by Spanish VAT according "
        "to articles 20.Uno.9 and 69.Dos of Spanish Law 37/1992 on Value Added Tax",
    )

    # --- Texto de pago ---
    estilo_pago = ParagraphStyle("pago", fontName="Helvetica", fontSize=10,
                                 leading=14, textColor=black, alignment=4)
    pago = Paragraph(
        "The amounts mentioned are to be paid through the following link, "
        "selecting the payment option that suits you best: "
        '<a href="https://campusspain.flywire.com/">'
        '<font color="#1155cc"><u>https://campusspain.flywire.com/</u></font></a>',
        estilo_pago,
    )
    pago_top = nota_y + 35
    w1, h1 = pago.wrap(RIGHT - LEFT, 100)
    pago.drawOn(c, LEFT, T(pago_top) - h1 + 10)

    c.setFont("Helvetica", 10)
    cuenta_intro_y = pago_top + h1 + 6
    c.drawString(LEFT, T(cuenta_intro_y),
                 "The amounts mentioned can also be paid to the following bank account:")

    # --- Datos bancarios ---
    bx = LEFT + 20
    by = cuenta_intro_y + 28
    banco = [
        ("Banco:", "Banco Santander (BSCH)"),
        ("Dirección:", "Policarpo Sanz 44, Vigo, 36202 (Pontevedra) España"),
        ("Nº de cuenta bancaria account number:", "0049 5470 50 2616149512"),
        ("IBAN:", "ES1900495470502616149512"),
        ("Código SWIFT", "BSCHESMM"),
        ("Destinatario:", "Campus Education Spain, S.L.U. NIF: B27847243"),
    ]
    for label, value in banco:
        label_value(c, bx, T(by), label, value, size=10)
        by += 16

    # --- Firma (abajo izquierda) ---
    firma_w = 230
    firma_h = firma_w * 132 / 479
    firma_top = by + 20
    c.drawImage(ImageReader(FIRMA), LEFT + 30, T(firma_top + firma_h),
                width=firma_w, height=firma_h, mask="auto",
                preserveAspectRatio=True)

    # --- Sello PAID (abajo derecha) ---
    box_w, box_h = 90, 50
    box_x = RIGHT - box_w
    box_y_top = firma_top + 5
    c.setStrokeColor(GREEN)
    c.setLineWidth(1.5)
    c.roundRect(box_x, T(box_y_top + box_h), box_w, box_h, 4, stroke=1, fill=0)
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(box_x + box_w / 2, T(box_y_top + 28), "PAID")
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
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
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

    c_factura = col("factura", "invoice")
    c_fecha = col("fecha", "date")
    c_alumno = col("alumno", "nombre/razón social", "nombre/razon social",
                   "nombre", "student", "estudiante")
    c_nac = col("nacionalidad", "nationality")
    c_concepto = col("concepto", "concept", "concepts")
    c_importe = col("€", "importe", "amount")

    if not c_factura:
        raise SystemExit("No se encontro la columna 'Factura' en el Excel.")

    os.makedirs(OUT_DIR, exist_ok=True)

    generadas = 0
    for row in ws.iter_rows(min_row=2):
        factura = row[c_factura - 1].value
        if factura is None or str(factura).strip() == "":
            continue  # solo filas con 'Factura' rellena

        nombre, pasaporte = parse_alumno(row[c_alumno - 1].value if c_alumno else "")
        datos = {
            "factura": str(factura).strip(),
            "fecha": fmt_fecha(row[c_fecha - 1].value if c_fecha else None),
            "fecha_completa": fmt_fecha(row[c_fecha - 1].value if c_fecha else None,
                                        con_anio_completo=True),
            "alumno": nombre,
            "pasaporte": pasaporte,
            "nacionalidad": cap_nacionalidad(row[c_nac - 1].value if c_nac else ""),
            "concepto": str(row[c_concepto - 1].value or "").strip() if c_concepto else "",
            "importe": fmt_importe(row[c_importe - 1].value if c_importe else 0),
        }

        nombre_archivo = re.sub(r"[^\w\-]", "_", datos["factura"]) + ".pdf"
        ruta = os.path.join(OUT_DIR, nombre_archivo)
        try:
            c = canvas.Canvas(ruta, pagesize=A4)
            c.setTitle(f"Invoice {datos['factura']}")
            generar_factura(c, datos)
            c.save()
        except PermissionError:
            print(f"  [OMITIDA] {nombre_archivo} (archivo abierto/bloqueado, "
                  "cierralo y vuelve a ejecutar)")
            continue
        generadas += 1
        print(f"  [OK] {nombre_archivo}")

    print(f"\nHecho. {generadas} facturas generadas en: {OUT_DIR}")


if __name__ == "__main__":
    main()
