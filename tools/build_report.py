#!/usr/bin/env python
"""Build the Word deliverable for the DAE S13 extra lab.

Every command output quoted here comes from docs/evidencia/*.txt, produced by
real runs during the work. The script only formats that material; it does not
invent results. Run it with the globally installed python-docx:

    python tools/build_report.py
"""
from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = ROOT / "docs" / "evidencia"
OUTPUT = ROOT / "entregable" / "DAE_S13_Extra_Entregable.docx"
REPO = "https://github.com/Fernando-Correa-27/DAE_Semana_13_Extra"

DARK = RGBColor(0x1F, 0x3A, 0x5F)
ACCENT = RGBColor(0x1F, 0x6F, 0x8B)
GRAY = RGBColor(0x55, 0x5F, 0x6B)
RED = RGBColor(0xA6, 0x1B, 0x1B)
GREEN = RGBColor(0x1E, 0x6B, 0x3A)


# ---------------------------------------------------------------------------
# Small docx helpers
# ---------------------------------------------------------------------------
PPR_AFTER_BORDER = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
PPR_AFTER_SHADE = PPR_AFTER_BORDER[1:]
TBLPR_AFTER_BORDERS = (
    "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
)
TCPR_AFTER_SHADE = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
    "w:tcPrChange",
)


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.insert_element_before(shd, *TCPR_AFTER_SHADE)


def table_borders(table, color="B7C4CF", size="4"):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    table._tbl.tblPr.insert_element_before(borders, *TBLPR_AFTER_BORDERS)


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def keep_row_together(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    cant.set(qn("w:val"), "true")
    tr_pr.append(cant)


def set_run_font(run, name, size=None, bold=None, color=None, italic=None):
    run.font.name = name
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:ascii"), name)
    fonts.set(qn("w:hAnsi"), name)
    fonts.set(qn("w:eastAsia"), name)


def para(doc, text="", size=None, bold=False, italic=False, color=None,
         align=None, space_after=4, space_before=0, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, "Calibri", size, bold, color, italic)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    return p


def rich(doc, parts, size=None, align=None, space_after=4, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    for text, bold in parts:
        run = p.add_run(text)
        set_run_font(run, "Calibri", size, bold)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    set_run_font(run, "Calibri", 10)
    p.paragraph_format.left_indent = Cm(0.65 + 0.55 * level)
    p.paragraph_format.space_after = Pt(2)
    return p


def numbered(doc, text):
    p = doc.add_paragraph(style="List Number")
    run = p.add_run(text)
    set_run_font(run, "Calibri", 10)
    p.paragraph_format.space_after = Pt(2)
    return p


def h1(doc, text, page_break=True):
    p = doc.add_paragraph()
    if page_break:
        p.paragraph_format.page_break_before = True
    run = p.add_run(text)
    set_run_font(run, "Calibri", 19, True, DARK)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    # bottom rule
    p_pr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1F6F8B")
    pbdr.append(bottom)
    p_pr.insert_element_before(pbdr, *PPR_AFTER_BORDER)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, "Calibri", 14, True, ACCENT)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    return p


def h3(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, "Calibri", 11.5, True, DARK)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    return p


def meta(doc, title):
    p = doc.add_paragraph()
    r1 = p.add_run("Alumno: [NOMBRE DEL ALUMNO]   ·   Desarrollo: ")
    set_run_font(r1, "Calibri", 8.5, True, GRAY)
    r2 = p.add_run(title)
    set_run_font(r2, "Calibri", 8.5, False, GRAY, italic=True)
    p.paragraph_format.space_after = Pt(4)
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "EEF3F7")
    p_pr.insert_element_before(shd, *PPR_AFTER_SHADE)


def evidence(doc, *names):
    label = ", ".join(f"docs/evidencia/{n}" for n in names)
    p = doc.add_paragraph()
    r = p.add_run("Evidencia: " + label)
    set_run_font(r, "Calibri", 8, False, GRAY, italic=True)
    p.paragraph_format.space_after = Pt(6)


def code(doc, text, size=8.2, fill="F4F6F8"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    shade(cell, fill)
    table_borders(table, "D0D7DE", "4")
    lines = text.strip("\n").split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        run = p.add_run(line if line else " ")
        set_run_font(run, "Consolas", size)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def table(doc, headers, rows, widths=None, size=8, header_fill="1F3A5F",
          first_col_bold=False, header_size=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    table_borders(t)
    header = t.rows[0]
    repeat_header(header)
    for i, text in enumerate(headers):
        cell = header.cells[i]
        shade(cell, header_fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        run = p.add_run(text)
        set_run_font(run, "Calibri", header_size or size, True, RGBColor(0xFF, 0xFF, 0xFF))
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        if widths:
            cell.width = Cm(widths[i])
    for index, row in enumerate(rows):
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cell = cells[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            if index % 2 == 1:
                shade(cell, "F3F7FA")
            p = cell.paragraphs[0]
            run = p.add_run(str(value))
            set_run_font(run, "Calibri", size, first_col_bold and i == 0)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.line_spacing = 1.0
            if widths:
                cell.width = Cm(widths[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def landscape(doc):
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    return section


def portrait(doc):
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(1.9)
    section.bottom_margin = Cm(1.9)
    return section


def page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    set_run_font(run, "Calibri", 8, False, GRAY)


def header_footer(doc):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DAE · Semana 13 extra · El carrito revisado por pruebas  |  ")
    set_run_font(r, "Calibri", 8, False, GRAY)
    page_number(p)
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("TECSUP · Desarrollo de Aplicaciones Empresariales")
    set_run_font(hr, "Calibri", 8, False, GRAY)


def usage_rows():
    text = (ROOT / "docs" / "uso_de_ia.md").read_text(encoding="utf-8")
    rows = []
    for line in text.splitlines():
        if re.match(r"^\|\s*\d+\s*\|", line):
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cells) == 6:
                rows.append(cells)
    return rows


# ---------------------------------------------------------------------------
# Document
# ---------------------------------------------------------------------------
doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)
style.paragraph_format.line_spacing = 1.12
style.paragraph_format.space_after = Pt(4)
doc.styles["List Bullet"].font.name = "Calibri"
doc.styles["List Number"].font.name = "Calibri"

section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
section.top_margin = Cm(1.9)
section.bottom_margin = Cm(1.9)
header_footer(doc)

# ------------------------------- Portada -----------------------------------
for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph()
r = p.add_run("TECSUP")
set_run_font(r, "Calibri", 13, True, ACCENT)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(0)

p = doc.add_paragraph()
r = p.add_run("Desarrollo de Aplicaciones Empresariales · 4.º ciclo")
set_run_font(r, "Calibri", 10, False, GRAY)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(28)

p = doc.add_paragraph()
r = p.add_run("El carrito revisado por pruebas")
set_run_font(r, "Calibri", 30, True, DARK)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)

p = doc.add_paragraph()
r = p.add_run("Laboratorio extra de la Semana 13")
set_run_font(r, "Calibri", 16, False, ACCENT)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)

p = doc.add_paragraph()
r = p.add_run("GLAB-S13-EXTRA-MMONTGOMERY-2026-02")
set_run_font(r, "Calibri", 10, False, GRAY, italic=True)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(34)

cover_rows = [
    ("Curso", "Desarrollo de Aplicaciones Empresariales (DAE)"),
    ("Actividad", "Laboratorio extra — Semana 13"),
    ("Docente", "Michael Montgomery Rosell"),
    ("Sección", "A"),
    ("Alumno", "[NOMBRE DEL ALUMNO]"),
    ("Integrantes", "[INTEGRANTES]"),
    ("Repositorio", REPO),
    ("Fecha", "2026"),
]
t = doc.add_table(rows=0, cols=2)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.autofit = False
table_borders(t, "C9D4DE")
for label, value in cover_rows:
    cells = t.add_row().cells
    cells[0].width = Cm(4.4)
    cells[1].width = Cm(12.2)
    shade(cells[0], "EAF1F7")
    pl = cells[0].paragraphs[0]
    rl = pl.add_run(label)
    set_run_font(rl, "Calibri", 9.5, True, DARK)
    pl.paragraph_format.space_after = Pt(2)
    pv = cells[1].paragraphs[0]
    rv = pv.add_run(value)
    set_run_font(rv, "Calibri", 9.5, label == "Repositorio", ACCENT if label == "Repositorio" else None)
    pv.paragraph_format.space_after = Pt(2)

doc.add_paragraph()
p = doc.add_paragraph()
r = p.add_run(
    "Documento generado a partir de las salidas reales guardadas en docs/evidencia/. "
    "Los nombres del alumno y de los integrantes se dejan como marcadores para completarlos. "
    "La captura del explorador de VS Code se inserta en el marcador del final."
)
set_run_font(r, "Calibri", 8.5, False, GRAY, italic=True)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_page_break()

# ------------------------------- Contenido ---------------------------------
h1(doc, "Contenido", page_break=False)
contents = [
    "1. Portada y datos de la actividad",
    "2. Pasos 1–9: entender el proyecto y construir el carrito del equipo",
    "3. Paso 10: el duelo contra el carrito de la IA",
    "4. Paso 11: los cinco fallos de shop/cart_ai.py",
    "5. Paso 12: las cuatro preguntas de revisión",
    "6. Paso 13: estado original y arreglos de la IA",
    "7. Pasos 14–15: auditoría de los seis carritos",
    "8. Paso 16: la compuerta antes y después",
    "9. Paso 17: sesión del agente y rechazo documentado",
    "10. Paso 18: bitácora de uso de IA",
    "11. Conclusiones",
    "12. Estructura del proyecto y captura pendiente",
]
for item in contents:
    p = doc.add_paragraph()
    r = p.add_run(item)
    set_run_font(r, "Calibri", 10.5, False, DARK)
    p.paragraph_format.space_after = Pt(3)

p = doc.add_paragraph()
r = p.add_run(
    "Resumen final: carrito del equipo 67/67 y compuerta ACEPTAR (código 0); carrito de la IA 67/67 "
    "y compuerta ACEPTAR (código 0); 6/6 carritos de carts/ con el resultado esperado en ambos modos."
)
set_run_font(r, "Calibri", 9.5, True, GREEN)
p.paragraph_format.space_before = Pt(10)

# ------------------------------- 1. Pasos 1-9 ------------------------------
h1(doc, "2. Pasos 1–9: entender el proyecto y construir el carrito del equipo")
para(doc,
     "Esta primera parte responde las preguntas conceptuales de la guía, mide el estado inicial y "
     "documenta el carrito del equipo (shop/cart_team.py) método por método, con la salida real de "
     "cada archivo de pruebas.")

h2(doc, "2.1 Las reglas de la tienda, con mis palabras")
table(doc,
      ["Regla", "Qué significa en este proyecto"],
      [
          ("Dinero",
           "Todo importe es Decimal, nunca float, porque el float guarda aproximaciones binarias y "
           "pierde céntimos. El redondeo es al céntimo con ROUND_HALF_UP: 0,845 pasa a 0,85 y "
           "452,205 pasa a 452,21. El helper money() de shop/domain.py concentra esa regla."),
          ("Precio",
           "El precio unitario lo decide el catálogo. La petición HTTP solo puede pedir sku y qty; "
           "si manda price, se ignora. lines() construye cada Line leyendo self._catalog[sku].price."),
          ("Cantidad",
           "Entero de 1 en adelante. Se rechazan 0, negativos, decimales, cadenas, None y booleanos "
           "(True es un int en Python, por eso se comprueba aparte). Quitar más de lo que hay deja la "
           "línea en cero y la borra: jamás una cantidad negativa."),
          ("Existencias",
           "Lo que ya está en el carrito más lo que se agrega no puede superar el stock. Si se intenta, "
           "OutOfStock y el carrito no cambia. Al pagar se vuelven a comprobar todas las líneas antes de "
           "restar ninguna; el stock nunca baja de cero."),
          ("Cupón",
           "Uno por carrito. BIENVENIDA10 descuenta 10 % del subtotal; MENOS20 descuenta S/ 20 sin pasar "
           "del subtotal. Aplicar otro, o el mismo otra vez, reemplaza al anterior: no se acumula. El "
           "código se normaliza con strip().upper()."),
          ("Totales",
           "subtotal = suma de precio × cantidad; descuento según el cupón vigente; igv = 18 % de "
           "(subtotal − descuento); total = (subtotal − descuento) + igv. Cada parte se redondea al "
           "céntimo con money()."),
      ],
      widths=[2.7, 14.3], size=8.8, first_col_bold=True)

h2(doc, "2.2 Línea base: pruebas y compuerta antes de escribir código")
meta(doc, "medición del estado inicial del repositorio")
para(doc, "Salida real de la suite completa antes de implementar el carrito:")
code(doc, """$ python manage.py test
Creating test database for alias 'default'...
EEEEEEEEEEEEEEEEE......FF.EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
...
FAILED (failures=2, errors=67)
Found 67 test(s).
[exit code: 1]""")
para(doc,
     "De las 67 pruebas, 7 pasaban y 60 fallaban. Las 7 que pasaban son las de shop/tests/test_audit.py "
     "que no usan Cart: audit_cart() trabaja sobre el JSON declarado y recalcula los totales con "
     "reference_totals() de shop/domain.py, una función independiente de toda implementación. Por eso "
     "puede detectar PRICE_MISMATCH, OVER_STOCK, BAD_QTY, MULTIPLE_COUPONS y TOTAL_MISMATCH aunque "
     "shop/cart_team.py sea un esqueleto. Las dos pruebas de replay sí fallan, porque --replay sí "
     "construye impl.Cart y recibe NotImplementedError, que se traduce en REPLAY_UNAVAILABLE.")
code(doc, """# shop/domain.py
def reference_totals(lines, coupons):
    \"\"\"Subtotal, discount, igv and total from (unit_price, qty) pairs and a list of coupon codes.

    Independent of any Cart: this is the formula in the README, used by the audit.
    \"\"\"
    subtotal = sum((Decimal(price) * qty for price, qty in lines), Decimal("0"))
    ...""")
para(doc, "Compuerta inicial de ambos modos:")
table(doc,
      ["Comando", "Pruebas", "Carritos", "Veredicto", "Evidencia"],
      [
          ("python manage.py gate", "7 pasan / 60 fallan", "0 de 6", "RECHAZAR (código 1)", "03_gate_inicial_equipo.txt"),
          ("python manage.py gate --modo ai", "49 pasan / 18 fallan", "4 de 6", "RECHAZAR (código 1)", "04_gate_inicial_ai.txt"),
      ],
      widths=[4.6, 3.4, 2.2, 3.7, 3.1], size=8.4)
evidence(doc, "02_test_inicial_equipo.txt", "03_gate_inicial_equipo.txt", "04_gate_inicial_ai.txt")

h2(doc, "2.3 Respuestas conceptuales")
h3(doc, "¿Por qué el carrito usa un catálogo de datos puros y no el modelo de Django?")
para(doc,
     "shop/domain.py no importa Django: define Product como un dataclass con sku, name, price y stock. "
     "Así la lógica del carrito es pura y se puede probar con catalog_data() (los ocho productos fijos) "
     "sin base de datos, sin migraciones y sin servidor. Además, checkout puede modificar el stock de "
     "una copia del catálogo y la prueba observa esa copia. La frontera con la base de datos vive en "
     "shop/catalog.py (catalog_from_db() y save_stock()), que traduce filas del modelo Product al "
     "ProductData que el carrito entiende.")
h3(doc, "¿Qué guarda _qty y por qué no guarda el precio?")
para(doc,
     "_qty es un diccionario sku → cantidad entera, en orden de inserción. No guarda el precio porque "
     "el precio pertenece al catálogo: si se copiara en el carrito podría quedar desactualizado o ser "
     "manipulado. lines() construye cada Line leyendo self._catalog[sku].price en ese momento; el "
     "carrito solo recuerda qué y cuánto, no a qué precio.")
h3(doc, "¿Qué hace shop/impl.py con SHOP_MODE?")
para(doc,
     "Lee la variable de entorno SHOP_MODE al importarse (por defecto team), valida que sea team, ai o "
     "solutions e importa el módulo correspondiente (shop.cart_team, shop.cart_ai o teacher.solutions.cart). "
     "Expone cart para que las pruebas, la API, la auditoría y la compuerta usen la misma implementación. "
     "Si el modo no existe o el módulo no está en esa copia, lanza RuntimeError con un mensaje claro.")

h2(doc, "2.4 Desarrollo del carrito del equipo")
para(doc,
     "Cada bloque muestra el resultado real, el código del turno (y el final cuando corresponde), la "
     "explicación y los casos de prueba que lo fijan. El orden sigue las dependencias reales: líneas, "
     "existencias, dinero, cupón y API.")

h3(doc, "Desarrollo 1 — add, remove, lines y validación de cantidades")
meta(doc, "carrito del equipo · líneas y cantidades (commit 0fe7b33)")
para(doc, "Resultado real del turno:", bold=True, space_after=2)
code(doc, """$ python manage.py test shop.tests.test_lines
......E.E....E.
...
Ran 15 tests in 0.015s
FAILED (errors=3)
[exit code: 1]""")
para(doc,
     "Las 12 pruebas propias de líneas ya pasaban; los 3 errores restantes llamaban a total() y "
     "apply_coupon(), todavía no implementados. No eran regresiones de este cambio, sino dependencias "
     "de los turnos siguientes.")
para(doc, "Código del turno (la comprobación de existencias se refina en el desarrollo 2):", bold=True, space_after=2)
code(doc, """def _check_qty(qty):
    # bool is an int subclass, so reject it explicitly.
    if isinstance(qty, bool) or not isinstance(qty, int) or qty < 1:
        raise InvalidQuantity(f"Invalid quantity: {qty!r}")

class Cart:
    def add(self, sku, qty=1):
        _check_qty(qty)
        product = self._catalog.get(sku)
        if product is None:
            raise UnknownProduct(f"Unknown product: {sku!r}")
        if qty > product.stock:
            raise OutOfStock(f"{sku}: requested {qty} and only {product.stock} available")
        self._qty[sku] = self._qty.get(sku, 0) + qty

    def remove(self, sku, qty=None):
        if sku not in self._qty:
            raise UnknownProduct(f"Product not in the cart: {sku!r}")
        if qty is None:
            del self._qty[sku]
            return
        _check_qty(qty)
        remaining = self._qty[sku] - qty
        if remaining <= 0:
            del self._qty[sku]
        else:
            self._qty[sku] = remaining

    def lines(self):
        return [
            Line(sku, self._catalog[sku].name, self._catalog[sku].price, qty)
            for sku, qty in self._qty.items()
        ]""")
para(doc,
     "Explicación: remove borra la línea al llegar a cero porque una línea con cantidad cero no "
     "representa nada comprable y aparecería en la API como qty: 0. Si se permitiera una cantidad "
     "negativa, el total restaría dinero: MOU-01 a −3 sumaría −119,70 al subtotal, el IGV bajaría y el "
     "cliente podría “ganar” saldo. La prueba exige cantidad 0, lista vacía y total 0,00.")
para(doc, "Casos de prueba:", bold=True, space_after=2)
for item in [
    "test_add_creates_a_line_with_the_catalog_price: el precio sale del catálogo.",
    "test_adding_the_same_product_accumulates y test_lines_keep_insertion_order.",
    "test_invalid_quantities_are_rejected: 0, −1, 1.5, \"2\", None y True.",
    "test_remove_more_than_there_is_never_leaves_a_negative_quantity.",
    "test_errors_share_a_base_class y test_state_survives_between_requests.",
]:
    bullet(doc, item)
evidence(doc, "05_turno1_test_lines.txt", "06_turno1_gate_json.txt")

h3(doc, "Desarrollo 2 — existencias acumuladas y checkout")
meta(doc, "carrito del equipo · stock y pago (commit c5b47b9)")
para(doc, "Resultado real del turno:", bold=True, space_after=2)
code(doc, """$ python manage.py test shop.tests.test_stock
.EE....E..
...
Ran 10 tests in 0.003s
FAILED (errors=3)
[exit code: 1]""")
para(doc,
     "Las 7 pruebas de stock y pago pasaban; los 3 errores eran llamadas a summary() y apply_coupon(), "
     "aún pendientes.")
para(doc, "Código final:", bold=True, space_after=2)
code(doc, """def add(self, sku, qty=1):
    _check_qty(qty)
    product = self._catalog.get(sku)
    if product is None:
        raise UnknownProduct(f"Unknown product: {sku!r}")
    new_qty = self._qty.get(sku, 0) + qty
    if new_qty > product.stock:
        raise OutOfStock(f"{sku}: requested {new_qty} and only {product.stock} available")
    self._qty[sku] = new_qty

def checkout(self):
    if not self._qty:
        raise CartError("The cart is empty")
    # Check every line before touching any stock: a failure leaves the whole cart untouched.
    for sku, qty in self._qty.items():
        if qty > self._catalog[sku].stock:
            raise OutOfStock(f"{sku}: cart has {qty} and only {self._catalog[sku].stock} available")
    order = self.summary()
    for sku, qty in self._qty.items():
        self._catalog[sku].stock -= qty
    self._qty.clear()
    self._coupon = None
    return order""")
para(doc,
     "Explicación: add garantiza el stock en el instante de agregar, pero entre add y checkout puede "
     "pasar tiempo y otro cliente puede comprar las mismas unidades. Por eso checkout verifica todas "
     "las líneas antes de restar ninguna: si una falla, el carrito y el stock quedan intactos. En la "
     "API, cart_checkout envuelve la operación en transaction.atomic() y llama a save_stock(catalog); "
     "la mitigación completa contra dos compradores simultáneos requiere bloquear las filas con "
     "select_for_update() o un UPDATE ... WHERE stock >= qty atómico. Sin ese bloqueo, dos procesos "
     "pueden leer el mismo stock y ambos pasar la comprobación: es una condición de carrera que SQLite "
     "y una sola petición por prueba no reproducen.")
para(doc, "Casos de prueba:", bold=True, space_after=2)
for item in [
    "test_you_can_add_up_to_the_stock y test_more_than_the_stock_is_rejected_and_the_cart_is_unchanged.",
    "test_the_stock_counts_what_is_already_in_the_cart y test_two_adds_cannot_get_around_the_stock.",
    "test_checkout_lowers_the_stock_and_empties_the_cart.",
    "test_checkout_checks_the_stock_again y test_the_stock_is_never_negative_after_any_sequence.",
    "test_checkout_clears_the_coupon.",
]:
    bullet(doc, item)
evidence(doc, "07_turno2_test_stock.txt", "08_turno2_gate.txt")

h3(doc, "Desarrollo 3 — subtotal, discount, igv, total y summary")
meta(doc, "carrito del equipo · dinero y totales (commit d0d86d7)")
para(doc, "Resultado real del turno:", bold=True, space_after=2)
code(doc, """$ python manage.py test shop.tests.test_totals
E....E..EE
...
Ran 10 tests in 0.003s
FAILED (errors=4)
[exit code: 1]""")
para(doc,
     "Las 6 pruebas que no usan cupón pasaban; las 4 restantes llamaban a apply_coupon(), que llegó en "
     "el turno siguiente.")
para(doc, "Código final:", bold=True, space_after=2)
code(doc, """def subtotal(self):
    return money(sum((line.total for line in self.lines()), Decimal("0")))

def discount(self):
    if self._coupon is None:
        return money(Decimal("0"))
    kind, amount = COUPONS[self._coupon]
    if kind == "percent":
        return money(self.subtotal() * amount / 100)
    return money(min(amount, self.subtotal()))

def igv(self):
    return money((self.subtotal() - self.discount()) * IGV_RATE)

def total(self):
    base = money(self.subtotal() - self.discount())
    return base + self.igv()

def summary(self):
    return {
        "lines": [
            {"sku": line.sku, "name": line.name, "unit_price": fmt(line.unit_price),
             "qty": line.qty, "line_total": fmt(line.total)}
            for line in self.lines()
        ],
        "coupon": self._coupon,
        "subtotal": fmt(self.subtotal()),
        "discount": fmt(self.discount()),
        "igv": fmt(self.igv()),
        "total": fmt(self.total()),
    }""")
para(doc,
     "Explicación: el descuento se calcula sobre el subtotal ya redondeado; el IGV se calcula sobre "
     "la base (subtotal − descuento) y el total suma la base redondeada más el IGV. Todo pasa por money(), "
     "que usa ROUND_HALF_UP. summary() entrega cadenas de dos decimales con fmt(), que es el contrato de "
     "la API.")
para(doc, "Casos de prueba:", bold=True, space_after=2)
for item in [
    "test_subtotal_is_the_sum_of_price_times_quantity y test_igv_and_total_without_coupon.",
    "test_money_is_decimal_not_float y test_ten_small_items_add_up_exactly.",
    "test_igv_rounds_half_up, test_the_discount_rounds_half_up y test_a_discount_of_one_and_a_half_cents.",
    "test_total_is_the_sum_of_its_rounded_parts.",
    "test_summary_has_the_shape_the_api_sends y test_summary_of_an_empty_cart.",
]:
    bullet(doc, item)
evidence(doc, "10_turno4_test_totals.txt", "11_turno4_gate.txt")

h3(doc, "Desarrollo 4 — apply_coupon idempotente")
meta(doc, "carrito del equipo · cupones (commit 8b78e26)")
para(doc, "Resultado real del turno:", bold=True, space_after=2)
code(doc, """$ python manage.py test shop.tests.test_coupons
.........
----------------------------------------------------------------------
Ran 9 tests in 0.001s
OK
[exit code: 0]""")
para(doc, "Código final:", bold=True, space_after=2)
code(doc, """def apply_coupon(self, code):
    if not isinstance(code, str):
        raise InvalidCoupon(f"Invalid coupon: {code!r}")
    normalized = code.strip().upper()
    if normalized not in COUPONS:
        raise InvalidCoupon(f"Unknown coupon: {normalized!r}")
    self._coupon = normalized""")
para(doc,
     "Explicación: apply_coupon guarda el código en un único atributo self._coupon. El segundo clic "
     "reemplaza el valor anterior por el mismo y discount() calcula el descuento una sola vez. El "
     "resultado es idempotente: dos clics dejan el mismo total que uno. Si el carrito guardara una lista "
     "y sumara el descuento de cada elemento, un doble clic o un reintento de red descontaría dos veces y "
     "la tienda perdería dinero. Un código desconocido lanza InvalidCoupon sin cambiar el carrito.")
para(doc, "Casos de prueba:", bold=True, space_after=2)
for item in [
    "test_percent_coupon y test_fixed_coupon.",
    "test_fixed_coupon_never_exceeds_the_subtotal.",
    "test_unknown_coupon_is_rejected_and_changes_nothing.",
    "test_code_is_trimmed_and_case_insensitive.",
    "test_applying_the_same_coupon_twice_discounts_once y test_a_second_coupon_replaces_the_first.",
    "test_the_discount_follows_the_cart y test_no_coupon_no_discount.",
]:
    bullet(doc, item)
evidence(doc, "12_turno5_test_coupons.txt", "13_turno5_regresion.txt", "14_turno5_gate.txt")

h3(doc, "Desarrollo 5 — add_from_payload y cierre del carrito")
meta(doc, "carrito del equipo · API y cierre (commit 3ffd078)")
para(doc, "Resultado real del turno:", bold=True, space_after=2)
code(doc, """$ python manage.py test shop.tests.test_api
..............
Ran 14 tests in 0.063s
OK
[exit code: 0]

$ python manage.py test
...................................................................
Ran 67 tests in 0.058s
OK
[exit code: 0]

$ python manage.py gate
  pruebas     pasan  67   fallan   0
  carritos    6 de 6 como se esperaba
  veredicto   ACEPTAR (código de salida 0)
[exit code: 0]""")
para(doc, "Código final:", bold=True, space_after=2)
code(doc, """def add_from_payload(cart, payload):
    if not isinstance(payload, dict) or not isinstance(payload.get("sku"), str):
        raise CartError("The body must be a JSON object with a string 'sku'")
    cart.add(payload["sku"], payload.get("qty", 1))""")
para(doc,
     "Explicación: el cuerpo de POST /api/cart/items/ es dato no confiable. add_from_payload solo acepta "
     "un diccionario con sku de tipo str y pasa sku/qty a cart.add; cualquier price del cuerpo se ignora. "
     "Si se aceptara el precio del cliente, bastaría enviar {\"sku\": \"LAP-01\", \"price\": 0.01} para "
     "comprar una laptop de S/ 2 499,90 por un céntimo.")
para(doc, "Casos de prueba:", bold=True, space_after=2)
for item in [
    "test_adding_an_item_returns_the_cart y test_the_cart_persists_between_requests.",
    "test_the_price_comes_from_the_catalog_not_from_the_client y la variante con precio negativo.",
    "test_unknown_product_is_404, test_more_than_the_stock_is_409 y test_bad_input_is_400.",
    "test_remove_more_than_there_is_leaves_no_negative_line.",
    "test_the_coupon_is_applied_once_even_if_the_client_repeats_it.",
    "test_checkout_creates_the_order_and_lowers_the_stock y test_checkout_when_the_stock_changed_is_409_and_keeps_the_cart.",
]:
    bullet(doc, item)
evidence(doc, "15_turno6_test_api.txt", "16_turno6_test_completo.txt",
         "17_turno6_gate_equipo.txt", "18_turno6_gate_json.txt")

h3(doc, "Revisión del carrito del equipo con las cuatro preguntas")
meta(doc, "cierre de la fase B · las cuatro preguntas de AGENTS.md")
for text in [
    "1. ¿Todo el dinero es Decimal de principio a fin? Sí. subtotal (línea 97), discount (99-107), igv "
    "(109-111) y total (113-116) usan money(); no hay ningún float en el archivo y "
    "test_money_is_decimal_not_float lo comprueba.",
    "2. ¿El servidor es la única fuente del precio? Sí. lines() construye cada Line con "
    "self._catalog[sku].price (línea 80) y add_from_payload (157-163) solo lee sku y qty; "
    "test_the_price_comes_from_the_catalog_not_from_the_client lo demuestra.",
    "3. ¿Cada regla de cantidad y existencias se cumple con cualquier secuencia? Sí. add suma lo que ya "
    "hay antes de comparar (54-57), remove borra la línea si el resto es <= 0 (68-72) y checkout "
    "re-verifica todas las líneas antes de restar (146-151); lo cubren test_two_adds_cannot_get_around_the_stock, "
    "test_remove_more_than_there_is_never_leaves_a_negative_quantity y test_checkout_checks_the_stock_again.",
    "4. ¿Repetir una acción da el mismo resultado? Sí. apply_coupon guarda un único código en "
    "self._coupon (línea 93) y discount() lo aplica una sola vez; "
    "test_applying_the_same_coupon_twice_discounts_once y test_a_second_coupon_replaces_the_first lo fijan.",
]:
    bullet(doc, text)

h2(doc, "2.5 Resultados por archivo de pruebas")
table(doc,
      ["Archivo", "Pruebas", "Resultado del turno", "Estado final", "Evidencia"],
      [
          ("test_lines.py", "15", "12 pasan; 3 errores por métodos aún no implementados", "15/15", "05_turno1_test_lines.txt"),
          ("test_stock.py", "10", "7 pasan; 3 errores por summary/apply_coupon", "10/10", "07_turno2_test_stock.txt"),
          ("test_totals.py", "10", "6 pasan; 4 errores por apply_coupon", "10/10", "10_turno4_test_totals.txt"),
          ("test_coupons.py", "9", "9/9", "9/9", "12_turno5_test_coupons.txt"),
          ("test_api.py", "14", "14/14", "14/14", "15_turno6_test_api.txt"),
          ("test_audit.py", "9", "7/9 al inicio (replay sin carrito)", "9/9", "02_test_inicial_equipo.txt"),
          ("Suite completa", "67", "OK al cerrar el carrito", "67/67", "16_turno6_test_completo.txt"),
      ],
      widths=[3.2, 1.7, 7.2, 2.2, 4.3], size=8.2)

h2(doc, "2.6 Float frente a Decimal")
meta(doc, "herramienta tools/decimal_check.py · paso 5")
para(doc, "Salida real de python tools/decimal_check.py:")
code(doc, """0.1 + 0.2 = 0.30000000000000004
Decimal('0.1') + Decimal('0.2') = 0.3
round(0.845, 2) = 0.84
Decimal('0.845').quantize(Decimal('0.01'), ROUND_HALF_UP) = 0.85""")
para(doc,
     "Un float guarda el número en binario con 53 bits de mantisa: 0,1 y 0,2 no tienen representación "
     "binaria finita, así que se guardan como la fracción más cercana y su suma arrastra el error "
     "0.30000000000000004. Tampoco 0,845 es exacto en binario: el valor almacenado queda ligeramente por "
     "debajo de 0,845 y round() de Python, además, redondea al par (banker's rounding), por eso devuelve "
     "0,84. Decimal guarda dígitos decimales exactos y quantize(..., ROUND_HALF_UP) aplica la regla de la "
     "tienda: 0,845 → 0,85.")
para(doc, "Cálculo a mano del mousepad PAD-01 (precio 8,45) con BIENVENIDA10:", bold=True, space_after=2)
table(doc,
      ["Paso", "Operación", "Resultado"],
      [
          ("1. Subtotal", "8,45 × 1", "8,45"),
          ("2. Descuento", "10 % de 8,45 = 0,845 → half-up", "0,85"),
          ("3. Base", "8,45 − 0,85", "7,60"),
          ("4. IGV", "18 % de 7,60 = 1,368 → half-up", "1,37"),
          ("5. Total", "7,60 + 1,37", "8,97"),
      ],
      widths=[3.2, 9.0, 4.8], size=9)
para(doc,
     "Coincide con test_the_discount_rounds_half_up, que exige discount() == 0.85 y total() == 8.97.")
evidence(doc, "09_turno3_decimales.txt")

h2(doc, "2.7 Tres llamadas reales a la API con la misma sesión")
meta(doc, "runserver en 127.0.0.1:8013 · tools/api_smoke.ps1")
para(doc,
     "Se levantó el servidor, se conservó una sesión de cookies y se hicieron tres llamadas: agregar dos "
     "mousepads, aplicar BIENVENIDA10 y pagar. Respuestas reales:")
code(doc, """POST /api/cart/items/  body={"sku":"MOU-01","qty":2}
{
  "lines": [{"sku": "MOU-01", "name": "Mouse inalámbrico", "unit_price": "39.90",
             "qty": 2, "line_total": "79.80"}],
  "coupon": null,
  "subtotal": "79.80", "discount": "0.00", "igv": "14.36", "total": "94.16"
}

POST /api/cart/coupon/  body={"code":"BIENVENIDA10"}
{
  "lines": [{"sku": "MOU-01", "name": "Mouse inalámbrico", "unit_price": "39.90",
             "qty": 2, "line_total": "79.80"}],
  "coupon": "BIENVENIDA10",
  "subtotal": "79.80", "discount": "7.98", "igv": "12.93", "total": "84.75"
}

POST /api/cart/checkout/  body={}
{
  "lines": [{"sku": "MOU-01", "name": "Mouse inalámbrico", "unit_price": "39.90",
             "qty": 2, "line_total": "79.80"}],
  "coupon": "BIENVENIDA10",
  "subtotal": "79.80", "discount": "7.98", "igv": "12.93", "total": "84.75"
}

GET /api/cart/  (después del pago)
{"lines": [], "coupon": null, "subtotal": "0.00", "discount": "0.00",
 "igv": "0.00", "total": "0.00"}""")
para(doc,
     "El servidor se detuvo al terminar y seed_shop restauró el catálogo de 8 productos. La sesión "
     "demuestra que el carrito vive entre peticiones y que el precio mostrado es el del catálogo.")
evidence(doc, "19_api_smoke.txt", "20_seed_restore.txt")

# ------------------------------- 2. Duelo ----------------------------------
h1(doc, "3. Paso 10: el duelo contra el carrito de la IA")
meta(doc, "diagnóstico comparativo · mismo oráculo, dos implementaciones")
para(doc, "Salida real de python manage.py duelo:")
code(doc, """Duelo: las mismas pruebas contra cada carrito
team       tu carrito                 pasan  67   fallan   0
ai         el carrito de la IA        pasan  49   fallan  18""")
table(doc,
      ["Modo", "Implementación", "Pasan", "Fallan"],
      [("team", "tu carrito", "67", "0"), ("ai", "el carrito de la IA", "49", "18")],
      widths=[3.0, 8.0, 3.0, 3.0], size=9.5)
para(doc, "Fallos de la IA agrupados por archivo:", bold=True, space_after=2)
table(doc,
      ["Archivo", "N.º", "Pruebas que fallan"],
      [
          ("test_api.py", "6",
           "test_checkout_when_the_stock_changed_is_409_and_keeps_the_cart (ERROR, CHECK constraint "
           "failed: stock); test_a_negative_price_from_the_client_changes_nothing_either; "
           "test_more_than_the_stock_is_409; test_remove_more_than_there_is_leaves_no_negative_line; "
           "test_the_coupon_is_applied_once_even_if_the_client_repeats_it; "
           "test_the_price_comes_from_the_catalog_not_from_the_client."),
          ("test_audit.py", "1", "test_replay_of_the_rounding_cart_matches_your_cart."),
          ("test_coupons.py", "2",
           "test_a_second_coupon_replaces_the_first; test_applying_the_same_coupon_twice_discounts_once."),
          ("test_lines.py", "1", "test_remove_more_than_there_is_never_leaves_a_negative_quantity."),
          ("test_stock.py", "4",
           "test_checkout_checks_the_stock_again; test_the_stock_counts_what_is_already_in_the_cart; "
           "test_the_stock_is_never_negative_after_any_sequence; test_two_adds_cannot_get_around_the_stock."),
          ("test_totals.py", "4",
           "test_a_discount_of_one_and_a_half_cents; test_igv_rounds_half_up; "
           "test_the_discount_rounds_half_up; test_total_is_the_sum_of_its_rounded_parts."),
      ],
      widths=[3.0, 1.4, 12.6], size=8.2)
para(doc,
     "Los 18 fallos se agrupan en cinco causas, que se clasifican en el paso 11. La IA compila y responde: "
     "el problema no es sintáctico, es de reglas de negocio.")
evidence(doc, "21_duelo.txt", "22_duelo_ai.txt")

# ------------------------------- 3. Cinco fallos ---------------------------
landscape(doc)
h1(doc, "4. Paso 11: los cinco fallos de shop/cart_ai.py", page_break=False)
meta(doc, "clasificación antes de tocar el archivo · líneas del estado original ai-original")
para(doc,
     "Las líneas corresponden al archivo original conservado en el commit 6978401 y la etiqueta "
     "ai-original. No se modificó shop/cart_ai.py durante esta clasificación.")
table(doc,
      ["#", "Archivo:línea", "Fallo", "Prueba que lo detecta", "Mensaje / diferencia real", "Causa", "¿Se vería a mano?", "Marca"],
      [
          ("1", "cart_ai.py:14, 17-18, 83-84, 92-94, 104", "Redondeo con flotantes",
           "test_totals (4); test_audit.test_replay_of_the_rounding_cart_matches_your_cart",
           "Decimal('0.84') != Decimal('0.85'); Decimal('452.2') != Decimal('452.21'); "
           "Decimal('0.1') != Decimal('0.11'); Decimal('71.63') != Decimal('71.64'); REPLAY_MISMATCH",
           "_money usa round(value, 2) sobre float; _subtotal suma float; IGV = 0.18 es float.",
           "Sí, con medios céntimos: un mousepad de S/ 8,45 con 10 % debe descontar 0,85 y la IA descuenta 0,84.",
           "Cálculo"),
          ("2", "cart_ai.py:45, 53, 139", "Precio tomado del cliente",
           "test_api.test_the_price_comes_from_the_catalog_not_from_the_client; "
           "test_api.test_a_negative_price_from_the_client_changes_nothing_either; "
           "PRICE_MISMATCH en 02-precio-del-cliente.json",
           "'0.01' != '2499.90'; '-100.00' != '2499.90'; la petición con price devuelve 201",
           "add(..., price=None) guarda Decimal(str(price)) en _prices y add_from_payload pasa "
           "payload.get(\"price\").",
           "Solo si alguien manipula la petición; en la pantalla normal no se nota, pero permite comprar "
           "una laptop de S/ 2 499,90 por S/ 0,01.",
           "SEGURIDAD y solo API"),
          ("3", "cart_ai.py:50 y 124-133", "Existencias que no cuentan lo ya agregado (y checkout no re-verifica)",
           "test_stock (4); test_api.test_more_than_the_stock_is_409; "
           "test_api.test_checkout_when_the_stock_changed_is_409_and_keeps_the_cart; OVER_STOCK en 03-stock.json",
           "OutOfStock not raised; 201 != 409; -2 not greater than or equal to 0; "
           "CHECK constraint failed: stock",
           "add compara solo qty > product.stock, no self._qty.get(sku, 0) + qty; checkout resta sin "
           "volver a comprobar.",
           "Sí: con 2 mousepads en stock, agregar 2 y luego 1 deja 3 en el carrito; al pagar se vende "
           "stock inexistente y el inventario queda negativo.",
           "Existencias"),
          ("4", "cart_ai.py:55-66 (línea 63)", "Cantidad negativa al quitar",
           "test_lines.test_remove_more_than_there_is_never_leaves_a_negative_quantity; "
           "test_api.test_remove_more_than_there_is_leaves_no_negative_line",
           "-3 != 0; la línea queda con qty: -2 y el total baja",
           "self._qty[sku] -= qty sin tope; solo borra si queda exactamente 0.",
           "Sí: quitar 5 unidades cuando hay 2 deja la línea en −3 y el total se vuelve negativo; el "
           "cliente “gana” dinero.",
           "Cantidad"),
          ("5", "cart_ai.py:77-81, 89-95", "Descuento acumulado dos veces",
           "test_coupons (2); test_api.test_the_coupon_is_applied_once_even_if_the_client_repeats_it; "
           "MULTIPLE_COUPONS en 05-cupon-doble.json",
           "Decimal('37.8') != Decimal('18.90'); '37.80' != '18.90'; Decimal('38.9') != Decimal('20.00'); "
           "REPLAY_MISMATCH",
           "apply_coupon hace append a una lista y _discount suma el descuento de todos los cupones.",
           "Sí: un doble clic en «Aplicar cupón» descuenta dos veces el 10 %; el cliente paga menos de lo "
           "que corresponde.",
           "Idempotencia"),
      ],
      widths=[0.8, 3.0, 3.4, 5.4, 5.4, 4.3, 4.3, 2.2], size=7.4, header_size=7.6)
para(doc,
     "El fallo de seguridad es el 2: acepta el precio que manda el cliente. Es también el único que solo "
     "se ve con la API, porque add_from_payload es el camino que traduce el cuerpo HTTP; las pruebas puras "
     "del carrito no envían price. Los fallos 1, 3, 4 y 5 se reproducen directamente sobre Cart y por eso "
     "aparecen en test_totals, test_stock, test_lines y test_coupons. El fallo 3 tiene además "
     "manifestaciones HTTP (el 409 que no llega y el CHECK constraint failed: stock), pero su raíz es la "
     "misma regla de existencias.")
evidence(doc, "22_duelo_ai.txt", "29_test_ai_antes.txt")

# ------------------------------- 4. Cuatro preguntas ---------------------
portrait(doc)
h1(doc, "5. Paso 12: las cuatro preguntas de revisión", page_break=False)
meta(doc, "aplicadas a shop/cart_ai.py original, con la línea que lo demuestra")
para(doc,
     "Las cuatro preguntas de AGENTS.md se responden sobre el diff o sobre el archivo revisado. Una "
     "respuesta «no sé» cuenta como «no». Estas son las respuestas al estado original de la IA.")

h3(doc, "1. ¿Todo el dinero es Decimal de principio a fin, con redondeo hacia arriba en la mitad?")
rich(doc, [("No. ", True),
           ("La línea 14 define IGV = 0.18 como float; la 18 hace Decimal(str(round(value, 2))), que "
            "primero redondea en binario; la 84 suma float(l.unit_price) * l.qty; las líneas 92 y 94 "
            "acumulan float; la 104 mezcla ese float con el IGV. El resultado es 0,84 donde la regla pide "
            "0,85.", False)])
code(doc, """IGV = 0.18

def _money(value):
    return Decimal(str(round(value, 2)))

def _subtotal(self):
    return sum(float(l.unit_price) * l.qty for l in self.lines())""")

h3(doc, "2. ¿El servidor es la única fuente del precio?")
rich(doc, [("No. ", True),
           ("La línea 45 acepta price=None; la 53 guarda Decimal(str(price)) si el cliente lo envía; la "
            "139 le pasa payload.get(\"price\"). El catálogo deja de ser la única fuente y la petición "
            "puede fijar el precio.", False)])
code(doc, """def add(self, sku, qty=1, price=None):
    ...
    self._prices[sku] = Decimal(str(price)) if price is not None else product.price

def add_from_payload(cart, payload):
    ...
    cart.add(payload["sku"], payload.get("qty", 1), payload.get("price"))""")

h3(doc, "3. ¿Cada regla de cantidad y existencias se cumple con cualquier secuencia de acciones?")
rich(doc, [("No. ", True),
           ("La línea 50 compara qty > product.stock sin sumar self._qty.get(sku, 0), así que dos add "
            "seguidos se saltan el stock. La línea 63 resta sin tope y deja cantidades negativas. Las "
            "líneas 124-133 no vuelven a comprobar el stock antes de restar en checkout.", False)])
code(doc, """if qty > product.stock:                 # ignora lo que ya está en el carrito
    raise OutOfStock(...)

self._qty[sku] -= qty                   # puede quedar negativo
if self._qty[sku] == 0:
    del self._qty[sku]

def checkout(self):
    ...
    order = self.summary()
    for sku, qty in self._qty.items():  # resta sin re-verificar
        self._catalog[sku].stock -= qty""")

h3(doc, "4. ¿Repetir una acción da el mismo resultado que hacerla una vez?")
rich(doc, [("No. ", True),
           ("La línea 81 hace self._coupons.append(code) y las líneas 89-95 recorren y suman todos los "
            "cupones; aplicar dos veces BIENVENIDA10 descuenta 37,80 en vez de 18,90. Un doble clic o un "
            "reintento de red cambia el total.", False)])
code(doc, """def apply_coupon(self, code):
    ...
    self._coupons.append(code)

def _discount(self):
    total_discount = 0.0
    for code in self._coupons:          # suma todos los cupones
        total_discount += subtotal * float(amount) / 100""")
para(doc,
     "Después de los arreglos, las cuatro respuestas cambian a «sí»: IGV y todos los cálculos son "
     "Decimal; add ya no acepta price; add acumula y checkout re-verifica; apply_coupon reemplaza el "
     "cupón y _discount usa solo uno. La sección 6 muestra el antes y el después de cada arreglo.")
evidence(doc, "29_test_ai_antes.txt")

# ------------------------------- 5. Arreglos -------------------------------
h1(doc, "6. Paso 13: estado original y arreglos de la IA")
meta(doc, "un fallo por turno · cinco commits · evidencia antes/después")
para(doc,
     "Antes de cambiar nada se guardó el estado original con un commit vacío y una etiqueta:")
code(doc, """$ git commit --allow-empty -m "chore: snapshot original AI cart before fixes"
[main 6978401] chore: snapshot original AI cart before fixes
$ git tag ai-original
$ git rev-parse HEAD
6978401a85d6d9c6a0481b20d0cf4de29b881ec4""")
para(doc, "Estado de la IA antes de tocar el archivo:")
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.046s
FAILED (failures=17, errors=1)
[exit code: 1]""")
para(doc,
     "Son 18 pruebas únicas en rojo (17 fallos + 1 error). La tabla siguiente resume los cinco arreglos, "
     "cada uno con su commit y el conteo real de fallas antes y después.")
table(doc,
      ["Arreglo", "Commit", "Antes", "Después", "Qué cambió"],
      [
          ("1. Redondeo con flotantes", "bbb4251", "18", "13",
           "IGV a Decimal(\"0.18\"); _money usa money; _subtotal y _discount en Decimal, con el descuento "
           "redondeado antes de calcular la base."),
          ("2. Precio del cliente", "2138c41", "13", "11",
           "add sin price; se elimina _prices; lines lee el catálogo; add_from_payload no pasa price."),
          ("3. Existencias", "e69e6db", "11", "5",
           "add compara lo_que_hay + qty; checkout re-verifica todas las líneas antes de restar."),
          ("4. Cantidad negativa", "32489e4", "5", "3",
           "remove calcula el resto y borra la línea si es <= 0."),
          ("5. Descuento acumulado", "e2f1be2", "3", "0",
           "apply_coupon reemplaza (_coupons = [code]); _discount usa solo el primer cupón; from_state "
           "conserva el último."),
      ],
      widths=[3.3, 1.9, 1.3, 1.5, 9.0], size=8.4, first_col_bold=True)

h3(doc, "Arreglo 1 — redondeo con flotantes")
meta(doc, "cart_ai.py · dinero Decimal · bbb4251")
para(doc, "Resultado real: 18 → 13 fallas.", bold=True, space_after=2)
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.045s
FAILED (failures=12, errors=1)     # 13 pruebas únicas
[exit code: 1]""")
para(doc, "Antes y después:", bold=True, space_after=2)
code(doc, """# ANTES
IGV = 0.18
def _money(value):
    return Decimal(str(round(value, 2)))
def _subtotal(self):
    return sum(float(l.unit_price) * l.qty for l in self.lines())

# DESPUÉS
IGV = Decimal("0.18")
def _money(value):
    return money(value)
def _subtotal(self):
    return money(sum((l.unit_price * l.qty for l in self.lines()), Decimal("0")))
def _discount(self):
    subtotal = self._subtotal()
    total_discount = Decimal("0")
    for code in self._coupons:
        kind, amount = COUPONS[code]
        if kind == "percent":
            total_discount += subtotal * amount / 100
        else:
            total_discount += min(amount, subtotal)
    return money(total_discount)""")
para(doc,
     "Por qué: el dinero deja de pasar por float y el descuento se redondea antes de calcular la base "
     "del total; sin ese redondeo, la base usaría 0,845 y el total no coincidiría con 8,97. Desaparecen "
     "los cuatro fallos de test_totals y el REPLAY_MISMATCH del carrito 06.")
evidence(doc, "30_test_ai_fix1_redondeo.txt", "31_gate_ai_fix1.txt")

h3(doc, "Arreglo 2 — el precio sale solo del catálogo")
meta(doc, "cart_ai.py · precio del servidor · 2138c41")
para(doc, "Resultado real: 13 → 11 fallas.", bold=True, space_after=2)
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.055s
FAILED (failures=10, errors=1)     # 11 pruebas únicas
[exit code: 1]""")
para(doc, "Antes y después:", bold=True, space_after=2)
code(doc, """# ANTES
def add(self, sku, qty=1, price=None):
    ...
    self._prices[sku] = Decimal(str(price)) if price is not None else product.price

def lines(self):
    return [Line(sku, self._catalog[sku].name, self._prices[sku], qty) ...]

def add_from_payload(cart, payload):
    ...
    cart.add(payload["sku"], payload.get("qty", 1), payload.get("price"))

# DESPUÉS
def add(self, sku, qty=1):
    ...
    # no price parameter and no _prices dictionary

def lines(self):
    return [Line(sku, self._catalog[sku].name, self._catalog[sku].price, qty) ...]

def add_from_payload(cart, payload):
    ...
    cart.add(payload["sku"], payload.get("qty", 1))""")
para(doc,
     "Por qué: el servidor vuelve a ser la única fuente del precio. Se eliminó _prices de __init__, "
     "to_state, from_state, remove, lines y checkout. Pasan las dos pruebas del precio del cliente y "
     "desaparece PRICE_MISMATCH como riesgo de cobro.")
evidence(doc, "32_test_ai_fix2_precio.txt", "33_gate_ai_fix2.txt")

h3(doc, "Arreglo 3 — existencias acumuladas y re-verificación en checkout")
meta(doc, "cart_ai.py · stock · e69e6db")
para(doc, "Resultado real: 11 → 5 fallas.", bold=True, space_after=2)
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.041s
FAILED (failures=5)
[exit code: 1]""")
para(doc, "Antes y después:", bold=True, space_after=2)
code(doc, """# ANTES
if qty > product.stock:
    raise OutOfStock(...)
self._qty[sku] = self._qty.get(sku, 0) + qty

def checkout(self):
    ...
    order = self.summary()
    for sku, qty in self._qty.items():
        self._catalog[sku].stock -= qty

# DESPUÉS
new_qty = self._qty.get(sku, 0) + qty
if new_qty > product.stock:
    raise OutOfStock(f"{sku}: pides {new_qty} y quedan {product.stock}")
self._qty[sku] = new_qty

def checkout(self):
    ...
    for sku, qty in self._qty.items():
        if qty > self._catalog[sku].stock:
            raise OutOfStock(f"{sku}: el carrito tiene {qty} y quedan {self._catalog[sku].stock}")
    order = self.summary()
    for sku, qty in self._qty.items():
        self._catalog[sku].stock -= qty""")
para(doc,
     "Por qué: dos add seguidos ya no pueden superar el stock y checkout vuelve a comprobar todas las "
     "líneas antes de tocar ninguna. Pasan las cuatro pruebas de stock y el error de CHECK constraint "
     "failed: stock desaparece.")
evidence(doc, "34_test_ai_fix3_stock.txt", "35_gate_ai_fix3.txt")

h3(doc, "Arreglo 4 — quitar nunca deja cantidad negativa")
meta(doc, "cart_ai.py · cantidades · 32489e4")
para(doc, "Resultado real: 5 → 3 fallas.", bold=True, space_after=2)
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.038s
FAILED (failures=3)
[exit code: 1]""")
para(doc, "Antes y después:", bold=True, space_after=2)
code(doc, """# ANTES
self._qty[sku] -= qty
if self._qty[sku] == 0:
    del self._qty[sku]

# DESPUÉS
remaining = self._qty[sku] - qty
if remaining <= 0:
    del self._qty[sku]
else:
    self._qty[sku] = remaining""")
para(doc,
     "Por qué: si se quita más de lo que hay, la línea se borra; nunca queda un −3 que reste dinero del "
     "total. Pasan la prueba de líneas y la de la API.")
evidence(doc, "36_test_ai_fix4_quitar.txt", "37_gate_ai_fix4.txt")

h3(doc, "Arreglo 5 — un solo cupón, descuento idempotente")
meta(doc, "cart_ai.py · cupón · e2f1be2")
para(doc, "Resultado real: 3 → 0 fallas; compuerta ACEPTAR.", bold=True, space_after=2)
code(doc, """$ python manage.py test          (con SHOP_MODE=ai)
Ran 67 tests in 0.043s
OK
[exit code: 0]

$ python manage.py gate --modo ai
  pruebas     pasan  67   fallan   0
  carritos    6 de 6 como se esperaba
  veredicto   ACEPTAR (código de salida 0)
[exit code: 0]""")
para(doc, "Antes y después:", bold=True, space_after=2)
code(doc, """# ANTES
def apply_coupon(self, code):
    ...
    self._coupons.append(code)

def _discount(self):
    for code in self._coupons:          # suma todos
        ...

@classmethod
def from_state(cls, catalog, state):
    cart._coupons = list(state.get("coupons", []))

# DESPUÉS
def apply_coupon(self, code):
    ...
    self._coupons = [code]              # reemplaza, no acumula

def _discount(self):
    for code in self._coupons[:1]:      # solo uno
        ...

@classmethod
def from_state(cls, catalog, state):
    cart._coupons = list(state.get("coupons", []))[-1:]""")
para(doc,
     "Por qué: aplicar el mismo cupón dos veces (doble clic o reintento) deja el mismo total que una "
     "sola vez. Pasa test_applying_the_same_coupon_twice_discounts_once y la compuerta queda en 0.")
evidence(doc, "38_test_ai_fix5_cupon.txt", "39_gate_ai_despues.txt",
         "40_gate_ai_json_despues.txt")

h2(doc, "6.1 Por qué una implementación correcta debe rechazar los carritos 03 y 04")
para(doc,
     "shop/gate.py define MUST_REJECT = {\"BAD_QTY\", \"UNKNOWN_SKU\", \"OVER_STOCK\"}. El carrito 03 "
     "declara PAD-01 con 4 unidades cuando el catálogo tiene 2 → OVER_STOCK; el 04 declara MOU-01 con −3 "
     "→ BAD_QTY. Para esos dos, expected_outcome() devuelve \"rechazar\" y judge_cart() solo da \"ok\" si "
     "el replay produce REPLAY_REJECTED. Una implementación correcta lanza OutOfStock al agregar 4 sobre 2 "
     "e InvalidQuantity al agregar −3, así que el replay las rechaza. Si las aceptara, el resultado sería "
     "\"acepta_lo_que_debia_rechazar\" y la compuerta no podría aceptar. Rechazar esos carritos no es un "
     "fallo: es la prueba de que las reglas de cantidad y stock están puestas.")

h2(doc, "6.2 Por qué el replay no detecta el precio del cliente y qué sí lo detecta")
para(doc,
     "shop/audit.py::_replay reconstruye con cart.add(item[\"sku\"], item[\"qty\"]), sin pasar "
     "unit_price. El carrito 02 declara LAP-01 a 0,01, pero el replay llama a add(\"LAP-01\", 1) y la IA "
     "usa el precio del catálogo; por eso no hay REPLAY_MISMATCH y el replay “pasa”. La parte que sí lo "
     "detecta es el escaneo de la declaración: audit_cart() compara item[\"unit_price\"] con "
     "product.price y emite PRICE_MISMATCH. Además, la compuerta lo detecta por el otro oráculo: la "
     "prueba test_the_price_comes_from_the_catalog_not_from_the_client falla mientras add acepte price. "
     "En resumen: el replay audita el cálculo, no la procedencia del precio; la prueba de la API y el "
     "PRICE_MISMATCH de la auditoría auditan la procedencia.")

# ------------------------------- 6. Auditoría -----------------------------
h1(doc, "7. Pasos 14–15: auditoría de los seis carritos")
meta(doc, "antes y después de los arreglos · con y sin --replay")
para(doc,
     "La auditoría se ejecutó con SHOP_MODE=ai y el catálogo fijo. La tabla compara el estado antes "
     "(IA rota) y después (IA corregida).")
table(doc,
      ["Carrito", "Antes, sin --replay", "Antes, con --replay", "Después, con --replay", "Fallo"],
      [
          ("01-correcto.json", "Sin hallazgos", "Sin hallazgos", "Sin hallazgos", "—"),
          ("02-precio-del-cliente.json", "PRICE_MISMATCH (0.01 vs 2499.90) + 3 TOTAL_MISMATCH",
           "Igual; no hay REPLAY_MISMATCH", "Igual; la declaración sigue siendo falsa", "2 (precio del cliente)"),
          ("03-stock.json", "OVER_STOCK (4 unidades, hay 2)", "REPLAY_REJECTED (la IA rechaza 4 > 2)",
           "REPLAY_REJECTED (correcto)", "3 (existencias)"),
          ("04-cantidad-negativa.json", "BAD_QTY (−3) + 3 TOTAL_MISMATCH",
           "REPLAY_REJECTED (la IA rechaza −3 en add)", "REPLAY_REJECTED (correcto)", "4 (cantidad negativa)"),
          ("05-cupon-doble.json", "MULTIPLE_COUPONS + 3 TOTAL_MISMATCH",
           "REPLAY_MISMATCH en descuento, IGV y total", "Sin REPLAY_MISMATCH; la declaración sigue mal", "5 (cupón acumulado)"),
          ("06-redondeo.json", "TOTAL_MISMATCH (0.84 vs 0.85)", "REPLAY_MISMATCH (0.84 vs 0.85)",
           "Sin REPLAY_MISMATCH; la declaración sigue mal", "1 (redondeo)"),
      ],
      widths=[3.6, 4.4, 4.0, 4.0, 3.0], size=7.8)

h2(doc, "7.1 TOTAL_MISMATCH frente a REPLAY_MISMATCH")
para(doc,
     "TOTAL_MISMATCH compara lo que el archivo declara con reference_totals(); señala que el carrito está "
     "mal documentado o que quien lo generó calculó mal. REPLAY_MISMATCH compara lo que calcula la "
     "implementación activa al reconstruir el carrito (--replay) con la misma fórmula; señala un defecto "
     "del código, aunque la declaración fuera correcta. En 06 se ven los dos: la declaración dice 0,84 y "
     "el replay de la IA también daba 0,84, mientras la fórmula manda 0,85. En 02 solo hay "
     "TOTAL_MISMATCH y PRICE_MISMATCH: el replay usa cart.add(sku, qty) sin precio, así que la IA cae en "
     "product.price y calcula bien; el defecto de aceptar el precio del cliente no se reproduce en el "
     "replay.")

h2(doc, "7.2 --json, --strict y el límite de la auditoría")
para(doc,
     "audit_cart ... --json devuelve un objeto con mode y carts[] (file, ok, findings[]). Un programa "
     "puede leer findings[].code y decidir; un agente puede tomar el primer hallazgo, ubicar la línea y "
     "proponer un parche.")
code(doc, """{
  "mode": "ai",
  "carts": [
    {
      "file": "02-precio-del-cliente.json",
      "ok": false,
      "findings": [
        {"code": "PRICE_MISMATCH", "detail": "LAP-01: el carrito cobra 0.01 y el catálogo dice 2499.90"}
      ]
    }
  ]
}""")
para(doc,
     "audit_cart ... --strict termina con CommandError: 5 carrito(s) con hallazgos. y código de salida 1, "
     "útil para un pipeline de CI. En el estado después también da 1: una implementación correcta no "
     "convierte en correctos los carritos mal declarados, solo deja de calcularlos mal.")
para(doc,
     "La auditoría no sustituye las reglas del servidor: examina un carrito ya declarado o reconstruido y "
     "avisa después de que el cálculo ocurrió; no impide que una petición con price o con cantidad "
     "negativa llegue a cobrarse. La defensa previa es add, remove, checkout y add_from_payload en el "
     "carrito: si la regla está en el servidor, la petición mala se rechaza antes de tocar el stock o el "
     "dinero. La auditoría es una red de seguridad para revisar código y datos, no la puerta de entrada.")
evidence(doc, "23_audit_ai_antes.txt", "24_audit_ai_replay_antes.txt",
         "25_audit_ai_json_antes.txt", "26_audit_ai_strict_antes.txt",
         "41_audit_ai_replay_despues.txt", "42_audit_ai_strict_despues.txt")

# ------------------------------- 7. Compuerta ------------------------------
h1(doc, "8. Paso 16: la compuerta antes y después")
meta(doc, "dos oráculos sin pantalla · veredicto y código de salida")
para(doc, "Salida real de la compuerta en los dos modos y en los dos estados:")
table(doc,
      ["Estado", "Comando", "Pruebas", "Carritos", "Veredicto", "Evidencia"],
      [
          ("Antes (equipo)", "gate", "7 / 60", "0 de 6", "RECHAZAR (1)", "03_gate_inicial_equipo.txt"),
          ("Antes (IA)", "gate --modo ai", "49 / 18", "4 de 6", "RECHAZAR (1)", "27_gate_ai_antes.txt"),
          ("Después (equipo)", "gate", "67 / 0", "6 de 6", "ACEPTAR (0)", "48_gate_final_equipo.txt"),
          ("Después (IA)", "gate --modo ai", "67 / 0", "6 de 6", "ACEPTAR (0)", "49_gate_final_ai.txt"),
      ],
      widths=[3.4, 3.4, 2.2, 2.2, 3.4, 4.4], size=8.2)
para(doc, "Salida textual del estado final de la IA:", bold=True, space_after=2)
code(doc, """$ python manage.py gate --modo ai

Compuerta: el carrito de la IA (modo ai)
  pruebas     pasan  67   fallan   0
  carritos    6 de 6 como se esperaba
  veredicto   ACEPTAR (código de salida 0)
[exit code: 0]""")
para(doc,
     "El JSON de la compuerta es la interfaz para un agente. Los campos que se leen son verdict "
     "(\"aceptar\" o \"rechazar\"), tests.failing (nombres dotted de las pruebas en rojo) y "
     "carts[].result (ok, calcula_distinto, acepta_lo_que_debia_rechazar, etc.). En el estado antes de la "
     "IA: verdict \"rechazar\", tests.failed 18 y carts 05/06 en calcula_distinto. En el estado después: "
     "verdict \"aceptar\", tests.failed 0 y 6/6 ok.")
code(doc, """{
  "mode": "ai",
  "tests": {"total": 67, "passed": 67, "failed": 0, "failing": []},
  "carts": [{"file": "01-correcto.json", "expected": "totales", "result": "ok"}, ...],
  "verdict": "aceptar",
  "exit_code": 0
}""")
evidence(doc, "27_gate_ai_antes.txt", "28_gate_ai_json_antes.txt",
         "48_gate_final_equipo.txt", "49_gate_final_ai.txt",
         "50_gate_final_equipo_json.txt", "51_gate_final_ai_json.txt")

# ------------------------------- 8. Rechazo -------------------------------
h1(doc, "9. Paso 17: sesión del agente y rechazo documentado")
meta(doc, "trabajo bajo AGENTS.md · un rechazo de ejercicio")
para(doc,
     "Durante todo el laboratorio se trabajó bajo AGENTS.md: solo se editó shop/cart_team.py y luego "
     "shop/cart_ai.py, un método o un fallo por turno; no se tocaron shop/tests/, shop/domain.py, "
     "shop/gate.py, carts/ ni config/. Cada avance se midió con las pruebas y con la compuerta, y cada "
     "turno quedó en docs/uso_de_ia.md con su evidencia. El dinero se mantuvo en Decimal y el precio "
     "siempre salió del catálogo.")
para(doc,
     "Con la IA ya en verde se probó un parche tentador: dejar que _subtotal sume con float y “arreglar” "
     "el redondeo convirtiendo a Decimal con Decimal(str(value)). Las 67 pruebas pasan y la compuerta "
     "acepta, pero el diff viola la pregunta 1. Es un rechazo de ejercicio, no espontáneo: se aplicó a "
     "propósito para demostrar que la compuerta no sustituye la revisión humana.", )
para(doc, "Diff completo guardado en docs/evidencia/rechazo_01.diff:", bold=True, space_after=2)
code(doc, """diff --git a/shop/cart_ai.py b/shop/cart_ai.py
index 8b0f5b9..7da94e4 100644
--- a/shop/cart_ai.py
+++ b/shop/cart_ai.py
@@ -1,5 +1,5 @@
 \"\"\"Shopping cart written by an AI assistant (version under review).\"\"\"
-from decimal import Decimal
+from decimal import ROUND_HALF_UP, Decimal
@@ -16,7 +16,7 @@ IGV = Decimal(\"0.18\")
 def _money(value):
-    return money(value)
+    return Decimal(str(value)).quantize(Decimal(\"0.01\"), rounding=ROUND_HALF_UP)
@@ -79,7 +79,7 @@ class Cart:
     def _subtotal(self):
-        return money(sum((l.unit_price * l.qty for l in self.lines()), Decimal(\"0\")))
+        return _money(sum((float(l.unit_price) * l.qty for l in self.lines()), 0.0))""")
para(doc, "Resultado con el parche aplicado:", bold=True, space_after=2)
code(doc, """$ python manage.py test
Ran 67 tests in 0.040s
OK
[exit code: 0]

$ python manage.py gate --modo ai
  pruebas     pasan  67   fallan   0
  carritos    6 de 6 como se esperaba
  veredicto   ACEPTAR (código de salida 0)
[exit code: 0]""")
para(doc,
     "Pregunta que lo rechaza: la 1. float(l.unit_price) * l.qty introduce error binario en una suma de "
     "dinero y Decimal(str(value)) solo enmascara el float ya contaminado; además se abandona el helper "
     "money de shop/domain.py. Se revirtió con git checkout -- shop/cart_ai.py y la compuerta volvió a "
     "ACEPTAR (código 0). Queda registrado como rechazo de ejercicio en la bitácora.")
evidence(doc, "rechazo_01.diff", "43_rechazo_01_test.txt", "44_rechazo_01_gate.txt",
         "45_rechazo_01_revertido_gate.txt")

# ------------------------------- 9. Uso de IA -----------------------------
landscape(doc)
h1(doc, "10. Paso 18: bitácora de uso de IA", page_break=False)
meta(doc, "copia de docs/uso_de_ia.md · una fila por turno")
para(doc,
     "La fila 0 es el prompt inicial. Las filas se agregaron al cerrar cada turno, no al final. Cada "
     "fila indica qué se pidió, qué cambió el agente, si se aceptó o rechazó, qué pregunta o prueba lo "
     "decidió y qué archivo de docs/evidencia/ lo demuestra.")
rows = usage_rows()
table(doc,
      ["N.º", "Pedido", "Qué respondió / cambió el agente", "Aceptado / rechazado", "Motivo", "Evidencia"],
      rows,
      widths=[0.9, 4.4, 7.6, 2.6, 5.2, 4.3], size=6.8, header_size=7.4)
evidence(doc, "docs/uso_de_ia.md")

# ------------------------------- 10. Conclusiones -------------------------
portrait(doc)
h1(doc, "11. Conclusiones", page_break=False)
for text in [
    "Las pruebas son el oráculo, pero no lo saben todo. La compuerta junta las pruebas y la auditoría "
    "con --replay y da un veredicto con código de salida; aun así, el rechazo de ejercicio demuestra que "
    "un diff puede quedar en verde y violar una regla que nadie escribió como prueba.",
    "El dinero exige Decimal de principio a fin. Un solo float en _subtotal o en _money basta para que "
    "0,845 se convierta en 0,84; el helper money() con ROUND_HALF_UP concentra la regla y evita "
    "repetirla en cada método.",
    "El precio solo puede venir del catálogo. Aceptar price del cliente no rompe la pantalla normal, pero "
    "permite comprar una laptop de S/ 2 499,90 por S/ 0,01: es el fallo de seguridad del carrito de la IA "
    "y el único que solo se ve por la API.",
    "Las reglas de cantidad y existencias deben cumplirse con cualquier secuencia: add tiene que contar "
    "lo que ya está en el carrito, remove no puede dejar negativos y checkout debe re-verificar todas las "
    "líneas antes de restar ninguna. En producción, esa última comprobación necesita transacción y "
    "bloqueo de fila (select_for_update o un UPDATE condicional) para evitar la carrera entre dos "
    "clientes.",
    "Repetir una acción debe dar el mismo resultado. Un cupón guardado en una lista y sumado dos veces "
    "descuenta doble ante un doble clic o un reintento de red; un único cupón que se reemplaza mantiene "
    "la operación idempotente.",
    "La auditoría es una red de seguridad, no la puerta de entrada: avisa después del cálculo y no impide "
    "que una petición mala se cobre. La regla tiene que estar en el servidor, en add, remove, checkout y "
    "add_from_payload.",
]:
    bullet(doc, text)

# ------------------------------- 11. Estructura ---------------------------
h1(doc, "12. Estructura del proyecto y captura pendiente")
para(doc,
     "Árbol real de los archivos versionados del repositorio (generado con git ls-files y ordenado como "
     "árbol; se omiten .venv, db.sqlite3 y los archivos ignorados):")
tree_text = (EVIDENCE / "53_tree.txt").read_text(encoding="utf-8")
code(doc, tree_text, size=6.6, fill="F7F9FB")
p = doc.add_paragraph()
r = p.add_run("[PEGAR CAPTURA DE VS CODE]")
set_run_font(r, "Consolas", 16, True, RED)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(18)
p.paragraph_format.space_after = Pt(6)
p = doc.add_paragraph()
r = p.add_run(
    "Marcador para la captura real del explorador de VS Code. No se generó ninguna imagen de pantalla "
    "falsa: la captura la inserta el alumno en este espacio."
)
set_run_font(r, "Calibri", 8.5, False, GRAY, italic=True)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

para(doc, "Verificación final del repositorio:", bold=True, space_before=12, space_after=2)
code(doc, """$ python manage.py test                         -> 67 tests, OK (exit 0)
$ SHOP_MODE=ai python manage.py test            -> 67 tests, OK (exit 0)
$ python manage.py gate                        -> ACEPTAR (exit 0)
$ python manage.py gate --modo ai              -> ACEPTAR (exit 0)
$ python manage.py gate --json                 -> verdict: aceptar, failed: 0
$ python manage.py gate --modo ai --json       -> verdict: aceptar, failed: 0
$ python manage.py seed_shop                   -> Catálogo cargado: 8 productos.""")
evidence(doc, "46_test_final_equipo.txt", "47_test_final_ai.txt",
         "48_gate_final_equipo.txt", "49_gate_final_ai.txt",
         "50_gate_final_equipo_json.txt", "51_gate_final_ai_json.txt",
         "52_seed_restore_final.txt")

# ------------------------------- Guardar -----------------------------------
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUTPUT)
print(f"Saved {OUTPUT}")
print(f"Size: {OUTPUT.stat().st_size} bytes")
