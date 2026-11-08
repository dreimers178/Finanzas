"""Hoja de calificación del juicio de expertos (para llenar en Word o a mano). Genera DOCX en ../juicio/."""
import pathlib
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = pathlib.Path(__file__).parent
OUT = HERE.parent / "juicio"; OUT.mkdir(exist_ok=True)
NAVY = RGBColor(0x14, 0x21, 0x3D)

DIMS = [
 ("Dimensión 1. Principios de alimentación saludable y selección de alimentos", [
  "Mejor descripción de una alimentación saludable", "Alimento más cercano a las recomendaciones",
  "Bebida de consumo más frecuente", "Producto de consumo menos frecuente",
  "Colación o snack escolar adecuado", "Importancia de variar los alimentos"]),
 ("Dimensión 2. Grupos de alimentos y clasificación", [
  "Los cinco grupos de alimentos", "Conjunto solo de verduras y frutas",
  "Alimento del grupo de cereales y tubérculos", "Conjunto solo de leguminosas",
  "Clasificación del pescado y el aguacate", "Grupo de frijoles, tortilla, huevo y manzana"]),
 ("Dimensión 3. Aportes y función nutricional básica de los alimentos", [
  "Nutrientes de las verduras y frutas", "Diferencia entre cereal integral y refinado",
  "Fuente de proteína de origen vegetal", "Combinación de leguminosa con cereal",
  "Característica del aguacate, semillas y aceites", "Relación entre alimento y aporte nutricional"]),
 ("Dimensión 4. Aplicación de las recomendaciones en situaciones cotidianas", [
  "Julia: comida variada y equilibrada", "Diego: alimento para variar su comida",
  "Desayuno con mayor variedad de grupos", "Sofía: qué agregar para más variedad",
  "Ana: cambio en su comida escolar", "Luis: aplicar varias recomendaciones a la vez"]),
]

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
sec.left_margin = sec.right_margin = Cm(1.4)
sec.top_margin = sec.bottom_margin = Cm(1.1)
st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(9.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
st.paragraph_format.space_after = Pt(0); st.paragraph_format.space_before = Pt(0)


def P(text="", bold=False, size=None, color=None, align=None, after=2, italic=False):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(after)
    if align is not None: p.alignment = align
    if text:
        r = p.add_run(text); r.bold = bold; r.italic = italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color
    return p


def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), hex_); tcPr.append(s)


def cell_text(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    p = cell.paragraphs[0]; p.alignment = align
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


P("Hoja de calificación · Juicio de expertos", bold=True, size=14, color=NAVY, after=0)
P("Cuestionario de conocimiento nutricional para estudiantes de 5.º de primaria (24 reactivos)", size=10, after=4)
P("Nombre: ________________________________   Perfil profesional: ________________________   Fecha: __________", after=4)
p = P(after=1)
p.add_run("Instrucciones: ").bold = True
p.add_run("con el instrumento a la vista, escriba en cada casilla un número del 1 al 4 según la escala. "
          "La columna Exactitud la llenan solo los jueces nutriólogos; los demás escriben N/A. "
          "Las observaciones son opcionales. Su nombre es confidencial y no aparecerá en ningún documento.")
P("Escala:  1 = No cumple   ·   2 = Cumple de manera insuficiente; requiere modificaciones importantes   ·   "
  "3 = Cumple adecuadamente; puede requerir ajustes menores   ·   4 = Cumple totalmente; no requiere modificaciones",
  size=8.5, after=4)

widths = [Cm(0.8), Cm(6.6), Cm(1.85), Cm(1.85), Cm(1.85), Cm(1.85), Cm(4.0)]
t = doc.add_table(rows=1, cols=7); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
for c, h in zip(t.rows[0].cells, ["N.º", "Reactivo (tema)", "Claridad", "Coherencia", "Pertinencia", "Exactitud*", "Observaciones"]):
    cell_text(c, h, bold=True, size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=RGBColor(0xFF, 0xFF, 0xFF)); shade(c, "14213D")
n = 0
for dim, temas in DIMS:
    row = t.add_row(); m = row.cells[0].merge(row.cells[6])
    cell_text(m, dim, bold=True, size=8.5); shade(m, "E6EAF2")
    for tema in temas:
        n += 1; row = t.add_row()
        cell_text(row.cells[0], str(n), align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(row.cells[1], tema, size=8.5)
        row.height = Cm(0.6)
for col, w in zip(t.columns, widths): col.width = w
for row in t.rows:
    for c, w in zip(row.cells, widths): c.width = w
tblPr = t._tbl.tblPr; lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
P("* Exactitud del contenido nutricional: solo jueces nutriólogos.", size=8, italic=True, after=6)

P("Valoración global (marque con una X)", bold=True, size=10, color=NAVY, after=2)
for q, ops in [("1. ¿El instrumento es adecuado para evaluar el conocimiento nutricional en 5.º de primaria?",
                "(   ) Sí, sin modificaciones   (   ) Sí, con modificaciones menores   (   ) Sí, con modificaciones importantes   (   ) No"),
               ("2. ¿Los reactivos representan adecuadamente las cuatro dimensiones?", "(   ) Sí   (   ) Parcialmente   (   ) No"),
               ("3. ¿El lenguaje es adecuado para estudiantes de 5.º de primaria?", "(   ) Sí   (   ) Parcialmente   (   ) No")]:
    P(q, after=0); P("      " + ops, size=9, after=3)
P("Observaciones generales: ____________________________________________________________________", after=4)
P("__________________________________________________________________________________________", after=0)

doc.save(OUT / "Hoja_calificacion_juicio_expertos.docx")
print("OK")
