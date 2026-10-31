"""Resumen para el CITEEA 2026 en la plantilla oficial (A4, Times New Roman 12, interlineado 1.5)."""
import pathlib
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

OUT = pathlib.Path(__file__).resolve().parents[1] / "congreso" / "Resumen_CITEEA_Reimers.docx"
OUT.parent.mkdir(exist_ok=True)

TITULO = "Nutrición en el aula: taller de microaprendizaje en línea para quinto de primaria (diseño pretest-postest)"
EJE = "Metodologías activas y pedagogías emergentes"
AUTORES = [
    ("Anakaren Ernestina Reimers Reyes", "Universidad Autónoma de Zacatecas", "México", "ana.reimers@gmail.com"),
    ("Eduardo Rivera Arteaga", "Universidad Autónoma de Zacatecas", "México", "edrivera@uaz.edu.mx"),
]
PATROCINIO = ("Proyecto desarrollado con el apoyo de una beca de posgrado otorgada a la primera autora por la Secretaría de "
              "Ciencia, Humanidades, Tecnología e Innovación (SECIHTI), México.")
RESUMEN = (
    "La educación nutricional en la infancia favorece la adquisición de conocimientos sobre alimentación saludable. El objetivo "
    "de esta investigación es incrementar el conocimiento nutricional de estudiantes de quinto grado de primaria mediante un "
    "taller educativo en línea basado en microaprendizaje. El taller se diseñó con el modelo ADDIE y organiza ocho cápsulas "
    "digitales breves con base en el Plato del Bien Comer y las Guías Alimentarias Saludables y Sostenibles para la Población "
    "Mexicana. Se trata de una investigación aplicada, con enfoque cuantitativo y diseño cuasiexperimental de preprueba y "
    "posprueba con un solo grupo. La muestra se integra por 25 estudiantes del Instituto Educativo Ammadeus, en Guadalupe, "
    "Zacatecas. Se aplicará un cuestionario de 24 reactivos antes y después del taller. Como avance, el instrumento se "
    "encuentra en validación de contenido mediante juicio de expertos. Posteriormente se implementará el taller y se "
    "determinará si hubo un cambio significativo en el conocimiento nutricional."
)
PALABRAS = "microaprendizaje, educación nutricional, educación primaria."

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.5))
st = doc.styles["Normal"]
st.font.name = "Times New Roman"
st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.space_after = Pt(0)
st.paragraph_format.space_before = Pt(0)
st.paragraph_format.line_spacing = 1.0


def p(text="", align=None, bold=False, italic=False, spacing=1.0, before=0, indent=None, hl=False):
    par = doc.add_paragraph()
    if align:
        par.alignment = align
    par.paragraph_format.line_spacing = spacing
    par.paragraph_format.space_before = Pt(before)
    if indent:
        par.paragraph_format.first_line_indent = Cm(indent)
    if text:
        r = par.add_run(text)
        r.bold, r.italic = bold, italic
        if hl:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return par


C, R, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY
p(TITULO, C, bold=True)
p(EJE, C, bold=True)
p()
for nombre, inst, pais, correo in AUTORES:
    p(nombre, R, bold=True)
    p(inst, R)
    p(pais, R)
    p(correo, R)
p()
p(PATROCINIO, J, italic=True)
p()
p(RESUMEN, J, spacing=1.5, indent=1.25)
p()
kw = p(align=J, spacing=1.5)
r = kw.add_run("Palabras clave: "); r.bold = True
kw.add_run(PALABRAS)
doc.save(OUT)
print("OK", OUT, "palabras del resumen:", len(RESUMEN.split()), "| palabras del título:", len(TITULO.split()))
