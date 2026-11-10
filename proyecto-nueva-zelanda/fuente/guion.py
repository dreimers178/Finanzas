import sys
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
GREEN = RGBColor(0x0E, 0x4D, 0x45)
FERN = RGBColor(0x2F, 0x8F, 0x5B)
OCEAN = RGBColor(0x1E, 0x6F, 0x8C)
GRAY = RGBColor(0x5B, 0x6B, 0x70)
JL, PB = "José Luis", "Pablo"

doc = Document()
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2)
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(11.5)
for name, size in (("Heading 1", 17), ("Heading 2", 13.5)):
    h = doc.styles[name]
    h.font.name = "Cambria"
    h.font.size = Pt(size)
    h.font.color.rgb = GREEN
    h.font.bold = True


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexcolor)
    tcPr.append(sh)


def para(text="", bold=False, size=None, color=None, align=None, italic=False, after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p


def bullet(text, boldlead=None):
    p = doc.add_paragraph(style="List Bullet")
    if boldlead:
        r = p.add_run(boldlead)
        r.bold = True
    p.add_run(text)
    p.paragraph_format.space_after = Pt(3)
    return p


# ---------------- PORTADA ----------------
lp = doc.add_paragraph(); lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
lp.add_run().add_picture("logo.png", width=Cm(3.5))
para("INSTITUTO EDUCATIVO AMMADEUS", bold=True, size=15, color=FERN, align=WD_ALIGN_PARAGRAPH.CENTER, after=30)
para("Países que construyen bienestar:", bold=True, size=26, color=GREEN, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
para("modelos para una mejor sociedad", bold=True, size=22, color=FERN, align=WD_ALIGN_PARAGRAPH.CENTER, after=24)
para("País investigado: Nueva Zelanda (Aotearoa)", bold=True, size=15, color=OCEAN, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
para("GUION DE LA PRESENTACIÓN", bold=True, size=14, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, after=40)
for lab, val in (("Materia", "Educación Financiera"), ("Grado", "2.º de secundaria"), ("Maestra", "Samantha Rubí Lara Ceh"),
                 ("Alumnos", "José Luis De Anda Reimers\nPablo De Santiago Hurtado Martin Del Campo"), ("Fecha", "Noviembre de 2026")):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(lab + ": ")
    r.bold = True
    r.font.color.rgb = GREEN
    r.font.size = Pt(13)
    r2 = p.add_run(val)
    r2.font.size = Pt(13)
doc.add_page_break()

# ---------------- INSTRUCCIONES ----------------
doc.add_heading("Cómo usar este guion", level=1)
bullet(" unos 10 minutos de exposición + 2 o 3 minutos de preguntas.", "Duración:")
bullet(" José Luis y Pablo se turnan diapositiva por diapositiva. Cada quien adelanta la diapositiva de la que va a hablar.", "Turnos:")
bullet(" no lean palabra por palabra. Apréndanse las ideas y los números clave; lo demás díganlo con sus propias palabras. Ensayen juntos al menos dos veces con cronómetro.", "Consejo:")
bullet(" los textos también vienen en las «Notas del orador» de cada diapositiva del PowerPoint.", "Notas:")
bullet(" «NZ$» = dólares neozelandeses; «US$» = dólares estadounidenses. Un dólar neozelandés vale menos que uno estadounidense.", "Monedas:")
para("")

# ---------------- GUION ----------------
import json
G = [tuple(g) for g in json.load(open("guion_corto.json", encoding="utf-8"))]

doc.add_heading("Guion diapositiva por diapositiva", level=1)
sections = {1: "Inicio", 3: "A. Conociendo el país", 8: "B. Modelos y propuestas de bienestar (apartado principal)", 15: "C. ¿Cómo podemos replicarlo en México?", 19: "D. Conclusiones"}
for num, title, t, lines in G:
    if num in sections:
        doc.add_heading(sections[num], level=2)
    tbl = doc.add_table(rows=1 + len(lines), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    m = hdr[0].merge(hdr[1])
    m.text = ""
    r = m.paragraphs[0].add_run(f"Diapositiva {num if num != 20 else '20–21'}: {title}")
    r.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r2 = m.paragraphs[0].add_run(f"   ⏱ {t}")
    r2.font.color.rgb = RGBColor(0xE2, 0xA2, 0x3B)
    r2.bold = True
    shade(m, "0E4D45")
    for i, (who, text) in enumerate(lines, start=1):
        c0, c1 = tbl.rows[i].cells
        c0.text = ""
        rr = c0.paragraphs[0].add_run(who)
        rr.bold = True
        rr.font.color.rgb = FERN if who == JL else OCEAN
        shade(c0, "EAF3EE")
        c1.text = text
    tbl.autofit = False
    tbl.columns[0].width = Cm(2.8)
    tbl.columns[1].width = Cm(13.8)
    for row in tbl.rows:
        row.cells[0].width = Cm(2.8)
        row.cells[1].width = Cm(13.8)
    para("", after=4)

# ---------------- PREGUNTAS ----------------
doc.add_heading("Posibles preguntas de la maestra o del grupo (para practicar)", level=1)
QA = [
    ("¿Por qué eligieron KiwiSaver como su propuesta principal?", "Porque ayuda en tres cosas a la vez (retiro, vivienda y hábito de ahorro), la persona también aporta y en México ya existen las Afores para apoyarse."),
    ("¿Qué es el interés compuesto?", "Es cuando los intereses que gana tu dinero también generan intereses. Por eso, entre más pronto empiezas a ahorrar, más crece tu dinero."),
    ("¿Cómo pagaría México esos programas si cobra pocos impuestos?", "Empezando en las zonas con más necesidad, con montos pequeños, y combinando el ahorro de la persona con un bono limitado del gobierno."),
    ("Si Nueva Zelanda tiene tanto bienestar, ¿por qué México es casi igual de feliz?", "Porque la felicidad también depende de la familia, los amigos y la comunidad, no solo del dinero o los servicios públicos."),
    ("¿Qué es el PIB?", "El valor de todo lo que produce un país en un año. El PIB por persona es ese valor dividido entre sus habitantes."),
    ("¿Qué es la economía informal?", "Trabajos sin contrato ni seguro social; por eso esas personas no tienen Afore ni IMSS."),
]
for q, a in QA:
    p = doc.add_paragraph()
    r = p.add_run("P: " + q)
    r.bold = True
    r.font.color.rgb = GREEN
    p.paragraph_format.space_after = Pt(1)
    p2 = doc.add_paragraph("R: " + a)
    p2.paragraph_format.space_after = Pt(8)

# ---------------- GLOSARIO ----------------
doc.add_heading("Mini glosario", level=1)
for t, d in [
    ("PIB:", " valor de todo lo que produce un país en un año."),
    ("IDH:", " Índice de Desarrollo Humano de la ONU; combina salud, educación e ingreso."),
    ("Esperanza de vida:", " años que vive en promedio una persona."),
    ("Edad mediana:", " edad que divide a la población en dos mitades iguales."),
    ("Pensión universal:", " pago para todos los adultos mayores, hayan trabajado o no."),
    ("Afore:", " cuenta de ahorro para el retiro de los trabajadores formales en México."),
    ("Energía renovable:", " energía que no se acaba, como la del agua, el sol, el viento o el calor de la tierra."),
    ("Maorí:", " pueblo originario de Nueva Zelanda."),
]:
    bullet(d, t)

# ---------------- BIBLIOGRAFÍA ----------------
doc.add_page_break()
doc.add_heading("Bibliografía (formato APA 7)", level=1)
BIB = [
    "ACC. (2025). Annual report 2025. Accident Compensation Corporation. https://www.acc.co.nz",
    "Employment New Zealand. (2026). Minimum wage is increasing on 1 April 2026. https://www.employment.govt.nz",
    "Fondo Monetario Internacional. (2025). World Economic Outlook Database. https://www.imf.org/en/Publications/WEO",
    "Helliwell, J. F., Layard, R., Sachs, J. D., De Neve, J.-E., Aknin, L. B. y Wang, S. (Eds.). (2026). World Happiness Report 2026. Wellbeing Research Centre, University of Oxford. https://worldhappiness.report",
    "Inland Revenue (IRD). (2025). KiwiSaver changes. https://www.ird.govt.nz/kiwisaver-changes",
    "INEGI. (2026). Indicadores de Ocupación y Empleo, septiembre de 2026 [Comunicado de prensa]. https://www.inegi.org.mx",
    "Ministry for the Environment. (2019). Climate Change Response (Zero Carbon) Amendment Act 2019. https://environment.govt.nz",
    "Ministry of Business, Innovation and Employment (MBIE). (2026). Energy in New Zealand 2026. https://www.mbie.govt.nz",
    "Ministry of Education. (2026). Healthy school lunches programme. https://www.education.govt.nz",
    "New Zealand Treasury. (2019). The Wellbeing Budget 2019. https://www.treasury.govt.nz",
    "OCDE. (2023). PISA 2022 Results (Volume I): The state of learning and equity in education. OECD Publishing. https://doi.org/10.1787/53f23881-en",
    "OCDE. (2025). Revenue Statistics 2025. OECD Publishing. https://www.oecd.org",
    "Pharmac. (2025). How Pharmac works. https://pharmac.govt.nz",
    "Programa de las Naciones Unidas para el Desarrollo (PNUD). (2025). Informe sobre Desarrollo Humano 2025: Una cuestión de elección. https://hdr.undp.org",
    "Secretaría de Educación Pública y Secretaría de Salud. (2024). Lineamientos generales para la preparación, distribución y expendio de alimentos y bebidas en escuelas. Diario Oficial de la Federación. https://www.dof.gob.mx",
    "Stats NZ. (2025). National and subnational period life tables: 2022–2024. https://www.stats.govt.nz",
    "Stats NZ. (2025). National population estimates: At 30 June 2025. https://www.stats.govt.nz",
    "Stats NZ. (2026). Labour market statistics: September 2026 quarter. https://www.stats.govt.nz",
    "Stats NZ. (2026). Overseas merchandise trade: December 2025. https://www.stats.govt.nz",
    "Tenancy Services. (2025). Healthy homes standards. Ministry of Housing and Urban Development. https://www.tenancy.govt.nz",
    "Transparency International. (2026). Corruption Perceptions Index 2025. https://www.transparency.org/cpi",
    "Work and Income New Zealand. (2026). Benefit rates at 1 April 2026: NZ Superannuation. https://www.workandincome.govt.nz",
]
for b in BIB:
    p = doc.add_paragraph(b)
    pf = p.paragraph_format
    pf.left_indent = Cm(1.27)
    pf.first_line_indent = Cm(-1.27)
    pf.space_after = Pt(6)
para("Nota: los ejemplos de KiwiSaver y de «Mi Ahorro Automático» son cálculos propios con tasas de interés supuestas (5 % y 4 % anual), solo para ilustrar el interés compuesto.", italic=True, size=10, color=GRAY)

doc.save(OUT)
words = sum(len(t.split()) for _, _, _, L in [(0,)+g[1:] for g in G] for _, t in L)
print("saved", OUT, "words in script:", words, "≈ min at 140 wpm:", round(words / 140, 1))
