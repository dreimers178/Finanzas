"""Genera: (1) guía con los textos para Classroom (PDF) y (2) plantilla de evidencias (DOCX)."""
import html, pathlib, subprocess
from contenido import *
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn

HERE = pathlib.Path(__file__).parent

# ---------------------------------------------------------------- 1. Guía HTML
def bloque(items):
    out = []
    for it in items:
        if isinstance(it, str):
            out.append(f"<p>{html.escape(it)}</p>")
        elif isinstance(it, tuple):
            k, v = it
            out.append(f"<p><b>{html.escape(k)}{':' if v else ':'}</b> {html.escape(v) if v else ''}</p>")
        else:
            out.append("<ul>" + "".join(f"<li>{html.escape(x)}</li>" for x in it) + "</ul>")
    return "".join(out)

cards = []
for i, (tit, tipo, items) in enumerate(APARTADOS, start=1):
    cards.append(f'''<section class="card"><div class="tag">Apartado {i} · {tipo}</div>
<div class="field">Título en Classroom</div><div class="title">{html.escape(tit)}</div>
<div class="field">Texto (copiar y pegar en la descripción)</div><div class="body">{bloque(items)}</div></section>''')

guia = f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Guía Unidad 3 Classroom</title><style>
@font-face{{font-family:"Source Sans 3";src:url(fonts/SourceSans3.woff2) format("woff2");font-weight:200 900}}
@font-face{{font-family:"Oswald";src:url(fonts/Oswald.woff2) format("woff2");font-weight:200 700}}
@page{{size:letter;margin:1.6cm 1.7cm 1.6cm}}
html{{font-family:"Source Sans 3",sans-serif;font-size:12pt;line-height:1.45;color:#111}} body{{margin:0}}
h1{{font-family:"Oswald";color:#8b1538;font-size:22pt;margin:0}} .sub{{color:#14213d;font-size:13pt;margin:.1cm 0 .4cm}}
.steps{{background:#f6f1f3;border-left:.15cm solid #8b1538;padding:.3cm .45cm;margin-bottom:.4cm}}
.steps ol{{margin:.1cm 0;padding-left:.6cm}}
.card{{border:1px solid #d9cfd3;border-radius:.2cm;padding:.35cm .45cm;margin:.35cm 0;break-inside:avoid}}
.tag{{display:inline-block;background:#14213d;color:#fff;font-family:"Oswald";font-size:10.5pt;padding:.03cm .25cm;border-radius:.15cm;letter-spacing:.03em}}
.field{{font-size:10pt;color:#13879c;font-weight:700;text-transform:uppercase;letter-spacing:.04em;margin-top:.22cm}}
.title{{font-weight:800;font-size:13.5pt;color:#8b1538}}
.body p{{margin:.12cm 0}} .body ul{{margin:.08cm 0 .12cm;padding-left:.6cm}} .body li{{margin:.05cm 0}}
.warn{{background:#fff3c4;padding:0 .1cm;border-radius:.08cm}}
</style></head><body>
<h1>Unidad 3 en Classroom: textos listos para pegar</h1>
<div class="sub">Curso: <b>{CURSO}</b> · Aplicaciones Tecnológicas en la Nube</div>
<div class="steps"><b>Orden para crear la unidad en Classroom</b><ol>
<li>En <b>Trabajo en clase → Crear → Tema</b>, escribe: <b>{TEMA}</b></li>
<li>Crea cada apartado con <b>Crear → Material</b> (apartados 1 a 4) y <b>Crear → Tarea</b> (apartado 5), y asígnalos al tema «Unidad 3».</li>
<li>En el material didáctico, agrega el enlace con el botón <b>Agregar → Vínculo</b>: <br><span style="font-size:10.5pt">{MATERIAL_ENLACE}</span></li>
<li>En el material complementario, agrega el enlace de tu concurso de preguntas con <b>Agregar → Vínculo</b>.</li>
<li>En la tarea, escribe <b>100 puntos</b>, agrega como recursos el enlace de la infografía y del concurso, y pon la fecha de entrega.</li>
<li>Revisa que todo quede publicado dentro del tema y que los enlaces abran; luego toma las capturas para la plantilla de evidencias.</li>
</ol>
<p style="margin:.1cm 0"><span class="warn">Revisa que el título de tu concurso en Flippity sea «{CONCURSO_TITULO}».</span></p></div>
{''.join(cards)}
</body></html>'''
(HERE / "guia.html").write_text(guia, encoding="utf-8")

# ---------------------------------------------------------------- 2. Plantilla DOCX de evidencias
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
for s in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, s, Cm(2.5))
st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.space_after = Pt(0)
RED = RGBColor(0xA0, 0x10, 0x10)

def P(text="", bold=False, align=None, size=None, color=None, italic=False, after=0, before=0):
    p = doc.add_paragraph()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
    if text:
        r = p.add_run(text); r.bold = bold; r.italic = italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color
    return p

def label_value(lbl, val):
    p = P(align=WD_ALIGN_PARAGRAPH.JUSTIFY, after=10)
    r = p.add_run(lbl + " "); r.bold = True
    p.add_run(val)

C = WD_ALIGN_PARAGRAPH.CENTER
# Portada (mismo modelo que la evidencia de la Unidad 2)
t = doc.add_table(rows=1, cols=3); t.autofit = False
for col, w in zip(t.columns, (4.3, 7.9, 4.3)): col.width = Cm(w)
for cell, w in zip(t.rows[0].cells, (4.3, 7.9, 4.3)): cell.width = Cm(w)
for cell, img, h in zip(t.rows[0].cells, ("escudo_uaz_w", "logo_mtie_w", "laberinto_uads_w"), (2.4, 2.0, 2.3)):
    cell.paragraphs[0].alignment = C
    cell.paragraphs[0].add_run().add_picture(str(HERE / "logos" / f"{img}.png"), height=Cm(h))
P(before=40)
for line in ("UNIVERSIDAD AUTÓNOMA DE ZACATECAS", "“FRANCISCO GARCÍA SALINAS”"):
    P(line, align=C)
P(before=18)
for line in ("UNIDAD ACADÉMICA DE DOCENCIA SUPERIOR", "MAESTRÍA EN TECNOLOGÍA INFORMÁTICA EDUCATIVA"):
    P(line, align=C)
P(before=18)
P("APLICACIONES TECNOLÓGICAS EN LA NUBE", align=C)
P("U12 Evidencia Unidad 3 en Classroom", align=C)
P(before=18); P("Facilitador:", align=C); P("Dra. Susana Cordero Dávila", align=C)
P(before=18); P("Alumna:", align=C); P("L.N.C.A. Anakaren Ernestina Reimers Reyes", align=C)
P(before=90); P("Zacatecas, Zac., Noviembre, 2026", align=C)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

P("Datos de la evidencia", bold=True, color=RED, after=10)
label_value("Curso en Classroom:", CURSO + " (5.º de primaria).")
label_value("Tema:", TEMA[:-1] + " (hábitos de alimentación y recomendaciones de las Guías Alimentarias).")
label_value("Apartados creados:", "Aprendizaje esperado, Instrucciones, Material didáctico (principal), "
            "Material complementario (concurso de preguntas) y Tarea.")

figs = [
 ("Estructura de la Unidad 3", "Tema “Unidad 3: Construyo hábitos saludables” con sus apartados en el curso Nutrición básica: Mi Plato Saludable.", ["1a_temas", "1b_unidad3"]),
 ("Aprendizaje esperado", "Apartado Aprendizaje esperado de la Unidad 3.", ["2_aprendizaje"]),
 ("Instrucciones", "Apartado Instrucciones de la Unidad 3.", ["3_instrucciones"]),
 ("Material principal", "Material didáctico: Hábitos saludables, las Guías Alimentarias, con el enlace a la infografía oficial de gob.mx.", ["4_material_didactico"]),
 ("Material de repaso", "Material complementario con el concurso de preguntas «" + CONCURSO_TITULO + "» y su enlace.", ["5_material_complementario"]),
 ("Tarea", "Tarea: Mi reto de hábitos saludables, con instrucciones, puntos y rúbrica.", ["6_tarea"]),
]
CAP = HERE / "capturas"
faltan = []
for i, (h, cap, imgs) in enumerate(figs, start=1):
    P(h, bold=True, color=RED, before=14, after=6).paragraph_format.keep_with_next = True
    files = [CAP / f"{x}.png" for x in imgs if (CAP / f"{x}.png").exists()]
    if files:
        for f in files:
            p = P(align=C, after=4); p.paragraph_format.keep_with_next = True
            from PIL import Image
            iw, ih = Image.open(f).size
            if 16.5 * ih / iw > 12.5:
                p.add_run().add_picture(str(f), height=Cm(12.5))
            else:
                p.add_run().add_picture(str(f), width=Cm(16.5))
    else:
        faltan.append(h)
        box = doc.add_table(rows=1, cols=1); box.style = "Table Grid"
        c = box.rows[0].cells[0]; c.height = Cm(8)
        cp = c.paragraphs[0]; cp.alignment = C
        r = cp.add_run("\n\n[Pega aquí la captura]\n\n"); r.italic = True; r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    P(f"Figura {i}. {cap}", italic=True, size=10, after=6, before=4)
doc.save(HERE / "U12_Evidencia_Unidad_3_Reimers.docx")
print("OK; faltan:", faltan)
