"""Actividad 4 (semana 10): presentación de avances del protocolo (16:9, editable en PowerPoint / Google Slides)."""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY = RGBColor(0x14, 0x21, 0x3D)
ORANGE = RGBColor(0xEF, 0x7D, 0x2C)
RED = RGBColor(0xC6, 0x28, 0x28)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
TEAL = RGBColor(0x13, 0x87, 0x9C)
GREY = RGBColor(0xF2, 0xF2, 0xF2)
SOFT = RGBColor(0xFD, 0xF0, 0xE6)
INK = RGBColor(0x1E, 0x1E, 0x1E)
MUTED = RGBColor(0x5B, 0x5B, 0x5B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YEL = RGBColor(0xFF, 0xF3, 0xC4)
HEAD, BODY = "Arial", "Arial"

prs = Presentation()
prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)
W, H = prs.slide_width, prs.slide_height
BL = prs.slide_layouts[6]


def rect(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = sl.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(1.25)
    else: s.line.fill.background()
    s.shadow.inherit = False
    return s


def text(sl, x, y, w, h, paras, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.1, font=BODY):
    tb = sl.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(.1); tf.margin_top = tf.margin_bottom = Cm(.05)
    paras = paras if isinstance(paras, list) else [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing
        for t, o in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run(); r.text = t; f = r.font
            f.name = o.get("font", font); f.size = Pt(o.get("size", size)); f.bold = o.get("bold", bold)
            f.color.rgb = o.get("color", color); f.italic = o.get("italic", False)
    return tb


def header(sl, title, n):
    rect(sl, 0, 0, W, Cm(2.5), NAVY)
    rect(sl, 0, Cm(2.5), W, Cm(.15), ORANGE)
    text(sl, Cm(1.3), Cm(.45), Cm(28), Cm(1.6), title, size=26, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, Cm(1.3), H - Cm(.95), Cm(26), Cm(.6), "Nutrición en el aula: taller de microaprendizaje en línea para quinto de primaria · A. E. Reimers Reyes",
         size=9.5, color=MUTED)
    text(sl, W - Cm(3), H - Cm(.95), Cm(1.8), Cm(.6), str(n), size=9.5, color=MUTED, align=PP_ALIGN.RIGHT)


def table(sl, x, y, w, rows, colw, size=13, rowh=Cm(1.0), head=True, first_col_fill=None, aligns=None):
    shp = sl.shapes.add_table(len(rows), len(rows[0]), x, y, w, rowh * len(rows)); t = shp.table
    for j, cw in enumerate(colw): t.columns[j].width = cw
    for i, row in enumerate(rows):
        t.rows[i].height = rowh
        for j, val in enumerate(row):
            c = t.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = c.margin_right = Cm(.22); c.margin_top = c.margin_bottom = Cm(.06)
            p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = str(val)
            p.alignment = (aligns[j] if aligns else PP_ALIGN.LEFT)
            f = r.font; f.name = BODY; f.size = Pt(size); c.fill.solid()
            if head and i == 0:
                c.fill.fore_color.rgb = NAVY; f.color.rgb = WHITE; f.bold = True
            elif first_col_fill is not None and j == 0:
                c.fill.fore_color.rgb = first_col_fill; f.color.rgb = WHITE; f.bold = True
            else:
                c.fill.fore_color.rgb = WHITE if i % 2 else GREY; f.color.rgb = INK
    return shp


def box(sl, x, y, w, h, title, lines, fill=SOFT, bar=ORANGE, tsize=15, size=13):
    rect(sl, x, y, w, h, fill); rect(sl, x, y, Cm(.22), h, bar)
    paras = [[(title, {"bold": True, "color": NAVY, "size": tsize})]] + [[(l, {})] if isinstance(l, str) else l for l in lines]
    text(sl, x + Cm(.5), y + Cm(.25), w - Cm(.7), h - Cm(.4), paras, size=size, spacing=1.15)


n = 0
# 1 Portada --------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1
rect(sl, 0, 0, W, Cm(10.4), GREY)
for img, x, h in (("escudo_uaz_w", 18.4, 3.0), ("logo_mtie_w", 21.7, 2.2), ("laberinto_uads_w", 29.3, 2.9)):
    sl.shapes.add_picture(f"logos/{img}.png", Cm(x), Cm(.7), height=Cm(h))
text(sl, Cm(1.5), Cm(4.6), W - Cm(3), Cm(4), [
    [("NUTRICIÓN EN EL AULA: TALLER DE MICROAPRENDIZAJE EN LÍNEA PARA QUINTO DE PRIMARIA (DISEÑO PRETEST-POSTEST)", {"bold": True, "size": 25, "color": NAVY})],
    [("Actividad 4 · Presentación de avances", {"size": 18, "color": ORANGE, "bold": True})]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15)
text(sl, Cm(3.2), Cm(11.4), Cm(24), Cm(6), [
    [("Alumna: ", {"bold": True}), ("L.N.C.A. Anakaren Ernestina Reimers Reyes", {})],
    [("Director: ", {"bold": True}), ("Dr. Eduardo Rivera Arteaga", {})],
    [("Facilitador: ", {"bold": True}), ("Dr. Leonel Ruvalcaba Arredondo", {})],
    [("Seminario de Proyecto Educativo y Estancia IV · Maestría en Tecnología Informática Educativa", {"color": MUTED, "size": 14})],
    [("Unidad Académica de Docencia Superior · UAZ · Noviembre de 2026", {"color": MUTED, "size": 14})]], size=17, spacing=1.3)
rect(sl, W - Cm(.45), Cm(15.5), Cm(.45), Cm(3.55), ORANGE)

# 2 Problema ---------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Planteamiento del problema", n)
for i, (big, small, src) in enumerate([("37.3 %", "de la población escolar en México presenta sobrepeso u obesidad", "Shamah-Levy et al., 2023"),
                                       ("45.9 %", "de escolares de una primaria de Guadalupe, Zac., con sobrepeso u obesidad", "López Hernández et al., 2024"),
                                       ("66.4 %", "de esos escolares con conocimientos insuficientes en nutrición", "López Hernández et al., 2024")]):
    x = Cm(1.3) + i * Cm(10.6)
    rect(sl, x, Cm(3.4), Cm(10), Cm(5.2), SOFT); rect(sl, x, Cm(3.4), Cm(10), Cm(.2), ORANGE)
    text(sl, x + Cm(.4), Cm(3.8), Cm(9.2), Cm(4.6), [[(big, {"bold": True, "size": 40, "color": RED})], [(small, {"size": 14})],
                                                       [(src, {"size": 11, "color": MUTED, "italic": True})]], align=PP_ALIGN.CENTER, spacing=1.15)
text(sl, Cm(1.3), Cm(9.3), Cm(31), Cm(7.5), [
    [("Problema: ", {"bold": True, "color": NAVY}), ("insuficiente conocimiento nutricional de estudiantes de primaria y falta de propuestas educativas "
      "sistemáticas, atractivas y apoyadas en recursos digitales.", {})],
    [("Vacío: ", {"bold": True, "color": NAVY}), ("no se ha evaluado el microaprendizaje aplicado a la educación nutricional en primarias de Zacatecas.", {})],
    [("Pregunta general: ", {"bold": True, "color": NAVY}), ("¿Se incrementa el nivel de conocimiento nutricional de los estudiantes de 5.º grado del Instituto "
      "Educativo Ammadeus después de su participación en un taller educativo en línea basado en microaprendizaje?", {})]],
    size=16, spacing=1.25)

# 3 Congruencia -------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Tabla de congruencia (actualizada)", n)
table(sl, Cm(1.3), Cm(3.2), Cm(31.2), [
    ["Elemento", "Contenido"],
    ["Objetivo general", "Incrementar el conocimiento nutricional de los estudiantes de 5.º grado del Instituto Educativo Ammadeus mediante la implementación de un taller educativo en línea basado en microaprendizaje."],
    ["Pregunta específica 2", "¿Qué elementos digitales se pueden incluir en un taller educativo en línea basado en microaprendizaje para abordar contenidos de educación nutricional en estudiantes de 5.º grado?"],
    ["Objetivos específicos", "1) Identificar el nivel inicial (preprueba). 2) Diseñar y desarrollar el taller con el modelo ADDIE. 3) Comparar el conocimiento, total y por dimensión, antes y después."],
    ["Hipótesis (H1)", "El puntaje de conocimiento nutricional será significativamente mayor en la posprueba que en la preprueba."],
    ["Variable independiente", "Taller educativo en línea basado en microaprendizaje (8 cápsulas, 4 semanas)."],
    ["Variable dependiente", "Nivel de conocimiento nutricional (puntaje 0–24)."],
    ["Instrumento", "Cuestionario de conocimiento nutricional: 24 reactivos de opción múltiple, 4 dimensiones."],
], [Cm(6.6), Cm(24.6)], size=12.5, rowh=Cm(1.55), first_col_fill=NAVY)

# 4 Marco teórico -----------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Fundamento teórico", n)
box(sl, Cm(1.3), Cm(3.3), Cm(15), Cm(10.8), "Microaprendizaje (variable independiente)", [
    "• Unidades breves con un objetivo cada una (Betancur-Chicué y García-Valcárcel Muñoz-Repiso, 2023).",
    "• Teoría de la carga cognitiva: segmentar reduce la carga extrínseca (Sweller et al., 2019).",
    "• Aprendizaje multimedia: coherencia, señalización y segmentación (Mayer, 2009).",
    "• Aprendizaje espaciado y recuperación activa (Kelley y Whatson, 2013).",
    "• Evidencia en educación básica (Silva et al., 2025)."], bar=TEAL, size=13.5)
box(sl, Cm(17.5), Cm(3.3), Cm(15), Cm(10.8), "Conocimiento nutricional (variable dependiente)", [
    "• Educación alimentaria y nutricional en la escuela (FAO, 2020; Pérez-Rodrigo y Aranceta, 2001).",
    "• Plato del Bien Comer Saludable y Sostenible: 5 grupos de alimentos.",
    "• Guías Alimentarias Saludables y Sostenibles para la Población Mexicana, 2.ª ed. (SSA et al., 2025).",
    "• Antecedentes: SaludArte (Shamah-Levy et al., 2017); taller en línea en Veracruz (Cruz Lara y Macossay Moreno, 2022); Plato del Bien Comer en 3.º y 4.º (Aguirre Vargas, 2019)."], bar=GREEN, size=13.5)
rect(sl, Cm(8), Cm(14.6), Cm(17.8), Cm(1.6), ORANGE, MSO_SHAPE.ROUNDED_RECTANGLE)
text(sl, Cm(8), Cm(14.6), Cm(17.8), Cm(1.6), "Cápsulas breves y visuales  →  mayor conocimiento nutricional", size=16, color=WHITE, bold=True,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 5 Metodología -------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Diseño metodológico", n)
cw = Cm(7.5)
items = [("Tipo y diseño", ["Investigación aplicada", "Enfoque cuantitativo", "Cuasiexperimental, preprueba y posprueba, un solo grupo intacto (O₁ X O₂)"]),
         ("Participantes", ["25 estudiantes de 5.º «B»", "Instituto Educativo Ammadeus, Guadalupe, Zac.", "Muestreo por conveniencia", "Consentimiento y asentimiento"]),
         ("Instrumento", ["Cuestionario de 24 reactivos", "4 dimensiones, puntaje 0–24", "Validez: juicio de expertos (V de Aiken)", "Confiabilidad: piloto con KR-20"]),
         ("Análisis", ["Estadística descriptiva", "Shapiro-Wilk", "t para muestras relacionadas o Wilcoxon (α = .05)", "Tamaño del efecto"])]
for i, (t, ls) in enumerate(items):
    x = Cm(1.3) + i * (cw + Cm(.4))
    rect(sl, x, Cm(3.3), cw, Cm(1.3), NAVY if i % 2 == 0 else RED)
    text(sl, x, Cm(3.3), cw, Cm(1.3), t, size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(sl, x, Cm(4.6), cw, Cm(8.4), GREY)
    text(sl, x + Cm(.3), Cm(4.9), cw - Cm(.6), Cm(8), [f"• {l}" for l in ls], size=14, spacing=1.3)
table(sl, Cm(1.3), Cm(13.5), Cm(31.2), [["D1. Principios de alimentación saludable", "D2. Grupos de alimentos", "D3. Aportes nutricionales", "D4. Aplicación en situaciones cotidianas"],
                                        ["Reactivos 1–6", "Reactivos 7–12", "Reactivos 13–18", "Reactivos 19–24"]],
      [Cm(7.8)] * 4, size=12.5, rowh=Cm(1.0), aligns=[PP_ALIGN.CENTER] * 4)

# 6 ADDIE -------------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Diseño instruccional del taller: modelo ADDIE", n)
fases = [("A", "Análisis", "Necesidades, perfil del grupo y recursos del aula digital"),
         ("D", "Diseño", "Objetivo, 8 cápsulas y carta descriptiva"),
         ("D", "Desarrollo", "Videos, infografías y juegos en Canva, Genially y Flippity"),
         ("I", "Implementación", "4 semanas, 2 cápsulas por semana en Google Classroom"),
         ("E", "Evaluación", "Diagnóstica, formativa (Kahoot) y sumativa (posprueba)")]
for i, (l, t, d) in enumerate(fases):
    x = Cm(1.3) + i * Cm(6.3)
    rect(sl, x, Cm(3.3), Cm(5.9), Cm(5.2), NAVY)
    text(sl, x, Cm(3.4), Cm(5.9), Cm(5), [[(l, {"size": 34, "bold": True, "color": ORANGE})], [(t, {"size": 15, "bold": True, "color": WHITE})],
                                          [(d, {"size": 11.5, "color": WHITE})]], align=PP_ALIGN.CENTER, spacing=1.1)
table(sl, Cm(1.3), Cm(9.1), Cm(31.2), [
    ["Semana", "Cápsulas", "Dimensión"],
    ["1", "C1 ¿Qué es comer saludable? · C2 Agua simple y sellos de advertencia", "D1"],
    ["2", "C3 Mi Plato del Bien Comer · C4 ¿A qué grupo pertenece?", "D2"],
    ["3", "C5 Vitaminas, minerales y fibra · C6 Proteína vegetal y grasas saludables", "D3"],
    ["4", "C7 Armo mi refrigerio escolar · C8 Reto final: mi comida del día", "D4"]],
    [Cm(3), Cm(24.2), Cm(4)], size=13, rowh=Cm(1.05), aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
text(sl, Cm(1.3), Cm(14.6), Cm(31), Cm(1.4), "Cada cápsula: recuperación de la anterior → microcontenido de 3 a 5 min → actividad interactiva con retroalimentación → reto de cierre.",
     size=13, color=MUTED)

# 7 Estado del proyecto -------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Estado del proyecto: ¿dónde estoy?", n)
steps = [("Protocolo revisado y corregido", "Planteamiento, marco teórico y metodología; referencias verificadas en APA 7.", GREEN, "Concluido"),
         ("Instrumento diseñado", "24 reactivos alineados con las Guías Alimentarias 2025 y el Plato del Bien Comer.", GREEN, "Concluido"),
         ("Diseño instruccional del taller", "Modelo ADDIE y carta descriptiva de 8 cápsulas.", GREEN, "Concluido"),
         ("Cartel aceptado en congreso internacional", "I Congreso Internacional de Tecnología Educativa y Ambientes de Aprendizaje, UAZ, 19 y 20 de noviembre.", GREEN, "Concluido"),
         ("Validación de contenido por juicio de expertos", "2 valoraciones recibidas (psicología clínica e investigación); en espera de docentes y nutriólogos.", ORANGE, "En proceso"),
         ("Prueba piloto, implementación y análisis", "Diciembre de 2026 a abril de 2027 (ver cronograma).", MUTED, "Pendiente")]
for i, (t, d, col, est) in enumerate(steps):
    y = Cm(3.25) + i * Cm(2.3)
    rect(sl, Cm(1.3), y + Cm(.45), Cm(1.1), Cm(1.1), col, MSO_SHAPE.OVAL)
    if i < len(steps) - 1: rect(sl, Cm(1.75), y + Cm(1.55), Cm(.2), Cm(1.2), RGBColor(0xCC, 0xCC, 0xCC))
    text(sl, Cm(3), y + Cm(.15), Cm(23), Cm(2), [[(t, {"bold": True, "size": 16, "color": NAVY})], [(d, {"size": 13, "color": MUTED})]], spacing=1.1)
    rect(sl, Cm(27.3), y + Cm(.4), Cm(5.2), Cm(1.2), col, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(sl, Cm(27.3), y + Cm(.4), Cm(5.2), Cm(1.2), est, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 8 Juicio de expertos: procedimiento ------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Validación de contenido: juicio de expertos", n)
box(sl, Cm(1.3), Cm(3.3), Cm(15.2), Cm(6.6), "Procedimiento", [
    "• Formato de valoración por reactivo (24 reactivos).",
    "• Jueces: psicología, nutrición, educación primaria y metodología.",
    "• Escala de 1 a 4 por criterio.",
    "• Observaciones y sugerencias de redacción por reactivo y globales."], size=14)
box(sl, Cm(17.3), Cm(3.3), Cm(15.2), Cm(6.6), "Criterios", [
    "• Claridad: lenguaje adecuado para 10 y 11 años.",
    "• Coherencia: relación con la dimensión.",
    "• Pertinencia: necesario para el propósito.",
    "• Exactitud nutricional: solo jueces nutriólogos."], bar=TEAL, size=14)
box(sl, Cm(1.3), Cm(10.4), Cm(31.2), Cm(5.8), "Criterio de decisión", [
    "• Se calculará la V de Aiken por reactivo y criterio, con intervalo de confianza al 95 % (Aiken, 1985; Penfield y Giacobbi, 2004).",
    "• Se conservan los reactivos con V ≥ .70 y límite inferior del IC > .50; los demás se reformulan o se eliminan.",
    "• Después: prueba piloto con un grupo no participante para estimar la confiabilidad (KR-20 ≥ .70) y analizar dificultad y discriminación."],
    bar=NAVY, size=14)

# 9 Avance del juicio -----------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Avance del juicio de expertos", n)
table(sl, Cm(1.3), Cm(3.3), Cm(14.2), [["Juez", "Perfil", "Estado"],
                                        ["Juez 1", "Psicóloga clínica", "Recibida"],
                                        ["Juez 2", "Experta en investigación (docente de posgrado)", "Recibida"],
                                        ["Pendientes", "Docentes de educación primaria", "En espera"],
                                        ["Pendientes", "Nutriólogos", "En espera"]],
      [Cm(2.9), Cm(7.3), Cm(4)], size=12.5, rowh=Cm(1.1))
box(sl, Cm(1.3), Cm(9.6), Cm(14.2), Cm(6.5), "Juez 1 · Psicóloga clínica", [
    "• Lenguaje claro y adecuado para 5.º grado; instrucciones fáciles.",
    "• Secuencia lógica y extensión breve.",
    "• Útil como referente para prevención e intervención.",
    [("Valoración global: adecuado, sin modificaciones.", {"bold": True, "color": GREEN})]], size=13)
box(sl, Cm(16.3), Cm(3.3), Cm(16.2), Cm(12.8), "Juez 2 · Experta en investigación", [
    "• Coherencia y pertinencia: 4 de 4 en los 24 reactivos.",
    "• Claridad: 4 en 20 reactivos; 3 en los reactivos 16 y 23; 1 en los reactivos 11 y 24.",
    "• Recomienda redactar los enunciados en forma aseverativa o de frase incompleta (DeVellis, 2017).",
    "• Un reactivo debe contener una sola idea (el 11 clasifica dos alimentos a la vez).",
    "• Dimensiones bien representadas y lenguaje adecuado para 5.º grado.",
    [("Valoración global: adecuado, con modificaciones menores.", {"bold": True, "color": ORANGE})],
    [("Exactitud nutricional: se toma solo de los jueces nutriólogos.", {"italic": True, "color": MUTED, "size": 11.5})]],
    bar=TEAL, size=13.5)

# 10 Ajustes al instrumento ------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Ajustes al instrumento a partir de la retroalimentación", n)
box(sl, Cm(1.3), Cm(3.3), Cm(15.2), Cm(8.6), "Sugerencias del Juez 2", [
    "• Enunciados en forma aseverativa, por ejemplo: «La opción que describe mejor una alimentación saludable es:».",
    "• Reactivo 11: separar la clasificación del pescado y la del aguacate.",
    "• Reactivo 24: precisar qué significa «aplicar diferentes recomendaciones al mismo tiempo».",
    "• Reactivos 16 y 23: ajustes menores de redacción.",
    "• Elegir el formato (aseverativo, interrogativo o solicitud directa) pensando en niños de 10 y 11 años."], size=13.5)
box(sl, Cm(17.3), Cm(3.3), Cm(15.2), Cm(8.6), "Ajustes de la revisión propia", [
    "• Unificar «agua simple» en todos los reactivos.",
    "• Evitar reactivos que miden lo mismo (6 y 1; 15 y 18; 20 y 22).",
    "• Fortalecer distractores débiles.",
    "• Simplificar términos como «respectivamente».",
    [("El director revisó el instrumento y lo consideró pertinente; para evitar sesgo, no participa como juez.", {"italic": True, "color": MUTED, "size": 11.5})]],
    bar=TEAL, size=13.5)
text(sl, Cm(1.3), Cm(12.8), Cm(31.2), Cm(1.6), [[("Siguiente paso: ", {"bold": True, "color": NAVY}),
     ("sigo a la espera de las valoraciones de docentes y nutriólogos; al recibirlas calcularé la V de Aiken y aplicaré los ajustes en una versión final antes del piloto.", {})]], size=13.5)

# 11 Cronograma ----------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Cronograma de actividades", n)
months = ["Nov 2026", "Dic 2026", "Ene 2027", "Feb 2027", "Mar 2027", "Abr 2027"]
acts = [("Juicio de expertos y ajuste del instrumento", 0, 1, ORANGE),
        ("Cartel en el I Congreso Internacional (19–20 nov)", 0, 0, ORANGE),
        ("Prueba piloto y confiabilidad (KR-20)", 1, 1, NAVY),
        ("Desarrollo de las 8 cápsulas", 0, 2, NAVY),
        ("Consentimientos y preprueba", 2, 2, NAVY),
        ("Implementación del taller (4 semanas)", 3, 3, RED),
        ("Posprueba", 3, 3, RED),
        ("Análisis de datos y redacción de resultados", 4, 5, GREEN)]
x0, lw, mw, y0, rh = Cm(1.3), Cm(12.2), Cm(3.15), Cm(3.3), Cm(1.38)
rect(sl, x0, y0, lw, Cm(1.1), NAVY)
text(sl, x0, y0, lw, Cm(1.1), "Actividad", size=13, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
for j, m in enumerate(months):
    rect(sl, x0 + lw + j * mw, y0, mw - Cm(.05), Cm(1.1), NAVY)
    text(sl, x0 + lw + j * mw, y0, mw, Cm(1.1), m, size=12.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, (a, s, e, col) in enumerate(acts):
    y = y0 + Cm(1.2) + i * rh
    rect(sl, x0, y, lw + mw * 6, rh - Cm(.08), GREY if i % 2 == 0 else WHITE)
    text(sl, x0 + Cm(.2), y, lw - Cm(.3), rh, a, size=12.5, anchor=MSO_ANCHOR.MIDDLE)
    rect(sl, x0 + lw + s * mw + Cm(.15), y + Cm(.3), (e - s + 1) * mw - Cm(.3), rh - Cm(.68), col, MSO_SHAPE.ROUNDED_RECTANGLE)
text(sl, Cm(1.3), Cm(16.6), Cm(31), Cm(.9), [[("Implementación en febrero de 2027: ", {"bold": True, "color": RED}),
     ("preprueba a finales de enero, taller durante febrero y posprueba al cierre. Fechas tentativas, sujetas al calendario escolar.", {})]], size=12.5)

# 12 Implicaciones -------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Implicaciones y siguientes pasos", n)
box(sl, Cm(1.3), Cm(3.3), Cm(15.2), Cm(12.8), "Aportes esperados", [
    "• Un taller replicable y de bajo costo para 5.º grado, alineado con las Guías Alimentarias 2025 y la Fase 5 de la SEP.",
    "• Un instrumento con evidencias de validez y confiabilidad para medir conocimiento nutricional en escolares.",
    "• Evidencia sobre el microaprendizaje aplicado a la educación nutricional en Zacatecas."], size=14.5)
box(sl, Cm(17.3), Cm(3.3), Cm(15.2), Cm(12.8), "Implicaciones para futuras investigaciones", [
    "• Incluir un grupo de comparación para fortalecer la validez interna.",
    "• Agregar una medición de seguimiento para valorar si el conocimiento se mantiene.",
    "• Evaluar cambios en conductas alimentarias, no solo en conocimiento.",
    "• Replicar el taller en otros grados y en escuelas públicas."], bar=TEAL, size=14.5)

# 13 Referencias ----------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1; header(sl, "Referencias", n)
refs = [
 "Aiken, L. R. (1985). Three coefficients for analyzing the reliability and validity of ratings. Educational and Psychological Measurement, 45(1), 131–142.",
 "Betancur-Chicué, V., y García-Valcárcel Muñoz-Repiso, A. (2023). Características del diseño de estrategias de microaprendizaje en escenarios educativos: Revisión sistemática. RIED, 26(1), 201–222. https://doi.org/10.5944/ried.26.1.34056",
 "Branch, R. M. (2009). Instructional design: The ADDIE approach. Springer.",
 "Cruz Lara, N. M., y Macossay Moreno, C. (2022). Intervención de educación nutricional en línea a población infantil en el estado de Veracruz. Horizonte Sanitario, 21(3), 403–409.",
 "DeVellis, R. F. (2017). Scale development: Theory and applications (4.ª ed.). SAGE.",
 "López Hernández, P. Y., Franco Trejo, C. S., Galicia Rodríguez, A. G., Argüelles Estrada, B. C., y Ornelas Pargas, A. C. (2024). Sobrepeso, obesidad y educación nutricional en una escuela primaria de Zacatecas. Revista Iberoamericana de Ciencias, 11(1).",
 "Mayer, R. E. (2009). Multimedia learning (2.ª ed.). Cambridge University Press.",
 "Penfield, R. D., y Giacobbi, P. R., Jr. (2004). Applying a score confidence interval to Aiken's item content-relevance index. Measurement in Physical Education and Exercise Science, 8(4), 213–225.",
 "Pérez-Rodrigo, C., y Aranceta, J. (2001). School-based nutrition education: Lessons learned and new perspectives. Public Health Nutrition, 4(1A), 131–139.",
 "Secretaría de Salud, Instituto Nacional de Salud Pública y UNICEF. (2025). Guías alimentarias saludables y sostenibles para la población mexicana (2.ª ed.).",
 "Shamah-Levy, T., Gaona-Pineda, E. B., Cuevas-Nasu, L., Morales-Ruan, C., Valenzuela-Bravo, D. G., Méndez-Gómez Humarán, I., y Ávila-Arcos, M. A. (2023). Prevalencias de sobrepeso y obesidad en población escolar y adolescente de México. Ensanut Continua 2020-2022. Salud Pública de México, 65(Supl. 1), S218–S224.",
 "Sweller, J., van Merriënboer, J. J. G., y Paas, F. (2019). Cognitive architecture and instructional design: 20 years later. Educational Psychology Review, 31(2), 261–292.",
]
text(sl, Cm(1.3), Cm(3.1), Cm(31.2), Cm(14.5), refs, size=10.5, spacing=1.05)

# 13 Gracias ------------------------------------------------------------------
sl = prs.slides.add_slide(BL); n += 1
rect(sl, 0, 0, W, H, NAVY); rect(sl, 0, Cm(12.2), W, Cm(.2), ORANGE)
text(sl, Cm(2), Cm(5.5), W - Cm(4), Cm(6), [[("¡Muchas gracias por su atención!", {"size": 36, "bold": True, "color": WHITE})],
     [("L.N.C.A. Anakaren Ernestina Reimers Reyes", {"size": 20, "bold": True, "color": WHITE})],
     [("Licenciada en Nutrición y Ciencia de los Alimentos · Maestría en Tecnología Informática Educativa, UAZ", {"size": 14, "color": RGBColor(0xDD, 0xE3, 0xF0)})],
     [("ana.reimers@gmail.com", {"size": 14, "color": RGBColor(0xDD, 0xE3, 0xF0)})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.4)

from guion import GUION
for sl_, (_, _, _, txt) in zip(prs.slides, GUION):
    sl_.notes_slide.notes_text_frame.text = txt

prs.core_properties.title = "Actividad 4 · Presentación de avances · Reimers Reyes"
prs.save("Actividad4_Presentacion_avances_Reimers.pptx")
print("OK", n)
