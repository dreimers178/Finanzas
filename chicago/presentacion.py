"""Presentación resumida del proyecto de desayunos Chicago (16:9, editable en PowerPoint)."""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

RED = RGBColor(0x8A, 0x1C, 0x16)
GREEN = RGBColor(0x0F, 0x3A, 0x2E)
GREEN2 = RGBColor(0x1D, 0x5A, 0x47)
CREAM = RGBColor(0xF7, 0xF1, 0xE6)
INK = RGBColor(0x1E, 0x1E, 0x1E)
MUTED = RGBColor(0x5B, 0x5B, 0x5B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xFF, 0xD9, 0xA8)
HEAD = "Arial Black"
BODY = "Arial"
LOGO = "logos/logo_chicago_color.png"

prs = Presentation()
prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)
W, H = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def rect(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background(); s.shadow.inherit = False
    return s


def text(sl, x, y, w, h, runs, size=18, color=INK, font=BODY, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.05):
    """runs: str o lista de párrafos; cada párrafo es str o lista de (texto, dict)."""
    tb = sl.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Cm(0.1); tf.margin_top = tf.margin_bottom = Cm(0.05)
    paras = runs if isinstance(runs, list) else [runs]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing
        for t, o in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run(); r.text = t
            f = r.font; f.name = o.get("font", font); f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold); f.color.rgb = o.get("color", color)
    return tb


def header(sl, title, kicker=None):
    rect(sl, 0, 0, W, Cm(2.6), GREEN)
    rect(sl, 0, Cm(2.6), W, Cm(0.18), RED)
    text(sl, Cm(1.4), Cm(0.55), Cm(24), Cm(1.6), title, size=28, color=WHITE, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
    pic = sl.shapes.add_picture(LOGO, 0, 0, height=Cm(1.9))
    pic.left = W - pic.width - Cm(1.0); pic.top = Cm(0.35)
    # fondo crema detrás del logo para legibilidad
    bg = rect(sl, pic.left - Cm(0.25), Cm(0.2), pic.width + Cm(0.5), Cm(2.2), CREAM, MSO_SHAPE.ROUNDED_RECTANGLE)
    sl.shapes._spTree.remove(bg._element); sl.shapes._spTree.insert(sl.shapes._spTree.index(pic._element), bg._element)
    if kicker:
        text(sl, Cm(1.4), Cm(3.1), Cm(31), Cm(1), kicker, size=15, color=MUTED)


def footer(sl, n):
    text(sl, Cm(1.4), H - Cm(1.0), Cm(25), Cm(0.7), "Chicago Deep Dish Pizza · Proyecto de desayunos · Plaza Auskara",
         size=10, color=MUTED)
    text(sl, W - Cm(3.4), H - Cm(1.0), Cm(2.4), Cm(0.7), str(n), size=10, color=MUTED, align=PP_ALIGN.RIGHT)


def card(sl, x, y, w, h, big, small, fill=CREAM, bigcolor=GREEN, bar=RED, bigsize=30):
    rect(sl, x, y, w, h, fill)
    rect(sl, x, y, Cm(0.25), h, bar)
    text(sl, x + Cm(0.6), y + Cm(0.25), w - Cm(0.8), Cm(1.6), big, size=bigsize, color=bigcolor, font=HEAD)
    text(sl, x + Cm(0.6), y + Cm(1.85), w - Cm(0.8), h - Cm(2.0), small, size=13, color=MUTED)


def table(sl, x, y, w, rows, colw, size=14, rowh=Cm(0.9), head=True, bold_last=False, right_cols=(1,)):
    shp = sl.shapes.add_table(len(rows), len(rows[0]), x, y, w, rowh * len(rows))
    t = shp.table
    for j, cw in enumerate(colw):
        t.columns[j].width = cw
    for i, row in enumerate(rows):
        t.rows[i].height = rowh
        for j, val in enumerate(row):
            c = t.cell(i, j); c.text = ""
            c.margin_left = c.margin_right = Cm(0.25); c.margin_top = c.margin_bottom = Cm(0.05)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = str(val)
            p.alignment = PP_ALIGN.RIGHT if j in right_cols else PP_ALIGN.LEFT
            f = r.font; f.name = BODY; f.size = Pt(size)
            c.fill.solid()
            if head and i == 0:
                c.fill.fore_color.rgb = GREEN; f.color.rgb = WHITE; f.bold = True
            elif bold_last and i == len(rows) - 1:
                c.fill.fore_color.rgb = CREAM; f.bold = True; f.color.rgb = GREEN
            else:
                c.fill.fore_color.rgb = WHITE if i % 2 else RGBColor(0xFA, 0xF7, 0xF1); f.color.rgb = INK
    return shp


n = 0
# 1 · Portada -------------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
rect(sl, 0, 0, W, H, CREAM)
rect(sl, 0, Cm(13.2), W, H - Cm(13.2), GREEN)
rect(sl, 0, Cm(13.0), W, Cm(0.2), RED)
pic = sl.shapes.add_picture(LOGO, 0, Cm(1.3), height=Cm(6.2)); pic.left = int((W - pic.width) / 2)
text(sl, Cm(2), Cm(8.0), W - Cm(4), Cm(2), "PROYECTO DE DESAYUNOS", size=40, color=GREEN, font=HEAD, align=PP_ALIGN.CENTER)
text(sl, Cm(2), Cm(10.2), W - Cm(4), Cm(1.2), "American Breakfast, Mexican Flavor", size=22, color=RED, bold=True, align=PP_ALIGN.CENTER)
text(sl, Cm(2), Cm(14.0), W - Cm(4), Cm(4), [
    [("Piloto de 30 días · Sucursal Plaza Auskara, Zacatecas", {"bold": True, "color": GOLD, "size": 20})],
    [("Propuesta para aprobación de socios · Noviembre de 2026", {"color": WHITE, "size": 16})]],
    align=PP_ALIGN.CENTER, spacing=1.3)

# 2 · La propuesta ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "La propuesta")
text(sl, Cm(1.4), Cm(3.3), Cm(31), Cm(2.4), [
    [("Abrir desayunos en Chicago Plaza Auskara ", {"bold": True}),
     ("con un menú americano adaptado al gusto mexicano, aprovechando el local, la marca y la cocina que ya tenemos.", {})]],
    size=19)
y = Cm(6.4); w = Cm(7.3); gap = Cm(0.55)
for i, (big, small) in enumerate([("Jue–Dom", "4 días por semana"), ("8:00–12:00", "Horario de desayuno"),
                                   ("3 personas", "Cocinero, auxiliar y servicio/caja"), ("$170", "Ticket promedio por cliente")]):
    card(sl, Cm(1.4) + i * (w + gap), y, w, Cm(3.6), big, small, bigsize=28)
text(sl, Cm(1.4), Cm(11.0), Cm(31), Cm(5.5), [
    [("¿Por qué ahora?", {"bold": True, "color": RED, "size": 20})],
    [("•  Hoy el local no vende en la mañana: son ventas nuevas sin cambiar la operación de pizza.", {})],
    [("•  La inversión es baja ($19,000) porque ya contamos con la cocina y el equipo principal.", {})],
    [("•  Es un piloto de 30 días con criterios claros para decidir si continúa.", {})]],
    size=18, spacing=1.35)
footer(sl, n)

# 3 · Menú ------------------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "Menú inicial y precios sugeridos")
table(sl, Cm(1.4), Cm(3.4), Cm(17.5), [["Platillo", "Precio"],
      ["Chicago American Breakfast", "$169"], ["Chicago Skillet", "$169"], ["Chilaquiles con arrachera o pollo", "$179"],
      ["Huevos con machaca", "$159"], ["Omelette Denver", "$149"], ["Chilaquiles con 2 huevos", "$139"],
      ["Burrito de desayuno Chicago", "$129"], ["Hotcakes Chicago", "$119"]], [Cm(13.5), Cm(4.0)], size=15, rowh=Cm(1.15))
table(sl, Cm(20.2), Cm(3.4), Cm(12.3), [["Bebida", "Sugerido"], ["Café americano", "$35"], ["Jugo de naranja natural", "$45"],
      ["Chocolate caliente", "$45"], ["Limonada o naranjada", "$40"], ["Refrescos", "$35"]],
      [Cm(8.8), Cm(3.5)], size=15, rowh=Cm(1.15))
rect(sl, Cm(20.2), Cm(11.0), Cm(12.3), Cm(3.8), GREEN)
text(sl, Cm(20.7), Cm(11.25), Cm(11.4), Cm(3.4), [
    [("Platillo promedio: $152", {"bold": True, "color": GOLD, "size": 18})],
    [("Con bebida el consumo típico supera $180; el ticket de $170 de las proyecciones es conservador.", {"color": WHITE, "size": 14})]],
    spacing=1.2)
text(sl, Cm(1.4), Cm(14.2), Cm(17.5), Cm(1.5), "Precios a validar con costeo por receta (meta: costo de alimentos ≤ 35 %).",
     size=12, color=MUTED)
footer(sl, n)

# 4 · Equipo e inversión ---------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "Equipo e inversión: $19,000")
text(sl, Cm(1.4), Cm(3.3), Cm(10), Cm(1), "Ya lo tenemos", size=20, color=RED, font=HEAD)
items = [("Freidora industrial", "Papas, totopos para chilaquiles"), ("Licuadora industrial", "Salsas, chocolate, limonada"),
         ("Microondas", "Recalentar frijoles y salsas"), ("Horno, plancha, estufa", "Skillet, huevos, hotcakes (confirmar)")]
for i, (a, b) in enumerate(items):
    yy = Cm(4.6) + i * Cm(2.15)
    rect(sl, Cm(1.4), yy, Cm(10.2), Cm(1.85), CREAM)
    rect(sl, Cm(1.4), yy, Cm(0.25), Cm(1.85), GREEN2)
    text(sl, Cm(2.0), yy + Cm(0.12), Cm(9.4), Cm(1.7), [[(a, {"bold": True, "size": 15, "color": GREEN})], [(b, {"size": 12.5, "color": MUTED})]])
text(sl, Cm(12.6), Cm(3.3), Cm(10), Cm(1), "Lo que compramos", size=20, color=RED, font=HEAD)
table(sl, Cm(12.6), Cm(4.6), Cm(10.2), [["Equipo menor", "Monto"], ["Cafetera de goteo", "$6,000"], ["Sartenes", "$1,500"],
      ["Platos, bowls y charolas", "$2,000"], ["Exprimidor de cítricos", "$800"], ["Tostador", "$800"],
      ["Utensilios, báscula, otros", "$2,300"], ["Imprevistos", "$600"], ["Total equipo", "$14,000"]],
      [Cm(7.2), Cm(3.0)], size=13, rowh=Cm(0.95), bold_last=True)
text(sl, Cm(23.8), Cm(3.3), Cm(9), Cm(1), "Uso de fondos", size=20, color=RED, font=HEAD)
table(sl, Cm(23.8), Cm(4.6), Cm(8.7), [["Concepto", "Monto"], ["Equipo y utensilios", "$14,000"], ["Inventario inicial", "$2,500"],
      ["Publicidad", "$1,000"], ["Reserva operativa", "$1,500"], ["Inversión total", "$19,000"]],
      [Cm(5.7), Cm(3.0)], size=13, rowh=Cm(0.95), bold_last=True)
text(sl, Cm(23.8), Cm(10.6), Cm(8.7), Cm(4), "Al tener freidora, licuadora y microondas, la inversión se concentra en equipo menor.",
     size=13, color=MUTED)
footer(sl, n)

# 5 · Números clave ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "¿Cuántos clientes necesitamos?")
cw = Cm(9.9); gx = Cm(0.65)
card(sl, Cm(1.4), Cm(3.5), cw, Cm(4.2), "$19,250", "Gastos fijos al mes: nómina $13,000 + reserva laboral $3,250 + otros $3,000")
card(sl, Cm(1.4) + cw + gx, Cm(3.5), cw, Cm(4.2), "$110.50", "Quedan por cliente después de ingredientes (ticket $170 − 35 %)")
card(sl, Cm(1.4) + 2 * (cw + gx), Cm(3.5), cw, Cm(4.2), "10.05", "Clientes por día para no perder (punto de equilibrio)")
rect(sl, Cm(1.4), Cm(8.6), W - Cm(2.8), Cm(4.4), RED)
text(sl, Cm(2.2), Cm(8.9), W - Cm(4.4), Cm(4), [
    [("Meta mínima: 11 clientes por día", {"font": HEAD, "size": 30, "color": WHITE})],
    [("Equivale a unos 3 clientes por hora en un horario de 4 horas.", {"size": 18, "color": GOLD, "bold": True})]],
    align=PP_ALIGN.CENTER, spacing=1.25, anchor=MSO_ANCHOR.MIDDLE)
text(sl, Cm(1.4), Cm(13.6), W - Cm(2.8), Cm(3), [
    [("Supuestos: ", {"bold": True}), ("17.33 días de apertura al mes, 3 trabajadores a $250 por turno de 4 horas "
      "(por encima del mínimo proporcional 2026), costo de alimentos de 35 %.", {})]], size=14, color=MUTED)
footer(sl, n)

# 6 · Escenarios -------------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "Utilidad mensual según clientes por día")
cd = CategoryChartData()
cats = ["8", "10", "12", "15", "18", "20", "25"]
vals = [-3927, -97, 3734, 9480, 15226, 19057, 28633]
cd.categories = cats
cd.add_series("Utilidad mensual (MXN)", vals)
gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Cm(1.2), Cm(3.2), Cm(20.5), Cm(13.6), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.name = BODY; ch.font.size = Pt(13)
plot = ch.plots[0]; plot.gap_width = 55
plot.has_data_labels = True
dl = plot.data_labels; dl.number_format = '"$"#,##0;"−$"#,##0'; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(12); dl.font.bold = True
ser = plot.series[0]
for i, v in enumerate(vals):
    pt = ser.points[i]; pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = RED if v < 0 else (GREEN if cats[i] == "15" else GREEN2)
va = ch.value_axis; va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
va.tick_labels.number_format = '"$"#,##0'; va.tick_labels.number_format_is_linked = False; va.tick_labels.font.size = Pt(11)
ca = ch.category_axis; ca.has_title = True; ca.axis_title.text_frame.text = "Clientes por día"
ca.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
from pptx.enum.chart import XL_TICK_LABEL_POSITION
ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
table(sl, Cm(22.6), Cm(3.6), Cm(10), [["Clientes/día", "Recupera $19,000"], ["12", "5.1 meses"], ["15", "2.0 meses"],
      ["18", "1.2 meses"], ["20", "1.0 mes"], ["25", "0.7 meses"]], [Cm(4.6), Cm(5.4)], size=14, rowh=Cm(1.05))
rect(sl, Cm(22.6), Cm(10.4), Cm(10), Cm(3.9), GREEN)
text(sl, Cm(23.0), Cm(10.6), Cm(9.2), Cm(3.6), [
    [("Escenario base: 15 clientes/día", {"bold": True, "color": GOLD, "size": 16})],
    [("Ventas de $44,200 y utilidad de $9,480 al mes; la inversión se recupera en 2 meses.", {"color": WHITE, "size": 14})]],
    spacing=1.2)
text(sl, Cm(22.6), Cm(14.6), Cm(10), Cm(2), "Utilidad antes de impuestos, depreciación y gastos generales existentes.",
     size=11, color=MUTED)
footer(sl, n)

# 7 · Plan y decisiones -------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "Plan de 30 días y criterios de decisión")
steps = [("Días 1–3", "Costear recetas y confirmar horario de la plaza"), ("Días 4–7", "Comprar equipo y montar la estación"),
         ("Días 8–10", "Capacitar y hacer pruebas"), ("Días 11–15", "Lanzamiento y promoción"), ("Días 16–30", "Medir y ajustar")]
for i, (a, b) in enumerate(steps):
    x = Cm(1.4) + i * Cm(6.3)
    rect(sl, x, Cm(3.5), Cm(5.9), Cm(1.2), GREEN if i < 4 else RED)
    text(sl, x, Cm(3.5), Cm(5.9), Cm(1.2), a, size=16, color=WHITE, font=HEAD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(sl, x, Cm(4.7), Cm(5.9), Cm(2.4), CREAM)
    text(sl, x + Cm(0.25), Cm(4.85), Cm(5.4), Cm(2.2), b, size=14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(sl, Cm(1.4), Cm(7.8), Cm(20), Cm(1), "Al terminar el piloto, decidimos con el promedio de clientes por día:",
     size=17, color=RED, bold=True)
table(sl, Cm(1.4), Cm(8.9), Cm(31.1), [["Clientes por día", "Decisión"], ["Menos de 10", "Revisar oferta, promoción, horario y personal"],
      ["10 a 14", "Cubre costos: mejorar margen y volumen"], ["15 a 17", "Continuar y optimizar"],
      ["18 o más", "Evaluar ampliar horario o días y replicar en otras sucursales"]],
      [Cm(7), Cm(24.1)], size=15, rowh=Cm(1.1), right_cols=())
footer(sl, n)

# 8 · Acuerdo ----------------------------------------------------------------
sl = prs.slides.add_slide(BLANK); n += 1
header(sl, "Acuerdo que solicitamos")
rect(sl, Cm(1.4), Cm(3.5), W - Cm(2.8), Cm(4.6), GREEN)
text(sl, Cm(2.2), Cm(3.7), W - Cm(4.4), Cm(4.2), [
    [("Autorizar un piloto de 30 días en Plaza Auskara", {"font": HEAD, "size": 26, "color": WHITE})],
    [("Inversión máxima de $19,000 · Jueves a domingo, de 8:00 a 12:00 · Evaluación semanal de ventas, costos y satisfacción",
      {"size": 17, "color": GOLD, "bold": True})]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)
text(sl, Cm(1.4), Cm(8.8), Cm(31), Cm(1), "Antes de arrancar confirmamos:", size=18, color=RED, bold=True)
text(sl, Cm(1.4), Cm(9.9), Cm(15.3), Cm(6), [
    [("•  Que la plaza permita abrir a las 8:00 a. m.", {})],
    [("•  Contratación, IMSS y prestaciones con un especialista", {})],
    [("•  Licencia de funcionamiento para el nuevo horario", {})]], size=16, spacing=1.45)
text(sl, Cm(17.2), Cm(9.9), Cm(15.3), Cm(6), [
    [("•  Costeo de recetas: alimentos ≤ 35 %", {})],
    [("•  Que la cocina no afecte el servicio de pizza", {})],
    [("•  Higiene conforme a la NOM-251-SSA1-2009", {})]], size=16, spacing=1.45)
footer(sl, n)

prs.core_properties.title = "Proyecto de desayunos · Chicago Plaza Auskara"
prs.core_properties.author = ""
prs.save("Presentacion_Desayunos_Chicago.pptx")
print("OK", n, "diapositivas")
