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
bullet(" unos 15 minutos de exposición + 2 o 3 minutos de preguntas.", "Duración:")
bullet(" José Luis y Pablo se turnan diapositiva por diapositiva. Cada quien adelanta la diapositiva de la que va a hablar.", "Turnos:")
bullet(" no lean palabra por palabra. Apréndanse las ideas y los números clave; lo demás díganlo con sus propias palabras. Ensayen juntos al menos dos veces con cronómetro.", "Consejo:")
bullet(" los textos también vienen en las «Notas del orador» de cada diapositiva del PowerPoint.", "Notas:")
bullet(" «NZ$» = dólares neozelandeses; «US$» = dólares estadounidenses. Un dólar neozelandés vale menos que uno estadounidense.", "Monedas:")
para("")

# ---------------- GUION ----------------
G = [
    (1, "Portada", "0:20", [
        (JL, "Buenos días, maestra Samantha y compañeros. Nuestro proyecto se llama «Países que construyen bienestar: modelos para una mejor sociedad»."),
        (PB, "Somos José Luis De Anda y Pablo De Santiago, y el país que investigamos es Nueva Zelanda, que en idioma maorí se llama Aotearoa."),
    ]),
    (2, "¿Qué vamos a presentar?", "0:35", [
        (JL, "Nuestra presentación tiene cuatro partes. A: conocer el país. B, la parte principal: los modelos que usa Nueva Zelanda para mejorar la vida de su gente en educación, salud, empleo y pensiones, y vivienda y medio ambiente. C: analizar si eso podría funcionar en México. Y D: nuestras conclusiones. La pregunta que nos guió fue: ¿qué hace Nueva Zelanda para que su gente viva mejor y qué podría copiar México?"),
    ]),
    (3, "Nueva Zelanda: ¿dónde está?", "0:50", [
        (JL, "Nueva Zelanda está en Oceanía, en el océano Pacífico Sur, a unos 11 mil kilómetros de México. Su capital es Wellington y su ciudad más grande es Auckland. Tiene dos islas principales, la Isla Norte y la Isla Sur, y es unas siete veces más pequeña que México."),
        (JL, "En su bandera aparece la bandera británica, porque fue colonia del Reino Unido, y cuatro estrellas que forman la Cruz del Sur, una constelación que se ve en el hemisferio sur. Tiene tres idiomas oficiales: inglés, maorí —el idioma del pueblo originario— y lengua de señas. Su moneda es el dólar neozelandés."),
    ]),
    (4, "Población", "0:45", [
        (PB, "En junio de 2025 Nueva Zelanda tenía 5 millones 324 mil 700 habitantes. Hay casi la misma cantidad de mujeres que de hombres: 50.3 % mujeres y 49.7 % hombres. La edad mediana, es decir, la edad que divide a la población en dos mitades, es de 39 años para mujeres y 37 para hombres."),
        (PB, "Para darnos una idea, México tiene más de 130 millones de habitantes, unas 25 veces más. Esto es importante: organizar programas para 5 millones de personas no es lo mismo que para 130 millones."),
    ]),
    (5, "Esperanza de vida y edades", "0:50", [
        (PB, "La esperanza de vida es cuántos años vive en promedio una persona. En Nueva Zelanda es de 81.8 años: 83.5 para las mujeres y 80.1 para los hombres. Pero hay una diferencia importante: el pueblo maorí vive en promedio 75.8 años, seis menos. Esto nos muestra que incluso en países con mucho bienestar hay desigualdad."),
        (PB, "En la gráfica vemos las edades: 18.7 % son niños, casi 65 % está en edad de trabajar y 16.6 % son adultos mayores. Como cada vez hay más adultos mayores, el país necesita un buen sistema de pensiones, del que hablaremos más adelante."),
    ]),
    (6, "¿De qué vive su economía?", "0:50", [
        (JL, "La economía de Nueva Zelanda se basa mucho en el campo. Su exportación número uno son los lácteos: leche en polvo, mantequilla y queso. En 2025 vendió al mundo 80.7 mil millones de dólares neozelandeses, y casi 30 % fueron lácteos. También exporta carne, madera y frutas como el kiwi, y recibe muchos turistas."),
        (JL, "El PIB, que es el valor de todo lo que produce un país en un año, fue de unos 259 mil millones de dólares estadounidenses. Si lo dividimos entre sus habitantes, da unos 48,600 dólares por persona, uno de los más altos del mundo. Por eso se considera un país desarrollado."),
    ]),
    (7, "Otros indicadores: NZ vs. México", "0:55", [
        (PB, "Además del PIB, revisamos otros indicadores. En el Índice de Desarrollo Humano de la ONU, que combina salud, educación e ingreso, Nueva Zelanda está en el lugar 17 y México en el 81. En percepción de corrupción, Nueva Zelanda tiene 81 de 100 puntos y es el cuarto país más limpio; México tiene 27."),
        (PB, "En impuestos, Nueva Zelanda junta casi 33 % de su PIB y México 18.3 %, el más bajo de la OCDE; con ese dinero los gobiernos pagan escuelas, hospitales y pensiones. Y lo que más nos sorprendió: en el Informe Mundial de la Felicidad 2026, Nueva Zelanda está en el lugar 11 y México en el 12. ¡Casi empatados! O sea, la felicidad no depende solo del dinero; también de la familia y las relaciones."),
    ]),
    (8, "La gran idea: medir el bienestar", "0:45", [
        (JL, "Ahora viene la parte principal. Antes de ver cada tema, hay una gran idea que hace diferente a Nueva Zelanda. Normalmente los países miden si les va bien con el PIB, o sea, cuánto creció la economía. En 2019, Nueva Zelanda fue el primer país en hacer un «Presupuesto de Bienestar»: antes de decidir en qué gastar, se preguntó si la gente vivía mejor. Por eso puso dinero en salud mental, pobreza infantil y apoyo al pueblo maorí."),
        (JL, "Desde 2020 una ley pide al gobierno informar sus metas de bienestar. Pero en 2023 llegó otro gobierno con otras prioridades, así que estos modelos también dependen de quién gobierna."),
    ]),
    (9, "Tema 1: Educación", "1:00", [
        (PB, "Nuestro primer tema es educación. Nueva Zelanda tiene un programa de almuerzos escolares saludables desde 2020. Consiste en dar comida nutritiva gratis todos los días de clases. El problema que busca resolver es que un niño con hambre no puede concentrarse ni aprender."),
        (PB, "El gobierno lo paga en las escuelas con más necesidad, y en 2026 ya llegaba a más de 240 mil alumnos en más de mil escuelas. Los estudios reportan menos hambre y mejor bienestar. Además, la escuela pública es gratuita de los 5 a los 19 años y los niños de 3 a 5 años tienen 20 horas gratis de preescolar a la semana."),
        (PB, "En la gráfica de la prueba PISA de la OCDE, que evalúa a jóvenes de 15 años, Nueva Zelanda queda arriba del promedio en las tres materias; México queda abajo."),
    ]),
    (10, "Tema 2: Salud", "1:05", [
        (JL, "Nuestro segundo tema es salud, y aquí encontramos algo único en el mundo: la ACC, que desde 1974 es un seguro de accidentes para todos. Si alguien se rompe una pierna jugando futbol o tiene un accidente en el trabajo o en la calle, la ACC paga su tratamiento y parte de su sueldo mientras se recupera, sin importar de quién fue la culpa."),
        (JL, "Se paga con cuotas que salen de los salarios, las empresas y los autos, más dinero del gobierno, y cubre incluso a los turistas. En un solo año aceptó más de 2.1 millones de casos. Lo diferente es que, a cambio, no se puede demandar a nadie por lesiones."),
        (JL, "Además, Pharmac compra las medicinas para todo el país y, como compra mucho, consigue precios bajos. Los hospitales públicos son gratuitos y los niños menores de 14 años van al médico sin costo en la mayoría de las clínicas."),
    ]),
    (11, "Tema 3: Empleo y pensiones", "1:00", [
        (PB, "Nuestro tercer tema es empleo y pensiones. Nueva Zelanda tiene dos pilares. El primero es NZ Super: una pensión para todas las personas de 65 años o más que viven ahí, hayan trabajado o no; una persona sola recibe unos 555 dólares neozelandeses a la semana."),
        (PB, "El segundo pilar es KiwiSaver, creado en 2007. Cuando alguien empieza un nuevo trabajo, lo inscriben automáticamente; si no quiere, puede salirse, pero si no hace nada, ¡ahorra! El trabajador pone 3.5 % de su sueldo, la empresa otro 3.5 % y el gobierno da un bono. Unos 3.4 millones de personas participan. Además, el salario mínimo es de 23.95 dólares neozelandeses por hora."),
    ]),
    (12, "KiwiSaver: ejemplo de interés compuesto", "0:55", [
        (JL, "Como nuestra materia es educación financiera, hicimos un ejemplo. Imaginemos a alguien que gana 50 mil dólares neozelandeses al año. Él ahorra 3.5 %, o sea 1,750; su empresa pone otros 1,750 y el gobierno le regala unos 261. En total junta 3,761 dólares al año, ¡y solo la mitad salió de su bolsillo!"),
        (JL, "En la gráfica vemos qué pasa en 10 años. Las barras claras muestran solo lo aportado: unos 37,600. Las barras verdes muestran lo que pasa si el dinero gana 5 % de interés al año: casi 47,300. Esa diferencia es el interés compuesto: el dinero que ganan tus ahorros también genera más dinero. Aclaramos que el 5 % es un supuesto para el ejemplo."),
    ]),
    (13, "Tema 4: Vivienda y medio ambiente", "1:00", [
        (PB, "Nuestro cuarto tema es vivienda y medio ambiente. En Nueva Zelanda hace frío y muchas casas eran húmedas, lo que enfermaba a los niños. Por eso en 2019 crearon las normas «Healthy Homes», o casas saludables: toda casa en renta debe tener calefacción, aislamiento, ventilación y control de humedad, y desde julio de 2025 es obligatorio para todas."),
        (PB, "Las familias de bajos ingresos que viven en vivienda pública pagan como máximo cerca del 25 % de su ingreso, y los ahorros de KiwiSaver se pueden usar para comprar la primera casa."),
        (PB, "En medio ambiente, 88.5 % de su electricidad en 2025 vino de fuentes renovables, como presas, el calor de la tierra, el viento y el sol. Tienen una ley para llegar a cero emisiones netas en 2050, y en 2017 reconocieron al río Whanganui como si fuera una persona con derechos, como pedía el pueblo maorí."),
    ]),
    (14, "¿Qué lo hace exitoso? (y sus retos)", "0:55", [
        (JL, "¿Por qué funciona el modelo de Nueva Zelanda? Encontramos cinco claves. Uno: confianza; hay muy poca corrupción, así que el dinero llega a donde debe. Dos: los programas son para todos, con menos trámites. Tres: son automáticos; por ejemplo, ahorrar es la opción por defecto. Cuatro: miden el bienestar y no solo el dinero. Y cinco: la gente paga impuestos, casi 33 % del PIB, para tener servicios públicos."),
        (JL, "Pero queremos ser honestos: no es perfecto. Las casas son muy caras, el desempleo subió a 5.7 % y a 12 % en jóvenes, la felicidad de los jóvenes está muy abajo en el ranking mundial y el pueblo maorí vive unos seis años menos."),
    ]),
    (15, "¿Qué podría adoptar México?", "0:55", [
        (PB, "Ahora la parte C: ¿qué podría copiar México? Elegimos cuatro ideas. Primera: el ahorro automático tipo KiwiSaver. En México existen las Afores, pero solo para quien tiene empleo formal, y más de la mitad de los trabajadores, 54.3 %, está en la informalidad."),
        (PB, "Segunda: un almuerzo escolar diario y nutritivo; México ya dio un paso al prohibir la comida chatarra en las escuelas en 2025. Tercera: un seguro de accidentes para todos, porque hoy el IMSS solo cubre a trabajadores formales. Y cuarta: reglas para que las casas sean saludables. Cada una tendría beneficios claros: más ahorro, menos hambre, menos deudas por accidentes y menos enfermedades."),
    ]),
    (16, "Obstáculos y adaptaciones", "1:00", [
        (JL, "Pero no basta con decir que un modelo es bueno; hay que ver si funcionaría aquí. Encontramos cuatro tipos de obstáculos. Económicos: México junta mucho menos en impuestos, 18.3 % de su PIB contra casi 33 % de Nueva Zelanda, así que hay menos dinero. Políticos: los programas suelen cambiar con cada gobierno. Culturales: mucha gente desconfía de las instituciones y no tiene costumbre de ahorrar. Y sociales: somos 25 veces más personas, con mucha informalidad y diferencias entre regiones."),
        (JL, "Por eso proponemos adaptaciones: inscribirse al ahorro desde el celular, aunque trabajes por tu cuenta; empezar con montos pequeños; comenzar en las zonas con más necesidad; comprar la comida escolar a productores locales; y hacer una ley para que el programa no desaparezca al cambiar de gobierno."),
    ]),
    (17, "Nuestra propuesta: «Mi Ahorro Automático»", "1:10", [
        (PB, "De todas las ideas, nosotros proponemos «Mi Ahorro Automático», un KiwiSaver a la mexicana. Funciona en cuatro pasos: uno, te inscriben solo al conseguir trabajo, o te inscribes con una app si trabajas por tu cuenta, y puedes salirte si quieres. Dos, ahorras poquito, desde 50 pesos a la semana. Tres, el gobierno te premia con 25 centavos por cada peso que ahorres, con un tope. Y cuatro, lo usas para tu retiro o para el enganche de tu primera casa."),
        (JL, "Hicimos la cuenta: si alguien ahorra 50 pesos a la semana más el bono, y su dinero gana 4 % al año, en 40 años tendría unos 309 mil pesos. Creemos que esta propuesta tendría el mayor impacto porque ataca tres problemas a la vez: la vejez sin ingresos, la falta de vivienda y la poca educación financiera. Además cuesta menos, porque la persona también aporta, y se apoya en algo que ya existe en México: las Afores."),
    ]),
    (18, "Infografía: resumen del proyecto", "0:45", [
        (PB, "Para resumir todo lo que vimos, hicimos esta infografía. A la izquierda están los pilares del modelo neozelandés: KiwiSaver y la pensión universal, el seguro de accidentes para todos y el 88.5 % de energía renovable. En el centro, la comparación con México: desarrollo humano, lugar 17 contra 81; corrupción, 81 puntos contra 27; e impuestos, 32.9 % contra 18.3 % del PIB."),
        (JL, "Y a la derecha, nuestra propuesta, «Mi Ahorro Automático»: inscripción automática y flexible desde 50 pesos a la semana, un bono del gobierno de 25 centavos por cada peso ahorrado, y un fondo con interés compuesto para el retiro o la primera vivienda."),
    ]),
    (19, "Conclusiones", "0:50", [
        (JL, "Para terminar, lo que más nos sorprendió fue que existe un seguro de accidentes que cubre hasta a los turistas, que un río tenga derechos como una persona y que México y Nueva Zelanda sean casi igual de felices."),
        (PB, "Y lo que creemos que México puede aprender es: que ahorrar sea automático y fácil; medir el éxito de un país por cómo vive su gente y no solo por el PIB; y que los buenos programas continúen aunque cambie el gobierno."),
        (JL, "Nuestra conclusión es: el bienestar no se mide solo en dinero, pero el dinero bien planeado construye bienestar."),
    ]),
    (20, "Bibliografía (2 diapositivas)", "0:15", [
        (PB, "Aquí están nuestras fuentes. Usamos fuentes oficiales: oficinas de estadística y ministerios del gobierno de Nueva Zelanda, la ONU, la OCDE, el FMI, el INEGI y la Secretaría de Educación Pública, entre otras. (Solo se muestran; no hace falta leerlas.)"),
    ]),
    (22, "¡Gracias! / Preguntas", "0:15", [
        (JL, "¡Gracias por su atención! Como dicen en Nueva Zelanda: kia ora, que en maorí significa «hola» y también «gracias»."),
        (PB, "¿Tienen alguna pregunta?"),
    ]),
]

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
