import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
doc = Document()
for s in doc.sections:
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2.54)

st = doc.styles["Normal"]
st.font.name = "Times New Roman"
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.font.size = Pt(12)
pf = st.paragraph_format
pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
pf.space_after = Pt(0)
pf.space_before = Pt(0)

# page number top-right (APA 7)
hdr = doc.sections[0].header.paragraphs[0]
hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = hdr.add_run()
for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
    if tag:
        el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag)
    else:
        el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt
    run._r.append(el)


def p(text="", align=None, bold=False, italic=False, indent=True, keep=False):
    para = doc.add_paragraph()
    if align is not None:
        para.alignment = align
    if indent and align is None:
        para.paragraph_format.first_line_indent = Cm(1.27)
    _runs(para, text, bold, italic)
    if keep:
        para.paragraph_format.keep_with_next = True
    return para


def _runs(para, text, bold=False, italic=False):
    # *cursiva* support for titles of works
    parts = text.split("*")
    for i, part in enumerate(parts):
        if part:
            r = para.add_run(part)
            r.bold = bold
            r.italic = italic or (i % 2 == 1)


def h1(text):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = para.add_run(text); r.bold = True
    para.paragraph_format.keep_with_next = True


def h2(text):
    para = doc.add_paragraph()
    r = para.add_run(text); r.bold = True
    para.paragraph_format.keep_with_next = True


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def table(num, title, header, rows, widths, note):
    a = doc.add_paragraph(); r = a.add_run(f"Tabla {num}"); r.bold = True
    a.paragraph_format.keep_with_next = True
    b = doc.add_paragraph(); r = b.add_run(title); r.italic = True
    b.paragraph_format.keep_with_next = True
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, w in enumerate(widths):
        t.columns[i].width = Cm(w)
    for ri, row in enumerate([header] + rows):
        for ci, val in enumerate(row):
            c = t.rows[ri].cells[ci]
            c.width = Cm(widths[ci])
            c.text = ""
            pp = c.paragraphs[0]
            pp.paragraph_format.line_spacing = 1.0
            pp.paragraph_format.first_line_indent = Cm(0)
            rr = pp.add_run(val); rr.font.size = Pt(10.5); rr.bold = ri == 0
    # APA-style horizontal rules only
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, val in (("top", "single"), ("bottom", "single"), ("left", "nil"), ("right", "nil"), ("insideH", "nil"), ("insideV", "nil")):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), val); e.set(qn("w:sz"), "8"); e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)
    for c in t.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr(); tb = OxmlElement("w:tcBorders")
        e = OxmlElement("w:bottom"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "8"); e.set(qn("w:color"), "000000"); tb.append(e); tcPr.append(tb)
    n = doc.add_paragraph()
    n.paragraph_format.line_spacing = 1.0
    n.paragraph_format.space_before = Pt(4)
    n.paragraph_format.space_after = Pt(12)
    r = n.add_run("Nota. "); r.italic = True; r.font.size = Pt(10.5)
    r = n.add_run(note); r.font.size = Pt(10.5)


# ======================= PORTADA (APA 7, estudiante) =======================
for _ in range(3):
    doc.add_paragraph()
p("Países que construyen bienestar: modelos para una mejor sociedad", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
p("El caso de Nueva Zelanda y sus lecciones para México", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
doc.add_paragraph()
for line in ["José Luis De Anda Reimers y Pablo De Santiago Hurtado Martin Del Campo",
             "Instituto Educativo Ammadeus",
             "Educación Financiera, 2.º de secundaria",
             "Mtra. Samantha Rubí Lara Ceh",
             "9 de noviembre de 2026"]:
    p(line, align=WD_ALIGN_PARAGRAPH.CENTER)
page_break()

# ======================= TEXTO =======================
h1("Países que construyen bienestar: modelos para una mejor sociedad")

h1("Introducción")
p("Cuando escuchamos que un país «va bien», casi siempre se habla de dinero: cuánto creció su economía o cuánto gana su gente. Sin embargo, una sociedad no vive mejor solo porque produce más; también importa que sus habitantes estén sanos, estudien, tengan una casa digna, puedan ahorrar para su futuro y vivan en un ambiente limpio. En este ensayo analizamos a Nueva Zelanda, un país que decidió medir su éxito con el bienestar de las personas y no únicamente con el crecimiento económico.")
p("El objetivo es responder una pregunta: ¿qué hace Nueva Zelanda para que su población viva mejor y qué de eso podría funcionar en México? Para lograrlo, primero describimos al país; después explicamos sus principales modelos de bienestar en cuatro temas —educación, salud, empleo y pensiones, y vivienda y medio ambiente—; luego hacemos un análisis crítico sobre cómo podrían adaptarse a México, y al final presentamos nuestras conclusiones. Todos los datos provienen de fuentes oficiales, como la oficina de estadística de Nueva Zelanda, la Organización de las Naciones Unidas (ONU), la Organización para la Cooperación y el Desarrollo Económicos (OCDE) y el Instituto Nacional de Estadística y Geografía (INEGI).")

h1("Conociendo Nueva Zelanda")
h2("Ubicación geográfica y bandera")
p("Nueva Zelanda, llamada *Aotearoa* en idioma maorí, se encuentra en Oceanía, en el océano Pacífico Sur, a unos 2,000 kilómetros de Australia y a cerca de 11,100 kilómetros de la Ciudad de México. Está formada por dos islas principales, la Isla Norte y la Isla Sur, y tiene una superficie aproximada de 268,000 km², es decir, unas siete veces menos que México. Su capital es Wellington y su ciudad más grande es Auckland. Tiene tres idiomas oficiales: el inglés, el maorí (lengua del pueblo originario) y la lengua de señas neozelandesa; su moneda es el dólar neozelandés (NZ$).")
p("Su bandera muestra dos elementos importantes de su identidad. En la esquina superior izquierda aparece la bandera del Reino Unido, que recuerda que el país fue colonia británica; a la derecha hay cuatro estrellas rojas con borde blanco que representan la Cruz del Sur, una constelación que solo se observa desde el hemisferio sur.")

h2("Población y distribución por sexo")
p("En junio de 2025, Nueva Zelanda tenía 5,324,700 habitantes y su población crecía 0.7 % al año. La distribución por sexo es casi igual: 2,678,500 mujeres (50.3 %) y 2,646,200 hombres (49.7 %). La edad mediana, que es la edad que divide a la población en dos mitades, era de 39.0 años para las mujeres y de 37.4 para los hombres (Stats NZ, 2025b). Para comparar, México tiene más de 130 millones de habitantes (Consejo Nacional de Población [CONAPO], 2023), alrededor de 25 veces más. Esta diferencia es importante, porque no es lo mismo organizar un programa social para cinco millones de personas que para más de cien millones.")

h2("Esperanza de vida y distribución por edades")
p("La esperanza de vida al nacer en Nueva Zelanda es de 81.8 años en promedio: 83.5 años para las mujeres y 80.1 para los hombres. Sin embargo, no todos los grupos viven lo mismo: el pueblo maorí tiene una esperanza de vida de 75.8 años, unos seis años menos que el promedio (Stats NZ, 2025a). Este dato muestra que, incluso en los países con más bienestar, sigue existiendo desigualdad.")
p("En cuanto a las edades, según el censo de 2023, el 18.7 % de la población tenía entre 0 y 14 años, el 64.7 % tenía entre 15 y 64 años (edad de trabajar) y el 16.6 % tenía 65 años o más (Stats NZ, 2024). Como la proporción de adultos mayores está creciendo, el país necesita un sistema de pensiones fuerte, tema que analizaremos más adelante.")

h2("Actividades económicas, PIB y nivel de desarrollo")
p("La economía neozelandesa depende mucho del campo y de la naturaleza. Su principal exportación son los productos lácteos, como la leche en polvo, la mantequilla y el queso. En 2025 el país exportó bienes por NZ$ 80.7 mil millones, de los cuales NZ$ 24.3 mil millones (cerca del 30 %) fueron lácteos y NZ$ 10.4 mil millones fueron carne (Stats NZ, 2026b). También son importantes la madera, frutas como el kiwi, el turismo y los servicios, como el comercio, la educación y la tecnología.")
p("El producto interno bruto (PIB), que es el valor de todo lo que produce un país en un año, fue de aproximadamente 259 mil millones de dólares estadounidenses en 2025, lo que equivale a unos US$ 48,600 por habitante (Fondo Monetario Internacional [FMI], 2025). Por esto, Nueva Zelanda es considerada una economía desarrollada con un nivel de desarrollo humano «muy alto».")

h2("Otros indicadores económicos y sociales")
p("Para entender mejor la situación del país, comparamos cuatro indicadores con México (ver Tabla 1). En el Índice de Desarrollo Humano de la ONU, que combina salud, educación e ingreso, Nueva Zelanda ocupa el lugar 17 y México el 81 (Programa de las Naciones Unidas para el Desarrollo [PNUD], 2025). En el Índice de Percepción de la Corrupción, Nueva Zelanda obtuvo 81 de 100 puntos y quedó en cuarto lugar mundial, mientras que México obtuvo 27 puntos (Transparency International, 2026). Además, Nueva Zelanda recauda en impuestos el 32.9 % de su PIB, mientras que México recauda el 18.3 %, la cifra más baja de los países de la OCDE (OCDE, 2025). Ese dinero es el que los gobiernos usan para pagar escuelas, hospitales y pensiones.")
p("El dato que más nos sorprendió fue el de la felicidad: en el *Informe Mundial de la Felicidad 2026*, Nueva Zelanda quedó en el lugar 11 y México en el 12, prácticamente empatados (Helliwell et al., 2026). Esto nos enseña que la satisfacción con la vida no depende solo del dinero o de los servicios públicos, sino también de la familia, los amigos y la comunidad.")
table(1, "Comparación de indicadores entre Nueva Zelanda y México",
      ["Indicador", "Nueva Zelanda", "México"],
      [["Índice de Desarrollo Humano (2025)", "Lugar 17 (0.938)", "Lugar 81 (0.789)"],
       ["Percepción de la corrupción (2025)", "81 puntos (4.º lugar)", "27 puntos (141.º lugar)"],
       ["Impuestos como % del PIB (2024)", "32.9 %", "18.3 %"],
       ["Informe Mundial de la Felicidad (2026)", "Lugar 11", "Lugar 12"]],
      [7.2, 4.4, 4.4],
      "Elaboración propia con datos de PNUD (2025), Transparency International (2026), OCDE (2025) y Helliwell et al. (2026).")

h1("Modelos y propuestas de bienestar social")
h2("La gran idea: el Presupuesto de Bienestar")
p("Antes de explicar cada tema, hay una idea que distingue a Nueva Zelanda. En 2019 se convirtió en el primer país en elaborar un «Presupuesto de Bienestar»: en lugar de preguntarse solamente cuánto crecería la economía, el gobierno se preguntó si la gente estaba viviendo mejor, y por eso destinó recursos a la salud mental, la pobreza infantil, el apoyo al pueblo maorí y una economía más sustentable (New Zealand Treasury, 2019). Desde 2020, una ley de finanzas públicas pide al gobierno informar sus objetivos de bienestar. No obstante, en 2023 llegó un nuevo gobierno con otras prioridades, lo que demuestra que estos modelos también dependen de las decisiones políticas.")

h2("Tema 1. Educación: almuerzos escolares saludables")
p("El programa de almuerzos escolares saludables, que comenzó en 2020, consiste en dar una comida nutritiva y gratuita a los alumnos cada día de clases. El problema que busca resolver es sencillo pero grave: un niño que llega con hambre a la escuela no puede concentrarse ni aprender. El gobierno lo financia en las escuelas que enfrentan más dificultades económicas y, en agosto de 2026, ya beneficiaba a más de 240,000 estudiantes en 1,023 escuelas (Ministry of Education, 2026). Diversos estudios reportan que el programa ha reducido el hambre y mejorado el bienestar de los alumnos.")
p("Lo que hace diferente a este modelo es que forma parte de un sistema educativo amplio: la escuela pública es gratuita de los 5 a los 19 años y los niños de 3 a 5 años tienen 20 horas gratuitas de educación inicial a la semana. Los resultados se reflejan en la prueba PISA 2022, que evalúa a estudiantes de 15 años: Nueva Zelanda obtuvo 479 puntos en matemáticas, 501 en lectura y 504 en ciencias, por encima del promedio de la OCDE (472, 476 y 485), mientras que México obtuvo 395, 415 y 410 puntos (OCDE, 2023).")

h2("Tema 2. Salud: un seguro de accidentes para todos")
p("En salud encontramos un modelo casi único en el mundo: la Corporación de Compensación por Accidentes (ACC, por sus siglas en inglés), que funciona desde 1974. Consiste en un seguro que cubre cualquier accidente —en el trabajo, en la calle, en la casa o practicando deporte— sin importar de quién haya sido la culpa. Busca evitar que un accidente deje a una familia sin atención médica o sin ingresos. Se financia con cuotas que se cobran sobre los salarios, las empresas y los vehículos, además de recursos del gobierno, y protege a los residentes e incluso a los turistas. En el periodo 2024-2025 aceptó más de 2.1 millones de casos y pagó NZ$ 5 mil millones en tratamiento y rehabilitación y NZ$ 3 mil millones en compensaciones (ACC, 2025). Lo que lo hace diferente es que, a cambio de esta protección, las personas generalmente no pueden demandar por lesiones personales.")
p("El sistema se completa con otras medidas. Pharmac, creada en 1993, es una sola agencia que compra los medicamentos para todo el país con un presupuesto fijo; como compra en grandes cantidades, negocia precios más bajos (Pharmac, 2025). Además, los hospitales públicos son gratuitos para los residentes y los niños menores de 14 años pueden consultar al médico general sin costo en la mayoría de las clínicas.")

h2("Tema 3. Empleo y pensiones: dos pilares para el retiro")
p("El sistema de pensiones neozelandés tiene dos pilares. El primero es NZ Superannuation, una pensión universal para todas las personas residentes de 65 años o más, sin importar si trabajaron o no; en 2026, una persona que vive sola recibe alrededor de NZ$ 555 a la semana después de impuestos (Work and Income New Zealand, 2026). El segundo pilar es KiwiSaver, creado en 2007: cuando una persona comienza un nuevo empleo, queda inscrita automáticamente en una cuenta de ahorro para el retiro. Desde abril de 2026, el trabajador aporta 3.5 % de su sueldo, la empresa aporta otro 3.5 % y el gobierno añade un bono de 25 centavos por cada dólar ahorrado, con un máximo de NZ$ 260.72 al año (Inland Revenue, 2025). Quien no quiera participar puede salirse, pero si no hace nada, ahorra. Alrededor de 3.4 millones de personas participan en este programa.")
p("Estos modelos buscan resolver dos problemas: la pobreza en la vejez y la falta de ahorro. Su resultado es que los adultos mayores tienen un ingreso seguro y que millones de personas ahorran casi sin darse cuenta. Además, el salario mínimo en 2026 es de NZ$ 23.95 por hora (Employment New Zealand, 2026).")
p("Como nuestra materia es Educación Financiera, hicimos un ejemplo. Una persona que gana NZ$ 50,000 al año ahorra NZ$ 1,750; su empresa aporta otros NZ$ 1,750 y el gobierno NZ$ 261, lo que da un total de NZ$ 3,761 al año, aunque solo la mitad sale de su bolsillo. En diez años habría aportado unos NZ$ 37,600; pero si el dinero gana un interés de 5 % anual (un supuesto para el ejemplo), tendría cerca de NZ$ 47,300. Esa diferencia se debe al interés compuesto: los intereses que gana el ahorro también generan nuevos intereses.")

h2("Tema 4. Vivienda y medio ambiente")
p("En vivienda, Nueva Zelanda enfrentaba un problema: muchas casas eran frías y húmedas, lo que provocaba enfermedades, sobre todo en los niños. Por eso, en 2019 se crearon las normas *Healthy Homes* («casas saludables»), que obligan a que toda casa en renta tenga calefacción, aislamiento, ventilación y control de humedad; desde julio de 2025 todas las viviendas en renta deben cumplirlas (Tenancy Services, 2025). Además, las familias de bajos ingresos que viven en vivienda pública pagan como máximo alrededor del 25 % de su ingreso como renta (Kāinga Ora, 2025), y los ahorros de KiwiSaver pueden usarse para dar el enganche de la primera casa.")
p("En medio ambiente, el 88.5 % de la electricidad generada en 2025 provino de fuentes renovables, como presas, energía geotérmica, viento y sol (Ministry of Business, Innovation and Employment [MBIE], 2026). La Ley Cero Carbono de 2019 establece la meta de llegar a cero emisiones netas en 2050 (Ministry for the Environment, 2019). Un caso especialmente interesante es el del río Whanganui, que en 2017 fue reconocido legalmente como una persona con derechos, tal como lo pedía el pueblo maorí, que lo considera un ancestro (Te Awa Tupua Act, 2017).")

h2("¿Qué hace exitoso a este modelo?")
p("Después de investigar, identificamos cinco claves del éxito neozelandés. La primera es la confianza: hay muy poca corrupción, por lo que el dinero público llega a donde debe. La segunda es que muchos programas son universales, es decir, para todos, lo que reduce trámites y evita que las personas se sientan señaladas. La tercera es que son automáticos, como KiwiSaver, en el que ahorrar es la opción «por defecto». La cuarta es que el país mide el bienestar y no solo el dinero. La quinta es que la población paga impuestos suficientes para sostener los servicios públicos.")
p("Sin embargo, el modelo no es perfecto. La vivienda es muy cara, el desempleo subió a 5.7 % en septiembre de 2026 y llegó a 12 % entre los jóvenes de 20 a 24 años (Stats NZ, 2026a). Además, en felicidad juvenil el país ocupa el lugar 126 de 136 (Helliwell et al., 2026), y, como mencionamos, el pueblo maorí vive unos seis años menos que el promedio.")

h1("¿Cómo podemos replicarlo en México?")
h2("Propuestas que podrían implementarse")
p("A partir de lo investigado, seleccionamos cuatro propuestas que podrían aplicarse en México (ver Tabla 2). La primera es un ahorro automático parecido a KiwiSaver. En México existen las Afores, pero solo cubren a quienes tienen un empleo formal, y en septiembre de 2026 el 54.3 % de los trabajadores estaba en la informalidad (INEGI, 2026). La segunda es un almuerzo escolar diario y nutritivo; México ya dio un primer paso al prohibir la venta de comida chatarra en las escuelas desde 2025 (Secretaría de Educación Pública y Secretaría de Salud, 2024). La tercera es un seguro de accidentes universal, ya que hoy el Instituto Mexicano del Seguro Social (IMSS) solo protege contra riesgos de trabajo a los trabajadores formales. La cuarta es establecer reglas de vivienda saludable adaptadas a los distintos climas del país.")
table(2, "Propuestas de Nueva Zelanda que podrían adaptarse en México",
      ["Propuesta", "Situación actual en México", "Beneficio esperado"],
      [["Ahorro automático (KiwiSaver)", "Afores solo para empleo formal; 54.3 % de informalidad", "Más personas con ahorro para la vejez y la vivienda"],
       ["Almuerzo escolar nutritivo", "Prohibición de comida chatarra en escuelas desde 2025", "Menos hambre y obesidad; mejor aprendizaje"],
       ["Seguro universal de accidentes (ACC)", "El IMSS cubre riesgos de trabajo solo a trabajadores formales", "Que nadie se endeude por un accidente"],
       ["Normas de vivienda saludable", "Créditos de vivienda, pero pocas reglas de calidad en rentas", "Menos enfermedades por humedad, frío o calor"]],
      [4.6, 6.0, 5.4],
      "Elaboración propia con datos de INEGI (2026) y Secretaría de Educación Pública y Secretaría de Salud (2024).")

h2("Beneficios, obstáculos y adaptaciones necesarias")
p("Adoptar estas propuestas traería beneficios claros: más personas con ahorro para su retiro y su vivienda, niños mejor alimentados que aprenden más, familias protegidas ante accidentes y hogares más sanos. Sin embargo, no basta con decir que un modelo es bueno; es necesario preguntarse si funcionaría en otro contexto. Identificamos cuatro tipos de obstáculos. Los económicos: México recauda en impuestos 18.3 % de su PIB, frente a 32.9 % de Nueva Zelanda, por lo que hay menos recursos (OCDE, 2025). Los políticos: muchos programas cambian o desaparecen con cada nuevo gobierno. Los culturales: existe desconfianza hacia las instituciones y poca costumbre de ahorrar. Y los sociales: México tiene 25 veces más habitantes, un alto nivel de informalidad y grandes diferencias entre regiones.")
p("Para que estas ideas funcionen, serían necesarias varias adaptaciones: permitir la inscripción al ahorro desde el teléfono celular, incluso para quienes trabajan por su cuenta; empezar con montos pequeños y flexibles; iniciar en las zonas con más necesidad e ir creciendo poco a poco; comprar los alimentos escolares a productores locales, lo que además ayudaría a la economía de cada región, y aprobar leyes para que los programas no desaparezcan al cambiar el gobierno.")

h2("Nuestra propuesta: «Mi Ahorro Automático»")
p("De todas las ideas, proponemos «Mi Ahorro Automático», un KiwiSaver adaptado a México que funcionaría en cuatro pasos. Primero, la persona quedaría inscrita automáticamente al conseguir un empleo, o podría inscribirse con una aplicación si trabaja por su cuenta, con libertad para salirse. Segundo, ahorraría una cantidad pequeña de forma automática, desde 50 pesos a la semana. Tercero, el gobierno la premiaría con 25 centavos por cada peso ahorrado, con un límite. Cuarto, el dinero se usaría para el retiro o para el enganche de una primera vivienda. Según nuestro cálculo, si una persona ahorra 50 pesos a la semana, recibe el bono y su dinero gana un interés de 4 % anual (supuesto), en 40 años tendría alrededor de 309,000 pesos.")
p("Consideramos que esta propuesta tendría el mayor impacto en la calidad de vida por tres razones. En primer lugar, ataca tres problemas a la vez: la vejez sin ingresos, la falta de vivienda y la poca educación financiera. En segundo lugar, es más barata que otros programas, porque la persona también aporta y el interés compuesto hace crecer el dinero con el tiempo. En tercer lugar, puede apoyarse en algo que ya existe en México, el sistema de Afores, por lo que no habría que empezar desde cero.")

h1("Conclusiones")
p("Lo que más nos sorprendió de Nueva Zelanda fue descubrir que existe un seguro de accidentes que protege a todos, incluso a los turistas; que un río puede tener derechos como una persona, y que México y Nueva Zelanda son casi igual de felices, a pesar de las grandes diferencias en sus ingresos y servicios públicos. También aprendimos que ningún país es perfecto: incluso Nueva Zelanda enfrenta problemas como la vivienda cara, el desempleo juvenil y la desigualdad que vive el pueblo maorí.")
p("Creemos que México puede aprender tres lecciones de este país. La primera es que ahorrar debe ser automático y fácil, porque así más personas lo hacen. La segunda es que el éxito de un país debe medirse por cómo vive su gente y no solo por su PIB. La tercera es que los buenos programas deben mantenerse aunque cambie el gobierno, porque los resultados del bienestar se ven a largo plazo. En resumen, el bienestar no se mide solo en dinero, pero el dinero bien planeado —en un país, en una familia o en nuestra propia alcancía— construye bienestar.")

# ======================= REFERENCIAS =======================
page_break()
h1("Referencias")
REFS = [
    "ACC. (2025). *Annual report 2025*. Accident Compensation Corporation. https://www.acc.co.nz",
    "Consejo Nacional de Población. (2023). *Proyecciones de la población de México y de las entidades federativas 2020-2070*. Gobierno de México. https://www.gob.mx/conapo",
    "Employment New Zealand. (2026). *Minimum wage is increasing on 1 April 2026*. https://www.employment.govt.nz",
    "Fondo Monetario Internacional. (2025). *World Economic Outlook database: October 2025* [Base de datos]. https://www.imf.org/en/Publications/WEO",
    "Helliwell, J. F., Layard, R., Sachs, J. D., De Neve, J.-E., Aknin, L. B. y Wang, S. (Eds.). (2026). *World happiness report 2026*. Wellbeing Research Centre, University of Oxford. https://worldhappiness.report",
    "Inland Revenue. (2025). *KiwiSaver changes*. https://www.ird.govt.nz/kiwisaver-changes",
    "Instituto Nacional de Estadística y Geografía. (2026). *Indicadores de ocupación y empleo: septiembre de 2026* [Comunicado de prensa]. https://www.inegi.org.mx",
    "Kāinga Ora – Homes and Communities. (2025). *Income-related rent*. https://kaingaora.govt.nz",
    "Ministry for the Environment. (2019). *Climate Change Response (Zero Carbon) Amendment Act 2019*. https://environment.govt.nz",
    "Ministry of Business, Innovation and Employment. (2026). *Energy in New Zealand 2026*. https://www.mbie.govt.nz",
    "Ministry of Education. (2026). *Healthy school lunches programme*. https://www.education.govt.nz",
    "New Zealand Treasury. (2019). *The wellbeing budget 2019*. https://www.treasury.govt.nz",
    "Organización para la Cooperación y el Desarrollo Económicos. (2023). *PISA 2022 results (Volume I): The state of learning and equity in education*. OECD Publishing. https://doi.org/10.1787/53f23881-en",
    "Organización para la Cooperación y el Desarrollo Económicos. (2025). *Revenue statistics 2025*. OECD Publishing. https://www.oecd.org",
    "Pharmac. (2025). *How Pharmac works*. https://pharmac.govt.nz",
    "Programa de las Naciones Unidas para el Desarrollo. (2025). *Informe sobre desarrollo humano 2025: Una cuestión de elección*. https://hdr.undp.org",
    "Secretaría de Educación Pública y Secretaría de Salud. (2024). *Lineamientos generales para la preparación, distribución y expendio de alimentos y bebidas en escuelas*. Diario Oficial de la Federación. https://www.dof.gob.mx",
    "Stats NZ. (2024). *2023 Census population, dwelling, and housing counts*. https://www.stats.govt.nz",
    "Stats NZ. (2025a). *National and subnational period life tables: 2022–2024*. https://www.stats.govt.nz",
    "Stats NZ. (2025b). *National population estimates: At 30 June 2025*. https://www.stats.govt.nz",
    "Stats NZ. (2026a). *Labour market statistics: September 2026 quarter*. https://www.stats.govt.nz",
    "Stats NZ. (2026b). *Overseas merchandise trade: December 2025*. https://www.stats.govt.nz",
    "Te Awa Tupua (Whanganui River Claims Settlement) Act 2017. (2017). New Zealand Legislation. https://www.legislation.govt.nz",
    "Tenancy Services. (2025). *Healthy homes standards*. Ministry of Housing and Urban Development. https://www.tenancy.govt.nz",
    "Transparency International. (2026). *Corruption perceptions index 2025*. https://www.transparency.org/cpi",
    "Work and Income New Zealand. (2026). *Benefit rates at 1 April 2026: New Zealand Superannuation*. https://www.workandincome.govt.nz",
]
for ref in REFS:
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(1.27)
    para.paragraph_format.first_line_indent = Cm(-1.27)
    _runs(para, ref)

doc.save(OUT)
import re
body = " ".join(pp.text for pp in doc.paragraphs)
print("saved", OUT, "words:", len(body.split()))
