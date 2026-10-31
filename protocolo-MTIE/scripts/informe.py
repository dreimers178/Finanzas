"""Informe de revisión metodológica y verificación de referencias."""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))
from apa import *  # noqa
from docx.enum.section import WD_ORIENT

OUT = pathlib.Path(__file__).resolve().parents[1] / "Informe_revision_metodologica_y_referencias.docx"
doc = new_document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = Cm(27.94), Cm(21.59)
st = doc.styles["Normal"].paragraph_format
st.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
for i in range(1, 5):
    doc.styles[f"Heading {i}"].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

heading(doc, "Informe de revisión metodológica y verificación de referencias", 1)
para(doc, "**Protocolo:** Nutrición en el aula: taller educativo basado en microaprendizaje en línea para niños de 5.º de primaria  "
          "·  **Autora:** L.N.C.A. Anakaren Ernestina Reimers Reyes  ·  **Director:** Dr. Eduardo Rivera Arteaga  ·  "
          "**Fecha de revisión:** 31 de octubre de 2026", indent=False)
para(doc, "Este informe acompaña al archivo *Protocolo_Reimers_corregido_APA7.docx*, que contiene todos los cambios aplicados y un "
          "comentario de Word en cada uno explicando el motivo. Lo resaltado en amarillo en el protocolo es información que la autora "
          "debe completar o confirmar.", indent=False)

heading(doc, "1. Resumen de hallazgos", 2)
bullets(doc, [
    "**Integridad de las referencias (prioridad crítica).** El protocolo cita 35 obras distintas, pero solo 21 estaban en la lista de referencias (14 se citaban sin referencia). Quince obras no pudieron localizarse con los datos indicados y otras cuatro tenían datos erróneos o insuficientes. Seis de las no localizadas eran antecedentes con datos de estudio detallados (muestras, métodos, resultados) que no se encontraron en ninguna base. Todas se sustituyeron por fuentes verificadas.",
    "**Clasificación del diseño (decisión de la autora).** Se conserva «cuasiexperimental» con un solo grupo. Algunos manuales clasifican este diseño como *preexperimental* (Campbell y Stanley, 1966; Hernández-Sampieri y Mendoza Torres, 2018), por lo que en el protocolo se explicitan las amenazas a la validez interna y se sugiere, como mejora opcional, agregar un grupo de comparación.",
    "**Análisis de datos (prioridad alta).** El protocolo solo proponía estadística descriptiva, lo que no permite contrastar la hipótesis. Se agregó la prueba *t* pareada o Wilcoxon con tamaño del efecto.",
    "**Congruencia del instrumento (prioridad alta).** Las dimensiones de la operacionalización no coincidían con las del instrumento enviado a jueces. Faltaba la confiabilidad (KR-20) y la prueba piloto.",
    "**Observaciones del asesor.** Se atendieron las seis: objetivo general, pregunta 2, término único, instrumento en anexos, desarrollo de ADDIE y cambio de juez.",
    "**Congreso (CITEEA).** Resumen reducido a 150 palabras exactas y tres palabras clave.",
])

heading(doc, "2. Atención a las observaciones del asesor (Dr. Eduardo Rivera Arteaga)", 2)
table(doc, 1, "Observaciones del asesor y acciones realizadas", ["Observación", "Acción realizada", "Ubicación"], [
    ["Pregunta 2: no iniciar con «cómo»", "Ya estaba atendida en la versión recibida («¿Qué elementos digitales…?»). Se mantuvo y se vinculó con el OE2 y la carta descriptiva.", "1.4.2; Tabla 4"],
    ["El objetivo general no debe ser «Evaluar…»; debe ser implementar o incrementar", "OG: «Incrementar el conocimiento nutricional… mediante la implementación de un taller educativo en línea basado en microaprendizaje». Es la misma redacción del resumen aceptado.", "1.5.1"],
    ["Usar un solo término (estrategia, propuesta o taller) en todo el documento", "Se usa «taller educativo en línea basado en microaprendizaje» en el título, los objetivos, las variables y el texto. Se eliminaron «estrategia educativa», «curso», «modelo educativo» y «taller virtual».", "Todo el documento"],
    ["Título: enfatizar el conocimiento y mencionar el taller", "El título del resumen aceptado ya menciona el taller. En un comentario se propone una alternativa que nombra el conocimiento nutricional.", "Portada"],
    ["El cuestionario va en los anexos", "Anexo A con los 24 reactivos y sus claves (la versión para estudiantes debe ir sin claves).", "Anexo A"],
    ["El modelo ADDIE no solo se describe: hay que desarrollar sus etapas", "Se desarrollaron las cinco fases con decisiones concretas y una carta descriptiva de ocho cápsulas alineadas con las cuatro dimensiones.", "3.8; Tabla 4"],
    ["Incluir otro juez en lugar del asesor (sugiere a la Dra. Verónica Torres)", "Se indica en 3.4.1 que los jueces deben ser externos a la dirección de tesis. Se agregó la V de Aiken con IC 95 % como criterio de decisión.", "3.4.1"],
    ["(Resumen) «Se realizará la evaluación pretest-postest para identificar si hubo un cambio significativo…»", "Incorporado en el resumen del congreso. En el protocolo se especifica la prueba estadística que permite afirmar «significativo».", "Resumen; 3.6"],
    ["(Resumen) Eje temático", "El resumen aceptado usa «Metodologías activas y pedagogías emergentes». Es coherente con el carácter de la propuesta y se mantiene.", "Cartel"],
], widths=[6.5, 13, 3.5])

heading(doc, "3. Lista de cotejo del congreso y resumen de 150 palabras", 2)
table(doc, 2, "Criterios no cumplidos en la lista de cotejo del CITEEA", ["Criterio", "Observación del comité", "Acción"], [
    ["3. Extensión (máx. 150 palabras)", "No cumple: 223 palabras", "Resumen reescrito: 150 palabras exactas (ver abajo)."],
    ["7. Palabras clave", "Son muchas; tres son suficientes", "microaprendizaje, educación nutricional, educación primaria"],
    ["11. Plantilla oficial", "Integrar al formato oficial del congreso", "==Pendiente: pegar el resumen en la plantilla oficial del CITEEA (no se recibió ese archivo).=="],
    ["Título", "La lista registra el título anterior («Evaluación de propuesta educativa…»)", "Usar el título actual en todos los materiales: «Nutrición en el aula: taller educativo basado en microaprendizaje en línea para niños de 5.º de primaria»."],
], widths=[5.5, 7, 10.5])
para(doc, "**Resumen para el congreso (150 palabras):**", indent=False)
para(doc, "La educación nutricional en la infancia favorece la adquisición de conocimientos sobre alimentación saludable. El objetivo "
          "de esta investigación es incrementar el conocimiento nutricional de estudiantes de quinto grado de primaria mediante un "
          "taller educativo en línea basado en microaprendizaje. El taller se diseñó con el modelo ADDIE y organiza ocho cápsulas "
          "digitales breves con base en el Plato del Bien Comer y las Guías Alimentarias Saludables y Sostenibles para la Población "
          "Mexicana. Se trata de una investigación aplicada, con enfoque cuantitativo y diseño cuasiexperimental de preprueba y "
          "posprueba con un solo grupo. La muestra se integra por 25 estudiantes del Instituto Educativo Ammadeus, en Guadalupe, "
          "Zacatecas. Se aplicará un cuestionario de 24 reactivos antes y después del taller. Como avance, el instrumento se "
          "encuentra en validación de contenido mediante juicio de expertos. Posteriormente se implementará el taller y se "
          "determinará si hubo un cambio significativo en el conocimiento nutricional.", indent=False)
para(doc, "*Palabras clave:* microaprendizaje, educación nutricional, educación primaria", indent=False)
para(doc, "*Nota.* Si se agrega un grupo de comparación, sustituye «con un solo grupo» por «con grupo de comparación».",
     indent=False, size=10)

heading(doc, "4. Observaciones metodológicas", 2)
table(doc, 3, "Observaciones metodológicas y correcciones", ["#", "Apartado", "Problema detectado", "Corrección aplicada", "Prioridad"], [
    ["1", "3.1 Diseño", "Se denomina «cuasiexperimental» a un diseño de un solo grupo sin control.", "Se conserva «cuasiexperimental» por decisión de la autora; se precisa que es un solo grupo intacto, se explicitan las amenazas a la validez interna y se sugiere, de forma opcional, agregar otro grupo de 5.º como comparación.", "Alta"],
    ["2", "3.3 / 3.6 Análisis", "Solo estadística descriptiva; no permite aceptar o rechazar H0. La plantilla sugería Chi cuadrado (no aplica).", "Shapiro-Wilk → t pareada unilateral o Wilcoxon; α = .05; tamaño del efecto; Bonferroni por dimensión.", "Alta"],
    ["3", "Operacionalización", "Las dimensiones de la tabla no coinciden con las del instrumento real.", "Tabla 3 alineada con D1–D4 del instrumento, reactivos por dimensión, escala dicotómica y puntaje 0–24.", "Alta"],
    ["4", "Instrumento", "No se mencionaba la confiabilidad ni la prueba piloto. La técnica se llamaba «encuesta».", "Prueba piloto con un grupo no participante, KR-20 ≥ .70, índices de dificultad y discriminación. Técnica: prueba de conocimientos.", "Alta"],
    ["5", "Hipótesis", "H1 direccional («mayor») con H0 bilateral («no existirán diferencias»). La pregunta general era bilateral.", "H0 como complemento de H1 («igual o menor»); pregunta general reformulada («¿Se incrementa…?»).", "Media"],
    ["6", "Muestra", "Se afirmaba que el grupo es «representativo» y la muestra es por conveniencia.", "Se eliminó la afirmación; se agregó un análisis de potencia (n = 25 → d ≈ 0.50 con potencia .80).", "Media"],
    ["7", "Criterios", "Se mezclaban criterios de exclusión y de eliminación; «no pertenecer al grupo» no es un criterio.", "Inclusión, exclusión y eliminación separados; umbral de exposición de 6/8 cápsulas.", "Media"],
    ["8", "Variable independiente", "Ítems y escala «No aplica»; sin control de fidelidad.", "Indicadores de fidelidad de implementación y lista de cotejo (Anexo C).", "Media"],
    ["9", "Ética", "Sin apartado propio (población menor de edad).", "Apartado 3.7: autorización, consentimiento, asentimiento, confidencialidad y riesgo mínimo.", "Media"],
    ["10", "Objetivos", "El OE3 «Evaluar los cambios» duplicaba el OG anterior; faltaba «desarrollar».", "OE2 «Diseñar y desarrollar… ADDIE»; OE3 «Comparar…». Matriz de congruencia (Tabla 1).", "Media"],
    ["11", "Planteamiento", "Dos encabezados idénticos; una afirmación atribuida a la ONU que no está en la Agenda 2030; una cita genérica de la SEP.", "Fusionados; paráfrasis corregida; cita de la SEP eliminada.", "Media"],
    ["12", "Antecedentes", "Sin párrafo de cierre que identifique el vacío de conocimiento.", "Párrafo de síntesis que justifica el estudio.", "Baja"],
    ["13", "Marco contextual", "El apartado «Misión» describía la infraestructura; frases valorativas sin fuente.", "Reorganizado en institución y grupo participante; se pide citar la fuente de los datos.", "Baja"],
    ["14", "Alcance temporal", "Implementación en octubre, que ya pasó y es incompatible con el avance reportado.", "Actualizado: preprueba del 18 al 22 de enero de 2027; taller del 25 de enero al 19 de febrero, con inicio en la semana del Día del Nutriólogo (27 de enero); posprueba del 22 al 26 de febrero de 2027 (Tabla 5).", "Media"],
    ["15", "Modalidad", "Taller «en línea», pero impartido en el aula con proyector.", "Se precisa que la plataforma es en línea y el acceso es en el aula (decisión pendiente).", "Media"],
    ["16", "Capítulos 2, 4 y 5", "Quedaba texto de la plantilla (ejemplo de «Motivación», «2.1 Tema 1», instrucciones).", "Eliminado; caps. 4 y 5 con el plan de presentación de resultados.", "Baja"],
    ["17", "Redacción", "Erratas: limkitará, condicyas, permitirpa, enlos, irra, alis, metodológicaquew, Teporía, ractios, instrumenetos, «sino» por «si no», contínua, «..».", "Corregidas.", "Baja"],
    ["18", "Formato APA 7", "Encabezados en MAYÚSCULAS, «Capítulo I» frente a «Capítulo 2», sin tablas numeradas, citas con 5 autores, «&» en el texto en español.", "Niveles de encabezado APA, tablas APA (número en negritas, título en cursivas), «et al.» desde la primera cita, «y».", "Baja"],
], widths=[0.8, 3.2, 7.2, 9.6, 1.8], size=9)

heading(doc, "5. Verificación de referencias", 2)
para(doc, "Cada referencia se verificó en bases académicas y sitios de editoriales (DOI, SciELO, Redalyc, Cambridge, Nature, "
          "Frontiers, Emerald, DOAJ y repositorios institucionales). «No localizada» significa que no se encontró registro con el "
          "título, los autores, la revista o el año indicados. No prueba de forma absoluta que la obra no exista, pero sí que no se "
          "puede citar con esos datos. Si tienes el PDF de alguna obra no localizada, conserva sus datos reales y restáurala.",
     indent=False)
table(doc, 4, "Estado de las referencias del protocolo original", ["Referencia original", "Estado", "Problema", "Acción"], [
    ["Betancur-Chicué y García-Valcárcel (2023), RIED", "Verificada", "Puntuación; páginas «201-222-»; nombre de la segunda autora incompleto en el texto", "Corregida"],
    ["Bermúdez y Tucker (2003), «FAO»", "Verificada con datos erróneos", "Es de *Cadernos de Saúde Pública* 19(Supl. 1), no de la FAO; la URL era de otro documento", "Corregida"],
    ["Leong et al. (2021)", "Verificada", "—", "Formato"],
    ["Luo et al. (2025)", "Verificada con datos erróneos", "Son dos autores (Luo y Li); título incompleto; sin volumen", "Corregida"],
    ["Silva et al. (2025)", "Verificada (título, revista, DOI)", "Confirmar las iniciales de los autores", "==Verificar=="],
    ["Pérez-Rodrigo y Aranceta (2001)", "Verificada", "DOI roto («doi.org.10…»)", "Corregida"],
    ["Pérez-Rodrigo y Aranceta (2011) y (2013)", "No localizadas", "Solo aparecen en el texto; no hay obra con esos años", "Sustituidas por (2003), EJCN"],
    ["López Hernández et al. (2024)", "Verificada (índice de la revista)", "Inicia en la p. 8, no en la 1; «&»; «Arguelles»", "Corregida; ==página final=="],
    ["NOM-043-SSA2-2012", "Verificada", "Año de publicación en el DOF: 2013; título incompleto", "Corregida"],
    ["SSA, INSP y UNICEF (2025), Guías Alimentarias", "Verificada (2.ª ed. 2025)", "Citada en el texto, ausente de la lista", "Agregada; ==URL=="],
    ["ONU (2015)", "Verificada", "Sigla duplicada en la referencia", "Corregida"],
    ["UNESCO (2017)", "Verificada", "—", "Formato"],
    ["BID (s. f.)", "Página institucional", "—", "Formato"],
    ["INSP (2020), ENSANUT 2020, «37 %»", "Cifra no verificada", "La cifra no se encontró en la ENSANUT 2020 COVID-19", "Sustituida por Shamah-Levy et al. (2023): 37.3 %"],
    ["SEP (2021), La Nueva Escuela Mexicana", "Insuficiente", "Remite a un portal general; no sustenta la afirmación", "Eliminada"],
    ["FAO (2013), Guía de EAN para escuelas", "No localizada", "La URL es una página web general, no la guía citada", "Sustituida por FAO (2020)"],
    ["García et al. (2019), Jóvenes en la Ciencia", "Datos erróneos", "Según tu Zotero, es Aguirre Vargas (2019), con escolares de 3.º y 4.º de primaria (n = 24); la URL estaba truncada (313 en lugar de 3137)", "Restaurada como Aguirre Vargas (2019)"],
    ["Gómez (2008), ResearchGate", "Autoría errónea", "Según tu Zotero, la obra es de Olivares y Snel (1998), FAO, Alimentación, Nutrición y Agricultura n.º 22", "Restaurada como Olivares y Snel (1998)"],
    ["Morales y Salazar (2013), CORE", "No localizada", "Título, autores y «15 estudios de caso» no encontrados", "Sustituida por Shamah-Levy et al. (2017)"],
    ["Díaz y Ramírez (2020), Rev. Mex. de Educación y Salud", "No localizada", "No se encontró; su descripción (3.º y 4.º grado) coincide con Aguirre Vargas (2019)", "Sustituida por Shamah-Levy et al. (2017) y Aguirre Vargas (2019)"],
    ["Hernández y Gómez (2018), Rev. Latinoam. de Salud Comunitaria", "No localizada", "No se encontraron ni el artículo ni la revista", "Sustituida por Martínez Villagrana (2023)"],
    ["Mendoza y Ruiz (2021), Educación y Desarrollo", "No localizada", "No se encontró el artículo", "Sustituida por Cruz Lara y Macossay Moreno (2022)"],
    ["Surbakti et al. (2024)", "No verificada", "Revista y DOI sin registro; autores duplicados en la referencia", "Sustituida por Sweller et al. (2019)"],
    ["Mirzaei (2020)", "No localizada", "Citada en el texto, ausente de la lista", "Sustituida por Leong et al. (2021)"],
    ["Vaca-Escobar (2024)", "No localizada", "Citada en el texto, ausente de la lista", "Sustituida por Mayer (2009)"],
    ["López González (2024)", "No localizada", "Citada en el texto, ausente de la lista", "Eliminada (la definición se sostiene con Betancur-Chicué)"],
    ["Jasinski (2022)", "No localizada", "Citada en el texto, ausente de la lista", "Sustituida por Kelley y Whatson (2013)"],
    ["Aranceta (2012)", "No localizada", "Citada en el texto, ausente de la lista", "Sustituida por Shamah-Levy et al. (2023)"],
    ["Ubillus-Sambachi et al. (2025)", "Verificada", "Citada en el texto, ausente de la lista", "Agregada; ==autores=="],
    ["UNESCO (2023)", "Verificada (Informe GEM 2023)", "Citada en el texto, ausente de la lista", "Agregada; ==URL=="],
    ["Sweller (1988); Mayer (2009); Siemens (2005); Branch (2009)", "Verificadas", "Citadas en el texto, ausentes de la lista", "Agregadas"],
], widths=[6.5, 3.8, 7.7, 5], size=9)
para(doc, "**Referencias nuevas** (todas verificadas): Aiken (1985); Campbell y Stanley (1966); Cohen (1988); Cruz Lara y Macossay Moreno (2022); "
          "Escobar-Pérez y Cuervo-Martínez (2008); FAO (2020); González Ruiz et al. (2024, opcional); Hernández-Sampieri y Mendoza "
          "Torres (2018); Kelley y Whatson (2013); Kuder y Richardson (1937); Martínez Villagrana (2023); Penfield y Giacobbi (2004); "
          "Pérez-Rodrigo y Aranceta (2003); y, desde el Zotero de la autora, Aguirre Vargas (2019), Lytle y Achterberg (1995), Olivares y Snel (1998), Reséndiz Martínez (2023), Ríos-Castillo et al. (2024), Story et al. (2006), Termes Escalé et al. (2020), Barquera et al. (2001) y Mojtahedzadeh et al. (2021); Haladyna et al. (2002); SEP (2024), programa sintético de la Fase 5; Reglamento de la LGS en Materia de Investigación (1987); Shamah-Levy et al. (2017, "
          "2023); Sweller et al. (2019).", indent=False)

heading(doc, "6. Observaciones al instrumento (para considerar junto con el juicio de expertos)", 2)
table(doc, 5, "Observaciones a reactivos específicos", ["Reactivo", "Observación"], [
    ["1 y 6", "Ambos miden «variedad». Diferenciar (concepto y razón) o sustituir uno."],
    ["3, 19–24", "Unificar «agua natural» con «agua simple», el término de las Guías."],
    ["5", "Evitar «colación/snack»: usar un solo término, p. ej., «refrigerio escolar»."],
    ["6 y 14", "Distractores poco plausibles (6B, 14C). El análisis de reactivos del piloto lo confirmará."],
    ["11", "Que el juez nutriólogo confirme el grupo del aguacate según las Guías 2025."],
    ["12", "«Respectivamente» puede ser difícil a los 10 y 11 años."],
    ["14", "El enunciado no está formulado como pregunta. Propuesta: «¿Cuál opción describe una diferencia entre un cereal integral y uno refinado?»."],
    ["15 y 18", "Miden el mismo contenido (frijol = proteína vegetal). Sustituir el 18 por otro aporte."],
    ["20 y 22", "Misma lógica y misma respuesta (agregar ensalada). Sustituir uno."],
    ["21", "Nombra a cuatro niños, pero las opciones no se asignan a ninguno."],
    ["General", "La distribución de las claves es equilibrada (A = 5, B = 7, C = 6, D = 6). Preparar una versión para estudiantes sin claves y con letra de mayor tamaño."],
], widths=[3, 20])

heading(doc, "7. Pendientes para la autora", 2)
bullets(doc, [
    "Leer cada fuente nueva y confirmar que la paráfrasis es fiel. Esto es obligatorio antes de entregar.",
    "Completar los datos resaltados: página final de López Hernández et al.; autores de Ubillus-Sambachi et al. (2025); revista de Martínez Villagrana (2023); URL de las Guías 2025 y de UNESCO (2023); páginas de Olivares y Snel (1998); apellidos de Aguirre Vargas (2019); tipo de tesis de Reséndiz Martínez (2023); objetivo, método y resultados de Ríos-Castillo et al. (2024), Termes Escalé et al. (2020) y Aguirre Vargas (2019).",
    "Valorar con el asesor si se agrega un grupo de comparación, lo que fortalecería la denominación de cuasiexperimental.",
    "Definir la modalidad (plataforma en línea con acceso en el aula), las herramientas digitales, y confirmar con la escuela las fechas tentativas (inicio del taller el 25 de enero de 2027).",
    "Registrar a los jueces expertos (incluida la Dra. Verónica Torres, si acepta) y calcular la V de Aiken.",
    "Elaborar los formatos de consentimiento y asentimiento (Anexo B) y verificar si se requiere dictamen de un comité de ética.",
    "En Word: actualizar la tabla de contenido (clic derecho → Actualizar campo) y revisar los comentarios uno por uno, aceptándolos o resolviéndolos.",
    "Pasar el resumen de 150 palabras a la plantilla oficial del CITEEA.",
])

heading(doc, "Referencias de este informe", 2)
for r in [
    "Campbell, D. T., y Stanley, J. C. (1966). *Experimental and quasi-experimental designs for research*. Rand McNally.",
    "Hernández-Sampieri, R., y Mendoza Torres, C. P. (2018). *Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta*. McGraw-Hill Interamericana.",
]:
    ref(doc, r)
doc.save(OUT)
print("OK", OUT)
