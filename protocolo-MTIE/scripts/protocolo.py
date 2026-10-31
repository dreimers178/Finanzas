"""Genera el protocolo corregido (APA 7) con comentarios de revisión."""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))
from apa import *  # noqa

OUT = pathlib.Path(__file__).resolve().parents[1] / "Protocolo_Reimers_corregido_APA7.docx"
TALLER = "taller educativo en línea basado en microaprendizaje"

doc = new_document()

# ---------------------------------------------------------------- Portada
for t, b in [("UNIVERSIDAD AUTÓNOMA DE ZACATECAS", True), ("«FRANCISCO GARCÍA SALINAS»", True), ("", False),
             ("UNIDAD ACADÉMICA DE DOCENCIA SUPERIOR", True), ("MAESTRÍA EN TECNOLOGÍA INFORMÁTICA EDUCATIVA", True),
             ("", False)]:
    para(doc, f"**{t}**" if b and t else t, indent=False, align="center")
para(doc, "**Nutrición en el aula: taller educativo basado en microaprendizaje en línea para niños de 5.º de primaria**",
     indent=False, align="center",
     c="Título homologado con el resumen aceptado en el CITEEA y con la observación del asesor: el título nombra la propuesta "
       "como «taller educativo» (término único en todo el documento). Se corrigió «de 5º» (mayúsculas inconsistentes y "
       "ordinal sin punto: en español se escribe 5.º). Opción alternativa, si el asesor prefiere enfatizar el resultado: "
       "«Taller educativo en línea basado en microaprendizaje para incrementar el conocimiento nutricional de estudiantes de 5.º de primaria».")
para(doc, "", indent=False)
para(doc, "Presenta:", indent=False, align="center")
para(doc, "L.N.C.A. Anakaren Ernestina Reimers Reyes", indent=False, align="center")
para(doc, "Director:", indent=False, align="center")
para(doc, "Dr. Eduardo Rivera Arteaga", indent=False, align="center")
para(doc, "Codirectores:", indent=False, align="center")
para(doc, "==‹Nombre de codirectores›==", indent=False, align="center", c="Completar o eliminar si no aplica.")
para(doc, "", indent=False)
para(doc, "Zacatecas, Zac., ==noviembre de 2026==", indent=False, align="center",
     c="Actualizar fecha de la versión (el original decía «Mayo, 2026»; en español los meses van en minúscula).")
page_break(doc)

for t in ("Oficio Responsable del Programa de la MTIE", "Oficio de la Responsable del Departamento Escolar Central",
          "Agradecimientos", "Dedicatoria"):
    heading(doc, t, 1)
    para(doc, "==[Sección institucional pendiente]==")
page_break(doc)

# ---------------------------------------------------------------- Resumen
heading(doc, "Resumen", 1)
para(doc, "La educación nutricional durante la infancia es una herramienta para favorecer la adquisición de conocimientos "
          "relacionados con una alimentación saludable. El objetivo de esta investigación es incrementar el conocimiento "
          f"nutricional de estudiantes de quinto grado de primaria mediante la implementación de un {TALLER}. "
          "El taller se sustenta en los principios del microaprendizaje, la teoría de la carga cognitiva y la teoría del "
          "aprendizaje multimedia; sus contenidos se estructuran con base en el Plato del Bien Comer Saludable y Sostenible y en las "
          "Guías Alimentarias Saludables y Sostenibles para la Población Mexicana. El diseño instruccional sigue el modelo ADDIE. "
          "Se desarrolla una investigación aplicada, con enfoque cuantitativo y diseño cuasiexperimental de preprueba y posprueba con "
          "un solo grupo. La muestra estará integrada por 25 estudiantes de quinto grado del Instituto Educativo Ammadeus "
          "(Guadalupe, Zacatecas), seleccionados mediante muestreo no probabilístico por conveniencia. Se aplicará un cuestionario de "
          "24 reactivos de opción múltiple, validado mediante juicio de expertos (V de Aiken) y prueba piloto (KR-20), antes y después "
          "del taller; las puntuaciones se compararán con la prueba *t* de Student para muestras relacionadas o con la prueba de "
          "rangos con signo de Wilcoxon, según la distribución de los datos. Actualmente el instrumento se encuentra en proceso de "
          "validación de contenido.",
     indent=False,
     c="APA 7: el resumen va en un solo párrafo, sin sangría y con un máximo de 250 palabras. Se unificó el término "
       "«taller educativo» y se agregó cómo se analizarán los datos, que es lo que permite responder a la hipótesis. "
       "La versión de 150 palabras para el congreso está en el informe de revisión.")
para(doc, "*Palabras clave:* microaprendizaje, educación nutricional, educación primaria",
     c="La lista de cotejo del CITEEA indicó que tres palabras clave son suficientes. Se eliminó «conocimiento nutricional» "
       "porque se repite con «educación nutricional». En APA 7 «Palabras clave:» va en cursiva y con sangría.")
page_break(doc)

heading(doc, "Tabla de contenido", 1)
add_toc(doc)
page_break(doc)

# ================================================================ CAPÍTULO 1
heading(doc, "Capítulo 1. Introducción", 1,
        c="Se uniformó la numeración de capítulos con números arábigos (el original mezclaba «Capítulo I» y «Capítulo 2»).")
heading(doc, "1.1 Antecedentes", 2,
        c="CAMBIO IMPORTANTE. Cinco antecedentes del original no pudieron localizarse con los datos citados en ninguna base de datos "
          "ni repositorio (García et al., 2019; Morales y Salazar, 2013; Díaz y Ramírez, 2020; Hernández y Gómez, 2018; Mendoza y "
          "Ruiz, 2021). Un sexto, «Gómez (2008)», es en realidad de Olivares y Snel (1998), según tu Zotero, y se restauró con su "
          "autoría correcta. Con tu exportación de Zotero se identificó además que «García et al. (2019)» es en realidad "
          "Aguirre Vargas (2019), que también se restauró. Los demás se sustituyeron por estudios verificados. Citar trabajos inexistentes o mal "
          "atribuidos es una falta grave de integridad académica, por lo que se sustituyeron por estudios verificados con DOI o URL "
          "institucional. Se organizan de lo internacional a lo nacional y a lo local. Detalle en el informe de verificación de "
          "referencias. Antes de entregar, LEE cada fuente nueva y confirma que la paráfrasis es fiel.")
para(doc, "La revisión de antecedentes se organizó en tres planos (internacional, nacional y local) y en dos líneas temáticas: "
          "la educación alimentaria y nutricional en el entorno escolar y el microaprendizaje como estrategia educativa mediada "
          "por tecnología.")
heading(doc, "Antecedentes internacionales", 3)
para(doc, "Pérez-Rodrigo y Aranceta (2001), en el artículo «School-based nutrition education: Lessons learned and new perspectives», "
          "analizaron la experiencia acumulada en programas de educación nutricional en la escuela. Concluyeron que las intervenciones "
          "con mejores resultados adoptan el enfoque de escuelas promotoras de salud, integran varios componentes (currículo, entorno "
          "alimentario escolar, familia y comunidad) y no se limitan a transmitir información, sino que desarrollan habilidades "
          "relacionadas con la alimentación. Este trabajo sustenta la conveniencia de integrar la educación nutricional a las "
          "actividades escolares y de evaluarla de forma sistemática.",
     c="Se eliminaron los datos metodológicos del original («metodología cualitativa», «12 programas», «entrevistas») porque "
       "no corresponden al artículo verificado. Confirma la paráfrasis en el texto completo (DOI 10.1079/PHN2000108).")
para(doc, "Bermudez y Tucker (2003), en «Trends in dietary patterns of Latin American populations», publicado en *Cadernos de "
          "Saúde Pública*, analizaron los cambios en los patrones alimentarios de la región a partir, principalmente, de las hojas de "
          "balance de alimentos de la FAO y de estudios nacionales. Identificaron un aumento en el consumo de grasas, productos de "
          "origen animal y azúcares, así como una disminución del consumo de cereales, frutas y algunas verduras. Concluyeron que "
          "esta transición alimentaria presiona a sistemas de salud que aún enfrentan desnutrición y enfermedades infecciosas, y "
          "propusieron la educación de la población sobre alimentación saludable como una de las respuestas prioritarias.",
     c="La referencia original atribuía este artículo a la FAO con una URL (fao.org/4/X0051T) que en realidad corresponde a un "
       "documento de la FAO de 1998 sobre educación nutricional en primarias (Olivares y Snel, 1998, que se agrega como antecedente aparte). Se corrigió la fuente: Cadernos de Saúde Pública, "
       "19(Supl. 1), S87–S99.")
para(doc, "Olivares y Snel (1998), en la serie *Alimentación, Nutrición y Agricultura* de la FAO, presentaron los resultados de "
          "una encuesta elaborada por la FAO y el Centro de Nutrición de los Países Bajos sobre la situación de la educación en "
          "nutrición en escuelas primarias de países en desarrollo. Documentaron las políticas oficiales de los ministerios, el "
          "tiempo dedicado al tema, que entre los países que respondieron iba de 3 a 32 horas al año, el papel de los docentes como "
          "principales responsables de impartirlo y los principales obstáculos para su aplicación. Este trabajo evidencia que la "
          "presencia de contenidos nutricionales en el currículo no garantiza su enseñanza sistemática.",
     c="RESTAURADO a partir de tu biblioteca de Zotero. El protocolo original lo citaba como «Gómez (2008)», pero el documento "
       "es de Olivares y Snel (1998), FAO, Alimentación, Nutrición y Agricultura n.º 22. Se ajustó la descripción a lo que el "
       "documento reporta (encuesta FAO y Centro de Nutrición de los Países Bajos). Verifica las cifras en el PDF.")
para(doc, "En el ámbito del microaprendizaje, Betancur-Chicué y García-Valcárcel Muñoz-Repiso (2023) realizaron una revisión "
          "sistemática de la literatura publicada entre 2018 y 2021 en Scopus y Web of Science. Señalaron como ventajas del "
          "microaprendizaje que ofrece una vía ágil de formación y que puede reducir la carga cognitiva, y como principal desventaja "
          "su menor aporte en temas o habilidades complejas. Las áreas que más lo han explorado son la salud, la educación y la "
          "ingeniería, y el diseño predominante se estructura en introducción, microcontenido y preguntas, con un uso destacado del "
          "video educativo. Por su parte, Silva et al. (2025), en una revisión sistemática sobre el microaprendizaje en la educación "
          "básica, identificaron que la organización del contenido en unidades breves y focalizadas favorece la motivación, la "
          "atención y la retención del conocimiento en estudiantes de este nivel.")
para(doc, "Luo y Li (2025) evaluaron el efecto de intervenciones de microaprendizaje en el desarrollo de habilidades blandas "
          "(trabajo en equipo, liderazgo, comunicación, gestión del tiempo e inteligencia emocional) en estudiantes universitarios de "
          "cuatro áreas disciplinares, y reportaron mejoras medibles que variaron entre disciplinas. Aunque la población es distinta, "
          "el estudio aporta evidencia sobre la eficacia del microaprendizaje y sobre la necesidad de adaptarlo a las características "
          "del grupo destinatario.",
     c="Antes se citaba como «Luo, H., et al.» con un título incompleto. El artículo tiene dos autores (Luo y Li), por lo que en el "
       "texto debe citarse «Luo y Li (2025)», no «Luo et al.».")
para(doc, "En el ámbito iberoamericano, Ríos-Castillo et al. (2024) realizaron una revisión narrativa de las estrategias "
          "innovadoras de educación alimentaria y nutricional empleadas en países hispanos para combatir el exceso de peso y la "
          "obesidad en niños en edad escolar, y Termes Escalé et al. (2020) documentaron la educación nutricional de escolares "
          "mediante el programa Nutriplato. ==[Completa con el objetivo, el método y los hallazgos principales de cada estudio.]==",
     c="NUEVO, desde tu Zotero. Ambas referencias se verificaron (DOI 10.4067/s0717-75182024000400333 y 10.20960/nh.03357), "
       "pero tu exportación RIS no incluye sus resúmenes. Completa el texto resaltado a partir de los PDF y no agregues datos que "
       "no estén en ellos.")
heading(doc, "Antecedentes nacionales", 3)
para(doc, "Shamah-Levy et al. (2017) evaluaron el efecto del programa SaludArte en los componentes de alimentación y nutrición de "
          "escolares de la Ciudad de México. Mediante un estudio comparativo de tres cohortes (2013-2015), con 1,620 escolares de 144 "
          "primarias, encontraron efectos atribuibles al programa en los conocimientos en nutrición (*p* = .003), la higiene personal "
          "y la actividad física. El estudio muestra que una intervención escolar estructurada puede incrementar el conocimiento "
          "nutricional de estudiantes de primaria, que es la variable dependiente de la presente investigación.",
     c="Este estudio sustituye a «Díaz y Ramírez (2020)» y a «Morales y Salazar (2013)», que no pudieron verificarse. "
       "DOI 10.21149/8116.")
para(doc, "Cruz Lara y Macossay Moreno (2022) implementaron una intervención de educación nutricional en línea con 30 estudiantes de una primaria "
          "pública del municipio de Veracruz, con un diseño cuasiexperimental de grupo control y grupo de intervención y mediciones "
          "antes y después. El grupo de intervención mejoró significativamente sus conocimientos sobre alimentación y nutrición "
          "(*p* = .04) y mostró cambios en el consumo de agua, leguminosas y lácteos, mientras que el grupo control no presentó "
          "mejoras significativas. No se observaron cambios significativos en el estado nutricional. Este antecedente es "
          "particularmente pertinente porque combina educación nutricional, modalidad en línea y población de educación primaria.",
     c="Sustituye a «Mendoza y Ruiz (2021)», que no pudo verificarse. Es el antecedente más cercano a tu estudio (en línea + "
       "primaria + conocimiento). Revisa el texto completo para precisar los grados escolares y el número de sesiones.")
para(doc, "Aguirre Vargas (2019) evaluó el conocimiento sobre el Plato del Bien Comer de escolares de tercero y cuarto grado de "
          "la Escuela Primaria Ignacio Zaragoza antes y después de una intervención educativa en nutrición. De los 33 escolares que "
          "aceptaron participar, con consentimiento de sus padres y asentimiento propio, 24 concluyeron el estudio. Las sesiones "
          "abordaron la alimentación en la etapa escolar con el Plato del Bien Comer como herramienta gráfica, y la evaluación tuvo "
          "una parte teórica y una práctica. ==[Completa con los resultados reportados en el artículo.]== Este antecedente comparte "
          "con el presente estudio el diseño de preprueba y posprueba, la población de primaria y el contenido del Plato del Bien "
          "Comer.",
     c="RESTAURADO desde tu Zotero. Es el artículo de Jóvenes en la Ciencia que el protocolo original citaba como «García et al. "
       "(2019)», con URL .../article/view/313 (truncada; la correcta termina en 3137) y con datos que no corresponden (adolescentes "
       "de secundaria en León, 30 alumnos). Es probable que también sea la fuente real de «Díaz y Ramírez (2020)», que hablaba de "
       "3.º y 4.º grado. VERIFICA EL NOMBRE: Zotero lo registró como «Vargas, Miriam Carolina Aguirre», pero lo más probable es "
       "que los apellidos sean Aguirre Vargas. Confírmalo en el PDF.")
para(doc, "En la línea de los recursos digitales, Reséndiz Martínez (2023) exploró, mediante un estudio cualitativo con grupos "
          "focales, los desafíos y alcances del material digital para la educación alimentaria de escolares de 9 a 11 años y sus "
          "familias. Mostró videos animados, infografías animadas y un videojuego a tres grupos de 12 escolares. Los niños mostraron "
          "agrado por el material y motivación para hacer cambios, pero la principal barrera fue la resistencia de los miembros de su "
          "familia; los padres señalaron barreras económicas y falta de tiempo. Estos hallazgos respaldan el uso de recursos "
          "digitales breves y visuales con escolares, y advierten sobre la influencia del entorno familiar.",
     c="NUEVO, desde tu Zotero (tesis del repositorio de la UAQ). Zotero registró la autora como «Martínez, Laura Reséndiz»; lo "
       "correcto en APA es Reséndiz Martínez, L. Confirma el grado de la tesis (maestría o licenciatura) para la referencia.")
para(doc, "En cuanto a la magnitud del problema, los resultados de la Encuesta Nacional de Salud y Nutrición Continua 2020-2022 "
          "muestran que el 37.3 % de la población escolar mexicana presenta sobrepeso (19.2 %) u obesidad (18.1 %), y que la "
          "prevalencia de obesidad aumentó poco más de cinco puntos porcentuales entre 2006 y 2020-2022 (Shamah-Levy et al., 2023).",
     c="El original citaba «ENSANUT (2020): aproximadamente 37 %». No fue posible verificar esa cifra en la ENSANUT 2020 sobre "
       "COVID-19. El dato verificado (37.3 %) proviene de la Ensanut Continua 2020-2022, publicada en Salud Pública de México "
       "(DOI 10.21149/14762). En APA 7 la encuesta se cita por sus autores, no por la sigla «ENSANUT».")
heading(doc, "Antecedentes locales", 3)
para(doc, "En Zacatecas, López Hernández et al. (2024) estudiaron a escolares de una primaria del municipio de Guadalupe y "
          "reportaron que el 45.9 % presentaba sobrepeso u obesidad y que el 66.4 % mostraba conocimientos insuficientes en nutrición. "
          "De manera complementaria, Martínez Villagrana (2023) documentó en una primaria rural de Monte Escobedo que el 46 % del "
          "alumnado presentaba sobrepeso, obesidad o desnutrición, y planteó que la escuela, como espacio físico y psicosocial, es un "
          "entorno propicio para impulsar una pedagogía para la vida saludable en la que participen docentes, profesionales de la "
          "salud, familias y comunidad.",
     c="APA 7: con tres o más autores se usa «et al.» desde la primera cita. El original listaba los cinco apellidos y además "
       "escribía «Arguelles» sin diéresis. Martínez Villagrana (2023) sustituye a «Hernández y Gómez (2018)», que no se encontró. "
       "Verifica en el artículo de López Hernández et al. que el municipio sea Guadalupe.")
para(doc, "En síntesis, los antecedentes muestran que las intervenciones escolares de educación nutricional pueden incrementar el "
          "conocimiento de los estudiantes y que el microaprendizaje es una estrategia prometedora en entornos digitales. Sin embargo, "
          "la evidencia que articula ambos campos en la educación primaria mexicana es escasa, en particular en Zacatecas, donde la "
          "prevalencia de sobrepeso y obesidad y el déficit de conocimientos nutricionales son elevados. El presente proyecto se "
          "propone atender ese vacío mediante un " + TALLER + ".",
     c="Se agregó un párrafo de cierre que identifica el vacío de conocimiento. Es un requisito de rigor: los antecedentes deben "
       "desembocar en la necesidad del estudio. Se eliminó la mención a UNESCO (2017), que no es un antecedente de investigación "
       "y ya aparece en el planteamiento del problema.")

heading(doc, "1.2 Marco contextual", 2)
para(doc, "El presente proyecto se llevará a cabo en el Instituto Educativo Ammadeus, institución educativa privada ubicada en "
          "Vialidad Vetagrande 468, Zona Conurbada, Guadalupe, Zacatecas, México. El instituto ofrece educación desde preescolar hasta "
          "preparatoria, con un modelo bilingüe que prioriza tanto el desarrollo académico como el crecimiento personal de sus "
          "estudiantes. Su misión es «el desarrollo integral de alumnos bilingües, con una gran autoestima, a través de un sistema "
          "educativo cálido y de excelencia».",
     c="Se reorganizó el marco contextual. El original tenía un subapartado «1.2.1 Misión» que en realidad describía la matrícula y "
       "la infraestructura, mientras que la misión aparecía al final y suelta. Se eliminaron frases valorativas sin sustento "
       "(«reconocida por su compromiso…», «condiciones óptimas»).")
heading(doc, "1.2.1 Características de la institución", 3)
para(doc, "Actualmente el instituto cuenta con 652 alumnos inscritos. El nivel primaria tiene 385 alumnos, de 6 a 12 años, "
          "distribuidos en tres grupos mixtos por grado. Este nivel dispone de 18 salones equipados con proyector, pintarrón, "
          "computadora y conexión a internet; un salón de computación; salón de usos múltiples; dos patios de juegos; una unidad "
          "deportiva con cancha de futbol; un auditorio con cancha de usos múltiples y una cafetería ubicada en el primer piso. En "
          "cada uno de los tres pisos del área de primaria hay sanitarios y oficinas de psicopedagogía y de prefectura.",
     c="Indica la fuente de estos datos (p. ej., «Dirección del Instituto Educativo Ammadeus, comunicación personal, fecha»). "
       "En APA 7 las comunicaciones personales se citan en el texto, pero no van en la lista de referencias.")
heading(doc, "1.2.2 Grupo participante", 3)
para(doc, "El estudio se realizará con el grupo 5.º «B» de primaria, integrado por 25 estudiantes (niñas y niños) de entre 10 y 11 "
          "años que cursan contenidos de salud y nutrición dentro del campo formativo correspondiente. El grupo se eligió por la "
          "disposición institucional para colaborar y por la accesibilidad para implementar la intervención. El grupo cuenta con "
          "acceso al aula digital, conectividad y acompañamiento docente, condiciones necesarias para un " + TALLER + ".",
     c="CORRECCIÓN METODOLÓGICA. El original afirmaba que el grupo se eligió «por su representatividad dentro del nivel educativo». "
       "Una muestra no probabilística por conveniencia NO es representativa, y afirmarlo contradice el apartado de muestreo. "
       "Además, en la NEM la asignatura «Ciencias Naturales» se integra en el campo formativo «Saberes y pensamiento científico». "
       "Verifica cómo la nombra la escuela.")
para(doc, "El taller se impartirá en español, tendrá una duración de cuatro semanas y se desarrollará con cápsulas digitales "
          "breves, dinámicas y accesibles. La elección del formato digital responde a su posibilidad de escalamiento y de adaptación "
          "a entornos híbridos o virtuales, en coherencia con los Objetivos de Desarrollo Sostenible 3 (salud y bienestar) y 4 "
          "(educación de calidad) de la Agenda 2030 (Organización de las Naciones Unidas [ONU], 2015).",
     c="APA 7: la sigla se introduce entre corchetes en la primera cita parentética: (Organización de las Naciones Unidas [ONU], "
       "2015), y después se usa (ONU, 2015). Se corrigió además el punto final que faltaba.")

heading(doc, "1.3 Planteamiento del problema", 2,
        c="El original tenía dos encabezados idénticos («1.3. Planteamiento del problema» y «1.3.1 Planteamiento del problema»). "
          "Se fusionaron: primero el contexto del problema y después su delimitación.")
para(doc, "El sobrepeso y la obesidad infantil son uno de los principales retos de salud pública en México. La Agenda 2030 "
          "establece, en el Objetivo de Desarrollo Sostenible 3, el compromiso de garantizar una vida sana y promover el bienestar "
          "de todas las personas a todas las edades (ONU, 2015), y la Organización de las Naciones Unidas para la Educación, la "
          "Ciencia y la Cultura (UNESCO, 2017) plantea que la educación debe proporcionar a los estudiantes conocimientos y "
          "competencias para llevar estilos de vida saludables y sostenibles. A su vez, el Banco Interamericano de Desarrollo "
          "(BID, s. f.) sostiene que la tecnología y la innovación son herramientas para cerrar brechas educativas y lograr "
          "aprendizajes pertinentes.",
     c="El original afirmaba que la ONU «advierte que la adopción de hábitos saludables desde la infancia es clave…». La Agenda "
       "2030 no contiene esa frase. Se reformuló con lo que el documento sí establece (ODS 3). Verifica la paráfrasis de UNESCO "
       "(2017) en el documento original.")
para(doc, "En el contexto nacional, el 37.3 % de la población escolar presenta sobrepeso u obesidad (Shamah-Levy et al., 2023). "
          "En Zacatecas, López Hernández et al. (2024) reportaron en una primaria de Guadalupe que el 45.9 % de los escolares "
          "presentaba sobrepeso u obesidad y que el 66.4 % tenía conocimientos insuficientes en nutrición. Estos datos evidencian la "
          "pertinencia de fortalecer, desde la escuela, el conocimiento sobre alimentación saludable.",
     c="Se eliminó la cita «SEP (2021)», que remitía a la página general de educación básica y no sustenta la afirmación de que "
       "«las prácticas de educación nutricional son limitadas y poco sistemáticas». Si quieres conservar esa idea, cita el "
       "documento específico (programa sintético de la fase 5 de la NEM o un estudio que lo documente).")
para(doc, "El problema que aborda esta investigación es el insuficiente conocimiento nutricional de estudiantes de educación "
          "primaria y la falta de propuestas educativas sistemáticas, atractivas y apoyadas en recursos digitales que faciliten la "
          "comprensión de contenidos nutricionales acordes con las recomendaciones alimentarias mexicanas vigentes. El "
          "microaprendizaje, que organiza los contenidos en unidades breves, estructuradas e interactivas, representa una "
          "alternativa para atender este problema; sin embargo, su efecto sobre el conocimiento nutricional de estudiantes de "
          "primaria en el contexto zacatecano no ha sido evaluado. Por ello, se requiere diseñar, implementar y evaluar un " + TALLER +
          " para estudiantes de quinto grado del Instituto Educativo Ammadeus.",
     c="La delimitación ahora nombra el problema (déficit de conocimiento + ausencia de propuestas), el vacío (falta de evaluación "
       "en este contexto) y lo que hará el estudio, en coherencia con el objetivo general.")

heading(doc, "1.4 Preguntas de investigación", 2)
heading(doc, "1.4.1 Pregunta general", 3)
para(doc, "¿Se incrementa el nivel de conocimiento nutricional de los estudiantes de 5.º grado de primaria del Instituto Educativo "
          "Ammadeus después de su participación en un " + TALLER + "?",
     c="Se ajustó para que se corresponda exactamente con el objetivo general («incrementar») y con la hipótesis de investigación "
       "(H1 unilateral). «¿Existen diferencias…?» es bilateral y no coincidía con una H1 direccional («será mayor»).")
heading(doc, "1.4.2 Preguntas específicas", 3)
bullets(doc, [
    "¿Cuál es el nivel de conocimiento nutricional de los estudiantes de 5.º grado de primaria antes de la implementación del taller?",
    "¿Qué elementos digitales se pueden incluir en un " + TALLER + " para abordar contenidos de educación nutricional en estudiantes de 5.º grado de primaria?",
    "¿Qué diferencias se presentan en el nivel de conocimiento nutricional de los estudiantes, total y por dimensión, al comparar los resultados obtenidos antes y después de la implementación del taller?",
], numbered=True,
    c="La pregunta 2 ya atiende la observación del asesor de no iniciar con «cómo». La pregunta 3 cambia «cambios» por "
      "«diferencias… total y por dimensión» para que sea medible con el instrumento de cuatro dimensiones. Cada pregunta "
      "específica corresponde a un objetivo específico (ver Tabla 1).")

heading(doc, "1.5 Objetivos", 2)
heading(doc, "1.5.1 Objetivo general", 3)
para(doc, "Incrementar el conocimiento nutricional de los estudiantes de 5.º grado de primaria del Instituto Educativo Ammadeus "
          "mediante la implementación de un " + TALLER + ".",
     c="Atiende la observación del asesor: el objetivo general no debe ser «evaluar», y la prioridad es incrementar el "
       "conocimiento. Se usa la misma redacción del resumen aceptado en el congreso. Se eliminó «de cápsulas educativas… con el "
       "fin de promover…» porque eran dos finalidades en un solo objetivo.")
heading(doc, "1.5.2 Objetivos específicos", 3)
bullets(doc, [
    "Identificar el nivel inicial de conocimiento nutricional de los estudiantes de 5.º grado de primaria mediante la aplicación de una preprueba.",
    "Diseñar y desarrollar, con base en el modelo ADDIE, las cápsulas digitales y las actividades interactivas del " + TALLER + ", alineadas con las Guías Alimentarias Saludables y Sostenibles para la Población Mexicana.",
    "Comparar el nivel de conocimiento nutricional de los estudiantes, total y por dimensión, antes y después de la implementación del taller.",
], numbered=True,
    c="El objetivo 2 incorpora «desarrollar» y el modelo ADDIE, que es donde se ubica la pregunta 2 (elementos digitales). El "
      "objetivo 3 cambia «Evaluar los cambios» por «Comparar», que es la operación estadística real (preprueba frente a "
      "posprueba). Se eliminó el doble punto final («..»).")
table(doc, 1, "Matriz de congruencia del estudio",
      ["Pregunta", "Objetivo", "Hipótesis", "Técnica / análisis"],
      [["PG. ¿Se incrementa el conocimiento nutricional… después del taller?", "OG. Incrementar el conocimiento nutricional mediante el taller.", "H1 / H0", "Preprueba-posprueba; t pareada o Wilcoxon"],
       ["PE1. Nivel inicial", "OE1. Identificar el nivel inicial", "—", "Cuestionario (preprueba); estadística descriptiva"],
       ["PE2. Elementos digitales", "OE2. Diseñar y desarrollar el taller (ADDIE)", "—", "Diseño instruccional; carta descriptiva; juicio de expertos"],
       ["PE3. Diferencias total y por dimensión", "OE3. Comparar preprueba y posprueba", "H1 / H0", "t pareada o Wilcoxon; tamaño del efecto"]],
      widths=[4.2, 4.2, 2.2, 5.4],
      c="Se agregó una matriz de congruencia. Permite al comité comprobar de un vistazo la alineación entre preguntas, objetivos, "
        "hipótesis y análisis, que es el criterio central del rigor metodológico.")

heading(doc, "1.6 Hipótesis", 2)
para(doc, "**Hipótesis de investigación (H1).** La media del puntaje de conocimiento nutricional de los estudiantes de 5.º grado "
          "de primaria será significativamente mayor en la posprueba que en la preprueba, tras su participación en el " + TALLER + ".")
para(doc, "**Hipótesis nula (H0).** La media del puntaje de conocimiento nutricional en la posprueba será igual o menor que en la "
          "preprueba.",
     c="Con una H1 direccional («mayor»), la H0 debe ser su complemento («igual o menor»), no «no existirán diferencias». La "
       "prueba será unilateral con α = .05. Si la prueba final es Wilcoxon, sustituye «media» por «mediana».")

heading(doc, "1.7 Justificación", 2)
para(doc, "La presente investigación se justifica por su relevancia teórica, práctica, social, económica y metodológica dentro del "
          "campo de la Tecnología Informática Educativa y de la educación para la salud.")
heading(doc, "Justificación teórica", 3)
para(doc, "La educación alimentaria y nutricional es un componente fundamental para la promoción de hábitos saludables desde la "
          "infancia (Food and Agriculture Organization of the United Nations [FAO], 2020; Pérez-Rodrigo y Aranceta, 2003). La "
          "literatura señala que las intervenciones escolares más exitosas están integradas a la vida escolar y se acompañan de "
          "evaluación sistemática (Pérez-Rodrigo y Aranceta, 2001), y estudios mexicanos han documentado incrementos en el "
          "conocimiento nutricional de escolares tras intervenciones estructuradas (Cruz Lara y Macossay Moreno, 2022; Shamah-Levy et al., 2017). Por "
          "otra parte, el microaprendizaje, fundamentado en la teoría de la carga cognitiva y en la teoría del aprendizaje multimedia, "
          "ha mostrado potencial para favorecer la atención y la retención (Betancur-Chicué y García-Valcárcel Muñoz-Repiso, 2023; "
          "Silva et al., 2025). El estudio aporta evidencia sobre la articulación de ambos campos en educación primaria.",
     c="Se corrigieron citas: «Pérez-Rodrigo y Aranceta (2013)» y «(2011)» no existen con esos datos; el trabajo verificable de "
       "esos autores sobre educación nutricional escolar es de 2003 (European Journal of Clinical Nutrition). «Díaz y Ramírez "
       "(2020)» y «Morales y Salazar (2013)» se sustituyeron por estudios verificados. ENSANUT es un dato epidemiológico, no un "
       "estudio de programas escolares, por lo que se retiró de esta frase.")
heading(doc, "Justificación práctica", 3)
para(doc, "El proyecto responde a la necesidad de contar con recursos educativos estructurados y accesibles que contribuyan a "
          "mejorar el conocimiento nutricional de estudiantes de educación básica. El " + TALLER + " puede implementarse en "
          "diversos contextos escolares sin requerir grandes recursos materiales, aprovechando plataformas digitales y la "
          "infraestructura con la que cuenten las instituciones. Además, el modelo de cápsulas puede replicarse y adaptarse a otros "
          "temas y grados, lo que favorece una educación continua y flexible sin modificaciones extensas al currículo formal.",
     c="Se corrigieron «pretende resultar ser práctica» (redundante) y «contínua» (no lleva tilde).")
heading(doc, "Justificación económica", 3)
para(doc, "Fortalecer los conocimientos alimentarios en la infancia puede tener efectos económicos positivos a largo plazo, al "
          "contribuir a la prevención de enfermedades crónicas asociadas con la mala alimentación. Además, el uso de recursos "
          "digitales gratuitos o de bajo costo reduce la necesidad de materiales impresos y facilita que otras instituciones adopten "
          "la propuesta sin gastos adicionales.",
     c="Se moderó la redacción («puede contribuir» en lugar de «al reducir los costos») porque el estudio no mide costos ni "
       "conductas. Se eliminó «garantizar la equidad», que el estudio no puede sostener.")
heading(doc, "Justificación social", 3)
para(doc, "El sobrepeso y la obesidad afectan al 37.3 % de la población escolar en México (Shamah-Levy et al., 2023) y, en una "
          "primaria de Guadalupe, Zacatecas, al 45.9 % (López Hernández et al., 2024). Este panorama destaca la importancia de "
          "fortalecer, desde el ámbito educativo, los conocimientos sobre alimentación saludable. El proyecto busca contribuir a la "
          "educación nutricional de estudiantes de quinto grado y se vincula con los ODS 3 (salud y bienestar) y 4 (educación de "
          "calidad) (ONU, 2015).",
     c="El original citaba «López-Hernández et al.» con guion. El apellido compuesto se escribe sin guion (López Hernández), igual "
       "que en la lista de referencias.")
heading(doc, "Justificación metodológica", 3)
para(doc, "El estudio emplea un enfoque cuantitativo con un diseño cuasiexperimental de preprueba y posprueba con un solo grupo, "
          "que permite comparar el nivel de conocimiento nutricional antes y después del taller. Aporta además un cuestionario de 24 "
          "reactivos alineado con las Guías Alimentarias Saludables y Sostenibles para la Población Mexicana (Secretaría de Salud "
          "[SSA] et al., 2025), con evidencias de validez de contenido (V de Aiken) y de confiabilidad (KR-20), que podrá servir de "
          "referencia para otras investigaciones sobre educación nutricional mediada por tecnología.",
     c="La principal contribución metodológica es el instrumento validado. Se resalta en lugar de apoyarse solo en Pérez-Rodrigo "
       "y Aranceta.")

heading(doc, "1.8 Alcances y limitaciones", 2)
heading(doc, "Alcance temático", 3)
para(doc, "El estudio se centra en la educación alimentaria y nutricional de estudiantes de 5.º de primaria y en la medición de "
          "cambios en su nivel de conocimiento nutricional antes y después del taller. Los contenidos abordan alimentación "
          "saludable, grupos de alimentos del Plato del Bien Comer Saludable y Sostenible, aportes nutrimentales básicos y la "
          "aplicación de las recomendaciones alimentarias en situaciones cotidianas. No contempla la medición de conductas "
          "alimentarias ni de indicadores clínicos.",
     c="APA 7: los encabezados no van en MAYÚSCULAS sostenidas. Los temas se alinearon con las cuatro dimensiones del instrumento.")
heading(doc, "Alcance poblacional", 3)
para(doc, "La población la integran las niñas y los niños de 5.º grado de primaria del Instituto Educativo Ammadeus. La muestra "
          "corresponderá a los estudiantes del grupo seleccionado que cuenten con el consentimiento informado de sus madres, padres "
          "o tutores y que otorguen su asentimiento.")
heading(doc, "Alcance espacial", 3)
para(doc, "La intervención y la recolección de datos se realizarán en el Instituto Educativo Ammadeus, con la infraestructura "
          "tecnológica de la institución (aula digital, proyector, computadoras y conexión a internet). El taller es «en línea» porque "
          "sus cápsulas y actividades se alojan en una plataforma digital, a la que los estudiantes acceden ==desde el aula digital "
          "durante el horario escolar y, de manera opcional, desde casa==.",
     c="DECISIÓN PENDIENTE. El texto original decía «taller en línea» pero también que se implementará «en el aula» con proyector. "
       "Precisa la modalidad, ya que afecta el control de la exposición a la intervención. Recomendación: acceso en el aula con "
       "acompañamiento docente y registro de cápsulas completadas.")
heading(doc, "Alcance temporal", 3)
para(doc, "La intervención durará cuatro semanas. Se realizarán dos mediciones: una preprueba antes del taller y una posprueba "
          "inmediatamente después. La preprueba se aplicará del 18 al 22 de enero de 2027; el taller iniciará el "
          "lunes 25 de enero de 2027, en la semana de la conmemoración del Día del Nutriólogo en México (27 de enero), y concluirá el "
          "19 de febrero de 2027; la posprueba se aplicará del 22 al 26 de febrero de 2027 (fechas tentativas; Tabla 5). El estudio no "
          "contempla una medición de seguimiento.",
     c="Fecha actualizada por decisión de la autora: el taller inicia el 25 de enero de 2027, en la semana del Día del "
       "Nutriólogo (miércoles 27 de enero). Con dos cápsulas por semana, la C2 coincide con el 27 de enero y puede usarse como "
       "actividad de lanzamiento. Confirma las fechas con el calendario escolar.")
heading(doc, "Limitaciones", 3)
para(doc, "El estudio presenta limitaciones que deberán considerarse al interpretar los resultados. En primer lugar, al no contar "
          "con un grupo control, los cambios observados no podrán atribuirse exclusivamente al taller, pues podrían influir factores "
          "como la historia (p. ej., actividades escolares o mediáticas por el Día del Nutriólogo), la maduración, el efecto de "
          "aplicar dos veces el mismo instrumento (efecto de prueba) y la regresión a la media (Campbell y Stanley, 1966).",
     c="Se nombran las amenazas a la validez interna propias del diseño de un solo grupo (Campbell y Stanley, 1966). Es justo lo "
       "que un comité espera ver cuando no hay grupo control.")
para(doc, "En segundo lugar, la muestra se integra por estudiantes de un solo grupo y de una institución privada, seleccionados "
          "por conveniencia, lo que impide generalizar los resultados a otros grados, instituciones o contextos.")
para(doc, "El estudio se limitará a medir el nivel de conocimiento nutricional y no evaluará conductas o hábitos alimentarios ni "
          "indicadores clínicos como el índice de masa corporal; por lo tanto, un incremento en el conocimiento no permitirá afirmar "
          "que se produjeron cambios en los hábitos alimentarios de los participantes.",
     c="Se corrigieron erratas del original: «limkitará», «condicyas», «permitirpa», «enlos», «Por irra parte», «alis», "
       "«metodológicaquew».")
para(doc, "Por otra parte, la aplicación del cuestionario a niñas y niños de 10 y 11 años puede verse afectada por dificultades de "
          "comprensión lectora o de vocabulario y por la deseabilidad social. Asimismo, al contar únicamente con una medición "
          "inmediata posterior, no será posible determinar si los cambios se mantienen a mediano o largo plazo. Finalmente, la "
          "disponibilidad de internet y dispositivos, los tiempos escolares, la exposición efectiva a las cápsulas y las prácticas "
          "alimentarias familiares quedan parcialmente fuera del control de la intervención.")
page_break(doc)

# ================================================================ CAPÍTULO 2
heading(doc, "Capítulo 2. Fundamento teórico", 1,
        c="Se eliminó el texto de la plantilla («Para la composición del presente capítulo…», el ejemplo de «Motivación» y el "
          "encabezado «2.1 Tema 1»), que se había quedado en el documento.")
heading(doc, "2.1 Variable independiente: taller educativo en línea basado en microaprendizaje", 2,
        c="Se usa el mismo nombre de la variable en la tabla de contenido, el título del apartado y la operacionalización. El "
          "original alternaba entre «estrategia educativa basada en microaprendizaje», «taller» y «microaprendizaje».")
heading(doc, "2.1.1 Concepto y características del microaprendizaje", 3)
para(doc, "El microaprendizaje es una estrategia educativa basada en la presentación de contenidos organizados en unidades breves, "
          "precisas y estructuradas, orientadas a objetivos específicos, con el fin de facilitar la comprensión y la retención "
          "progresiva del conocimiento (Betancur-Chicué y García-Valcárcel Muñoz-Repiso, 2023). Desde una perspectiva pedagógica, "
          "organiza la información en fragmentos que permiten abordar un objetivo a la vez sin sobrecargar al estudiante, lo que "
          "resulta valioso en contextos que requieren adaptarse a los ritmos atencionales del alumnado.",
     c="Se eliminó la cita «López González (2024)», que no pudo localizarse. La definición queda sustentada en Betancur-Chicué y "
       "García-Valcárcel Muñoz-Repiso (2023).")
para(doc, "El interés académico y práctico por el microaprendizaje ha crecido de forma sostenida: Leong et al. (2021) documentaron "
          "que, desde 2006, el aumento de publicaciones sobre el tema ha corrido paralelo al incremento de búsquedas en internet "
          "relacionadas con él, lo que refleja una demanda creciente de formación en unidades breves. Este enfoque promueve un "
          "aprendizaje flexible, accesible y centrado en el estudiante, que avanza paso a paso mediante contenidos que responden a "
          "necesidades específicas. La incorporación de recursos multimedia, que combinan estímulos visuales, textuales y auditivos, "
          "puede potenciar la comprensión cuando se diseña de acuerdo con los principios del aprendizaje multimedia (Mayer, 2009).",
     c="Se sustituyeron «Mirzaei (2020)» y «Vaca-Escobar (2024)», que no pudieron localizarse, por Leong et al. (2021) y Mayer "
       "(2009), ambas verificadas y ya presentes en tu marco.")
para(doc, "Desde el ámbito tecnológico, el microaprendizaje se apoya en plataformas digitales que facilitan la presentación de "
          "cápsulas interactivas, videos cortos, infografías y actividades breves. La UNESCO (2023) advierte que la tecnología puede "
          "contribuir a la equidad y la calidad educativas solo cuando su uso responde a necesidades pedagógicas claras y a contextos "
          "específicos. Asimismo, la formación docente en tecnologías educativas es un factor clave: Ubillus-Sambachi et al. (2025) "
          "encontraron, en un estudio mixto con 26 profesores, que el microaprendizaje se encontraba entre las metodologías menos "
          "utilizadas (alrededor del 14 %) y que los docentes demandaban capacitación en herramientas como Powtoon, Canva y Kahoot.",
     c="UNESCO (2023) y Ubillus-Sambachi et al. (2025) no estaban en la lista de referencias; se agregaron. Verifica en el informe "
       "GEM 2023 la paráfrasis: el informe es crítico del uso de la tecnología sin propósito pedagógico, por lo que se matizó la "
       "afirmación original («contribuyen a mejorar la equidad»).")
para(doc, "La investigación sobre la memoria aporta bases adicionales: la presentación del contenido en episodios breves, "
          "distribuidos en el tiempo y acompañados de recuperación activa favorece la consolidación del aprendizaje en la memoria a "
          "largo plazo (Kelley y Whatson, 2013). Por ello, el taller distribuye las cápsulas a lo largo de cuatro semanas e incluye "
          "actividades breves de recuperación al inicio de cada cápsula.",
     c="Se sustituyó «Jasinski (2022)», que no pudo localizarse, por Kelley y Whatson (2013), un estudio verificado sobre "
       "aprendizaje espaciado (Frontiers in Human Neuroscience). Se vinculó el principio con una decisión de diseño concreta.")
para(doc, "A pesar del consenso sobre las ventajas del microaprendizaje en la segmentación del contenido y la retención, gran parte "
          "de los estudios se ha desarrollado en contextos de educación superior o de formación profesional (Betancur-Chicué y "
          "García-Valcárcel Muñoz-Repiso, 2023; Luo y Li, 2025), con menos evidencia en educación básica y, en particular, en "
          "contenidos de educación alimentaria y nutricional en escuelas mexicanas. El presente trabajo busca aportar evidencia "
          "empírica en ese ámbito.")
heading(doc, "2.1.2 Fundamentos teóricos del microaprendizaje", 3)
para(doc, "Uno de los principales referentes del microaprendizaje es la teoría de la carga cognitiva, propuesta por Sweller (1988), "
          "que parte de la capacidad limitada de la memoria de trabajo para procesar información nueva. Cuando los contenidos se "
          "presentan de forma extensa o desorganizada se incrementa la carga cognitiva extrínseca, lo que dificulta la comprensión y "
          "la retención (Sweller et al., 2019). La segmentación deliberada del contenido en unidades breves contribuye a reducir esa "
          "carga innecesaria y a focalizar la atención en los elementos esenciales; en esta línea, Luo y Li (2025) asocian las "
          "intervenciones de microaprendizaje con mejores resultados de aprendizaje.",
     c="Se sustituyó «Surbakti et al. (2024)», cuya revista y DOI no pudieron verificarse, por Sweller et al. (2019), la "
       "actualización de referencia de la teoría (Educational Psychology Review). Se corrigió «Surbakti et al,» (coma en lugar de "
       "punto). Verifica en Luo y Li (2025) la afirmación sobre la carga «germana» antes de reincorporarla.")
para(doc, "Asimismo, la teoría del aprendizaje multimedia de Mayer (2009) establece que se aprende mejor con palabras e imágenes "
          "que solo con palabras, siempre que se eliminen elementos irrelevantes (principio de coherencia), se resalte la "
          "información esencial (señalización) y el contenido se divida en segmentos que el estudiante pueda controlar "
          "(segmentación). Estos principios orientan el diseño de cápsulas breves y visualmente estructuradas.",
     c="Se corrigió «Teporía». Se desarrollaron los principios concretos de Mayer que justifican el diseño; el original solo "
       "mencionaba la teoría. Se eliminó un párrafo posterior que repetía esta idea sin cita.")
para(doc, "Desde una perspectiva complementaria, el conectivismo (Siemens, 2005) plantea que, en entornos digitales, el aprendizaje "
          "se construye mediante la interacción con recursos tecnológicos, redes de información y experiencias compartidas. En este "
          "sentido, el microaprendizaje favorece procesos flexibles y accesibles en los que los estudiantes interactúan con contenidos "
          "digitales breves.")
para(doc, "No obstante, la simple fragmentación de la información no garantiza un aprendizaje significativo si no se acompaña de un "
          "diseño instruccional coherente y alineado con objetivos claros (Betancur-Chicué y García-Valcárcel Muñoz-Repiso, 2023). "
          "Por ello, el microaprendizaje se entiende aquí como una estrategia que requiere una estructuración didáctica intencional, "
          "que en este estudio se organiza mediante el modelo ADDIE (Branch, 2009).",
     c="El original decía «algunos autores advierten…» sin citar a ninguno. En APA toda atribución requiere cita. Se corrigió "
       "también «sino se acompaña» por «si no se acompaña». Se eliminó la frase «eficiencia institucional», que no tenía sustento.")
heading(doc, "2.1.3 Aplicaciones del microaprendizaje en educación básica", 3)
para(doc, "La revisión sistemática de Silva et al. (2025) identifica que, en educación básica, la organización del contenido en "
          "unidades breves y focalizadas favorece la motivación, la atención sostenida y la retención a corto plazo, y que puede "
          "integrarse en entornos virtuales o híbridos, lo que favorece la autonomía del estudiante y la posibilidad de repetir los "
          "contenidos según sus necesidades (Leong et al., 2021). En la educación primaria, la segmentación contribuye a mantener el "
          "interés y a evitar la saturación informativa.",
     c="Se corrigió «Silva el at.» por «Silva et al.» y «Leong et al. (2021) destacan que el aprendizaje se consolida…» "
       "(faltaba «micro»). Leong et al. (2021) es una revisión de tendencias en el ámbito laboral; verifica que la afirmación "
       "sobre autonomía provenga de esa fuente.")
para(doc, "Sin embargo, la mayoría de estas investigaciones se centra en habilidades digitales o en formación general, y deja en "
          "segundo plano contenidos específicos como la educación nutricional; además, muchas intervenciones se desarrollan en "
          "contextos controlados. Existe, por tanto, la necesidad de generar evidencia contextualizada sobre el efecto del "
          "microaprendizaje en contenidos relevantes para la salud y el desarrollo infantil.")
heading(doc, "2.1.4 Diseño de cápsulas digitales y recursos tecnológicos", 3)
para(doc, "El diseño de cápsulas digitales requiere una planificación pedagógica intencional que considere tanto los principios "
          "cognitivos como los recursos tecnológicos disponibles. No se trata solo de reducir la duración del contenido, sino de "
          "estructurarlo para que cada cápsula cumpla un objetivo específico. Desde la teoría de la carga cognitiva, las cápsulas "
          "deben eliminar elementos irrelevantes, presentar un solo concepto central, cuidar la claridad visual y evitar la "
          "saturación de texto (Mayer, 2009; Sweller et al., 2019). En este proyecto, cada cápsula integra un video breve, un recurso "
          "visual y una actividad interactiva alineados con una dimensión del conocimiento nutricional (ver Tabla 4).")
heading(doc, "2.1.5 Ventajas del microaprendizaje para el aprendizaje activo", 3,
        c="Se eliminó el doble punto final del encabezado original («activo..»).")
para(doc, "El microaprendizaje puede favorecer la participación activa al presentar objetivos concretos y alcanzables en cada "
          "cápsula. Silva et al. (2025) señalan que incrementa la motivación y la atención sostenida en contextos escolares, y "
          "Betancur-Chicué y García-Valcárcel Muñoz-Repiso (2023) destacan que su diseño suele combinar microcontenido y preguntas, "
          "lo que promueve la implicación cognitiva del estudiante. No obstante, existe el riesgo de fragmentar en exceso el "
          "contenido y dificultar una visión integral; por ello, el taller incluye una cápsula final de integración y actividades que "
          "conectan las unidades entre sí.",
     c="Se eliminó el primer párrafo del original, que se repetía casi palabra por palabra en el segundo. Se agregó cómo el "
       "diseño atiende la limitación señalada.")

heading(doc, "2.2 Variable dependiente: nivel de conocimiento nutricional", 2)
heading(doc, "2.2.1 Conocimiento nutricional y educación alimentaria", 3)
para(doc, "La educación alimentaria y nutricional es un proceso educativo orientado a desarrollar conocimientos, habilidades y "
          "actitudes que permitan a las personas adoptar prácticas alimentarias saludables. La FAO (2020) plantea que la educación "
          "alimentaria y nutricional en la escuela, apoyada en un entorno alimentario saludable, contribuye a que escolares, "
          "adolescentes y sus comunidades mejoren su alimentación.",
     c="Se sustituyó «FAO (2013). Guía de educación alimentaria y nutricional para escuelas», cuyo título no pudo verificarse (la "
       "URL remite a una página web general), por el documento técnico verificado FAO (2020) sobre educación alimentaria y "
       "nutricional en la escuela.")
para(doc, "En el contexto escolar, la educación alimentaria se vincula con el desarrollo del conocimiento nutricional, entendido "
          "como la comprensión de los principios básicos de una alimentación saludable: la identificación de los grupos de "
          "alimentos, la importancia del equilibrio y la variedad, y la relación entre alimentación y salud ==(Pérez-Rodrigo y "
          "Aranceta, 2003)==.",
     c="VERIFICAR. El original atribuía esta definición a «Pérez-Rodrigo y Aranceta (2011)», que no se encontró. Se cambió al "
       "artículo verificado de 2003; confirma que la definición está en él o cita la fuente correcta.")
para(doc, "**Definición operacional.** Para efectos del presente estudio, el nivel de conocimiento nutricional es el puntaje "
          "(0 a 24) que obtiene el estudiante en el *Cuestionario de conocimiento nutricional para estudiantes de quinto grado de "
          "primaria*, el cual evalúa cuatro dimensiones: (a) principios de alimentación saludable y selección de alimentos, (b) "
          "grupos de alimentos y clasificación, (c) aportes y función nutricional básica de los alimentos y (d) aplicación de las "
          "recomendaciones alimentarias mexicanas en situaciones cotidianas.",
     c="Se agregó la definición operacional de la variable dependiente. Faltaba, y es indispensable en un estudio cuantitativo.")
para(doc, "Si bien el conocimiento nutricional se relaciona con la adopción de hábitos saludables, su incremento no se traduce "
          "necesariamente en cambios conductuales sostenidos (Pérez-Rodrigo y Aranceta, 2001). Por esta razón, el presente estudio se "
          "limita a medir el conocimiento.")
heading(doc, "2.2.2 Importancia del conocimiento nutricional en la infancia", 3)
para(doc, "La infancia es una etapa relevante para la construcción de conocimientos sobre alimentación y salud. La educación "
          "nutricional escolar puede ayudar a que los estudiantes comprendan principios básicos de la alimentación saludable y "
          "desarrollen herramientas para tomar decisiones informadas (FAO, 2020; Pérez-Rodrigo y Aranceta, 2003). La escuela ofrece, "
          "además, múltiples oportunidades para la prevención de la obesidad, tanto por la oferta de alimentos como por la educación "
          "y los servicios de salud que puede proporcionar (Story et al., 2006). Su relevancia "
          "aumenta ante la magnitud del sobrepeso y la obesidad en la población escolar mexicana (Shamah-Levy et al., 2023).",
     c="Se sustituyeron «Aranceta (2012)» y «Pérez-Rodrigo y Aranceta (2011)», que no pudieron verificarse.")
heading(doc, "2.2.3 El Plato del Bien Comer y las guías alimentarias mexicanas", 3)
para(doc, "En México, el Plato del Bien Comer ha sido la herramienta gráfica de orientación alimentaria. La versión de la "
          "NOM-043-SSA2-2012 organizaba los alimentos en tres grupos (Secretaría de Salud, 2013). Posteriormente, las *Guías "
          "Alimentarias Saludables y Sostenibles para la Población Mexicana* incorporaron el Plato del Bien Comer Saludable y "
          "Sostenible, que integra cinco grupos: verduras y frutas; cereales integrales o de granos enteros y tubérculos; "
          "leguminosas; alimentos de origen animal, y aceites y grasas saludables. Las guías recomiendan, además, el consumo de agua "
          "simple y la disminución de productos ultraprocesados y bebidas azucaradas (SSA et al., 2025).",
     c="La norma NOM-043-SSA2-2012 se publicó en el DOF el 22 de enero de 2013, por lo que se cita como Secretaría de Salud "
       "(2013). El original usaba 2012 en la lista y 2013 en el texto. La cita de las Guías se abrevió después de la primera "
       "mención: (SSA et al., 2025).")
para(doc, "La segunda edición de las guías, publicada en 2025, integra recomendaciones basadas en evidencia y contextualizadas a las "
          "características alimentarias, sociales y culturales de la población mexicana, con criterios de sostenibilidad, "
          "accesibilidad y pertinencia cultural (SSA et al., 2025). En el presente proyecto, los contenidos del taller y los reactivos "
          "del instrumento toman como referencia estas guías y el Plato del Bien Comer Saludable y Sostenible.",
     c="Se eliminó la cita triple repetida de las Guías al final de tres párrafos consecutivos. En APA basta con citar donde se "
       "usa la información.")
heading(doc, "2.2.4 Relación entre conocimiento nutricional y hábitos saludables", 3)
para(doc, "El conocimiento nutricional puede orientar las decisiones alimentarias, pero es solo uno de los factores que intervienen "
          "en ellas. El contexto familiar, el entorno socioeconómico y la disponibilidad de alimentos también influyen (Pérez-Rodrigo y "
          "Aranceta, 2001). Por ello, los programas que combinan información con participación activa y cambios en el entorno escolar "
          "obtienen mejores resultados. En consecuencia, el presente estudio se limita a evaluar el conocimiento nutricional, sin "
          "evaluar modificaciones en los hábitos.",
     c="Se compactaron tres párrafos del original que repetían la misma idea.")
heading(doc, "2.2.5 Programas escolares de educación nutricional", 3)
para(doc, "Los programas escolares de educación nutricional pueden mejorar el conocimiento de los estudiantes cuando se integran de "
          "manera sistemática a las actividades educativas (Pérez-Rodrigo y Aranceta, 2001; Shamah-Levy et al., 2017). En México, "
          "estas estrategias pueden apoyarse en la NOM-043-SSA2-2012 y en las Guías Alimentarias Saludables y Sostenibles para la "
          "Población Mexicana (Secretaría de Salud, 2013; SSA et al., 2025). Experiencias recientes muestran, además, la viabilidad de "
          "la modalidad en línea con escolares de primaria (Cruz Lara y Macossay Moreno, 2022).")
para(doc, "Lytle y Achterberg (1995) identificaron seis elementos de los programas exitosos de educación nutricional infantil: "
          "estar basados en la conducta y fundamentados en teoría, involucrar a la familia en el nivel primaria, incluir la "
          "autoevaluación de los patrones de alimentación en secundaria y bachillerato, intervenir en el entorno escolar y en la "
          "comunidad, y dedicar un tiempo de instrucción intensivo. Señalaron, además, que las lecciones solo expositivas son "
          "insuficientes; que el aprendizaje experiencial, la resolución de problemas y las oportunidades interactivas con computadora "
          "son más eficaces, y que la tecnología interactiva permite a los niños avanzar a su propio ritmo. También destacaron la "
          "necesidad de una evaluación rigurosa de resultados. Estos principios sustentan el diseño del taller: fundamentación "
          "teórica, actividades interactivas y evaluación preprueba-posprueba.",
     c="NUEVO, desde tu Zotero, con base en el resumen ejecutivo del artículo. Es un referente clásico que vincula la educación "
       "nutricional con la tecnología interactiva.")
heading(doc, "2.3 Relación entre el microaprendizaje y el conocimiento nutricional", 2)
para(doc, "La relación entre ambas variables se fundamenta en que la organización del contenido en unidades breves y estructuradas "
          "reduce la carga cognitiva extrínseca y focaliza la atención en un objetivo por vez (Sweller et al., 2019), lo que favorece "
          "el procesamiento progresivo de contenidos nutricionales como los grupos de alimentos o sus aportes. Diversos estudios "
          "indican que el microaprendizaje favorece la retención en educación básica (Leong et al., 2021; Silva et al., 2025); sin "
          "embargo, la evidencia sobre su aplicación a la educación alimentaria y nutricional en escuelas mexicanas es limitada. A "
          "partir de estos fundamentos, el estudio examina si la participación en un " + TALLER + " se asocia con un incremento del "
          "nivel de conocimiento nutricional de estudiantes de quinto grado, mediante la comparación de las mediciones previas y "
          "posteriores a la intervención.")
page_break(doc)

# ================================================================ CAPÍTULO 3
heading(doc, "Capítulo 3. Diseño metodológico", 1)
heading(doc, "3.1 Tipo, enfoque, alcance y diseño de la investigación", 2)
para(doc, "**Tipo.** La investigación es aplicada, ya que se orienta a atender una problemática específica del contexto escolar "
          "mediante la implementación de un " + TALLER + ".")
para(doc, "**Enfoque.** Cuantitativo, porque mide la variable dependiente mediante un instrumento estandarizado y analiza los datos "
          "con procedimientos estadísticos (Hernández-Sampieri y Mendoza Torres, 2018).",
     c="Faltaba una cita metodológica. Se usa Hernández-Sampieri y Mendoza Torres (2018), el referente más común en posgrados "
       "mexicanos. Si tu programa usa otro manual, sustitúyelo.")
para(doc, "**Alcance.** Descriptivo y explicativo: describe el nivel de conocimiento nutricional antes y después de la intervención "
          "y analiza si la diferencia se asocia con la implementación del taller.")
para(doc, "**Diseño.** Cuasiexperimental de preprueba y posprueba con un solo grupo intacto (O₁ X O₂), en el que se realiza una "
          "medición inicial (O₁), se implementa el taller (X) y se aplica una medición final (O₂) con el mismo instrumento. El grupo "
          "no se conforma por asignación aleatoria, sino que corresponde a un grupo escolar ya constituido.",
     c="Se conserva «cuasiexperimental» por decisión de la autora, en coherencia con el resumen aceptado en el CITEEA. Ten "
       "presente que algunos manuales (Campbell y Stanley, 1966; Hernández-Sampieri y Mendoza Torres, 2018) clasifican el diseño "
       "de un solo grupo sin grupo de comparación como preexperimental, y un sinodal podría señalarlo. Para blindar la "
       "denominación hay dos opciones: (a) agregar como grupo de comparación otro grupo de 5.º que no reciba el taller hasta "
       "después de la posprueba (el instituto tiene tres), lo que la hace indiscutible y fortalece la validez interna, o (b) "
       "conservar el diseño actual, con las amenazas a la validez interna explícitas en las limitaciones, como ya está redactado.")
heading(doc, "3.2 Población y muestra", 2)
para(doc, "**Población.** Estudiantes de 5.º grado de primaria del Instituto Educativo Ammadeus, distribuidos en tres grupos "
          "(==N = número total de estudiantes de 5.º==).",
     c="Indica el tamaño de la población (suma de los tres grupos).")
para(doc, "**Muestra.** Un grupo intacto de 5.º grado (n = 25), seleccionado mediante muestreo no probabilístico por conveniencia, "
          "considerando la accesibilidad y la disposición institucional. Con n = 25, una prueba *t* pareada unilateral (α = .05) "
          "tiene una potencia de .80 para detectar un tamaño del efecto de aproximadamente *d* = 0.50 (efecto mediano; Cohen, 1988).",
     c="Se agregó un análisis de potencia a priori. Justifica que 25 participantes son suficientes para detectar un efecto "
       "mediano. Cálculo: n ≈ [(1.645 + 0.842)/0.5]² ≈ 25. Puedes confirmarlo con G*Power (t tests, means: difference between two "
       "dependent means, one tail).")
heading(doc, "3.2.1 Criterios de selección", 3)
para(doc, "**Criterios de inclusión:** (a) estar inscrito en el grupo seleccionado; (b) contar con el consentimiento informado "
          "firmado por la madre, el padre o el tutor, y (c) otorgar asentimiento verbal y escrito.", indent=False)
para(doc, "**Criterios de exclusión:** estudiantes con alguna condición que les impida responder el cuestionario de forma "
          "autónoma, aun con los ajustes razonables disponibles.", indent=False)
para(doc, "**Criterios de eliminación:** (a) no responder la preprueba o la posprueba; (b) completar menos de seis de las ocho "
          "cápsulas del taller, y (c) dejar sin responder más del 20 % de los reactivos.", indent=False,
     c="Se separaron los criterios de exclusión (antes de iniciar) de los de eliminación (durante o después del estudio). El "
       "original mezclaba ambos e incluía como exclusión «no pertenecer al grupo», que no es un criterio porque ya está implícito "
       "en la muestra. Se agregó un umbral de exposición mínima a la intervención (6 de 8 cápsulas). Reporta cuántos casos se "
       "eliminan y por qué.")
heading(doc, "3.3 Variables y su operacionalización", 2)
para(doc, "La variable independiente es el " + TALLER + " y la variable dependiente es el nivel de conocimiento nutricional. Las "
          "Tablas 2 y 3 presentan su operacionalización.")
table(doc, 2, "Operacionalización de la variable independiente",
      ["Variable", "Dimensión", "Indicador", "Técnica / instrumento"],
      [["Taller educativo en línea basado en microaprendizaje", "Diseño del contenido", "8 cápsulas de 3 a 5 min con un objetivo cada una", "Carta descriptiva (Tabla 4); juicio de expertos"],
       ["", "Recursos digitales", "Video, infografía y actividad interactiva por cápsula", "Lista de cotejo de implementación (Anexo C)"],
       ["", "Implementación (fidelidad)", "Cápsulas completadas por estudiante (meta ≥ 6/8); asistencia", "Registro de la plataforma; lista de cotejo"],
       ["", "Participación", "Actividades interactivas resueltas", "Registro de la plataforma"]],
      widths=[3.6, 3.2, 5.0, 4.2],
      c="En el original, la variable independiente tenía «No aplica» en ítems y escala. Una variable manipulada no se mide con un "
        "cuestionario, pero sí debe documentarse su fidelidad de implementación: qué recibió realmente cada estudiante. Se "
        "agregaron indicadores verificables.")
table(doc, 3, "Operacionalización de la variable dependiente",
      ["Variable", "Dimensión", "Indicador", "Reactivos", "Escala / medición"],
      [["Nivel de conocimiento nutricional", "D1. Principios de alimentación saludable y selección de alimentos", "Identifica alimentos y bebidas acordes con las recomendaciones", "1–6", "Dicotómica (1 = correcto, 0 = incorrecto)"],
       ["", "D2. Grupos de alimentos y clasificación", "Clasifica alimentos en los cinco grupos del Plato del Bien Comer Saludable y Sostenible", "7–12", "Dicotómica"],
       ["", "D3. Aportes y función nutricional básica", "Relaciona alimentos con sus aportes nutrimentales", "13–18", "Dicotómica"],
       ["", "D4. Aplicación de las recomendaciones en situaciones cotidianas", "Resuelve casos de selección y combinación de alimentos", "19–24", "Dicotómica"],
       ["", "Puntaje total", "Suma de aciertos", "1–24", "0–24; niveles: bajo (0–12), medio (13–18), alto (19–24)"]],
      widths=[2.8, 4.0, 4.4, 1.6, 3.2],
      note="Los puntos de corte de los niveles corresponden a ≤ 50 %, 51 %–75 % y > 75 % de aciertos. ==Son una propuesta y "
           "deben justificarse o ajustarse con el asesor.==",
      c="CORRECCIÓN DE CONGRUENCIA. Las dimensiones de la tabla original (conceptos de educación alimentaria, grupos de alimentos, "
        "guías alimentarias, hábitos saludables) NO coincidían con las cuatro dimensiones del instrumento que se envió a juicio de "
        "expertos. Se alineó la tabla con el instrumento real. Además, la técnica no es «encuesta»: para medir conocimiento es una "
        "prueba de conocimientos con cuestionario de opción múltiple.")
heading(doc, "3.4 Técnica e instrumento", 2)
para(doc, "Se empleará la técnica de prueba de conocimientos, mediante el *Cuestionario de conocimiento nutricional para estudiantes "
          "de quinto grado de primaria* (Anexo A), elaborado ex profeso. Está integrado por 24 reactivos de opción múltiple con cuatro "
          "opciones y una sola respuesta correcta, distribuidos en cuatro dimensiones de seis reactivos cada una (Tabla 3). Los "
          "reactivos se redactaron con lenguaje adecuado para niñas y niños de 10 y 11 años y con base en el Plato del Bien Comer "
          "Saludable y Sostenible (SSA et al., 2025). Cada respuesta correcta vale un punto, por lo que el puntaje total va de 0 a 24.",
     c="El texto original describía un instrumento «de opción múltiple y/o respuesta cerrada» con dimensiones distintas. Se "
       "actualizó con las características reales del instrumento (24 reactivos, 4 dimensiones). Se corrigió «ractios» por "
       "«reactivos». Atiende la observación del asesor: el cuestionario completo va en el Anexo A.")
heading(doc, "3.4.1 Validez de contenido", 3)
para(doc, "La validez de contenido se valorará mediante juicio de expertos (Escobar-Pérez y Cuervo-Martínez, 2008). ==Tres a cinco "
          "jueces== con experiencia en nutrición, educación primaria y metodología de la investigación calificarán cada reactivo en "
          "claridad, coherencia, pertinencia y, en el caso de los jueces nutriólogos, exactitud del contenido nutricional, con una "
          "escala de 1 a 4. Para cada reactivo y criterio se calculará el coeficiente V de Aiken (Aiken, 1985) con su intervalo de "
          "confianza al 95 % (Penfield y Giacobbi, 2004). Se conservarán los reactivos con V ≥ .70 cuyo límite inferior del "
          "intervalo sea mayor que .50; los demás se revisarán o se eliminarán según las observaciones de los jueces.",
     c="Atiende la observación del asesor: debe incluirse otro juez en su lugar (sugirió a la Dra. Verónica Torres), porque el "
       "director de tesis no debe validar su propio instrumento. Registra en el documento final los nombres, grados y perfiles "
       "de los jueces. Los umbrales (V ≥ .70, IC 95 %) son convencionales; ajústalos si tu programa usa otros.")
heading(doc, "3.4.2 Confiabilidad y análisis de reactivos", 3)
para(doc, "Antes de la intervención se aplicará una prueba piloto a ==un grupo de 5.º grado que no participe en el estudio== "
          "(aproximadamente 25 estudiantes). Por tratarse de reactivos dicotómicos, la consistencia interna se estimará con el "
          "coeficiente Kuder-Richardson 20 (KR-20; Kuder y Richardson, 1937), con un valor mínimo aceptable de .70. Se calcularán "
          "además el índice de dificultad y el índice de discriminación de cada reactivo para depurar los que resulten demasiado "
          "fáciles, demasiado difíciles o poco discriminativos. La prueba piloto también permitirá estimar el tiempo de aplicación y "
          "detectar problemas de comprensión.",
     c="FALTABA en el original: no se mencionaba la confiabilidad. Es un requisito para un instrumento elaborado por la autora. "
       "El alfa de Cronbach no es el indicado para reactivos correcto/incorrecto; KR-20 es su equivalente para datos dicotómicos. "
       "El piloto debe hacerse con estudiantes que NO formen parte de la muestra.")
heading(doc, "3.5 Procedimiento", 2)
bullets(doc, [
    "Obtención de la autorización de la dirección del Instituto Educativo Ammadeus.",
    "Validación de contenido del instrumento mediante juicio de expertos y ajuste de los reactivos.",
    "Prueba piloto con un grupo no participante; cálculo de KR-20 y análisis de reactivos.",
    "Entrega y recepción del consentimiento informado y obtención del asentimiento (Anexo B).",
    "Aplicación de la preprueba en un formulario digital, en el aula, de forma individual, con instrucciones estandarizadas y un tiempo estimado de 30 minutos.",
    "Implementación del taller durante cuatro semanas (dos cápsulas por semana; Tabla 4) y registro de la fidelidad de implementación (Anexo C).",
    "Aplicación de la posprueba con el mismo instrumento y en las mismas condiciones, en la semana siguiente al cierre del taller.",
    "Captura, depuración y análisis de los datos.",
], numbered=True, c="Se agregó un procedimiento paso a paso que hace replicable el estudio. Faltaba en el original.")
heading(doc, "3.6 Análisis de datos", 2)
para(doc, "Los datos se analizarán con ==JASP o IBM SPSS Statistics==. Primero se obtendrá la estadística descriptiva de la "
          "preprueba y de la posprueba (media, desviación estándar, mediana, rango intercuartílico, mínimo y máximo) para el puntaje "
          "total y por dimensión, así como la distribución de los estudiantes por nivel de conocimiento y el porcentaje de aciertos "
          "por reactivo (OE1).",
     c="Indica el software que usarás.")
para(doc, "Para contrastar la hipótesis (OE3) se evaluará la normalidad de las diferencias (posprueba − preprueba) con la prueba "
          "de Shapiro-Wilk, adecuada para muestras pequeñas. Si las diferencias se distribuyen normalmente, se aplicará la prueba *t* "
          "de Student para muestras relacionadas, unilateral; en caso contrario, la prueba de rangos con signo de Wilcoxon. El nivel "
          "de significancia será α = .05. Se reportará el tamaño del efecto (*d* de Cohen para muestras relacionadas o *r* = *Z*/√*N* "
          "para Wilcoxon), interpretado según Cohen (1988). El mismo procedimiento se aplicará a cada dimensión, con corrección de "
          "Bonferroni (α = .0125) por realizar cuatro comparaciones.",
     c="CORRECCIÓN METODOLÓGICA IMPORTANTE. El original proponía solo estadística descriptiva («frecuencias, porcentajes y, en "
       "caso de ser pertinente, comparación de puntajes»), lo cual NO permite aceptar o rechazar la hipótesis planteada. Además, "
       "la plantilla del capítulo 4 sugería Chi cuadrado, que no aplica a puntajes de un mismo grupo medidos dos veces. Las pruebas "
       "pertinentes son la t pareada o Wilcoxon. El tamaño del efecto es imprescindible con n pequeña.")
heading(doc, "3.7 Consideraciones éticas", 2)
para(doc, "La investigación se realizará conforme a los principios de respeto a las personas, beneficencia y justicia. Se "
          "solicitará la autorización de la dirección escolar, el consentimiento informado de madres, padres o tutores y el "
          "asentimiento de las niñas y los niños, a quienes se explicará en lenguaje sencillo que su participación es voluntaria, que "
          "pueden retirarse en cualquier momento sin consecuencias y que el cuestionario no afecta sus calificaciones. Los datos se "
          "identificarán con códigos, se resguardarán de forma confidencial y se usarán solo con fines académicos. Por tratarse de "
          "una intervención educativa sin procedimientos invasivos, se considera una investigación de riesgo mínimo ==(Reglamento de "
          "la Ley General de Salud en Materia de Investigación para la Salud, 1987, art. 17)==. Al concluir el estudio, los "
          "materiales del taller se pondrán a disposición de los demás grupos de 5.º grado.",
     c="El original solo mencionaba confidencialidad y consentimiento dentro de los criterios de inclusión. Al trabajar con "
       "menores, el comité espera un apartado ético propio. Ofrecer el taller a los demás grupos atiende el principio de justicia. "
       "VERIFICA si tu programa o la UAZ exige dictamen de un comité de ética y la versión vigente del Reglamento.")
heading(doc, "3.8 Diseño instruccional: modelo ADDIE", 2,
        c="ATIENDE LA OBSERVACIÓN DEL ASESOR: «El modelo ADDIE no solo se describe, también hay que desarrollar las etapas de su "
          "programa educativo». El original agrupaba ADDIE en tres etapas genéricas. Ahora se desarrollan las cinco fases con las "
          "decisiones concretas del taller, una carta descriptiva por cápsula (Tabla 4) y la alineación de cada cápsula con una "
          "dimensión del instrumento. Revisa que las herramientas propuestas estén permitidas en la escuela.")
para(doc, "El taller se diseñó con el modelo ADDIE (análisis, diseño, desarrollo, implementación y evaluación), que organiza de "
          "manera sistemática la elaboración de una propuesta formativa y asegura la coherencia entre necesidades, objetivos, "
          "contenidos, actividades y evaluación (Branch, 2009). A continuación se desarrolla cada fase.")
heading(doc, "3.8.1 Fase de análisis", 3)
bullets(doc, [
    "**Necesidad educativa:** déficit de conocimiento nutricional documentado en escolares zacatecanos (López Hernández et al., 2024), que se confirmará en el grupo con la preprueba.",
    "**Perfil de los estudiantes:** 25 estudiantes de 10 y 11 años, con lectura autónoma, familiarizados con dispositivos digitales; atención sostenida limitada, lo que justifica cápsulas de 3 a 5 minutos.",
    "**Contexto y recursos:** aula digital con proyector, computadoras y conexión a internet; acompañamiento de la docente titular; dos sesiones semanales de 20 a 25 minutos dentro del horario escolar ==(confirmar con la escuela)==.",
    "**Contenidos de referencia:** Plato del Bien Comer Saludable y Sostenible y recomendaciones para población escolar de las Guías Alimentarias (SSA et al., 2025); vinculación con el campo formativo de la fase 5 del plan de estudios.",
    "**Restricciones:** calendario escolar, disponibilidad de equipos y tiempos de la docente.",
])
heading(doc, "3.8.2 Fase de diseño", 3)
para(doc, "**Objetivo de aprendizaje del taller:** que el estudiante identifique, clasifique y combine alimentos de acuerdo con el "
          "Plato del Bien Comer Saludable y Sostenible para elegir opciones saludables en situaciones de su vida diaria.")
para(doc, "**Estructura:** ocho cápsulas distribuidas en cuatro semanas (dos por semana), cada una con un objetivo específico y "
          "alineada con una dimensión del instrumento. Cada cápsula sigue una secuencia de cuatro momentos, derivada de la estructura "
          "introducción–microcontenido–preguntas (Betancur-Chicué y García-Valcárcel Muñoz-Repiso, 2023): (a) activación o "
          "recuperación de la cápsula anterior (2 min), (b) microcontenido en video o infografía (3 a 5 min), (c) actividad "
          "interactiva con retroalimentación inmediata (5 a 8 min) y (d) reto o pregunta de cierre (2 min). La Tabla 4 presenta la "
          "carta descriptiva.")
table(doc, 4, "Carta descriptiva de las cápsulas del taller",
      ["Sem.", "Cápsula", "Objetivo específico", "Recursos y actividad", "Dim."],
      [["1", "C1. ¿Qué es comer saludable?", "Explicar por qué se recomienda una alimentación variada", "Video animado; actividad de unir parejas", "D1"],
       ["1", "C2. Agua simple y sellos de advertencia", "Elegir agua simple y reconocer productos con sellos", "Infografía; juego de clasificar «diario / de vez en cuando»", "D1"],
       ["2", "C3. Mi Plato del Bien Comer", "Nombrar los cinco grupos del Plato del Bien Comer Saludable y Sostenible", "Video; rompecabezas interactivo del plato", "D2"],
       ["2", "C4. ¿A qué grupo pertenece?", "Clasificar alimentos cotidianos en su grupo", "Juego de arrastrar y soltar", "D2"],
       ["3", "C5. Vitaminas, minerales y fibra", "Relacionar verduras, frutas y cereales integrales con sus aportes", "Video; cuestionario con retroalimentación", "D3"],
       ["3", "C6. Proteína vegetal y grasas saludables", "Identificar leguminosas y grasas saludables y la combinación leguminosa + cereal", "Infografía; ruleta de alimentos", "D3"],
       ["4", "C7. Armo mi refrigerio escolar", "Seleccionar un refrigerio saludable", "Simulador de lonchera", "D4"],
       ["4", "C8. Reto final: mi comida del día", "Integrar las recomendaciones en un menú del día", "Caso interactivo y repaso gamificado", "D1–D4"]],
      widths=[1.3, 3.5, 4.6, 4.6, 1.6],
      note="D1 = principios de alimentación saludable; D2 = grupos de alimentos; D3 = aportes nutricionales; D4 = aplicación de "
           "recomendaciones. ==Ajustar títulos y recursos a la versión final del taller.==",
      c="Esta tabla responde también a la pregunta específica 2 (¿qué elementos digitales incluir?). Cada dimensión del "
        "instrumento queda cubierta por al menos dos cápsulas, con lo que se asegura la alineación entre enseñanza y evaluación.")
heading(doc, "3.8.3 Fase de desarrollo", 3)
para(doc, "Se producirán los materiales de cada cápsula: videos de 3 a 5 minutos e infografías elaborados ==en Canva==, actividades "
          "interactivas ==en Wordwall o Genially== y un repaso gamificado ==en Kahoot==, alojados en ==Google Classroom== como espacio "
          "único del taller. Los materiales se diseñarán con los principios de coherencia, señalización y segmentación (Mayer, 2009) "
          "y se revisarán con una lista de verificación de calidad (exactitud nutricional, claridad del lenguaje, duración y "
          "accesibilidad). Antes de implementarlos, ==dos especialistas (nutrición y tecnología educativa)== revisarán una cápsula "
          "prototipo y se harán los ajustes necesarios.",
     c="Las herramientas son una propuesta coherente con las que mencionan Ubillus-Sambachi et al. (2025). Sustitúyelas por las "
       "que realmente usarás y verifica su política de uso con menores (cuentas, privacidad).")
heading(doc, "3.8.4 Fase de implementación", 3)
para(doc, "El taller se implementará durante cuatro semanas, con dos sesiones semanales en el aula digital y acompañamiento de la "
          "docente titular y de la investigadora. En cada sesión se registrará la asistencia y las cápsulas completadas por cada "
          "estudiante (Anexo C). Las cápsulas permanecerán disponibles en la plataforma para su consulta posterior.")
heading(doc, "3.8.5 Fase de evaluación", 3)
para(doc, "La evaluación se realizará en tres niveles: (a) **diagnóstica**, con la preprueba; (b) **formativa**, mediante los "
          "resultados de las actividades interactivas de cada cápsula, que permitirán reforzar los contenidos con menor desempeño, y "
          "(c) **sumativa**, con la posprueba y su comparación estadística con la preprueba. Además, se evaluará el propio taller con "
          "un breve cuestionario de satisfacción para los estudiantes y con el registro de fidelidad de implementación, a fin de "
          "identificar mejoras para versiones futuras.",
     c="El original decía que se mediría el conocimiento «antes, durante y después» con instrumentos de evaluación. La evaluación "
       "«durante» es formativa y NO forma parte del contraste de hipótesis. Se distinguieron los tres tipos de evaluación.")
heading(doc, "3.9 Cronograma", 2)
table(doc, 5, "Cronograma de actividades",
      ["Actividad", "Periodo"],
      [["Juicio de expertos y ajuste del instrumento", "Noviembre de 2026"],
       ["Presentación del cartel en el I CITEEA (UAZ)", "19 y 20 de noviembre de 2026"],
       ["Prueba piloto y análisis de confiabilidad", "1 al 11 de diciembre de 2026"],
       ["Desarrollo de las cápsulas (fase de desarrollo)", "Noviembre de 2026 – enero de 2027"],
       ["Entrega y recepción de consentimientos y asentimientos", "7 al 15 de enero de 2027"],
       ["Preprueba", "18 al 22 de enero de 2027"],
       ["Implementación del taller (8 cápsulas; inicio en la semana del Día del Nutriólogo, 27 de enero)", "25 de enero al 19 de febrero de 2027"],
       ["Posprueba", "22 al 26 de febrero de 2027"],
       ["Análisis de datos y redacción de resultados", "Marzo – abril de 2027"]],
      widths=[10, 6],
      note="Fechas tentativas. ==Confirmar con el calendario escolar del Instituto Educativo Ammadeus (vacaciones, suspensiones y evaluaciones).==",
      c="Faltaba un cronograma, que es un componente estándar de todo protocolo.")
page_break(doc)

heading(doc, "Capítulo 4. Resultados", 1)
para(doc, "Este capítulo se desarrollará una vez concluida la intervención. Los resultados se presentarán en el orden de los "
          "objetivos específicos: (a) nivel inicial de conocimiento nutricional (estadística descriptiva de la preprueba); (b) "
          "descripción del taller desarrollado y de su fidelidad de implementación, y (c) comparación entre preprueba y posprueba, "
          "total y por dimensión, con el contraste de hipótesis y el tamaño del efecto. Posteriormente, en la discusión, los "
          "resultados se compararán con los antecedentes (Cruz Lara y Macossay Moreno, 2022; Shamah-Levy et al., 2017).",
     c="Se sustituyeron las instrucciones de la plantilla (incluida la indicación de «agregar la prueba de Chi cuadrado», que no "
       "aplica a este diseño) por el plan de presentación de resultados de este estudio.")
heading(doc, "Capítulo 5. Conclusiones", 1)
para(doc, "Este capítulo se redactará al concluir el análisis. Responderá en orden a la pregunta general y a las preguntas "
          "específicas, e incluirá las limitaciones y recomendaciones para futuras investigaciones.")
page_break(doc)

# ================================================================ REFERENCIAS
heading(doc, "Referencias", 1,
        c="Lista reconstruida y verificada. Todas las fuentes citadas en el texto aparecen aquí y todas las de la lista se citan en "
          "el texto. Se ordenaron alfabéticamente, con sangría francesa de 1.27 cm, «y» antes del último autor (APA en español), "
          "DOI en formato https://doi.org/ y sin «doi:» ni «Recuperado de». Lo resaltado en amarillo requiere un dato que no pudo "
          "confirmarse en línea.")
REFS = [
    ("==Aguirre Vargas, M. C.== (2019). Evaluación de la mejora del conocimiento de los escolares que cursan tercer y cuarto grado de primaria sobre «El Plato del Bien Comer» después de la intervención educativa en nutrición. *Jóvenes en la Ciencia, 6*. https://www.jovenesenlaciencia.ugto.mx/index.php/jovenesenlaciencia/article/view/3137",
     "Desde tu Zotero. Corresponde a la cita errónea «García et al. (2019)». Confirma el orden de los apellidos de la autora."),
    ("Aiken, L. R. (1985). Three coefficients for analyzing the reliability and validity of ratings. *Educational and Psychological Measurement, 45*(1), 131–142. https://doi.org/10.1177/0013164485451012", None),
    ("Banco Interamericano de Desarrollo. (s. f.). *Educación*. https://www.iadb.org/es/quienes-somos/tematicas/educacion", None),
    ("Bermudez, O. I., y Tucker, K. L. (2003). Trends in dietary patterns of Latin American populations. *Cadernos de Saúde Pública, 19*(Supl. 1), S87–S99. https://doi.org/10.1590/S0102-311X2003000700010",
     "Corregida: no es una publicación de la FAO, sino de Cadernos de Saúde Pública. El apellido se escribe Bermudez, sin tilde, como en la fuente."),
    ("Betancur-Chicué, V., y García-Valcárcel Muñoz-Repiso, A. (2023). Características del diseño de estrategias de microaprendizaje en escenarios educativos: Revisión sistemática. *RIED. Revista Iberoamericana de Educación a Distancia, 26*(1), 201–222. https://doi.org/10.5944/ried.26.1.34056",
     "Corregida: faltaba el punto tras el año, la coma tras «Revisión sistemática» debía ser punto y el rango de páginas terminaba en «222-»."),
    ("Branch, R. M. (2009). *Instructional design: The ADDIE approach*. Springer. https://doi.org/10.1007/978-0-387-09506-6", "Agregada: se citaba en el texto pero no estaba en la lista."),
    ("Campbell, D. T., y Stanley, J. C. (1966). *Experimental and quasi-experimental designs for research*. Rand McNally.", "Nueva: sustenta la clasificación del diseño."),
    ("Cohen, J. (1988). *Statistical power analysis for the behavioral sciences* (2.ª ed.). Lawrence Erlbaum Associates.", "Nueva: tamaño del efecto y potencia."),
    ("Cruz Lara, N. M., y Macossay Moreno, C. (2022). Intervención de educación nutricional en línea a población infantil en el estado de Veracruz. *Horizonte Sanitario, 21*(3), 403–409. https://doi.org/10.19136/hs.a21n3.4893",
     "Nueva (sustituye a Mendoza y Ruiz, 2021). Datos completados con tu Zotero: dos autoras, por lo que en el texto se cita «Cruz Lara y Macossay Moreno (2022)», no «et al.». En tu Zotero el registro tiene las autoras duplicadas; conviene limpiarlo."),
    ("Olivares, S., y Snel, J. (1998). Educación en nutrición en las escuelas primarias. *Alimentación, Nutrición y Agricultura*, (22), ==[páginas]==. Organización de las Naciones Unidas para la Alimentación y la Agricultura. https://www.fao.org/4/X0051T/X0051t08.htm",
     "Restaurada desde tu Zotero (antes citada por error como «Gómez, 2008»). Iniciales confirmadas; completa las páginas."),
    ("Escobar-Pérez, J., y Cuervo-Martínez, Á. (2008). Validez de contenido y juicio de expertos: Una aproximación a su utilización. *Avances en Medición, 6*, 27–36.", "Nueva: fundamenta el juicio de expertos."),
    ("Food and Agriculture Organization of the United Nations. (2020). *School-based food and nutrition education: A white paper on the current state, principles, challenges and recommendations for low- and middle-income countries*. https://www.fao.org/platforms/school-food/technical-resources/publications/school-based-food-and-nutrition-education/en",
     "Sustituye a «FAO (2013). Guía de educación alimentaria y nutricional para escuelas», cuyo título no pudo verificarse. Si cuentas con la guía de 2013, verifica sus datos y restáurala."),
    ("González Ruiz, C. J., Martín Gómez, S., y Cabrera Hernández, D. J. (2024). Microlearning en la formación de docentes y familias de Educación Infantil: Una propuesta de aprendizaje híbrido. *Edutec. Revista Electrónica de Tecnología Educativa*, (88). https://doi.org/10.21556/edutec.2024.88.3127",
     "Opcional: no se cita en el texto. Es un ejemplo verificado de microaprendizaje en educación básica (Educación Infantil) que puedes incorporar en el apartado 2.1.3. Si no lo usas, elimínalo de la lista."),
    ("Hernández-Sampieri, R., y Mendoza Torres, C. P. (2018). *Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta*. McGraw-Hill Interamericana.", "Nueva: referente metodológico."),
    ("Kelley, P., y Whatson, T. (2013). Making long-term memories in minutes: A spaced learning pattern from memory research in education. *Frontiers in Human Neuroscience, 7*, Artículo 589. https://doi.org/10.3389/fnhum.2013.00589", "Sustituye a Jasinski (2022), que no pudo localizarse."),
    ("Kuder, G. F., y Richardson, M. W. (1937). The theory of the estimation of test reliability. *Psychometrika, 2*(3), 151–160. https://doi.org/10.1007/BF02288391", "Nueva: confiabilidad KR-20."),
    ("Leong, K., Sung, A., Au, D., y Blanchard, C. (2021). A review of the trend of microlearning. *Journal of Work-Applied Management, 13*(1), 88–102. https://doi.org/10.1108/JWAM-10-2020-0044", "Verificada."),
    ("López Hernández, P. Y., Franco Trejo, C. S., Galicia Rodríguez, A. G., Argüelles Estrada, B. C., y Ornelas Pargas, A. C. (2024). Sobrepeso, obesidad y educación nutricional en una escuela primaria de Zacatecas. *Revista Iberoamericana de Ciencias, 11*(1), ==8–[página final]==. https://www.reibci.org/publicados/2024/mzo/5300113.pdf",
     "Verificada en el índice del número de marzo de 2024 de la revista: inicia en la página 8, no en la 1. Completa la página final. Se corrigió «&» por «y» y «Arguelles» por «Argüelles»."),
    ("Luo, H., y Li, W. (2025). Impact of microlearning on developing soft skills of university students across disciplines. *Frontiers in Psychology, 16*, Artículo 1491265. https://doi.org/10.3389/fpsyg.2025.1491265",
     "Corregida: eran dos autores, no «Luo, H., et al.» (APA 7 no admite «et al.» en la lista de referencias). El título estaba incompleto, faltaba el volumen y había un carácter oculto al final del DOI."),
    ("Martínez Villagrana, M. (2023). Educación nutricional en la escuela: Pedagogía para la vida saludable. Caso: primaria rural en municipio zacatecano. ==*[Nombre completo de la revista RIEIPD]*==, *1*(1). https://doi.org/10.71770/rieipd.v1i1.1793",
     "Nueva (contexto local; revista de la UAZ). Completa el nombre de la revista y las páginas."),
    ("Lytle, L., y Achterberg, C. (1995). Changing the diet of America's children: What works and why? *Journal of Nutrition Education, 27*(5), 250–260. https://doi.org/10.1016/S0022-3182(12)80793-X", "Nueva, desde tu Zotero."),
    ("Mayer, R. E. (2009). *Multimedia learning* (2.ª ed.). Cambridge University Press. https://doi.org/10.1017/CBO9780511811678", "Agregada: se citaba en el texto pero no estaba en la lista."),
    ("Organización de las Naciones Unidas. (2015). *Transformar nuestro mundo: La Agenda 2030 para el Desarrollo Sostenible* (A/RES/70/1). https://sdgs.un.org/2030agenda", "Corregida: no se repite la sigla en la referencia; se agregó la signatura del documento."),
    ("Penfield, R. D., y Giacobbi, P. R., Jr. (2004). Applying a score confidence interval to Aiken's item content-relevance index. *Measurement in Physical Education and Exercise Science, 8*(4), 213–225. https://doi.org/10.1207/s15327841mpee0804_3", "Nueva: intervalo de confianza de la V de Aiken."),
    ("Pérez-Rodrigo, C., y Aranceta, J. (2001). School-based nutrition education: Lessons learned and new perspectives. *Public Health Nutrition, 4*(1A), 131–139. https://doi.org/10.1079/PHN2000108",
     "Corregida: el DOI decía «doi: https://doi.org.10.1079…», un enlace roto (punto en lugar de barra)."),
    ("Pérez-Rodrigo, C., y Aranceta, J. (2003). Nutrition education in schools: Experiences and challenges. *European Journal of Clinical Nutrition, 57*(Supl. 1), S82–S85. https://doi.org/10.1038/sj.ejcn.1601824",
     "Sustituye las citas «Pérez-Rodrigo y Aranceta (2011)» y «(2013)», que no pudieron localizarse."),
    ("Reséndiz Martínez, L. (2023). *Desafíos y alcances de la educación alimentaria a partir del uso de material digital en niños de edad escolar y sus familias* ==[Tesis de maestría]==, Universidad Autónoma de Querétaro. Repositorio Institucional UAQ. https://ri-ng.uaq.mx/handle/123456789/8429", "Nueva, desde tu Zotero. Confirma el tipo de tesis."),
    ("Reglamento de la Ley General de Salud en Materia de Investigación para la Salud. (1987, 6 de enero). Diario Oficial de la Federación. ==Última reforma publicada [fecha]. [URL de la Cámara de Diputados o del DOF]==", "Nueva: consideraciones éticas. Completa la última reforma y la URL."),
    ("Ríos-Castillo, I., Lizárraga-Quintero, A., Ortega, L., Fontes, F., y Valdés, V. (2024). Estrategias innovadoras de educación alimentaria y nutricional para combatir el exceso de peso y la obesidad en niños de edad escolar empleadas en países hispanos: Una revisión narrativa. *Revista Chilena de Nutrición, 51*(4), 333–339. https://doi.org/10.4067/s0717-75182024000400333", "Nueva, desde tu Zotero."),
    ("Secretaría de Salud. (2013, 22 de enero). Norma Oficial Mexicana NOM-043-SSA2-2012, Servicios básicos de salud. Promoción y educación para la salud en materia alimentaria. Criterios para brindar orientación. *Diario Oficial de la Federación*. https://www.gob.mx/cms/uploads/attachment/file/138258/NOM-043-servicios-basicos-salud-educacion-alimentaria.pdf",
     "Corregida: el año de publicación en el DOF es 2013; el título oficial de la norma estaba incompleto."),
    ("Secretaría de Salud, Instituto Nacional de Salud Pública y Fondo de las Naciones Unidas para la Infancia. (2025). *Guías alimentarias saludables y sostenibles para la población mexicana* (2.ª ed.). ==[URL oficial; p. ej., la publicación en el DOF]==",
     "Agregada: se citaba cuatro veces, pero no estaba en la lista. La 2.ª edición (2025) está confirmada. Agrega la URL de la que obtuviste el documento."),
    ("Shamah-Levy, T., Cuevas-Nasu, L., Gómez-Acosta, L. M., Morales-Ruan, M. C., Méndez Gómez-Humarán, I., Robles-Villaseñor, M. N., y Hernández-Ávila, M. (2017). Efecto del programa de servicios SaludArte en los componentes de alimentación y nutrición en escolares de la Ciudad de México. *Salud Pública de México, 59*(6), 621–629. https://doi.org/10.21149/8116",
     "Nueva (sustituye a Díaz y Ramírez, 2020, y a Morales y Salazar, 2013)."),
    ("Shamah-Levy, T., Gaona-Pineda, E. B., Cuevas-Nasu, L., ==[demás autores]==. (2023). Prevalencias de sobrepeso y obesidad en población escolar y adolescente de México. Ensanut Continua 2020-2022. *Salud Pública de México, 65*(Supl. 1), S218–S224. https://doi.org/10.21149/14762",
     "Sustituye a «INSP (2020). ENSANUT 2020». APA 7 exige listar hasta 20 autores: copia la lista completa desde el artículo."),
    ("Siemens, G. (2005). Connectivism: A learning theory for the digital age. *International Journal of Instructional Technology and Distance Learning, 2*(1), 3–10.", "Agregada: se citaba en el texto pero no estaba en la lista."),
    ("Silva, E. S., Pires da Costa, W., de Lima, J. C., y Ferreira, J. C. (2025). Contribution of microlearning in basic education: A systematic review. *Education Sciences, 15*(3), Artículo 302. https://doi.org/10.3390/educsci15030302",
     "Se confirmaron título, revista, volumen y DOI. Verifica las iniciales de los autores en la página del artículo (MDPI); en una cita secundaria aparecen como «Silva, Costa, Lima y Ferreira»."),
    ("Story, M., Kaphingst, K. M., y French, S. (2006). The role of schools in obesity prevention. *The Future of Children, 16*(1), 109–142. https://muse.jhu.edu/pub/11/article/194627", "Nueva, desde tu Zotero. Se omite «Ph. D», que Zotero agregó al nombre de la autora: APA no incluye grados académicos."),
    ("Sweller, J. (1988). Cognitive load during problem solving: Effects on learning. *Cognitive Science, 12*(2), 257–285. https://doi.org/10.1207/s15516709cog1202_4", "Agregada: se citaba en el texto pero no estaba en la lista."),
    ("Sweller, J., van Merriënboer, J. J. G., y Paas, F. (2019). Cognitive architecture and instructional design: 20 years later. *Educational Psychology Review, 31*(2), 261–292. https://doi.org/10.1007/s10648-019-09465-5", "Sustituye a Surbakti et al. (2024), que no pudo verificarse."),
    ("Termes Escalé, M., Martínez Chicano, D., Egea Castillo, N., Gutiérrez Sánchez, A., García Arenas, D., Llata Vidal, N., y Martín de Carpi, J. (2020). Educación nutricional en niños en edad escolar a través del Programa Nutriplato®. *Nutrición Hospitalaria, 37*(Spe. 2), 47–51. https://doi.org/10.20960/nh.03357", "Nueva, desde tu Zotero."),
    ("Ubillus-Sambachi, ==[iniciales y demás autores]== (2025). Capacitación docente en TICS para el fomento de la metodología de micro aprendizaje. *593 Digital Publisher CEIT, 10*(3), 219–230. https://doi.org/10.33386/593dp.2025.3.3128",
     "Agregada: se citaba en el texto pero no estaba en la lista. Completa los nombres de los autores."),
    ("UNESCO. (2017). *Educación para los Objetivos de Desarrollo Sostenible: Objetivos de aprendizaje*. https://unesdoc.unesco.org/ark:/48223/pf0000252423", "Verificada. Se eliminó «UNESCO Publishing», que es el mismo autor."),
    ("UNESCO. (2023). *Informe de seguimiento de la educación en el mundo 2023: Tecnología en la educación: ¿Una herramienta en los términos de quién?* ==[URL o DOI de UNESDOC]==", "Agregada: se citaba en el texto pero no estaba en la lista. Confirma el título en español y agrega el enlace."),
]
import unicodedata
def _key(t):
    t = t[0].replace("==", "").replace("*", "")
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower().replace(",", "~")
for text, c in sorted(REFS, key=_key):
    ref(doc, text, c)
page_break(doc)

# ================================================================ ANEXOS
heading(doc, "Anexos", 1)
heading(doc, "Anexo A. Cuestionario de conocimiento nutricional para estudiantes de quinto grado de primaria", 2,
        c="Atiende la observación del asesor: el instrumento va completo en los anexos. Las claves de respuesta se incluyen solo en "
          "la versión del protocolo; la versión que contestarán los estudiantes no debe mostrarlas. Las observaciones a reactivos "
          "específicos están en los comentarios y en el informe de revisión.")
para(doc, "**Instrucciones para el estudiante:** Lee con atención cada pregunta y elige solo una respuesta. No es un examen con "
          "calificación: queremos saber lo que conoces sobre alimentación. ¡Gracias por participar!", indent=False)
import json
ITEMS = json.loads((pathlib.Path(__file__).parent / "items.json").read_text(encoding="utf-8"))
for dim in ITEMS:
    heading(doc, dim["dim"], 3)
    for it in dim["items"]:
        p = para(doc, f"**{it['n']}.** {it['q']}", indent=False, spacing=1, c=it.get("obs"))
        p.paragraph_format.space_before = Pt(6)
        for letter, opt in zip("ABCD", it["opts"]):
            q = para(doc, f"{letter}) {opt}", indent=False, spacing=1)
            q.paragraph_format.left_indent = Cm(1.27)
        k = para(doc, f"*Clave: {it['key']}*", indent=False, spacing=1, size=10)
        k.paragraph_format.left_indent = Cm(1.27)

heading(doc, "Anexo B. Formato de consentimiento informado y asentimiento", 2)
para(doc, "==El formato debe incluir:== título del proyecto; nombre y contacto de la investigadora y del director; propósito del "
          "estudio en lenguaje claro; procedimientos (preprueba, ocho cápsulas en cuatro semanas, posprueba); duración; riesgos "
          "(mínimos) y beneficios; carácter voluntario y derecho a retirarse sin consecuencias; confidencialidad y uso de códigos; "
          "uso de los datos solo con fines académicos; espacio para nombre y firma de la madre, el padre o el tutor y fecha. El "
          "asentimiento del menor debe redactarse en lenguaje adecuado para niñas y niños de 10 y 11 años, con opciones «Sí quiero "
          "participar» y «No quiero participar».",
     indent=False, c="Elabora los formatos completos con base en esta lista y anéxalos.")
heading(doc, "Anexo C. Lista de cotejo de fidelidad de implementación", 2)
table(doc, 6, "Lista de cotejo de implementación por sesión",
      ["Criterio", "Sí", "No", "Observaciones"],
      [["La cápsula se presentó completa (video/infografía)", "", "", ""],
       ["Se realizó la actividad interactiva", "", "", ""],
       ["Se realizó la actividad de recuperación de la cápsula anterior", "", "", ""],
       ["Número de estudiantes que completaron la cápsula", "", "", "__ / 25"],
       ["Incidencias técnicas", "", "", ""],
       ["Duración real de la sesión (min)", "", "", ""]],
      widths=[8, 1.5, 1.5, 5])

doc.save(OUT)
print("OK", OUT)
