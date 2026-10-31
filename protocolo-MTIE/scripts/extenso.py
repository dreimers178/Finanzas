"""Trabajo en extenso para el CITEAA 2026: versión con autores y versión anónima (doble ciego)."""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))
import apa
from apa import *  # noqa
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH

OUTDIR = pathlib.Path(__file__).resolve().parents[1] / "congreso"
TITULO = "Nutrición en el aula: taller de microaprendizaje en línea para quinto de primaria (diseño pretest-postest)"
EJE = "Eje temático: Metodologías activas y pedagogías emergentes"
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

TALLER = "taller educativo en línea basado en microaprendizaje"


def setup():
    apa.FONT = "Arial"
    doc = new_document()  # tamaño carta, márgenes de 2.54 cm
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for i in range(1, 5):
        hs = doc.styles[f"Heading {i}"]
        hs.font.size = Pt(11)
        hs.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        hs.paragraph_format.space_before = Pt(10)
        hs.paragraph_format.space_after = Pt(4)
    doc.styles["Heading 2"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return doc


def J(doc, text, indent=True):
    p = para(doc, text, indent=indent)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def build(anonimo: bool):
    doc = setup()
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.core_properties.title = TITULO
    para(doc, "I Congreso Internacional de Tecnología Educativa y Ambientes de Aprendizaje", indent=False, align="center")
    para(doc, f"**{TITULO.upper()}**", indent=False, align="center")
    if not anonimo:
        para(doc, "Anakaren Ernestina Reimers Reyes, Eduardo Rivera Arteaga", indent=False, align="center")
        para(doc, "Universidad Autónoma de Zacatecas «Francisco García Salinas», México", indent=False, align="center")
        para(doc, "ana.reimers@gmail.com · edrivera@uaz.edu.mx", indent=False, align="center", size=10)
    para(doc, EJE, indent=False, align="center", size=10)
    if not anonimo:
        p = para(doc, f"*{PATROCINIO}*", indent=False, align="center", size=9)
    para(doc, "", indent=False)

    heading(doc, "1. INTRODUCCIÓN", 2)
    J(doc, "El sobrepeso y la obesidad infantil son uno de los principales retos de salud pública en México. De acuerdo con la "
           "Encuesta Nacional de Salud y Nutrición Continua 2020-2022, el 19.2 % de la población escolar presenta sobrepeso y el "
           "18.1 % obesidad, es decir, 37.3 % en conjunto; además, en los escolares varones la prevalencia de obesidad aumentó 5.8 "
           "puntos porcentuales entre 2006 y 2020-2022 (Shamah-Levy et al., 2023). Esta situación persiste a pesar de la larga "
           "trayectoria de políticas y programas de alimentación y nutrición en el país (Barquera et al., 2001). En Zacatecas, "
           "López Hernández et al. (2024) reportaron en una primaria del municipio de Guadalupe que el 45.9 % de los escolares "
           "presentaba sobrepeso u obesidad y que el 66.4 % tenía conocimientos insuficientes en nutrición.")
    J(doc, "Ante este panorama, resulta pertinente diseñar propuestas educativas sistemáticas, atractivas y apoyadas en recursos "
           "digitales que fortalezcan, desde la escuela, el conocimiento nutricional de los estudiantes. La propuesta se inscribe "
           "en el ámbito de la tecnología educativa porque articula la educación para la salud con el microaprendizaje y con "
           "recursos digitales de bajo costo, replicables en otros contextos escolares, y se vincula con los Objetivos de "
           "Desarrollo Sostenible 3 (salud y bienestar) y 4 (educación de calidad) de la Agenda 2030 (Organización de las "
           "Naciones Unidas [ONU], 2015).")

    heading(doc, "2. MARCO TEÓRICO", 2)
    J(doc, "La escuela es un entorno privilegiado para la educación alimentaria y nutricional. Pérez-Rodrigo y Aranceta (2001) "
           "señalan que las intervenciones escolares con mejores resultados integran varios componentes y no se limitan a "
           "transmitir información, y Lytle y Achterberg (1995) identificaron, entre los elementos de los programas exitosos, "
           "la fundamentación teórica, la participación de la familia, la intervención en el entorno escolar y el uso de "
           "estrategias interactivas, incluidas las mediadas por computadora. En México, el programa SaludArte mejoró los "
           "conocimientos en nutrición de 1,620 escolares de la Ciudad de México (Shamah-Levy et al., 2017), y una intervención "
           "de educación nutricional en línea con escolares de Veracruz incrementó significativamente su conocimiento sobre "
           "alimentación y nutrición frente a un grupo control (Cruz Lara y Macossay Moreno, 2022).")
    J(doc, "Por su parte, el microaprendizaje organiza los contenidos en unidades breves, con un objetivo cada una, y se ha "
           "asociado con una formación ágil y con la reducción de la carga cognitiva (Betancur-Chicué y García-Valcárcel "
           "Muñoz-Repiso, 2023). Se fundamenta en la teoría de la carga cognitiva, según la cual segmentar el contenido y eliminar "
           "elementos irrelevantes libera la memoria de trabajo para procesar lo esencial (Sweller et al., 2019); en los principios "
           "de coherencia, señalización y segmentación del aprendizaje multimedia (Mayer, 2009), y en la evidencia sobre el "
           "aprendizaje espaciado, que favorece la consolidación de la memoria cuando los contenidos breves se distribuyen en el "
           "tiempo (Kelley y Whatson, 2013). En educación básica, la revisión de Silva et al. (2025) identifica efectos favorables "
           "sobre la motivación, la atención y la retención.")
    J(doc, "Los contenidos del taller se basan en el Plato del Bien Comer Saludable y Sostenible, que integra cinco grupos de "
           "alimentos y recomienda el consumo de agua simple y la reducción de productos ultraprocesados y bebidas azucaradas "
           "(Secretaría de Salud [SSA] et al., 2025), y se articulan con el campo formativo «Saberes y pensamiento científico» de "
           "la Fase 5 del plan de estudios, que incluye las características de una dieta correcta y los riesgos del consumo de "
           "ultraprocesados (Secretaría de Educación Pública [SEP], 2024).")
    J(doc, "Si bien existe evidencia sobre la eficacia de la educación nutricional escolar y sobre el potencial del "
           "microaprendizaje, son escasos los estudios que articulan ambos campos con estudiantes de primaria en México, en "
           "particular en Zacatecas. Por ello, la pregunta de investigación es: ¿se incrementa el nivel de conocimiento "
           "nutricional de los estudiantes de quinto grado de primaria después de su participación en un " + TALLER + "? La "
           "hipótesis de investigación plantea que el puntaje de conocimiento nutricional será significativamente mayor en la "
           "posprueba que en la preprueba.")

    heading(doc, "3. METODOLOGÍA", 2)
    heading(doc, "Objetivos", 3)
    J(doc, "**Objetivo general:** incrementar el conocimiento nutricional de los estudiantes de 5.º grado de primaria del "
           "Instituto Educativo Ammadeus mediante la implementación de un " + TALLER + ".")
    J(doc, "**Objetivos específicos:** (a) identificar el nivel inicial de conocimiento nutricional de los estudiantes mediante "
           "una preprueba; (b) diseñar y desarrollar, con base en el modelo ADDIE, las cápsulas digitales y las actividades "
           "interactivas del taller, alineadas con las Guías Alimentarias Saludables y Sostenibles para la Población Mexicana, y "
           "(c) comparar el nivel de conocimiento nutricional, total y por dimensión, antes y después de la implementación del "
           "taller.")

    heading(doc, "Diseño", 3)
    J(doc, "Investigación aplicada, con enfoque cuantitativo (Hernández-Sampieri y Mendoza Torres, 2018) y diseño "
           "cuasiexperimental de preprueba y posprueba con un solo grupo intacto (O₁ X O₂): se mide el conocimiento nutricional, se "
           "implementa el taller y se vuelve a medir con el mismo instrumento. El grupo corresponde a un grupo escolar ya "
           "constituido, sin asignación aleatoria.")
    heading(doc, "Participantes", 3)
    J(doc, "Participarán 25 estudiantes (niñas y niños de 10 y 11 años) del grupo 5.º «B» del Instituto Educativo Ammadeus, "
           "institución privada de Guadalupe, Zacatecas, seleccionados mediante muestreo no probabilístico por conveniencia. Con "
           "n = 25, una prueba *t* para muestras relacionadas unilateral (α = .05) tiene una potencia de .80 para detectar un "
           "efecto de magnitud mediana (*d* ≈ 0.50; Cohen, 1988). Se incluirá a quienes cuenten con el consentimiento informado de "
           "sus madres, padres o tutores y otorguen su asentimiento; se eliminará a quienes no respondan alguna de las dos "
           "mediciones o completen menos de seis de las ocho cápsulas.")
    heading(doc, "Técnica e instrumento", 3)
    J(doc, "Se aplicará una prueba de conocimientos mediante el *Cuestionario de conocimiento nutricional para estudiantes de "
           "quinto grado de primaria*, elaborado ex profeso: 24 reactivos de opción múltiple con cuatro opciones y una sola "
           "respuesta correcta, redactados conforme a las pautas de Haladyna et al. (2002) y organizados en cuatro dimensiones "
           "(Tabla 1). El puntaje total va de 0 a 24.")
    table(doc, 1, "Dimensiones del cuestionario de conocimiento nutricional", ["Dimensión", "Reactivos"], [
        ["D1. Principios de alimentación saludable y selección de alimentos", "1–6"],
        ["D2. Grupos de alimentos y clasificación", "7–12"],
        ["D3. Aportes y función nutricional básica de los alimentos", "13–18"],
        ["D4. Aplicación de las recomendaciones alimentarias en situaciones cotidianas", "19–24"],
    ], widths=[12.5, 3])
    J(doc, "La validez de contenido se valora mediante juicio de expertos en nutrición, educación primaria y metodología "
           "(Escobar-Pérez y Cuervo-Martínez, 2008), con el coeficiente V de Aiken (Aiken, 1985) y su intervalo de confianza al "
           "95 % (Penfield y Giacobbi, 2004). La confiabilidad se estimará en una prueba piloto con un grupo no participante "
           "mediante el coeficiente Kuder-Richardson 20 (Kuder y Richardson, 1937).")
    heading(doc, "Procedimiento: diseño instruccional del taller", 3)
    J(doc, "El taller se diseñó con el modelo ADDIE (Branch, 2009). En el **análisis** se identificaron las necesidades, el "
           "perfil de los estudiantes y los recursos tecnológicos del plantel. En el **diseño** se definieron ocho cápsulas, dos por "
           "semana durante cuatro semanas, cada una con un objetivo y alineada con una dimensión del instrumento (Tabla 2). Cada "
           "cápsula sigue una secuencia de cuatro momentos: recuperación de la cápsula anterior, microcontenido de 3 a 5 minutos "
           "(video o infografía), actividad interactiva con retroalimentación inmediata y reto de cierre. En el **desarrollo** se "
           "producen los videos, las infografías y las actividades interactivas. La **implementación** se realizará en el aula "
           "digital, con registro de las cápsulas completadas por cada estudiante, y la **evaluación** será diagnóstica (preprueba), "
           "formativa (actividades de cada cápsula) y sumativa (posprueba).")
    table(doc, 2, "Cápsulas del taller", ["Semana", "Cápsulas", "Dimensión"], [
        ["1", "C1. ¿Qué es comer saludable?\nC2. Agua simple y sellos de advertencia", "D1"],
        ["2", "C3. Mi Plato del Bien Comer\nC4. ¿A qué grupo pertenece?", "D2"],
        ["3", "C5. Vitaminas, minerales y fibra\nC6. Proteína vegetal y grasas saludables", "D3"],
        ["4", "C7. Armo mi refrigerio escolar\nC8. Reto final: mi comida del día", "D4"],
    ], widths=[2, 10.5, 3])
    heading(doc, "Análisis de datos", 3)
    J(doc, "Se obtendrá la estadística descriptiva de la preprueba y la posprueba, total y por dimensión. La normalidad de las "
           "diferencias se evaluará con la prueba de Shapiro-Wilk; según el resultado, se aplicará la prueba *t* de Student para "
           "muestras relacionadas o la prueba de rangos con signo de Wilcoxon (α = .05), y se reportará el tamaño del efecto "
           "(Cohen, 1988).")
    heading(doc, "Consideraciones éticas", 3)
    J(doc, "Se contará con la autorización de la dirección escolar, el consentimiento informado de madres, padres o tutores y el "
           "asentimiento de los estudiantes. La participación es voluntaria, los datos se identificarán con códigos y se usarán "
           "solo con fines académicos. Al concluir el estudio, los materiales del taller se compartirán con los demás grupos de "
           "5.º grado.")

    heading(doc, "4. RESULTADOS Y DISCUSIÓN", 2)
    J(doc, "El proyecto se encuentra en la etapa de planteamiento metodológico. Se concluyeron el planteamiento del problema y el "
           "fundamento teórico; se diseñó el instrumento de 24 reactivos, alineado con las Guías Alimentarias 2025, y el diseño "
           "instruccional del taller, con ocho cápsulas alineadas con las cuatro dimensiones del conocimiento nutricional. El "
           "instrumento se encuentra en validación de contenido mediante juicio de expertos. De forma tentativa, la prueba piloto "
           "se realizará en diciembre de 2026, la preprueba del 18 al 22 de enero de 2027, el taller del 25 de enero al 19 de "
           "febrero de 2027 (iniciando en la semana del Día del Nutriólogo) y la posprueba del 22 al 26 de febrero de 2027.")
    J(doc, "Una vez obtenidos los datos, los resultados se compararán con los de intervenciones previas que reportaron "
           "incrementos significativos en el conocimiento nutricional de escolares, como la intervención en línea de Cruz Lara y "
           "Macossay Moreno (2022) y el programa SaludArte (Shamah-Levy et al., 2017). La comparación permitirá valorar si el "
           "formato de microaprendizaje produce cambios de magnitud similar con una intervención breve, de cuatro semanas, y "
           "discutir las condiciones del contexto zacatecano que pudieran explicar las diferencias.")

    heading(doc, "5. CONCLUSIONES", 2)
    J(doc, "La pregunta de investigación se responderá con la comparación de la preprueba y la posprueba, por lo que aún no es "
           "posible afirmar si el taller incrementa el conocimiento nutricional. En cuanto al cumplimiento del objetivo, se ha "
           "avanzado en el diseño del instrumento y del taller (objetivo específico b); la identificación del nivel inicial y la "
           "comparación pretest-postest se realizarán en 2027.")
    J(doc, "Se espera que la propuesta aporte un taller replicable y de bajo costo para la educación primaria, un instrumento con "
           "evidencias de validez y confiabilidad para estudiantes de 5.º grado, y evidencia sobre el microaprendizaje aplicado a la "
           "educación nutricional en el contexto zacatecano. Entre las limitaciones destacan la ausencia de grupo control, que "
           "impide descartar amenazas como la historia, la maduración o el efecto de prueba (Campbell y Stanley, 1966); la "
           "muestra por conveniencia, y la medición del conocimiento y no de las conductas alimentarias. Se recomienda que "
           "estudios posteriores incluyan un grupo de comparación y una medición de seguimiento, y que exploren el efecto del "
           "microaprendizaje sobre las conductas alimentarias y su aplicación en otros grados y contextos escolares.")

    heading(doc, "6. REFERENCIAS BIBLIOGRÁFICAS", 2)
    refs = [
        "Aiken, L. R. (1985). Three coefficients for analyzing the reliability and validity of ratings. *Educational and Psychological Measurement, 45*(1), 131–142. https://doi.org/10.1177/0013164485451012",
        "Barquera, S., Rivera-Dommarco, J., y Gasca-García, A. (2001). Políticas y programas de alimentación y nutrición en México. *Salud Pública de México, 43*(5), 464–477. https://doi.org/10.1590/S0036-36342001000500011",
        "Betancur-Chicué, V., y García-Valcárcel Muñoz-Repiso, A. (2023). Características del diseño de estrategias de microaprendizaje en escenarios educativos: Revisión sistemática. *RIED. Revista Iberoamericana de Educación a Distancia, 26*(1), 201–222. https://doi.org/10.5944/ried.26.1.34056",
        "Branch, R. M. (2009). *Instructional design: The ADDIE approach*. Springer. https://doi.org/10.1007/978-0-387-09506-6",
        "Campbell, D. T., y Stanley, J. C. (1966). *Experimental and quasi-experimental designs for research*. Rand McNally.",
        "Cohen, J. (1988). *Statistical power analysis for the behavioral sciences* (2.ª ed.). Lawrence Erlbaum Associates.",
        "Cruz Lara, N. M., y Macossay Moreno, C. (2022). Intervención de educación nutricional en línea a población infantil en el estado de Veracruz. *Horizonte Sanitario, 21*(3), 403–409. https://doi.org/10.19136/hs.a21n3.4893",
        "Escobar-Pérez, J., y Cuervo-Martínez, Á. (2008). Validez de contenido y juicio de expertos: Una aproximación a su utilización. *Avances en Medición, 6*, 27–36.",
        "Haladyna, T. M., Downing, S. M., y Rodriguez, M. C. (2002). A review of multiple-choice item-writing guidelines for classroom assessment. *Applied Measurement in Education, 15*(3), 309–333. https://doi.org/10.1207/S15324818AME1503_5",
        "Hernández-Sampieri, R., y Mendoza Torres, C. P. (2018). *Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta*. McGraw-Hill Interamericana.",
        "Kelley, P., y Whatson, T. (2013). Making long-term memories in minutes: A spaced learning pattern from memory research in education. *Frontiers in Human Neuroscience, 7*, Artículo 589. https://doi.org/10.3389/fnhum.2013.00589",
        "Kuder, G. F., y Richardson, M. W. (1937). The theory of the estimation of test reliability. *Psychometrika, 2*(3), 151–160. https://doi.org/10.1007/BF02288391",
        "López Hernández, P. Y., Franco Trejo, C. S., Galicia Rodríguez, A. G., Argüelles Estrada, B. C., y Ornelas Pargas, A. C. (2024). Sobrepeso, obesidad y educación nutricional en una escuela primaria de Zacatecas. *Revista Iberoamericana de Ciencias, 11*(1). https://www.reibci.org/publicados/2024/mzo/5300113.pdf",
        "Lytle, L., y Achterberg, C. (1995). Changing the diet of America's children: What works and why? *Journal of Nutrition Education, 27*(5), 250–260. https://doi.org/10.1016/S0022-3182(12)80793-X",
        "Mayer, R. E. (2009). *Multimedia learning* (2.ª ed.). Cambridge University Press. https://doi.org/10.1017/CBO9780511811678",
        "Organización de las Naciones Unidas. (2015). *Transformar nuestro mundo: La Agenda 2030 para el Desarrollo Sostenible* (A/RES/70/1). https://sdgs.un.org/2030agenda",
        "Penfield, R. D., y Giacobbi, P. R., Jr. (2004). Applying a score confidence interval to Aiken's item content-relevance index. *Measurement in Physical Education and Exercise Science, 8*(4), 213–225. https://doi.org/10.1207/s15327841mpee0804_3",
        "Pérez-Rodrigo, C., y Aranceta, J. (2001). School-based nutrition education: Lessons learned and new perspectives. *Public Health Nutrition, 4*(1A), 131–139. https://doi.org/10.1079/PHN2000108",
        "Secretaría de Educación Pública. (2024). *Programa de estudio para la educación primaria: Programa sintético de la Fase 5*. https://educacionbasica.sep.gob.mx/wp-content/uploads/2024/06/Programa_Sintetico_Fase_5.pdf",
        "Secretaría de Salud, Instituto Nacional de Salud Pública y Fondo de las Naciones Unidas para la Infancia. (2025). *Guías alimentarias saludables y sostenibles para la población mexicana* (2.ª ed.).",
        "Shamah-Levy, T., Cuevas-Nasu, L., Gómez-Acosta, L. M., Morales-Ruan, M. C., Méndez Gómez-Humarán, I., Robles-Villaseñor, M. N., y Hernández-Ávila, M. (2017). Efecto del programa de servicios SaludArte en los componentes de alimentación y nutrición en escolares de la Ciudad de México. *Salud Pública de México, 59*(6), 621–629. https://doi.org/10.21149/8116",
        "Shamah-Levy, T., Gaona-Pineda, E. B., Cuevas-Nasu, L., Morales-Ruan, C., Valenzuela-Bravo, D. G., Méndez-Gómez Humarán, I., y Ávila-Arcos, M. A. (2023). Prevalencias de sobrepeso y obesidad en población escolar y adolescente de México. Ensanut Continua 2020-2022. *Salud Pública de México, 65*(Supl. 1), S218–S224. https://doi.org/10.21149/14762",
        "Silva, E. S., Pires da Costa, W., de Lima, J. C., y Ferreira, J. C. (2025). Contribution of microlearning in basic education: A systematic review. *Education Sciences, 15*(3), Artículo 302. https://doi.org/10.3390/educsci15030302",
        "Sweller, J., van Merriënboer, J. J. G., y Paas, F. (2019). Cognitive architecture and instructional design: 20 years later. *Educational Psychology Review, 31*(2), 261–292. https://doi.org/10.1007/s10648-019-09465-5",
    ]
    for r in refs:
        p = ref(doc, r)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    name = ("MetodologiasActivasPedagogiasEmergentes_SinAutores_Extenso.docx" if anonimo
            else "MetodologiasActivasPedagogiasEmergentes_Reimers_Extenso.docx")
    doc.save(OUTDIR / name)
    return OUTDIR / name


for a in (False, True):
    print("OK", build(a))
