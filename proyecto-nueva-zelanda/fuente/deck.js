const pptxgen = require("pptxgenjs");
const fs = require("fs");
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/fe2298e2-c83b-4749-b1e4-9d67924edffe_5c4f3dbb-923f-40a2-b725-558fb0298d22/pptx/scripts/apply_theme.js");
const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Aotearoa",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "14262B", lt1: "FFFFFF", dk2: "0E4D45", lt2: "EAF3EE",
    accent1: "2F8F5B", accent2: "1E6F8C", accent3: "E2A23B", accent4: "C8553D",
    accent5: "7FB69A", accent6: "5B6B70", hlink: "1E6F8C", folHlink: "5B6B70",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Países que construyen bienestar: Nueva Zelanda";
pres.author = "José Luis De Anda Reimers y Pablo De Santiago Hurtado Martin Del Campo";
pres.company = "Instituto Educativo Ammadeus";
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "tag", type: "body", x: 0.5, y: 0.22, w: 6, h: 0.3, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2, margin: 0, valign: "middle" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.5, w: 9, h: 0.65, fontSize: 26, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { text: { text: "Instituto Educativo Ammadeus · Educación Financiera · Nueva Zelanda", options: { x: 0.5, y: 5.28, w: 7, h: 0.25, fontSize: 9, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 9.1, y: 5.28, w: 0.4, h: 0.25, fontSize: 9, color: C.accent6, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "OSCURA",
  background: { color: C.text2 },
  objects: [],
});

// ---------- Helpers ----------
async function icon(Comp, hex) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function iconCircle(slide, Comp, x, y, d, fill, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " círculo" });
  const pad = d * 0.24;
  slide.addImage({ data: await icon(Comp, "FFFFFF"), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name, objectName: name + " ícono" });
}
function card(slide, x, y, w, h, fill, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { type: "none" }, objectName: name || "tarjeta" });
}
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top" }, opts));
}
function source(slide, text) {
  txt(slide, "Fuente: " + text, { x: 0.5, y: 5.02, w: 9, h: 0.24, fontSize: 9, color: C.accent6, italic: true, objectName: "fuente" });
}
function header(slide, tag, title) {
  slide.addText(tag, { placeholder: "tag" });
  slide.addText(title, { placeholder: "title" });
}
function addLogo(slide, x, y, d) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.background1 }, line: { type: "none" }, objectName: "fondo logo" });
  const pad = d * 0.13;
  slide.addImage({ data: "image/png;base64," + fs.readFileSync("logo.png").toString("base64"), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: "Logo del Instituto Educativo Ammadeus", objectName: "logo Ammadeus" });
}
// five-question "ficha" used on every model slide
function ficha(slide, x, y, w, h, items) {
  card(slide, x, y, w, h, C.background2, "ficha del modelo");
  const runs = [];
  items.forEach(([q, a], i) => {
    runs.push({ text: q, options: { bold: true, color: C.accent1, fontSize: 11, breakLine: true } });
    runs.push({ text: a, options: { color: C.text1, fontSize: 12, breakLine: i < items.length - 1, paraSpaceAfter: 5 } });
  });
  txt(slide, runs, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: h - 0.3, objectName: "ficha texto" });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  catAxisLabelColor: H.dk1, valAxisLabelColor: H.accent6, dataLabelColor: H.dk1, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 11 };

(async () => {
  const flag = "image/png;base64," + fs.readFileSync("flag.png").toString("base64");
  const map = "image/png;base64," + fs.readFileSync("map.png").toString("base64");

  // ================= 1. PORTADA =================
  pres.addSection({ title: "Portada" });
  let s = pres.addSlide({ masterName: "OSCURA", sectionTitle: "Portada" });
  addLogo(s, 0.6, 0.3, 0.95);
  txt(s, "INSTITUTO EDUCATIVO AMMADEUS", { x: 1.75, y: 0.3, w: 5, h: 0.95, valign: "middle", fontSize: 13, bold: true, color: C.accent3, charSpacing: 3 });
  txt(s, [
    { text: "Países que construyen bienestar:", options: { breakLine: true } },
    { text: "modelos para una mejor sociedad", options: { color: C.accent5 } },
  ], { x: 0.6, y: 1.45, w: 5.7, h: 1.2, fontSize: 27, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "top" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.85, w: 4.2, h: 0.45, rectRadius: 0.22, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "etiqueta país" });
  txt(s, "País investigado: Nueva Zelanda (Aotearoa)", { x: 0.6, y: 2.85, w: 4.2, h: 0.45, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, [
    { text: "Materia: ", options: { bold: true, color: C.accent5 } }, { text: "Educación Financiera · 2.º de secundaria", options: { breakLine: true } },
    { text: "Maestra: ", options: { bold: true, color: C.accent5 } }, { text: "Samantha Rubí Lara Ceh", options: { breakLine: true } },
    { text: "Alumnos: ", options: { bold: true, color: C.accent5 } }, { text: "José Luis De Anda Reimers", options: { breakLine: true } },
    { text: "Pablo De Santiago Hurtado Martin Del Campo", options: { breakLine: true, indentLevel: 0 } },
    { text: "Noviembre de 2026", options: { color: C.accent5 } },
  ], { x: 0.6, y: 3.5, w: 5.6, h: 1.65, fontSize: 13, color: C.background1, paraSpaceAfter: 3 });
  s.addImage({ data: flag, x: 6.7, y: 1.8, w: 2.8, h: 1.4, altText: "Bandera de Nueva Zelanda", objectName: "bandera", shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.35 } });
  txt(s, "Bandera de Nueva Zelanda", { x: 6.7, y: 3.3, w: 2.8, h: 0.3, fontSize: 10, italic: true, color: C.accent5, align: "center" });
  s.addNotes("JOSÉ LUIS: Buenos días, maestra Samantha y compañeros. Nuestro proyecto se llama 'Países que construyen bienestar: modelos para una mejor sociedad'. PABLO: Somos José Luis De Anda y Pablo De Santiago, y el país que investigamos es Nueva Zelanda, que en idioma maorí se llama Aotearoa.");

  // ================= 2. CONTENIDO =================
  pres.addSection({ title: "Introducción" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Introducción" });
  header(s, "INTRODUCCIÓN", "¿Qué vamos a presentar?");
  txt(s, [{ text: "Pregunta guía: ", options: { bold: true, color: C.accent1 } }, { text: "¿Qué hace Nueva Zelanda para que su gente viva mejor… y qué podría copiar México?" }],
    { x: 0.5, y: 1.3, w: 9, h: 0.4, fontSize: 15 });
  const agenda = [
    ["A", "Conociendo el país", "Ubicación, población, economía e indicadores", fa.FaGlobeAmericas, C.accent2],
    ["B", "Modelos de bienestar", "Educación, salud, empleo y pensiones, vivienda y medio ambiente", fa.FaHandsHelping, C.accent1],
    ["C", "¿Funcionaría en México?", "Beneficios, obstáculos y nuestra propuesta", fa.FaBalanceScale, C.accent4],
    ["D", "Conclusiones", "Lo que aprendimos y lo que México puede aprender", fa.FaLightbulb, C.accent3],
  ];
  for (let i = 0; i < 4; i++) {
    const [L, t, d, ic, col] = agenda[i];
    const x = 0.5 + i * 2.3;
    card(s, x, 1.95, 2.1, 2.85, C.background2, "tarjeta " + L);
    await iconCircle(s, ic, x + 0.2, 2.15, 0.7, col, "sección " + L);
    txt(s, L, { x: x + 1.3, y: 2.15, w: 0.6, h: 0.7, fontSize: 36, bold: true, color: col, align: "right", fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, t, { x: x + 0.2, y: 3.0, w: 1.7, h: 0.6, fontSize: 15, bold: true, color: C.text2 });
    txt(s, d, { x: x + 0.2, y: 3.65, w: 1.7, h: 1.0, fontSize: 12, color: C.text1 });
  }
  s.addNotes("JOSÉ LUIS: Nuestra presentación tiene cuatro partes. A: conocer el país. B, la parte principal: los modelos que Nueva Zelanda usa para mejorar la vida de su gente en educación, salud, empleo y pensiones, y vivienda y medio ambiente. C: analizar si eso podría funcionar en México. D: nuestras conclusiones. La pregunta que nos guió fue: ¿qué hace Nueva Zelanda para que su gente viva mejor y qué podría copiar México?");

  // ================= 3. UBICACIÓN Y BANDERA =================
  pres.addSection({ title: "A. Conociendo el país" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "A. Conociendo el país" });
  header(s, "A · CONOCIENDO EL PAÍS", "Nueva Zelanda: ¿dónde está?");
  s.addImage({ data: map, x: 0.5, y: 1.3, w: 4.9, h: 3.5, altText: "Mapa del océano Pacífico con México y Nueva Zelanda; distancia aproximada de 11,100 km", objectName: "mapa", rounding: false });
  s.addImage({ data: flag, x: 5.75, y: 1.3, w: 1.5, h: 0.75, altText: "Bandera de Nueva Zelanda", objectName: "bandera" });
  txt(s, "La bandera: la «Union Jack» recuerda su historia británica; las 4 estrellas son la Cruz del Sur.", { x: 7.4, y: 1.27, w: 2.1, h: 0.85, fontSize: 10.5, color: C.text1, valign: "middle" });
  txt(s, [
    { text: "Continente: ", options: { bold: true } }, { text: "Oceanía, en el Pacífico Sur", options: { breakLine: true } },
    { text: "Capital: ", options: { bold: true } }, { text: "Wellington · Ciudad mayor: Auckland", options: { breakLine: true } },
    { text: "Territorio: ", options: { bold: true } }, { text: "2 islas principales (Norte y Sur), ≈ 268,000 km² (¡7 veces más chico que México!)", options: { breakLine: true } },
    { text: "Vecino: ", options: { bold: true } }, { text: "Australia, a unos 2,000 km", options: { breakLine: true } },
    { text: "Idiomas oficiales: ", options: { bold: true } }, { text: "inglés, maorí y lengua de señas", options: { breakLine: true } },
    { text: "Moneda: ", options: { bold: true } }, { text: "dólar neozelandés (NZ$)" },
  ], { x: 5.75, y: 2.25, w: 3.75, h: 2.6, fontSize: 12.5, paraSpaceAfter: 5 });
  source(s, "Stats NZ; Gobierno de Nueva Zelanda (govt.nz). Mapa elaborado con datos de Natural Earth.");
  s.addNotes("JOSÉ LUIS: Nueva Zelanda está en Oceanía, en el océano Pacífico Sur, a unos 11 mil kilómetros de México. Su capital es Wellington y su ciudad más grande es Auckland. Tiene dos islas principales, la Isla Norte y la Isla Sur, y es unas siete veces más pequeña que México. En su bandera aparece la bandera británica, porque fue colonia del Reino Unido, y cuatro estrellas que forman la Cruz del Sur, una constelación que se ve en el hemisferio sur. Tiene tres idiomas oficiales: inglés, maorí —el idioma del pueblo originario— y lengua de señas. Su moneda es el dólar neozelandés.");

  // ================= 4. POBLACIÓN =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "A. Conociendo el país" });
  header(s, "A · CONOCIENDO EL PAÍS", "Población: pocos habitantes, mitad y mitad");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Población por sexo", labels: ["Mujeres", "Hombres"], values: [2678500, 2646200] }], {
    x: 0.5, y: 1.3, w: 4.0, h: 3.6, holeSize: 55, chartColors: [H.accent1, H.accent2], showPercent: true, showValue: false, dataLabelFormatCode: "0.0%", showLegend: true, legendPos: "b",
    showTitle: true, title: "Población por sexo (junio 2025)", titleFontSize: 12, titleColor: H.dk1, ...chartText, dataLabelColor: "FFFFFF", dataLabelFontSize: 13, dataLabelFontBold: true, legendFontSize: 11,
    objectName: "gráfica población por sexo",
  });
  card(s, 4.9, 1.3, 4.6, 1.3, C.background2, "total población");
  txt(s, "5,324,700", { x: 5.1, y: 1.4, w: 4.2, h: 0.7, fontSize: 40, bold: true, color: C.accent1, fontFace: THEME.headFontFace });
  txt(s, "habitantes en junio de 2025 (crece 0.7 % al año)", { x: 5.1, y: 2.1, w: 4.2, h: 0.35, fontSize: 12.5 });
  txt(s, [
    { text: "Mujeres: ", options: { bold: true, color: C.accent1 } }, { text: "2,678,500 (50.3 %)", options: { breakLine: true } },
    { text: "Hombres: ", options: { bold: true, color: C.accent2 } }, { text: "2,646,200 (49.7 %)", options: { breakLine: true } },
    { text: "Edad mediana: ", options: { bold: true } }, { text: "39.0 años mujeres · 37.4 años hombres", options: { breakLine: true } },
    { text: "Comparación: ", options: { bold: true, color: C.accent4 } }, { text: "México tiene más de 130 millones; ¡unas 25 veces más personas!" },
  ], { x: 4.9, y: 2.8, w: 4.6, h: 2.1, fontSize: 13, paraSpaceAfter: 6 });
  source(s, "Stats NZ, National population estimates: At 30 June 2025 (2025); CONAPO, proyecciones de población de México.");
  s.addNotes("PABLO: En junio de 2025 Nueva Zelanda tenía 5 millones 324 mil 700 habitantes. Hay casi la misma cantidad de mujeres que de hombres: 50.3 % mujeres y 49.7 % hombres. La edad mediana, es decir, la edad que divide a la población en dos mitades, es de 39 años para mujeres y 37 para hombres. Para darnos una idea, México tiene más de 130 millones de habitantes, unas 25 veces más. Esto es importante: organizar programas para 5 millones de personas no es lo mismo que para 130 millones.");

  // ================= 5. ESPERANZA DE VIDA Y EDADES =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "A. Conociendo el país" });
  header(s, "A · CONOCIENDO EL PAÍS", "¿Cuánto viven y qué edades tienen?");
  const le = [["81.8", "años · promedio general", C.accent1], ["83.5", "años · mujeres", C.accent1], ["80.1", "años · hombres", C.accent2]];
  le.forEach(([n, l, col], i) => {
    const y = 1.3 + i * 1.05;
    card(s, 0.5, y, 3.6, 0.9, C.background2, "esperanza de vida " + l);
    txt(s, n, { x: 0.7, y: y + 0.1, w: 1.5, h: 0.7, fontSize: 32, bold: true, color: col, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, l, { x: 2.2, y: y + 0.1, w: 1.8, h: 0.7, fontSize: 13, valign: "middle" });
  });
  txt(s, [{ text: "Ojo: ", options: { bold: true, color: C.accent4 } }, { text: "el pueblo maorí vive en promedio 75.8 años; seis menos que el promedio." }],
    { x: 0.5, y: 4.45, w: 3.6, h: 0.55, fontSize: 11.5 });
  s.addChart(pres.charts.BAR, [{ name: "Porcentaje de la población", labels: ["0–14 años", "15–64 años", "65 años o más"], values: [18.7, 64.7, 16.6] }], {
    x: 4.5, y: 1.3, w: 5.0, h: 3.6, barDir: "col", chartColors: [H.accent2, H.accent1, H.accent3], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"',
    showTitle: true, title: "Distribución por edades (% de la población, Censo 2023)", titleFontSize: 12, titleColor: H.dk1, showLegend: false, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 75, ...chartText, objectName: "gráfica edades",
  });
  source(s, "Stats NZ, National and subnational period life tables: 2022–2024 (2025); Stats NZ, Censo 2023 (vía EHINZ).");
  s.addNotes("PABLO: La esperanza de vida es cuántos años vive en promedio una persona. En Nueva Zelanda es de 81.8 años: 83.5 para las mujeres y 80.1 para los hombres. Pero hay una diferencia importante: el pueblo maorí vive en promedio 75.8 años, seis menos. Esto nos muestra que incluso en países con mucho bienestar hay desigualdad. En la gráfica vemos las edades: 18.7 % son niños, casi 65 % está en edad de trabajar y 16.6 % son adultos mayores. Como cada vez hay más adultos mayores, el país necesita un buen sistema de pensiones, del que hablaremos más adelante.");

  // ================= 6. ECONOMÍA =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "A. Conociendo el país" });
  header(s, "A · CONOCIENDO EL PAÍS", "¿De qué vive su economía?");
  const acts = [
    [fa.FaCheese, "Lácteos", "Leche en polvo, mantequilla y queso: su exportación n.º 1"],
    [fa.FaDrumstickBite, "Carne", "Res y cordero, segunda exportación"],
    [fa.FaPlaneArrival, "Turismo", "Naturaleza, montañas y cultura maorí"],
    [fa.FaTree, "Bosques y frutas", "Madera y el famoso kiwi"],
    [fa.FaBriefcase, "Servicios", "Comercio, educación, salud, tecnología"],
  ];
  for (let i = 0; i < acts.length; i++) {
    const [ic, t, d] = acts[i];
    const y = 1.3 + i * 0.73;
    await iconCircle(s, ic, 0.5, y, 0.55, C.accent1, t);
    txt(s, [{ text: t, options: { bold: true, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 11.5 } }], { x: 1.2, y: y - 0.03, w: 3.5, h: 0.65, fontSize: 13.5, valign: "middle" });
  }
  const eco = [["US$ 259 mil millones", "PIB 2025: valor de todo lo que el país produce en un año"], ["US$ 48,600", "PIB por persona: de los más altos del mundo"], ["NZ$ 80.7 mil millones", "Exportaciones 2025; los lácteos fueron NZ$ 24.3 mil millones (≈ 30 %)"]];
  eco.forEach(([n, d], i) => {
    const y = 1.3 + i * 1.12;
    card(s, 5.0, y, 4.5, 0.98, C.background2, "dato económico " + (i + 1));
    txt(s, n, { x: 5.2, y: y + 0.08, w: 4.1, h: 0.45, fontSize: 21, bold: true, color: C.accent2, fontFace: THEME.headFontFace });
    txt(s, d, { x: 5.2, y: y + 0.52, w: 4.1, h: 0.4, fontSize: 11.5 });
  });
  txt(s, [{ text: "Nivel de desarrollo: ", options: { bold: true, color: C.accent1 } }, { text: "economía desarrollada, con desarrollo humano «muy alto»." }], { x: 5.0, y: 4.62, w: 4.5, h: 0.35, fontSize: 12 });
  source(s, "FMI, World Economic Outlook (2025, estimación); Stats NZ, Overseas merchandise trade: December 2025 (2026).");
  s.addNotes("JOSÉ LUIS: La economía de Nueva Zelanda se basa mucho en el campo. Su exportación número uno son los lácteos: leche en polvo, mantequilla y queso. En 2025 vendió al mundo 80.7 mil millones de dólares neozelandeses, y casi 30 % fueron lácteos. También exporta carne, madera y frutas como el kiwi, y recibe muchos turistas. El PIB, que es el valor de todo lo que produce un país en un año, fue de unos 259 mil millones de dólares estadounidenses. Si lo dividimos entre todos sus habitantes, da unos 48,600 dólares por persona, uno de los más altos del mundo. Por eso se considera un país desarrollado.");

  // ================= 7. OTROS INDICADORES =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "A. Conociendo el país" });
  header(s, "A · CONOCIENDO EL PAÍS", "Otros indicadores: Nueva Zelanda vs. México");
  const th = (t) => ({ text: t, options: { bold: true, color: H.lt1, fill: { color: H.dk2 }, fontSize: 12.5, valign: "middle" } });
  const td = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 12.5, color: H.dk1, valign: "middle" }, o) });
  const rows = [
    [th("Indicador"), th("¿Qué mide?"), th("Nueva Zelanda"), th("México")],
    [td("Índice de Desarrollo Humano (ONU, 2025)", { bold: true }), td("Salud, educación e ingreso"), td("Lugar 17 (0.938)", { bold: true, color: H.accent1 }), td("Lugar 81 (0.789)")],
    [td("Percepción de corrupción (2025)", { bold: true }), td("0 = muy corrupto · 100 = muy limpio"), td("81 puntos (4.º lugar)", { bold: true, color: H.accent1 }), td("27 puntos (141.º)")],
    [td("Impuestos como % del PIB (OCDE, 2024)", { bold: true }), td("Dinero que el gobierno junta para servicios"), td("32.9 %", { bold: true, color: H.accent1 }), td("18.3 % (el más bajo de la OCDE)")],
    [td("Felicidad (Informe Mundial, 2026)", { bold: true }), td("Qué tan satisfecha está la gente con su vida"), td("Lugar 11", { bold: true, color: H.accent1 }), td("Lugar 12", { bold: true, color: H.accent4 })],
  ];
  s.addTable(rows, { x: 0.5, y: 1.3, w: 9, colW: [2.6, 2.6, 1.9, 1.9], rowH: [0.4, 0.55, 0.55, 0.55, 0.55], border: { type: "solid", pt: 0.75, color: "C9D8CF" }, fill: { color: H.lt1 }, margin: [0.04, 0.08, 0.04, 0.08], objectName: "tabla indicadores" });
  card(s, 0.5, 4.15, 9, 0.75, C.background2, "dato sorpresa");
  await iconCircle(s, fa.FaRegSmile, 0.65, 4.25, 0.55, C.accent3, "sorpresa");
  txt(s, [{ text: "¡Sorpresa! ", options: { bold: true, color: C.accent4 } }, { text: "México es casi tan feliz como Nueva Zelanda, pero tiene menos dinero público y más corrupción. La felicidad no depende solo del dinero." }],
    { x: 1.35, y: 4.2, w: 8.0, h: 0.65, fontSize: 13, valign: "middle" });
  source(s, "PNUD, Informe sobre Desarrollo Humano 2025; Transparency International, CPI 2025; OCDE, Revenue Statistics 2025; World Happiness Report 2026.");
  s.addNotes("PABLO: Además del PIB, revisamos otros indicadores. En el Índice de Desarrollo Humano de la ONU, que combina salud, educación e ingreso, Nueva Zelanda está en el lugar 17 y México en el 81. En percepción de corrupción, Nueva Zelanda tiene 81 de 100 puntos y es el cuarto país más limpio; México tiene 27. En impuestos, Nueva Zelanda junta casi 33 % de su PIB y México 18.3 %, el más bajo de la OCDE; con ese dinero los gobiernos pagan escuelas, hospitales y pensiones. Y lo que más nos sorprendió: en el Informe Mundial de la Felicidad 2026, Nueva Zelanda está en el lugar 11 y México en el 12. ¡Casi empatados! O sea, la felicidad no depende solo del dinero; también de la familia y las relaciones.");

  // ================= 8. PRESUPUESTO DE BIENESTAR =================
  pres.addSection({ title: "B. Modelos de bienestar" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · MODELOS DE BIENESTAR (APARTADO PRINCIPAL)", "La gran idea: medir el bienestar, no solo el dinero");
  card(s, 0.5, 1.3, 4.3, 1.6, C.background2, "antes");
  txt(s, [{ text: "Forma tradicional", options: { bold: true, color: C.accent6, fontSize: 13, breakLine: true } }, { text: "«¿Cuánto creció la economía (el PIB)?»", options: { fontSize: 17, bold: true, color: C.text2, fontFace: THEME.headFontFace } }],
    { x: 0.75, y: 1.45, w: 3.8, h: 1.7 });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.88, y: 1.85, w: 0.25, h: 0.5, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "flecha" });
  card(s, 5.2, 1.3, 4.3, 1.6, C.accent1, "con bienestar");
  txt(s, [{ text: "Presupuesto de Bienestar (2019)", options: { bold: true, color: C.background2, fontSize: 13, breakLine: true } }, { text: "«¿La gente vive mejor: más sana, más preparada, más segura?»", options: { fontSize: 17, bold: true, color: C.background1, fontFace: THEME.headFontFace } }],
    { x: 5.45, y: 1.45, w: 3.8, h: 1.7 });
  txt(s, [
    { text: "Fue el primer país en hacer su presupuesto nacional pensando en el bienestar: salud mental, pobreza infantil, apoyo al pueblo maorí y economía sustentable.", options: { bullet: true, breakLine: true } },
    { text: "Desde 2020, una ley pide al gobierno informar sus metas de bienestar cada año.", options: { bullet: true, breakLine: true } },
    { text: "En 2023 llegó otro gobierno con otras prioridades: estos modelos dependen de decisiones políticas.", options: { bullet: true } },
  ], { x: 0.5, y: 3.15, w: 9, h: 1.8, fontSize: 14, paraSpaceAfter: 8 });
  source(s, "The Treasury NZ, The Wellbeing Budget 2019; Public Finance Amendment Act 2020.");
  s.addNotes("JOSÉ LUIS: Ahora viene la parte principal. Antes de ver cada tema, hay una gran idea que hace diferente a Nueva Zelanda. Normalmente los países miden si les va bien con el PIB, o sea, cuánto creció la economía. En 2019, Nueva Zelanda fue el primer país en hacer un 'Presupuesto de Bienestar': antes de decidir en qué gastar, se preguntó si la gente vivía mejor. Por eso puso dinero en salud mental, pobreza infantil y apoyo al pueblo maorí. Desde 2020 una ley pide al gobierno informar sus metas de bienestar. Pero en 2023 llegó otro gobierno con otras prioridades, así que estos modelos también dependen de quién gobierna.");

  // ================= 9. EDUCACIÓN =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · TEMA 1: EDUCACIÓN", "Comida saludable en la escuela para aprender mejor");
  ficha(s, 0.5, 1.3, 4.6, 3.65, [
    ["¿En qué consiste?", "Almuerzo nutritivo gratis cada día de clases (desde 2020)."],
    ["¿Qué problema resuelve?", "Niños que llegan con hambre no se concentran ni aprenden."],
    ["¿Cómo funciona y a quién beneficia?", "El gobierno lo paga en escuelas con más necesidad: +240,000 alumnos en 1,023 escuelas (2026)."],
    ["¿Qué resultados ha dado?", "Estudios reportan menos hambre y mejor bienestar de los alumnos."],
    ["¿Qué lo hace diferente?", "Se suma a escuela pública gratis (5 a 19 años) y 20 h gratis de preescolar (3 a 5 años)."],
  ]);
  s.addChart(pres.charts.BAR, [
    { name: "Nueva Zelanda", labels: ["Matemáticas", "Lectura", "Ciencias"], values: [479, 501, 504] },
    { name: "Promedio OCDE", labels: ["Matemáticas", "Lectura", "Ciencias"], values: [472, 476, 485] },
    { name: "México", labels: ["Matemáticas", "Lectura", "Ciencias"], values: [395, 415, 410] },
  ], {
    x: 5.3, y: 1.3, w: 4.2, h: 3.65, barDir: "col", barGapWidthPct: 60, chartColors: [H.accent1, H.accent6, H.accent4], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9,
    showTitle: true, title: "Prueba PISA 2022 (puntos, alumnos de 15 años)", titleFontSize: 12, titleColor: H.dk1, showLegend: true, legendPos: "b", legendFontSize: 10,
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 600, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, dataLabelFontSize: 9, objectName: "gráfica PISA",
  });
  source(s, "Ministry of Education NZ, Healthy School Lunches (2026); OCDE, Resultados PISA 2022 (2023).");
  s.addNotes("PABLO: Nuestro primer tema es educación. Nueva Zelanda tiene un programa de almuerzos escolares saludables desde 2020. Consiste en dar comida nutritiva gratis todos los días de clases. El problema que busca resolver es que un niño con hambre no puede concentrarse ni aprender. El gobierno lo paga en las escuelas con más necesidad, y en 2026 ya llegaba a más de 240 mil alumnos en más de mil escuelas. Los estudios reportan menos hambre y mejor bienestar. Además, la escuela pública es gratuita de los 5 a los 19 años y los niños de 3 a 5 años tienen 20 horas gratis de preescolar a la semana. En la gráfica de la prueba PISA de la OCDE, que evalúa a jóvenes de 15 años, Nueva Zelanda queda arriba del promedio en las tres materias; México queda abajo.");

  // ================= 10. SALUD =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · TEMA 2: SALUD", "ACC: un seguro de accidentes para todos");
  ficha(s, 0.5, 1.3, 4.6, 3.65, [
    ["¿En qué consiste?", "ACC (desde 1974) cubre cualquier accidente, sin importar de quién fue la culpa."],
    ["¿Qué problema resuelve?", "Que un accidente deje a una familia sin atención médica ni ingresos."],
    ["¿Cómo funciona y a quién beneficia?", "Se paga con cuotas de salarios, empresas y autos, más dinero del gobierno. Cubre a residentes y ¡hasta turistas!"],
    ["¿Qué resultados ha dado?", "En 2024/25 aceptó +2.1 millones de casos: NZ$ 5 mil millones en tratamiento y NZ$ 3 mil millones en pagos."],
    ["¿Qué lo hace diferente?", "Casi ningún país lo tiene; a cambio, no se demanda por lesiones."],
  ]);
  const health = [
    [fa.FaPills, "Pharmac (desde 1993)", "Una sola agencia compra las medicinas para todo el país y negocia precios bajos."],
    [fa.FaHospital, "Hospitales públicos", "Gratis para residentes; se pagan con impuestos."],
    [fa.FaChild, "Niños menores de 14", "Consulta con el médico general sin costo en la mayoría de clínicas."],
  ];
  for (let i = 0; i < health.length; i++) {
    const [ic, t, d] = health[i];
    const y = 1.3 + i * 1.25;
    card(s, 5.3, y, 4.2, 1.1, C.background2, t);
    await iconCircle(s, ic, 5.45, y + 0.25, 0.6, C.accent2, t);
    txt(s, [{ text: t, options: { bold: true, color: C.text2, breakLine: true, fontSize: 13.5 } }, { text: d, options: { fontSize: 11.5 } }], { x: 6.2, y: y + 0.08, w: 3.15, h: 0.95, valign: "middle" });
  }
  source(s, "ACC, Annual Report 2025; Pharmac, How Pharmac works; Health New Zealand – Te Whatu Ora.");
  s.addNotes("JOSÉ LUIS: Nuestro segundo tema es salud, y aquí encontramos algo único en el mundo: la ACC, que desde 1974 es un seguro de accidentes para todos. Si alguien se rompe una pierna jugando futbol o tiene un accidente en el trabajo o en la calle, la ACC paga su tratamiento y parte de su sueldo mientras se recupera, sin importar de quién fue la culpa. Se paga con cuotas que salen de los salarios, las empresas y los autos, más dinero del gobierno, y cubre incluso a los turistas. En un solo año aceptó más de 2.1 millones de casos. Lo diferente es que, a cambio, no se puede demandar a nadie por lesiones. Además, Pharmac compra las medicinas para todo el país y, como compra mucho, consigue precios bajos. Los hospitales públicos son gratuitos y los niños menores de 14 años van al médico sin costo en la mayoría de las clínicas.");

  // ================= 11. EMPLEO Y PENSIONES =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · TEMA 3: EMPLEO Y PENSIONES", "Trabajar hoy y vivir tranquilo mañana");
  ficha(s, 0.5, 1.3, 4.6, 3.65, [
    ["¿En qué consiste?", "Dos pilares: pensión universal (NZ Super) + ahorro automático (KiwiSaver)."],
    ["¿Qué problema resuelve?", "Pobreza en la vejez y que la gente no ahorre para su retiro."],
    ["¿Cómo funciona y a quién beneficia?", "Toda persona residente de 65 años o más recibe pensión. KiwiSaver: unos 3.4 millones de personas."],
    ["¿Qué resultados ha dado?", "Los adultos mayores tienen un ingreso seguro y millones ahorran casi sin darse cuenta."],
    ["¿Qué lo hace diferente?", "Si no haces nada, ¡ahorras! Y la empresa y el gobierno te ponen dinero extra."],
  ]);
  const pil = [
    ["NZ Super", "Pensión para todos desde los 65 años, hayan trabajado o no.", "NZ$ 555 a la semana", "(persona sola, después de impuestos, 2026)"],
    ["KiwiSaver (2007)", "Al entrar a un trabajo te inscriben solo; tú y tu empresa ahorran 3.5 % del sueldo cada uno.", "+ bono del gobierno", "(25 centavos por cada dólar, con tope)"],
  ];
  pil.forEach(([t, d, n, nn], i) => {
    const y = 1.3 + i * 1.5;
    card(s, 5.3, y, 4.2, 1.35, i ? C.accent2 : C.accent1, t);
    txt(s, [{ text: t, options: { bold: true, fontSize: 15, breakLine: true, fontFace: THEME.headFontFace } }, { text: d, options: { fontSize: 11.5 } }], { x: 5.5, y: y + 0.1, w: 3.8, h: 0.8, color: C.background1 });
    txt(s, [{ text: n + " ", options: { bold: true, fontSize: 14, color: C.accent3 } }, { text: nn, options: { fontSize: 10, color: C.background2 } }], { x: 5.5, y: y + 0.92, w: 3.8, h: 0.35, color: C.background1, valign: "middle" });
  });
  txt(s, [{ text: "Salario mínimo 2026: ", options: { bold: true, color: C.accent1 } }, { text: "NZ$ 23.95 por hora." }], { x: 5.3, y: 4.35, w: 4.2, h: 0.5, fontSize: 12.5, valign: "middle" });
  source(s, "Work and Income NZ, tarifas al 1 de abril de 2026; Inland Revenue, KiwiSaver changes (2025); Employment NZ (2026).");
  s.addNotes("PABLO: Nuestro tercer tema es empleo y pensiones. Nueva Zelanda tiene dos pilares. El primero es NZ Super: una pensión para todas las personas de 65 años o más que viven ahí, hayan trabajado o no; una persona sola recibe unos 555 dólares neozelandeses a la semana. El segundo pilar es KiwiSaver, creado en 2007. Cuando alguien empieza un nuevo trabajo, lo inscriben automáticamente; si no quiere, puede salirse, pero si no hace nada, ¡ahorra! El trabajador pone 3.5 % de su sueldo, la empresa otro 3.5 % y el gobierno da un bono. Unos 3.4 millones de personas participan. Además, el salario mínimo es de 23.95 dólares neozelandeses por hora.");

  // ================= 12. KIWISAVER EJEMPLO =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · EDUCACIÓN FINANCIERA EN ACCIÓN", "Así crece un ahorro KiwiSaver (ejemplo)");
  txt(s, "Una persona gana NZ$ 50,000 al año. Cada año se ahorra:", { x: 0.5, y: 1.25, w: 9, h: 0.35, fontSize: 14 });
  const parts = [["NZ$ 1,750", "Tú (3.5 %)", C.accent1], ["NZ$ 1,750", "Tu empresa (3.5 %)", C.accent2], ["NZ$ 261", "Gobierno (bono)", C.accent3], ["NZ$ 3,761", "Total en un año", C.text2]];
  parts.forEach(([n, l, col], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.7, 1.9, 1.0, col, l);
    txt(s, [{ text: n, options: { bold: true, fontSize: 20, breakLine: true, fontFace: THEME.headFontFace } }, { text: l, options: { fontSize: 11.5 } }], { x: x + 0.1, y: 1.75, w: 1.7, h: 0.9, color: C.background1, align: "center", valign: "middle" });
    if (i < 3) txt(s, i < 2 ? "+" : "=", { x: x + 1.9, y: 1.7, w: 0.4, h: 1.0, fontSize: 26, bold: true, color: C.text2, align: "center", valign: "middle" });
  });
  const yrs = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  const A = 3760.72, r = 0.05;
  const noInt = yrs.map((y) => Math.round(A * y));
  const withInt = yrs.map((y) => Math.round(A * ((Math.pow(1 + r, y) - 1) / r)));
  s.addChart(pres.charts.BAR, [
    { name: "Solo lo aportado", labels: yrs.map((y) => "Año " + y), values: noInt },
    { name: "Con interés del 5 % anual", labels: yrs.map((y) => "Año " + y), values: withInt },
  ], {
    x: 0.5, y: 2.85, w: 6.0, h: 2.15, barDir: "col", barGapWidthPct: 50, chartColors: [H.accent5, H.accent1], showValue: false, showLegend: true, legendPos: "t", legendFontSize: 10,
    valAxisLabelFormatCode: '#,##0', valGridLine: { color: "E3E8E6", size: 0.5 }, catGridLine: { style: "none" }, ...chartText, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, objectName: "gráfica interés compuesto",
  });
  card(s, 6.8, 2.85, 2.7, 2.15, C.background2, "resultado 10 años");
  txt(s, [
    { text: "En 10 años", options: { bold: true, color: C.accent6, fontSize: 12, breakLine: true } },
    { text: "NZ$ " + withInt[9].toLocaleString("en-US"), options: { bold: true, color: C.accent1, fontSize: 22, fontFace: THEME.headFontFace, breakLine: true } },
    { text: "Aportó NZ$ " + noInt[9].toLocaleString("en-US") + "; el resto lo ganó el interés compuesto.", options: { fontSize: 11.5 } },
  ], { x: 6.95, y: 2.95, w: 2.4, h: 1.95 });
  source(s, "Cálculo propio con reglas de Inland Revenue (2025–2026). Ejemplo ilustrativo, antes de impuestos; el 5 % es un supuesto.");
  s.addNotes("JOSÉ LUIS: Como nuestra materia es educación financiera, hicimos un ejemplo. Imaginemos a alguien que gana 50 mil dólares neozelandeses al año. Él ahorra 3.5 %, o sea 1,750; su empresa pone otros 1,750 y el gobierno le regala unos 261. En total junta 3,761 dólares al año, ¡y solo la mitad salió de su bolsillo! En la gráfica vemos qué pasa en 10 años. Las barras claras muestran solo lo aportado: unos 37,600. Las barras verdes muestran lo que pasa si el dinero gana 5 % de interés al año: casi 47,300. Esa diferencia es el interés compuesto: el dinero que ganan tus ahorros también genera más dinero. Aclaramos que el 5 % es un supuesto para el ejemplo.");

  // ================= 13. VIVIENDA Y MEDIO AMBIENTE =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · TEMA 4: VIVIENDA Y MEDIO AMBIENTE", "Casas sanas y un país con energía limpia");
  card(s, 0.5, 1.3, 4.4, 3.65, C.background2, "vivienda");
  await iconCircle(s, fa.FaHome, 0.7, 1.45, 0.55, C.accent2, "vivienda");
  txt(s, "Vivienda", { x: 1.4, y: 1.45, w: 3.3, h: 0.55, fontSize: 17, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle" });
  txt(s, [
    { text: "Problema: ", options: { bold: true, color: C.accent4 } }, { text: "casas frías y húmedas enferman, sobre todo a niños.", options: { breakLine: true } },
    { text: "Solución: ", options: { bold: true, color: C.accent1 } }, { text: "«Healthy Homes» (2019): toda casa en renta debe tener calefacción, aislamiento, ventilación y control de humedad (obligatorio desde julio 2025).", options: { breakLine: true } },
    { text: "Vivienda pública: ", options: { bold: true, color: C.accent1 } }, { text: "las familias de bajos ingresos pagan como máximo cerca del 25 % de su ingreso.", options: { breakLine: true } },
    { text: "KiwiSaver: ", options: { bold: true, color: C.accent1 } }, { text: "se puede usar para comprar la primera casa." },
  ], { x: 0.7, y: 2.15, w: 4.0, h: 2.7, fontSize: 12.5, paraSpaceAfter: 6 });
  card(s, 5.1, 1.3, 4.4, 3.65, C.background2, "medio ambiente");
  await iconCircle(s, fa.FaLeaf, 5.3, 1.45, 0.55, C.accent1, "medio ambiente");
  txt(s, "Medio ambiente", { x: 6.0, y: 1.45, w: 3.3, h: 0.55, fontSize: 17, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle" });
  txt(s, [{ text: "88.5 %", options: { bold: true, fontSize: 34, color: C.accent1, fontFace: THEME.headFontFace, breakLine: true } }, { text: "de su electricidad en 2025 vino de fuentes renovables (agua, calor de la tierra, viento y sol).", options: { fontSize: 11.5 } }],
    { x: 5.3, y: 2.1, w: 4.0, h: 1.25 });
  txt(s, [
    { text: "Ley Cero Carbono (2019): ", options: { bold: true, color: C.accent1 } }, { text: "meta de cero emisiones netas en 2050.", options: { breakLine: true } },
    { text: "Río Whanganui (2017): ", options: { bold: true, color: C.accent1 } }, { text: "primer río del mundo con derechos de «persona», como pidió el pueblo maorí." },
  ], { x: 5.3, y: 3.45, w: 4.0, h: 1.4, fontSize: 12.5, paraSpaceAfter: 5 });
  source(s, "Tenancy Services NZ, Healthy homes standards; Kāinga Ora; MBIE, Energy in New Zealand 2026; Ministry for the Environment.");
  s.addNotes("PABLO: Nuestro cuarto tema es vivienda y medio ambiente. En Nueva Zelanda hace frío y muchas casas eran húmedas, lo que enfermaba a los niños. Por eso en 2019 crearon las normas 'Healthy Homes', o casas saludables: toda casa en renta debe tener calefacción, aislamiento, ventilación y control de humedad, y desde julio de 2025 es obligatorio para todas. Las familias de bajos ingresos que viven en vivienda pública pagan como máximo cerca del 25 % de su ingreso, y los ahorros de KiwiSaver se pueden usar para comprar la primera casa. En medio ambiente, 88.5 % de su electricidad en 2025 vino de fuentes renovables, como presas, el calor de la tierra, el viento y el sol. Tienen una ley para llegar a cero emisiones netas en 2050, y en 2017 reconocieron al río Whanganui como si fuera una persona con derechos, como pedía el pueblo maorí.");

  // ================= 14. POR QUÉ FUNCIONA =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "B. Modelos de bienestar" });
  header(s, "B · ¿QUÉ LO HACE EXITOSO?", "Las claves del modelo… y sus retos");
  const keys = [
    [fa.FaHandshake, "Confianza", "Poca corrupción: el dinero llega a donde debe."],
    [fa.FaUsers, "Para todos", "Programas universales: menos trámites y menos estigma."],
    [fa.FaSyncAlt, "Automático", "Ahorrar es la opción «por defecto»."],
    [fa.FaChartPie, "Mide bienestar", "No solo cuenta dinero; cuenta vidas."],
    [fa.FaCoins, "Impuestos", "Junta casi 33 % del PIB para servicios públicos."],
  ];
  for (let i = 0; i < keys.length; i++) {
    const [ic, t, d] = keys[i];
    const y = 1.3 + i * 0.74;
    await iconCircle(s, ic, 0.5, y, 0.55, C.accent1, t);
    txt(s, [{ text: t + ": ", options: { bold: true, color: C.text2 } }, { text: d }], { x: 1.2, y: y, w: 4.1, h: 0.55, fontSize: 13, valign: "middle" });
  }
  card(s, 5.6, 1.3, 3.9, 3.65, C.background2, "retos");
  txt(s, "Pero no es perfecto", { x: 5.8, y: 1.42, w: 3.5, h: 0.45, fontSize: 17, bold: true, color: C.accent4, fontFace: THEME.headFontFace });
  txt(s, [
    { text: "Las casas son muy caras.", options: { bullet: true, breakLine: true } },
    { text: "Desempleo de 5.7 % (sep. 2026); 12 % en jóvenes de 20 a 24 años.", options: { bullet: true, breakLine: true } },
    { text: "La felicidad de los jóvenes está en el lugar 126 de 136 países.", options: { bullet: true, breakLine: true } },
    { text: "El pueblo maorí vive unos 6 años menos.", options: { bullet: true } },
  ], { x: 5.8, y: 1.95, w: 3.5, h: 2.9, fontSize: 12.5, paraSpaceAfter: 7 });
  source(s, "Transparency International (2026); OCDE (2025); Stats NZ, Labour market statistics (2026); World Happiness Report 2026.");
  s.addNotes("JOSÉ LUIS: ¿Por qué funciona el modelo de Nueva Zelanda? Encontramos cinco claves. Uno: confianza; hay muy poca corrupción, así que el dinero llega a donde debe. Dos: los programas son para todos, con menos trámites. Tres: son automáticos; por ejemplo, ahorrar es la opción por defecto. Cuatro: miden el bienestar y no solo el dinero. Y cinco: la gente paga impuestos, casi 33 % del PIB, para tener servicios públicos. Pero queremos ser honestos: no es perfecto. Las casas son muy caras, el desempleo subió a 5.7 % y a 12 % en jóvenes, la felicidad de los jóvenes está muy abajo en el ranking mundial y el pueblo maorí vive unos seis años menos.");

  // ================= 15. QUÉ PODRÍA APLICAR MÉXICO =================
  pres.addSection({ title: "C. ¿Funcionaría en México?" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "C. ¿Funcionaría en México?" });
  header(s, "C · ¿CÓMO PODEMOS REPLICARLO?", "¿Qué propuestas podría adoptar México?");
  const th2 = (t) => ({ text: t, options: { bold: true, color: H.lt1, fill: { color: H.dk2 }, fontSize: 12.5, valign: "middle" } });
  const td2 = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 12.5, color: H.dk1, valign: "middle" }, o) });
  s.addTable([
    [th2("Idea de Nueva Zelanda"), th2("¿Qué hay hoy en México?"), th2("Beneficio para México")],
    [td2("Ahorro automático (KiwiSaver)", { bold: true, color: H.accent1 }), td2("Afores solo para empleos formales; 54.3 % trabaja en la informalidad (sep. 2026)."), td2("Más personas con ahorro para su vejez y su casa.")],
    [td2("Almuerzo escolar diario y nutritivo", { bold: true, color: H.accent1 }), td2("Desde 2025 está prohibida la comida chatarra en escuelas."), td2("Menos hambre y obesidad; mejor aprendizaje.")],
    [td2("Seguro de accidentes para todos (ACC)", { bold: true, color: H.accent1 }), td2("El IMSS cubre accidentes de trabajo solo a trabajadores formales."), td2("Que nadie se endeude por un accidente.")],
    [td2("Casas saludables (Healthy Homes)", { bold: true, color: H.accent1 }), td2("Hay créditos de vivienda (Infonavit), pero pocas reglas de calidad en rentas."), td2("Menos enfermedades por humedad, calor o frío.")],
  ], { x: 0.5, y: 1.3, w: 9, colW: [2.6, 3.5, 2.9], rowH: [0.4, 0.75, 0.65, 0.65, 0.75], border: { type: "solid", pt: 0.75, color: "C9D8CF" }, fill: { color: H.lt1 }, margin: [0.04, 0.08, 0.04, 0.08], objectName: "tabla propuestas" });
  source(s, "INEGI, Indicadores de Ocupación y Empleo, sept. 2026; SEP-Salud, lineamientos de alimentación escolar (2025); IMSS; Infonavit.");
  s.addNotes("PABLO: Ahora la parte C: ¿qué podría copiar México? Elegimos cuatro ideas. Primera: el ahorro automático tipo KiwiSaver. En México existen las Afores, pero solo para quien tiene empleo formal, y más de la mitad de los trabajadores, 54.3 %, está en la informalidad. Segunda: un almuerzo escolar diario y nutritivo; México ya dio un paso al prohibir la comida chatarra en escuelas en 2025. Tercera: un seguro de accidentes para todos, porque hoy el IMSS solo cubre a trabajadores formales. Y cuarta: reglas para que las casas sean saludables. Cada una tendría beneficios claros: más ahorro, menos hambre, menos deudas por accidentes y menos enfermedades.");

  // ================= 16. OBSTÁCULOS Y ADAPTACIONES =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "C. ¿Funcionaría en México?" });
  header(s, "C · ANÁLISIS CRÍTICO", "¿Qué obstáculos habría y cómo adaptarlo?");
  const obs = [
    [fa.FaMoneyBillWave, "Económicos", "México junta 18.3 % del PIB en impuestos; NZ, 32.9 %. Hay menos dinero."],
    [fa.FaLandmark, "Políticos", "Los programas cambian con cada gobierno."],
    [fa.FaComments, "Culturales", "Desconfianza en las instituciones y poca costumbre de ahorrar."],
    [fa.FaUsers, "Sociales", "25 veces más personas, informalidad y desigualdad entre regiones."],
  ];
  for (let i = 0; i < 4; i++) {
    const [ic, t, d] = obs[i];
    const x = 0.5 + (i % 2) * 2.3, y = 1.3 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 2.1, 1.7, C.background2, t);
    await iconCircle(s, ic, x + 0.15, y + 0.15, 0.5, C.accent4, t);
    txt(s, t, { x: x + 0.75, y: y + 0.15, w: 1.3, h: 0.5, fontSize: 14, bold: true, color: C.accent4, valign: "middle" });
    txt(s, d, { x: x + 0.15, y: y + 0.75, w: 1.85, h: 0.9, fontSize: 11 });
  }
  card(s, 5.3, 1.3, 4.2, 3.55, C.accent1, "adaptaciones");
  txt(s, "Adaptaciones necesarias", { x: 5.5, y: 1.42, w: 3.8, h: 0.45, fontSize: 17, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  txt(s, [
    { text: "Inscripción al ahorro desde el celular, también para quien trabaja por su cuenta.", options: { bullet: true, breakLine: true } },
    { text: "Montos pequeños y flexibles (desde $50 a la semana).", options: { bullet: true, breakLine: true } },
    { text: "Empezar en las zonas con más necesidad e ir creciendo.", options: { bullet: true, breakLine: true } },
    { text: "Comida escolar con productores locales.", options: { bullet: true, breakLine: true } },
    { text: "Una ley para que no desaparezca al cambiar de gobierno.", options: { bullet: true } },
  ], { x: 5.5, y: 1.95, w: 3.8, h: 2.85, fontSize: 12, color: C.background1, paraSpaceAfter: 6 });
  source(s, "OCDE, Revenue Statistics 2025; INEGI (2026); análisis propio.");
  s.addNotes("JOSÉ LUIS: Pero no basta con decir que un modelo es bueno; hay que ver si funcionaría aquí. Encontramos cuatro tipos de obstáculos. Económicos: México junta mucho menos en impuestos, 18.3 % de su PIB contra casi 33 % de Nueva Zelanda, así que hay menos dinero. Políticos: los programas suelen cambiar con cada gobierno. Culturales: mucha gente desconfía de las instituciones y no tiene costumbre de ahorrar. Y sociales: somos 25 veces más personas, con mucha informalidad y diferencias entre regiones. Por eso proponemos adaptaciones: inscribirse al ahorro desde el celular, aunque trabajes por tu cuenta; empezar con montos pequeños; comenzar en las zonas con más necesidad; comprar la comida escolar a productores locales; y hacer una ley para que el programa no desaparezca al cambiar de gobierno.");

  // ================= 17. NUESTRA PROPUESTA =================
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "C. ¿Funcionaría en México?" });
  header(s, "C · NUESTRA PROPUESTA", "«Mi Ahorro Automático»: un KiwiSaver a la mexicana");
  const steps = [
    ["1", "Te inscribes solo", "Al conseguir trabajo o con una app si trabajas por tu cuenta. Puedes salirte si quieres."],
    ["2", "Ahorras poquito", "Desde $50 a la semana, automático."],
    ["3", "El gobierno te premia", "25 centavos por cada peso, con un tope."],
    ["4", "Lo usas bien", "Para tu retiro o el enganche de tu primera casa."],
  ];
  steps.forEach(([n, t, d], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.3, 2.1, 1.75, C.background2, "paso " + n);
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.42, w: 0.48, h: 0.48, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "número " + n });
    txt(s, n, { x: x + 0.15, y: 1.42, w: 0.48, h: 0.48, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.72, y: 1.42, w: 1.3, h: 0.48, fontSize: 12.5, bold: true, color: C.text2, valign: "middle" });
    txt(s, d, { x: x + 0.15, y: 2.0, w: 1.85, h: 1.0, fontSize: 11 });
  });
  card(s, 0.5, 3.25, 3.4, 1.7, C.accent1, "ejemplo pesos");
  txt(s, [
    { text: "$50 a la semana + bono", options: { fontSize: 12, breakLine: true } },
    { text: "≈ $309,000", options: { bold: true, fontSize: 28, fontFace: THEME.headFontFace, color: C.accent3, breakLine: true } },
    { text: "en 40 años (supuesto: 4 % de interés anual)", options: { fontSize: 11 } },
  ], { x: 0.65, y: 3.35, w: 3.1, h: 1.5, color: C.background1, valign: "middle" });
  txt(s, "¿Por qué tendría el mayor impacto?", { x: 4.1, y: 3.25, w: 5.4, h: 0.4, fontSize: 15, bold: true, color: C.text2, fontFace: THEME.headFontFace });
  txt(s, [
    { text: "Ataca tres problemas a la vez: vejez sin ingresos, falta de vivienda y poca educación financiera.", options: { bullet: true, breakLine: true } },
    { text: "Es más barato: la persona también aporta y el interés compuesto hace crecer el dinero.", options: { bullet: true, breakLine: true } },
    { text: "Se apoya en algo que ya existe: las Afores.", options: { bullet: true } },
  ], { x: 4.1, y: 3.7, w: 5.4, h: 1.3, fontSize: 12, paraSpaceAfter: 4 });
  source(s, "Propuesta del equipo basada en KiwiSaver (Inland Revenue NZ) y el sistema de Afores (CONSAR). Cálculo ilustrativo.");
  s.addNotes("PABLO: De todas las ideas, nosotros proponemos 'Mi Ahorro Automático', un KiwiSaver a la mexicana. Funciona en cuatro pasos: uno, te inscriben solo al conseguir trabajo, o te inscribes con una app si trabajas por tu cuenta, y puedes salirte si quieres. Dos, ahorras poquito, desde 50 pesos a la semana. Tres, el gobierno te premia con 25 centavos por cada peso que ahorres, con un tope. Y cuatro, lo usas para tu retiro o para el enganche de tu primera casa. JOSÉ LUIS: Hicimos la cuenta: si alguien ahorra 50 pesos a la semana más el bono, y su dinero gana 4 % al año, en 40 años tendría unos 309 mil pesos. Creemos que esta propuesta tendría el mayor impacto porque ataca tres problemas a la vez: la vejez sin ingresos, la falta de vivienda y la poca educación financiera. Además cuesta menos, porque la persona también aporta, y se apoya en algo que ya existe en México: las Afores.");

  // ================= INFOGRAFÍA (resumen) =================
  s = pres.addSlide({ masterName: "OSCURA", sectionTitle: "C. ¿Funcionaría en México?" });
  s.addImage({ data: "image/png;base64," + fs.readFileSync("infografia.png").toString("base64"), x: 0.2, y: 0.134, w: 9.6, h: 5.357,
    altText: "Infografía: Modelos de bienestar, lecciones de Nueva Zelanda para México. Pilares del modelo neozelandés (KiwiSaver y retiro universal, seguro universal de accidentes, 88.5 % de energía renovable), comparativa de indicadores con México y la propuesta «Mi Ahorro Automático».",
    objectName: "infografía resumen" });
  s.addNotes("PABLO: Para resumir todo lo que vimos, hicimos esta infografía. A la izquierda están los pilares del modelo neozelandés: KiwiSaver y la pensión universal, el seguro de accidentes para todos y el 88.5 % de energía renovable. En el centro, la comparación con México: desarrollo humano, lugar 17 contra 81; corrupción, 81 puntos contra 27; e impuestos, 32.9 % contra 18.3 % del PIB. JOSÉ LUIS: Y a la derecha, nuestra propuesta, «Mi Ahorro Automático»: inscripción automática y flexible desde 50 pesos a la semana, un bono del gobierno de 25 centavos por cada peso ahorrado, y un fondo con interés compuesto para el retiro o la primera vivienda.");

  // ================= 18. CONCLUSIONES =================
  pres.addSection({ title: "D. Conclusiones" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "D. Conclusiones" });
  header(s, "D · CONCLUSIONES", "Lo que aprendimos");
  card(s, 0.5, 1.3, 4.4, 2.75, C.background2, "sorprendente");
  await iconCircle(s, fa.FaStar, 0.7, 1.45, 0.55, C.accent3, "sorprendente");
  txt(s, "Lo más sorprendente", { x: 1.4, y: 1.45, w: 3.3, h: 0.55, fontSize: 16, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle" });
  txt(s, [
    { text: "Un seguro de accidentes que cubre ¡hasta a los turistas!", options: { bullet: true, breakLine: true } },
    { text: "Un río con derechos de persona.", options: { bullet: true, breakLine: true } },
    { text: "México y Nueva Zelanda son casi igual de felices.", options: { bullet: true } },
  ], { x: 0.7, y: 2.15, w: 4.0, h: 1.8, fontSize: 14, paraSpaceAfter: 6 });
  card(s, 5.1, 1.3, 4.4, 2.75, C.background2, "aprender");
  await iconCircle(s, fa.FaGraduationCap, 5.3, 1.45, 0.55, C.accent1, "aprender");
  txt(s, "¿Qué puede aprender México?", { x: 6.0, y: 1.45, w: 3.4, h: 0.55, fontSize: 16, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "middle" });
  txt(s, [
    { text: "Que el ahorro sea automático y fácil.", options: { bullet: true, breakLine: true } },
    { text: "Medir el éxito por cómo vive la gente, no solo por el PIB.", options: { bullet: true, breakLine: true } },
    { text: "Que los buenos programas duren aunque cambie el gobierno.", options: { bullet: true } },
  ], { x: 5.3, y: 2.15, w: 4.0, h: 1.8, fontSize: 14, paraSpaceAfter: 6 });
  card(s, 0.5, 4.2, 9, 0.75, C.text2, "frase final");
  txt(s, "«El bienestar no se mide solo en dinero, pero el dinero bien planeado construye bienestar.»", { x: 0.7, y: 4.2, w: 8.6, h: 0.75, fontSize: 15, italic: true, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: THEME.headFontFace });
  s.addNotes("JOSÉ LUIS: Para terminar, lo que más nos sorprendió fue que existe un seguro de accidentes que cubre hasta a los turistas, que un río tenga derechos como una persona y que México y Nueva Zelanda sean casi igual de felices. PABLO: Y lo que creemos que México puede aprender es: que ahorrar sea automático y fácil; medir el éxito de un país por cómo vive su gente y no solo por el PIB; y que los buenos programas continúen aunque cambie el gobierno. JOSÉ LUIS: Nuestra conclusión es: el bienestar no se mide solo en dinero, pero el dinero bien planeado construye bienestar.");

  // ================= 19-20. BIBLIOGRAFÍA =================
  pres.addSection({ title: "Bibliografía" });
  const bib1 = [
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
  ];
  const bib2 = [
    "OCDE. (2025). Revenue Statistics 2025. OECD Publishing. https://www.oecd.org",
    "Pharmac. (2025). How Pharmac works. https://pharmac.govt.nz",
    "Programa de las Naciones Unidas para el Desarrollo (PNUD). (2025). Informe sobre Desarrollo Humano 2025: Una cuestión de elección. https://hdr.undp.org",
    "Stats NZ. (2025). National population estimates: At 30 June 2025. https://www.stats.govt.nz",
    "Stats NZ. (2025). National and subnational period life tables: 2022–2024. https://www.stats.govt.nz",
    "Stats NZ. (2026). Overseas merchandise trade: December 2025. https://www.stats.govt.nz",
    "Stats NZ. (2026). Labour market statistics: September 2026 quarter. https://www.stats.govt.nz",
    "Tenancy Services. (2025). Healthy homes standards. Ministry of Housing and Urban Development. https://www.tenancy.govt.nz",
    "Transparency International. (2026). Corruption Perceptions Index 2025. https://www.transparency.org/cpi",
    "Work and Income New Zealand. (2026). Benefit rates at 1 April 2026: NZ Superannuation. https://www.workandincome.govt.nz",
    "Secretaría de Educación Pública y Secretaría de Salud. (2024). Lineamientos generales para la preparación, distribución y expendio de alimentos y bebidas en escuelas. Diario Oficial de la Federación. https://www.dof.gob.mx",
  ];
  [[bib1, "1 de 2"], [bib2, "2 de 2"]].forEach(([list, pg]) => {
    const sl = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Bibliografía" });
    header(sl, "FUENTES CONSULTADAS (FORMATO APA 7)", "Bibliografía (" + pg + ")");
    txt(sl, list.map((t, i) => ({ text: t, options: { breakLine: i < list.length - 1, paraSpaceAfter: 4, indentLevel: 0 } })),
      { x: 0.5, y: 1.3, w: 9, h: 3.85, fontSize: 11, color: C.text1 });
    sl.addNotes("PABLO: Aquí están todas nuestras fuentes. Usamos fuentes oficiales: oficinas de estadística y ministerios del gobierno de Nueva Zelanda, la ONU, la OCDE, el FMI, el INEGI y la Secretaría de Educación Pública, entre otras. (No es necesario leer esta diapositiva completa.)");
  });

  // ================= 21. GRACIAS =================
  s = pres.addSlide({ masterName: "OSCURA", sectionTitle: "Bibliografía" });
  addLogo(s, 0.6, 0.35, 0.85);
  txt(s, "¡Gracias!", { x: 0.6, y: 1.5, w: 6, h: 1.0, fontSize: 48, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  txt(s, "¿Tienen alguna pregunta?", { x: 0.6, y: 2.5, w: 6, h: 0.6, fontSize: 22, color: C.accent5 });
  txt(s, "José Luis De Anda Reimers · Pablo De Santiago Hurtado Martin Del Campo\nInstituto Educativo Ammadeus · Educación Financiera · 2.º de secundaria",
    { x: 0.6, y: 3.6, w: 7.5, h: 0.8, fontSize: 12, color: C.background1 });
  s.addImage({ data: flag, x: 7.3, y: 1.6, w: 2.0, h: 1.0, altText: "Bandera de Nueva Zelanda", objectName: "bandera" });
  txt(s, "Kia ora = «hola / gracias» en maorí", { x: 7.0, y: 2.7, w: 2.6, h: 0.35, fontSize: 11, italic: true, color: C.accent5, align: "center" });
  s.addNotes("AMBOS: ¡Gracias por su atención! JOSÉ LUIS: Como dicen en Nueva Zelanda: kia ora, que en maorí significa 'hola' y también 'gracias'. PABLO: ¿Tienen alguna pregunta?");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
