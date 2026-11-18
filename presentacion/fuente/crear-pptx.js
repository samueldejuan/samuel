// Genera Pitch_Merco_Guatemala.pptx (editable) a partir del diseño del PDF original.
// Uso: node crear-pptx.js <carpeta-de-recursos> <salida.pptx> <apply_theme.js>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");

const [ASSETS, OUT, APPLY_THEME] = process.argv.slice(2);
const { applyTheme } = require(APPLY_THEME);
const A = (f) => path.join(ASSETS, f);

const THEME = {
  name: "Merco Guatemala",
  headFontFace: "Arial Narrow",
  bodyFontFace: "Arial",
  colors: {
    dk1: "0F2D4E", lt1: "FFFFFF", dk2: "2A6FB5", lt2: "F7FAFC",
    accent1: "2A6FB5", accent2: "4A97D3", accent3: "0F2D4E", accent4: "DCEBF7",
    accent5: "F28C28", accent6: "9EC5F0", hlink: "2A6FB5", folHlink: "0F2D4E",
  },
};
const HEX = {
  navy: "0F2D4E", blue: "2A6FB5", mid: "4A97D3", pale: "DCEBF7", card: "F5F8FC",
  muted: "4A5D7A", soft: "D9E6F5", light: "F7FAFC", orange: "F28C28", sky: "8EC0EC",
};

// Medidas tomadas sobre el original a 1920×1080 px
const P = (v) => v / 192;
const HEAD = "+mj-lt";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Pitch Merco: Guatemala";
pres.author = "Samuel de Juan Ortiz";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

async function icon(name, hex) {
  const Comp = lu[name];
  if (!Comp) throw new Error(`Icono no encontrado: ${name}`);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { size: 256, color: `#${hex}`, strokeWidth: 1.6 }),
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Diseños (layouts) ----------
function layout(name, bg, { kicker, title, source, number, numberX = 1690, kickerSize = 15, kickerW = 1100, kickerSpacing = 2 }) {
  // El pie llega hasta justo antes del número de página
  const objects = [
    { placeholder: { options: { name: "kicker", type: "body", x: P(130), y: P(122), w: P(kickerW), h: P(50), fontFace: HEAD, fontSize: kickerSize, bold: true, charSpacing: kickerSpacing, color: kicker, margin: 0, valign: "middle" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: P(126), y: P(178), w: P(1120), h: P(180), fontSize: 31, bold: true, color: title, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.86 }, text: "" } },
    { placeholder: { options: { name: "source", type: "body", x: P(130), y: P(982), w: P(numberX - 150), h: P(40), fontSize: 10, color: source, margin: 0, valign: "middle" }, text: "" } },
  ];
  const def = { title: name, background: { color: bg }, objects };
  if (number) {
    def.slideNumber = { x: P(numberX), y: P(982), w: P(100), h: P(40), fontSize: 10, color: number, align: "right", margin: 0 };
  }
  pres.defineSlideMaster(def);
}
layout("Portada", HEX.navy, { kicker: HEX.soft, title: "FFFFFF", source: HEX.soft, numberX: 1030, kickerSize: 13.5, kickerW: 1250, kickerSpacing: 1.5 });
layout("Oscuro", "0F2E4D", { kicker: HEX.soft, title: "FFFFFF", source: HEX.soft, number: HEX.soft });
layout("Claro", HEX.light, { kicker: HEX.blue, title: HEX.blue, source: HEX.muted, number: HEX.muted });
layout("Azul", HEX.blue, { kicker: HEX.pale, title: "FFFFFF", source: HEX.pale, number: HEX.pale });
layout("Celeste", HEX.pale, { kicker: HEX.blue, title: HEX.blue, source: HEX.muted, number: HEX.muted });
layout("Fuentes", HEX.light, { kicker: HEX.blue, title: HEX.blue, source: HEX.muted });

// ---------- Ayudantes ----------
function text(slide, runs, o) {
  slide.addText(runs, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), margin: 0, isTextBox: true,
    valign: o.valign || "top", align: o.align || "left", fontSize: o.fontSize || 12,
    color: o.color || C.text1, bold: o.bold, fontFace: o.fontFace, rotate: o.rotate,
    charSpacing: o.charSpacing, lineSpacingMultiple: o.lineSpacingMultiple, objectName: o.name,
    fill: o.fill, line: o.line, shape: o.shape, rectRadius: o.rectRadius, paraSpaceAfter: o.paraSpaceAfter,
  });
}
function box(slide, o) {
  slide.addShape(o.shape || pres.shapes.ROUNDED_RECTANGLE, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), rectRadius: o.radius ?? 0.12,
    fill: { color: o.fill }, line: o.line ? { color: o.line, width: o.lineW || 1.5 } : { type: "none" },
    rotate: o.rotate, objectName: o.name,
  });
}
async function card(slide, o) {
  // Tarjeta con icono, título en negrita y texto
  box(slide, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill, line: o.line, lineW: o.lineW, name: o.name });
  if (o.icon) {
    slide.addImage({ data: await icon(o.icon, o.iconColor), x: P(o.x + 34), y: P(o.y + o.h / 2 - 19), w: P(38), h: P(38), altText: o.title });
  }
  const tx = o.x + (o.icon ? 92 : 30);
  text(slide, [
    { text: o.title, options: { bold: true, fontSize: o.titleSize || 12, color: o.titleColor, breakLine: true } },
    { text: o.body, options: { fontSize: o.bodySize || 10, color: o.bodyColor } },
  ], { x: tx, y: o.y + 14, w: o.w - (tx - o.x) - 18, h: o.h - 24, valign: "middle", name: `${o.name} texto` });
}
function stamp(slide, label, o) {
  text(slide, label, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: HEAD, bold: true, fontSize: o.fontSize, color: o.color,
    align: "center", valign: "middle", rotate: o.rotate, line: { color: o.color, width: 4.5 },
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, lineSpacingMultiple: 0.95, name: o.name,
  });
}
function photoCaption(slide, label) {
  text(slide, label, {
    x: 1272, y: 820, w: 470, h: 36, fontSize: 10.5, color: "FFFFFF", valign: "middle",
    fill: { color: "000000", transparency: 45 }, name: "Pie de foto",
  });
}

(async () => {
  // 1 · Portada
  pres.addSection({ title: "Inicio" });
  let s = pres.addSlide({ masterName: "Portada", sectionTitle: "Inicio" });
  s.addImage({ path: A("portada-foto.jpg"), x: P(1127), y: 0, w: P(793), h: P(1080), altText: "Arco de Santa Catalina, Antigua Guatemala" });
  s.addImage({ path: A("fundido-h.png"), x: P(1126), y: 0, w: P(330), h: P(1080), altText: "Fundido decorativo" });
  s.addText("PROPUESTA DE INTERNACIONALIZACIÓN · SUPERMERCADOS MERCO", { placeholder: "kicker" });
  s.addText([
    { text: "¿Y SI EL PRÓXIMO", options: { breakLine: true } },
    { text: "SÚPER DE", options: { breakLine: true } },
    { text: "GUATEMALA ", options: { color: HEX.sky } },
    { text: "FUERA" },
  ], { placeholder: "title", fontSize: 44, y: P(196), h: P(300) });
  text(s, " MERCO? ", { x: 128, y: 482, w: 420, h: 116, fontFace: HEAD, bold: true, fontSize: 40, lineSpacingMultiple: 0.9, color: HEX.blue, fill: { color: "F5F8FC" }, valign: "middle", name: "Merco destacado" });
  box(s, { x: 128, y: 630, w: 598, h: 260, fill: HEX.card, line: HEX.mid, lineW: 3, rotate: -3, name: "Dato tienda de barrio" });
  box(s, { x: 155, y: 740, w: 30, h: 30, fill: HEX.navy, shape: pres.shapes.OVAL, name: "Punto" });
  text(s, [
    { text: "1 de cada 2", options: { fontFace: HEAD, bold: true, fontSize: 32, color: HEX.blue, breakLine: true } },
    { text: "compras en Guatemala se hacen en la tienda de barrio", options: { bold: true, fontSize: 12.5, color: HEX.navy, breakLine: true } },
    { text: "Kantar", options: { fontSize: 11, color: HEX.muted } },
  ], { x: 210, y: 655, w: 490, h: 210, rotate: -3, valign: "middle", name: "Dato texto" });
  s.addImage({ path: A("mapa.png"), x: P(762), y: P(632), w: P(322), h: P(261), altText: "Mapa: de Monterrey a Guatemala" });
  s.addText("Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504)", { placeholder: "source" });
  text(s, "Fuente: Kantar · Foto: Chad Davis, CC BY-SA 2.0 · Mapa: Natural Earth", { x: 1020, y: 982, w: 770, h: 40, fontSize: 9, color: "FFFFFF", valign: "middle", align: "right", name: "Créditos portada" });

  // 2 · Diagnóstico
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 1 · DIAGNÓSTICO", { placeholder: "kicker" });
  s.addText("¿QUÉ TAN LISTA ESTÁ MERCO?", { placeholder: "title" });
  stamp(s, "PREPARACIÓN MEDIA:\nLISTA PARA EMPEZAR", { x: 1270, y: 120, w: 530, h: 170, fontSize: 19, color: HEX.blue, rotate: -6, name: "Sello preparación" });
  const colHead = { fontFace: HEAD, bold: true, fontSize: 22, color: HEX.blue, valign: "middle" };
  text(s, "FORTALEZA", { ...colHead, x: 130, y: 315, w: 500, h: 55 });
  text(s, "CÓMO AYUDA", { ...colHead, x: 780, y: 315, w: 360, h: 55, align: "center" });
  text(s, "DEBILIDAD QUE CUBRE", { ...colHead, x: 1185, y: 315, w: 600, h: 55 });
  const F = [
    { y: 378, h: 148, icon: "LuChartColumn", title: "F3 · Tiendas que rinden", body: "$118,000 por m², más que Soriana y Walmart", mid: "Tienda chica = poca inversión" },
    { y: 544, h: 118, icon: "LuHouse", title: "F2 · Experta en cercanía", body: "80 años con familias que cuidan su dinero", mid: "El trato vale más que una aplicación" },
    { y: 679, h: 150, icon: "LuUsers", title: "F1 · Cuida a su gente", body: "Rotación de 98% a 41%; Great Place to Work", mid: "Universidad Merco + socio local" },
  ];
  const D = [
    { y: 393, icon: "LuZap", title: "D1 · Es chica", body: "39 tiendas en 2 estados" },
    { y: 544, icon: "LuCodeXml", title: "D3 · Atrás en lo digital", body: "Sin programa de lealtad" },
    { y: 695, icon: "LuGlobe", title: "D2 · Nunca ha salido", body: "Sin experiencia fuera de México" },
  ];
  for (const [i, f] of F.entries()) {
    await card(s, { x: 128, y: f.y, w: 608, h: f.h, fill: HEX.blue, icon: f.icon, iconColor: "FFFFFF", title: f.title, body: f.body, titleColor: "FFFFFF", bodyColor: HEX.pale, name: `Fortaleza ${i + 1}` });
    const ay = f.y + f.h / 2;
    text(s, f.mid, { x: 760, y: ay - 62, w: 400, h: 50, bold: true, fontSize: 10.5, color: HEX.navy, align: "center", valign: "bottom", name: `Cómo ayuda ${i + 1}` });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: P(810), y: P(ay - 4), w: P(300), h: P(34), fill: { color: HEX.mid }, line: { type: "none" }, objectName: `Flecha ${i + 1}` });
  }
  for (const [i, d] of D.entries()) {
    await card(s, { x: 1184, y: d.y, w: 608, h: 118, fill: HEX.pale, icon: d.icon, iconColor: HEX.blue, title: d.title, body: d.body, titleColor: HEX.navy, bodyColor: HEX.muted, name: `Debilidad ${i + 1}` });
  }
  s.addText("Fuentes: mi diagnóstico interno; caso Merco (Carreón et al., 2026)", { placeholder: "source" });

  // 3 · El mercado
  s = pres.addSlide({ masterName: "Oscuro", sectionTitle: "Análisis" });
  s.addImage({ path: A("ciudad.jpg"), x: 0, y: P(810), w: P(1920), h: P(175), altText: "Vista de Ciudad de Guatemala" });
  s.addImage({ path: A("fundido-v.png"), x: 0, y: P(808), w: P(1920), h: P(110), altText: "Fundido decorativo" });
  s.addText("PASO 2 · EL MERCADO", { placeholder: "kicker" });
  s.addText("¿POR QUÉ GUATEMALA?", { placeholder: "title" });
  const stats = [
    { x: 128, rot: -2.5, big: "18.7 M", desc: "personas, la mayor población de Centroamérica", src: "Banco Mundial, 2025" },
    { x: 695, rot: 2, big: "+4.1%", desc: "crecerá su economía, más que América Latina", src: "Banguat, 2025" },
    { x: 1262, rot: -2, big: "+18.7%", desc: "subieron las remesas en 2025: US$25,530 M", src: "Banguat, 2026" },
  ];
  for (const [i, st] of stats.entries()) {
    box(s, { x: st.x, y: 308, w: 530, h: 278, fill: HEX.card, line: HEX.mid, lineW: 3, rotate: st.rot, name: `Dato ${i + 1}` });
    box(s, { x: st.x + 28, y: 432, w: 30, h: 30, fill: HEX.blue, shape: pres.shapes.OVAL, name: `Punto ${i + 1}` });
    text(s, [
      { text: st.big, options: { fontFace: HEAD, bold: true, fontSize: 38, color: HEX.blue, breakLine: true } },
      { text: st.desc, options: { bold: true, fontSize: 12.5, color: HEX.navy, breakLine: true } },
      { text: st.src, options: { fontSize: 11, color: HEX.muted } },
    ], { x: st.x + 86, y: 326, w: 424, h: 242, rotate: st.rot, valign: "middle", name: `Dato ${i + 1} texto` });
  }
  const info = [
    { x: 128, icon: "LuUsers", title: "Cliente que cuida su dinero", body: "US$6,506 de PIB por persona (FocusEconomics, 2025)" },
    { x: 694, icon: "LuSearch", title: "Competidores", body: "Walmart (276 tiendas), La Torre y Súper del Barrio (200)" },
    { x: 1260, icon: "LuLock", title: "Barreras", body: "Tamaño de Walmart, extorsiones y +2,100 km de distancia" },
  ];
  for (const [i, c] of info.entries()) {
    await card(s, { x: c.x, y: 640, w: 532, h: 153, fill: HEX.card, line: HEX.mid, lineW: 2, icon: c.icon, iconColor: HEX.blue, title: c.title, body: c.body, titleColor: HEX.navy, bodyColor: HEX.muted, titleSize: 12.5, name: `Contexto ${i + 1}` });
  }
  s.addText("Fuentes: Banco Mundial; Banguat; FocusEconomics; Walmart; Prensa Latina · Foto: Juan Francisco, CC BY-SA 2.0", { placeholder: "source", w: P(1540) });

  // 4 · La estrategia
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addImage({ path: A("mercado.png"), x: P(1256), y: P(132), w: P(532), h: P(736), altText: "Mercado de Guatemala con frutas y especias" });
  photoCaption(s, "Foto: Francisco Anzola, CC BY 2.0");
  s.addText("PASO 3 · LA ESTRATEGIA", { placeholder: "kicker" });
  s.addText("SEMIGLOBAL: EL CORAZÓN DE\nMERCO CON SABOR CHAPÍN", { placeholder: "title" });
  const sub = { fontFace: HEAD, bold: true, fontSize: 22, color: HEX.blue, valign: "middle" };
  text(s, "SE QUEDA IGUAL", { ...sub, x: 130, y: 362, w: 500, h: 55 });
  await card(s, { x: 128, y: 427, w: 520, h: 148, fill: HEX.blue, icon: "LuUsers", iconColor: "FFFFFF", title: "Su cultura y su gente", body: "Ingreso digno y Universidad Merco", titleColor: "FFFFFF", bodyColor: HEX.pale, name: "Igual 1" });
  await card(s, { x: 668, y: 427, w: 520, h: 148, fill: HEX.blue, icon: "LuBadgeCheck", iconColor: "FFFFFF", title: "Trato, garantías y cercanía", body: "Frescura garantizada, 30 días para devolver y tiendas de proximidad", titleColor: "FFFFFF", bodyColor: HEX.pale, name: "Igual 2" });
  text(s, "SE ADAPTA", { ...sub, x: 130, y: 592, w: 500, h: 55 });
  await card(s, { x: 128, y: 657, w: 520, h: 148, fill: HEX.pale, icon: "LuBriefcase", iconColor: HEX.blue, title: "1 · Producto", body: "Frijol, maíz y café de Guatemala + lo mejor de México", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Adapta 1" });
  await card(s, { x: 668, y: 657, w: 520, h: 148, fill: HEX.pale, icon: "LuSend", iconColor: HEX.blue, title: "2 · Distribución", body: "Almacén en Ciudad de Guatemala: hoy todo sale de Nuevo León", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Adapta 2" });
  text(s, "¿POR QUÉ SEMIGLOBAL?", { fontFace: HEAD, bold: true, fontSize: 12, color: HEX.blue, x: 128, y: 827, w: 335, h: 50, valign: "middle" });
  const pill = (label, x, y, w) => text(s, label, { x, y, w, h: 50, bold: true, fontSize: 9, color: HEX.navy, align: "center", valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: label });
  pill("Capacidades: su cultura sí viaja", 470, 827, 420);
  pill("Mercado: gusto local distinto", 128, 890, 360);
  pill("Producto: surtido chapín", 499, 890, 325);
  pill("Recursos: menos inversión", 835, 890, 353);
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO)", { placeholder: "source" });

  // 5 · Oportunidades y desafíos
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Análisis" });
  s.addImage({ path: A("chichicastenango.png"), x: P(1256), y: P(132), w: P(532), h: P(736), altText: "Mercado de Chichicastenango" });
  photoCaption(s, "Foto: Toby Argüelles, CC BY-SA 2.0");
  s.addText("PASO 4 · OPORTUNIDADES Y DESAFÍOS", { placeholder: "kicker" });
  s.addText("MUCHO POR GANAR, RETOS\nQUE SÍ SE PUEDEN MANEJAR", { placeholder: "title" });
  const subW = { ...sub, color: "FFFFFF" };
  text(s, "OPORTUNIDADES", { ...subW, x: 130, y: 375, w: 500, h: 55 });
  await card(s, { x: 128, y: 440, w: 518, h: 152, fill: HEX.card, line: HEX.mid, lineW: 2, icon: "LuStar", iconColor: HEX.blue, title: "O1 · Prioridad 1: más dinero", body: "Remesas +18.7% y economía +4.1%", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Oportunidad 1" });
  await card(s, { x: 668, y: 440, w: 518, h: 152, fill: HEX.card, line: HEX.mid, lineW: 2, icon: "LuHouse", iconColor: HEX.blue, title: "O2 · Del barrio al súper", body: "El 48% aún compra en tiendas pequeñas", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Oportunidad 2" });
  text(s, "DESAFÍOS", { ...subW, x: 130, y: 617, w: 500, h: 55 });
  await card(s, { x: 128, y: 682, w: 518, h: 185, fill: HEX.pale, icon: "LuTriangleAlert", iconColor: HEX.blue, title: "D1 · Prioridad 1: Walmart", body: "276 tiendas; Merco gana con trato y frescura", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Desafío 1" });
  await card(s, { x: 668, y: 682, w: 518, h: 185, fill: HEX.pale, icon: "LuKeyRound", iconColor: HEX.blue, title: "D2 · Lejos y con inseguridad", body: "+2,100 km y extorsiones: almacén local y socio", titleColor: HEX.navy, bodyColor: HEX.muted, name: "Desafío 2" });
  s.addText("Fuentes: Banguat; Kantar; Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });

  // 6 · Acción estratégica (con el botón a la tienda)
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Celeste", sectionTitle: "Propuesta" });
  s.addText("PASO 4 · ACCIÓN ESTRATÉGICA", { placeholder: "kicker" });
  s.addText("PASO A PASO: UN SOCIO LOCAL\nY 2 TIENDAS PILOTO", { placeholder: "title" });
  text(s, "Demo: así se vería Merco GT", { x: 1348, y: 160, w: 440, h: 36, bold: true, fontSize: 10.5, color: HEX.blue, align: "center", valign: "middle" });
  text(s, [{ text: "VER LA TIENDA →", options: { hyperlink: { url: "merco-guatemala.html", tooltip: "Abrir la tienda Merco GT" } } }], {
    x: 1368, y: 206, w: 400, h: 88, fontFace: HEAD, bold: true, fontSize: 18, color: "FFFFFF", align: "center", valign: "middle",
    fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: "Botón ver la tienda",
  });
  const steps = [
    { cx: 357, fill: HEX.mid, num: HEX.navy, when: "MESES 0 A 6", what: "Socio guatemalteco, almacén y proveedores locales" },
    { cx: 960, fill: HEX.blue, num: "FFFFFF", when: "MESES 6 A 24", what: "Abrir 2 tiendas de cercanía en la capital y medir" },
    { cx: 1564, fill: HEX.navy, num: "FFFFFF", when: "MES 24", what: "Si cumple las metas, crecer a más ciudades" },
  ];
  for (const [i, st] of steps.entries()) {
    text(s, String(i + 1), { x: st.cx - 66, y: 389, w: 132, h: 132, fontFace: HEAD, bold: true, fontSize: 40, color: st.num, align: "center", valign: "middle", fill: { color: st.fill }, shape: pres.shapes.OVAL, name: `Paso ${i + 1}` });
    text(s, st.when, { x: st.cx - 230, y: 535, w: 460, h: 50, fontFace: HEAD, bold: true, fontSize: 21, color: HEX.blue, align: "center", valign: "middle" });
    text(s, st.what, { x: st.cx - 230, y: 590, w: 460, h: 80, bold: true, fontSize: 11.5, color: HEX.navy, align: "center" });
  }
  for (const ax of [612, 1216]) {
    s.addShape(pres.shapes.RIGHT_ARROW, { x: P(ax), y: P(442), w: P(96), h: P(28), fill: { color: HEX.mid }, line: { type: "none" }, objectName: "Flecha" });
  }
  text(s, "¡OFERTAS DE\nAPERTURA!", { x: 982, y: 352, w: 244, h: 106, fontFace: HEAD, bold: true, fontSize: 16, color: "FFFFFF", align: "center", valign: "middle", fill: { color: "102E50" }, rotate: 7, lineSpacingMultiple: 0.95, name: "Etiqueta ofertas" });
  await card(s, { x: 130, y: 714, w: 816, h: 151, fill: HEX.blue, icon: "LuLightbulb", iconColor: "FFFFFF", title: "Por qué esta acción", body: "Aprovecha O1 y O2 y resuelve D2; el socio ayuda a competir con Walmart (D1)", titleColor: "FFFFFF", bodyColor: HEX.pale, titleSize: 12.5, name: "Por qué" });
  await card(s, { x: 976, y: 714, w: 816, h: 151, fill: HEX.blue, icon: "LuCheck", iconColor: "FFFFFF", title: "Metas para crecer", body: "Ventas por m², clientes que regresan y rotación menor al 41%, como en México", titleColor: "FFFFFF", bodyColor: HEX.pale, titleSize: 12.5, name: "Metas" });
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Propuesta" });
  s.addText("RECOMENDACIÓN FINAL", { placeholder: "kicker" });
  s.addText("B: SÍ A GUATEMALA, PASO\nA PASO, CON UN\nSOCIO LOCAL", { placeholder: "title", h: P(260), w: P(900) });
  stamp(s, "APROBADO CON CONDICIONES", { x: 1080, y: 182, w: 708, h: 104, fontSize: 19, color: "FFFFFF", rotate: -5, name: "Sello aprobado" });
  const reasons = [
    { x: 128, icon: "LuThumbsUp", title: "1 · El mercado crece", body: "Economía +4.1% y remesas +18.7% (Banguat)" },
    { x: 694, icon: "LuStar", title: "2 · Merco sabe ganar", body: "Rotación de 98% a 41% y $118,000 por m² (caso Merco)" },
    { x: 1258, icon: "LuCheck", title: "3 · Riesgo real, pero acotado", body: "Walmart e inseguridad: por eso solo 2 tiendas (5% de su tamaño)" },
  ];
  for (const [i, r] of reasons.entries()) {
    await card(s, { x: r.x, y: 463, w: 534, h: 190, fill: HEX.card, line: HEX.mid, lineW: 2, icon: r.icon, iconColor: HEX.blue, title: r.title, body: r.body, titleColor: HEX.navy, bodyColor: HEX.muted, titleSize: 12.5, name: `Razón ${i + 1}` });
  }
  text(s, "SIGUIENTES PASOS", { ...subW, x: 130, y: 690, w: 600, h: 55 });
  const next = [
    { x: 128, when: "Ene–jun 2027", what: "Socio, almacén y proveedores" },
    { x: 692, when: "Jul 2027–dic 2028", what: "2 tiendas con ofertas de apertura" },
    { x: 1256, when: "2029", what: "Crecer a Quetzaltenango y Escuintla" },
  ];
  for (const [i, n] of next.entries()) {
    box(s, { x: n.x, y: 755, w: 536, h: 130, fill: HEX.pale, name: `Siguiente ${i + 1}` });
    text(s, String(i + 1), { x: n.x + 22, y: 770, w: 50, h: 100, fontFace: HEAD, bold: true, fontSize: 42, color: HEX.blue, align: "center", valign: "middle" });
    text(s, [
      { text: n.when, options: { bold: true, fontSize: 14, color: HEX.navy, breakLine: true } },
      { text: n.what, options: { fontSize: 11.5, color: HEX.muted } },
    ], { x: n.x + 84, y: 770, w: 430, h: 100, valign: "middle", name: `Siguiente ${i + 1} texto` });
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });

  // 8 · Fuentes
  pres.addSection({ title: "Anexo" });
  s = pres.addSlide({ masterName: "Fuentes", sectionTitle: "Anexo" });
  s.addText("NO SE EXPONE", { placeholder: "kicker", y: P(92) });
  s.addText("FUENTES", { placeholder: "title", y: P(148), h: P(80) });
  const ref = (t, last) => ({ text: t, options: { breakLine: !last, paraSpaceAfter: 5 } });
  const left = [
    "Banco de Guatemala (2026). Ingreso de divisas por remesas familiares 2025. Vía AGN y Emisoras Unidas.",
    "Banco de Guatemala (2025, dic.). Proyección de crecimiento económico 2025–2026 e inflación. banguat.gob.gt",
    "Banco Mundial (2025). Población total, Guatemala. data.worldbank.org",
    "Centroamérica 360 (2025). Inflación de Guatemala a noviembre: 1.73% (datos INE/Banguat).",
    "SEGEPLAN y UNFPA (2024). Análisis de situación de población. guatemala.unfpa.org",
    "Kantar (2022). Canales de consumo masivo en Guatemala.",
    "FocusEconomics (2026). Guatemala: PIB per cápita en US$.",
    "Prensa Latina (2025, 27 ago.). Cámaras empresariales de Guatemala señalaron aumento de extorsiones.",
    "Walmart (2025). Central America: retail units. corporate.walmart.com",
    "MINECO Guatemala (2025). Informe de comercio exterior.",
  ];
  text(s, left.map((t, i) => ref(t, i === left.length - 1)), { x: 128, y: 255, w: 805, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 1" });
  text(s, [
    ref("Carreón Flores, H. et al. (2026). Supermercados Merco: Crecimiento centrado en las personas (RC32-11-270). Tec de Monterrey."),
    ref("Great Place to Work México (2025). Merco."),
    ref("Emisoras Unidas (2025). Súper del Barrio celebra su tienda 200 en Guatemala."),
    ref("Dess, G. G. et al. (2021). Strategic management (10.ª ed.). McGraw-Hill."),
    ref("[Passport, Statista y EMIS: agregar]"),
    { text: "Fotos (Wikimedia Commons): ", options: { bold: true } },
    ref("Chad Davis, CC BY-SA 2.0 (Antigua); Juan Francisco, CC BY-SA 2.0 (Ciudad de Guatemala); Francisco Anzola, CC BY 2.0 (mercado); Toby Argüelles, CC BY-SA 2.0 (Chichicastenango). Fotos de productos de la tienda demo: créditos en la página. Mapa: Natural Earth."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 984, y: 255, w: 805, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
