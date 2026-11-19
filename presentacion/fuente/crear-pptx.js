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
function photoCaption(slide, label, x, w) {
  text(slide, label, {
    x: x + 14, y: 796, w: w - 28, h: 60, fontSize: 10, color: "FFFFFF", valign: "middle",
    fill: { color: "000000", transparency: 45 }, name: "Pie de foto",
  });
}
// Encabezado de sección dentro de una lámina
function sectionHead(slide, label, o) {
  text(slide, label, { fontFace: HEAD, bold: true, fontSize: o.fontSize || 20, color: o.color || HEX.blue, valign: "middle", x: o.x, y: o.y, w: o.w, h: o.h || 50 });
}
// Cifra grande con su leyenda debajo
function bigStat(slide, value, caption, o) {
  text(slide, [
    { text: value, options: { fontFace: HEAD, bold: true, fontSize: o.size || 30, color: o.color, breakLine: true } },
    { text: caption, options: { fontSize: 9.5, color: o.captionColor || o.color } },
  ], { x: o.x, y: o.y, w: o.w, h: o.h, valign: "middle", lineSpacingMultiple: 0.95, name: o.name });
}

(async () => {
  // 1 · Portada
  pres.addSection({ title: "Inicio" });
  let s = pres.addSlide({ masterName: "Portada", sectionTitle: "Inicio" });
  s.addImage({ path: A("portada-foto.jpg"), x: P(1127), y: 0, w: P(793), h: P(1080), altText: "Arco de Santa Catalina, Antigua Guatemala" });
  s.addImage({ path: A("fundido-h.png"), x: P(1126), y: 0, w: P(330), h: P(1080), altText: "Fundido decorativo" });
  s.addText("PROPUESTA DE INTERNACIONALIZACIÓN · SUPERMERCADOS MERCO", { placeholder: "kicker" });
  s.addText([
    { text: "EL SÚPER DE ", options: {} },
    { text: "BARRIO", options: { color: HEX.sky, breakLine: true } },
    { text: "QUE GUATEMALA" },
  ], { placeholder: "title", fontSize: 44, y: P(196), h: P(200) });
  text(s, " TODAVÍA NO TIENE ", { x: 128, y: 400, w: 780, h: 108, fontFace: HEAD, bold: true, fontSize: 34, color: HEX.blue, fill: { color: "F5F8FC" }, valign: "middle", lineSpacingMultiple: 0.9, name: "Gancho destacado" });
  text(s, "Merco lleva 80 años siendo el súper de cercanía de las familias del norte de México. Proponemos llevarlo a Ciudad de Guatemala.", { x: 130, y: 526, w: 990, h: 70, fontSize: 13, color: HEX.soft, name: "Subtítulo" });
  box(s, { x: 128, y: 662, w: 598, h: 240, fill: HEX.card, line: HEX.mid, lineW: 3, rotate: -3, name: "Dato tienda de barrio" });
  box(s, { x: 155, y: 765, w: 30, h: 30, fill: HEX.navy, shape: pres.shapes.OVAL, name: "Punto" });
  text(s, [
    { text: "1 de cada 2", options: { fontFace: HEAD, bold: true, fontSize: 32, color: HEX.blue, breakLine: true } },
    { text: "compras en Guatemala se hacen en la tienda de barrio", options: { bold: true, fontSize: 12.5, color: HEX.navy, breakLine: true } },
    { text: "Kantar", options: { fontSize: 11, color: HEX.muted } },
  ], { x: 210, y: 680, w: 490, h: 205, rotate: -3, valign: "middle", name: "Dato texto" });
  s.addImage({ path: A("mapa.png"), x: P(762), y: P(655), w: P(310), h: P(251), altText: "Mapa: de Monterrey a Guatemala" });
  s.addText("Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504)", { placeholder: "source" });
  text(s, "Fuente: Kantar · Foto: Chad Davis, CC BY-SA 2.0 · Mapa: Natural Earth", { x: 1020, y: 982, w: 770, h: 40, fontSize: 9, color: "FFFFFF", valign: "middle", align: "right", name: "Créditos portada" });
  s.addNotes("En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio. Merco lleva 80 años siendo justo eso en el norte de México: un súper cercano para familias que cuidan su dinero. Hoy les proponemos llevar ese modelo a Ciudad de Guatemala. Les vamos a mostrar si Merco está lista, por qué Guatemala, cómo entraríamos y qué recomendamos.");

  // 2 · Fortalezas y debilidades
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 1 · DIAGNÓSTICO", { placeholder: "kicker" });
  s.addText("¿QUÉ TAN LISTA ESTÁ MERCO?", { placeholder: "title" });
  stamp(s, "PREPARACIÓN MEDIA:\nLISTA PARA EMPEZAR", { x: 1270, y: 120, w: 530, h: 170, fontSize: 19, color: HEX.blue, rotate: -6, name: "Sello preparación" });
  sectionHead(s, "LO QUE MERCO HACE BIEN", { x: 130, y: 318, w: 800, fontSize: 22 });
  sectionHead(s, "LO QUE LE FALTA", { x: 979, y: 318, w: 800, fontSize: 22 });
  const good = [
    { v: "41%", c: "rotación (antes 98%)", t: "Su gente se queda", b: "Paga un ingreso digno y forma a su equipo en la Universidad Merco." },
    { v: "$118,000", c: "ventas por m² (2023)", t: "Tiendas chicas que rinden", b: "Vende más por metro cuadrado que Soriana y Walmart." },
    { v: "80 años", c: "en el norte de México", t: "Experta en el súper de barrio", b: "Conoce a las familias que cuidan su dinero: el mismo cliente de Guatemala." },
  ];
  const bad = [
    { v: "39", c: "tiendas en 2 estados", t: "Es chica", b: "Walmart tiene 3,316 tiendas en México: menos dinero y fuerza para crecer." },
    { v: "0", c: "países fuera de México", t: "Nunca ha salido", b: "No tiene experiencia operando en otro país ni con otras leyes." },
    { v: "0", c: "programas de lealtad", t: "Atrás en lo digital", b: "Tiene app y tienda en línea, pero no un programa de puntos." },
  ];
  for (const [i, g] of good.entries()) {
    const y = 380 + i * 166;
    box(s, { x: 128, y, w: 815, h: 150, fill: HEX.blue, name: `Fortaleza ${i + 1}` });
    bigStat(s, g.v, g.c, { x: 156, y: y + 15, w: 260, h: 120, size: 24, color: "FFFFFF", captionColor: HEX.pale, name: `Fortaleza ${i + 1} cifra` });
    text(s, [
      { text: g.t, options: { bold: true, fontSize: 13, color: "FFFFFF", breakLine: true } },
      { text: g.b, options: { fontSize: 10.5, color: HEX.pale } },
    ], { x: 436, y: y + 15, w: 485, h: 120, valign: "middle", name: `Fortaleza ${i + 1} texto` });
  }
  for (const [i, d] of bad.entries()) {
    const y = 380 + i * 166;
    box(s, { x: 977, y, w: 815, h: 150, fill: HEX.pale, name: `Debilidad ${i + 1}` });
    bigStat(s, d.v, d.c, { x: 1005, y: y + 15, w: 260, h: 120, size: 24, color: HEX.blue, captionColor: HEX.muted, name: `Debilidad ${i + 1} cifra` });
    text(s, [
      { text: d.t, options: { bold: true, fontSize: 13, color: HEX.navy, breakLine: true } },
      { text: d.b, options: { fontSize: 10.5, color: HEX.muted } },
    ], { x: 1285, y: y + 15, w: 485, h: 120, valign: "middle", name: `Debilidad ${i + 1} texto` });
  }
  text(s, [
    { text: "En resumen: ", options: { bold: true } },
    { text: "su modelo funciona, pero es chica y nunca ha salido → debe entrar pequeño y con un socio local." },
  ], { x: 128, y: 880, w: 1664, h: 62, fontSize: 12, color: "FFFFFF", valign: "middle", fill: { color: HEX.navy }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, name: "Conclusión diagnóstico" });
  s.addText("Fuentes: diagnóstico interno del Equipo 2; caso Merco (Carreón et al., 2026); Great Place to Work (2025)", { placeholder: "source" });
  s.addNotes("A la izquierda, lo que Merco hace bien: su gente se queda (la rotación bajó de 98% a 41%), sus tiendas chicas venden más por metro cuadrado que Soriana y Walmart, y lleva 80 años como súper de barrio. A la derecha, lo que le falta: es chica, nunca ha operado fuera de México y no tiene programa de lealtad. Conclusión: está lista para empezar, pero en pequeño y con ayuda local.");

  // 3 · El mercado
  s = pres.addSlide({ masterName: "Oscuro", sectionTitle: "Análisis" });
  s.addImage({ path: A("ciudad.jpg"), x: 0, y: P(810), w: P(1920), h: P(175), altText: "Vista de Ciudad de Guatemala" });
  s.addImage({ path: A("fundido-v.png"), x: 0, y: P(808), w: P(1920), h: P(110), altText: "Fundido decorativo" });
  s.addText("PASO 2 · EL MERCADO", { placeholder: "kicker" });
  s.addText("¿POR QUÉ GUATEMALA?", { placeholder: "title" });
  text(s, "MERCADO ELEGIDO: GUATEMALA", { x: 1212, y: 188, w: 580, h: 56, fontFace: HEAD, bold: true, fontSize: 13, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: "Mercado elegido" });
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
    await card(s, { x: c.x, y: 630, w: 532, h: 150, fill: HEX.card, line: HEX.mid, lineW: 2, icon: c.icon, iconColor: HEX.blue, title: c.title, body: c.body, titleColor: HEX.navy, bodyColor: HEX.muted, titleSize: 12, bodySize: 9.5, name: `Contexto ${i + 1}` });
  }
  text(s, [
    { text: "Conclusión: ", options: { bold: true } },
    { text: "un mercado grande, que crece y recibe dinero de fuera. El reto es Walmart y la distancia." },
  ], { x: 128, y: 830, w: 1664, h: 70, fontSize: 13, color: "FFFFFF", valign: "middle", fill: { color: HEX.navy, transparency: 15 }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, name: "Conclusión mercado" });
  s.addText("Fuentes: Banco Mundial; Banguat; FocusEconomics; Walmart; Prensa Latina · Foto: Juan Francisco, CC BY-SA 2.0", { placeholder: "source" });
  s.addNotes("Elegimos Guatemala por tres datos: es el país más poblado de Centroamérica, su economía crecerá 4.1% y las remesas subieron 18.7%, así que las familias tienen más dinero para el súper. Además el cliente cuida su dinero, igual que el de Merco. Los retos son claros: Walmart ya tiene 276 tiendas, hay extorsiones y estamos a más de 2,100 km.");

  // 4 · Estrategia semiglobal
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 3 · LA ESTRATEGIA", { placeholder: "kicker" });
  s.addText("SEMIGLOBAL: EL CORAZÓN DE\nMERCO CON SABOR CHAPÍN", { placeholder: "title" });
  text(s, [
    { text: "Semiglobal = ", options: { bold: true, color: HEX.blue } },
    { text: "lo que hace única a Merco se queda igual; solo adaptamos 2 áreas al cliente guatemalteco." },
  ], { x: 128, y: 384, w: 1664, h: 54, fontSize: 12, color: HEX.navy, valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: "Definición" });
  const colW = 535;
  const cols = [128, 128 + colW + 29, 128 + 2 * (colW + 29)];
  // Columna 1: lo que se queda igual
  text(s, "SE QUEDA IGUAL", { x: cols[0], y: 460, w: colW, h: 58, fontFace: HEAD, bold: true, fontSize: 16, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.navy }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1 });
  box(s, { x: cols[0], y: 530, w: colW, h: 330, fill: HEX.blue, name: "Igual" });
  const same = [
    { icon: "LuUsers", t: "Su gente", b: "Ingreso digno y Universidad Merco" },
    { icon: "LuBadgeCheck", t: "Sus garantías", b: "Frescura garantizada y 30 días para devolver" },
    { icon: "LuHouse", t: "Su formato", b: "Súper de cercanía para familias" },
  ];
  for (const [i, it] of same.entries()) {
    const y = 542 + i * 104;
    s.addImage({ data: await icon(it.icon, "FFFFFF"), x: P(cols[0] + 30), y: P(y + 30), w: P(36), h: P(36), altText: it.t });
    text(s, [
      { text: it.t, options: { bold: true, fontSize: 12.5, color: "FFFFFF", breakLine: true } },
      { text: it.b, options: { fontSize: 10, color: HEX.pale } },
    ], { x: cols[0] + 86, y, w: colW - 110, h: 96, valign: "middle" });
  }
  // Columnas 2 y 3: las dos áreas de adaptación
  const adapt = [
    { head: "ADAPTACIÓN 1 · PRODUCTO", icon: "LuShoppingCart", hoy: "Surtido para el norte de México", gt: "Frijol negro, maíz y café de Guatemala + lo mejor de México", por: "El cliente busca sus productos de siempre." },
    { head: "ADAPTACIÓN 2 · DISTRIBUCIÓN", icon: "LuTruck", hoy: "Todo sale de 3 centros en Nuevo León", gt: "Almacén en Ciudad de Guatemala y proveedores locales", por: "+2,100 km es demasiado lejos para lo fresco." },
  ];
  for (const [i, a] of adapt.entries()) {
    const x = cols[i + 1];
    text(s, a.head, { x, y: 460, w: colW, h: 58, fontFace: HEAD, bold: true, fontSize: 14, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1 });
    box(s, { x, y: 530, w: colW, h: 330, fill: "FFFFFF", line: HEX.pale, lineW: 1.5, name: a.head });
    text(s, [
      { text: "HOY EN MÉXICO", options: { fontFace: HEAD, bold: true, fontSize: 12, color: HEX.muted, breakLine: true } },
      { text: a.hoy, options: { fontSize: 10.5, color: HEX.muted } },
    ], { x: x + 30, y: 544, w: colW - 60, h: 80 });
    s.addShape(pres.shapes.DOWN_ARROW, { x: P(x + colW / 2 - 16), y: P(632), w: P(32), h: P(28), fill: { color: HEX.orange }, line: { type: "none" }, objectName: "Flecha" });
    text(s, [
      { text: "EN GUATEMALA", options: { fontFace: HEAD, bold: true, fontSize: 12, color: HEX.blue, breakLine: true } },
      { text: a.gt, options: { bold: true, fontSize: 11.5, color: HEX.navy } },
    ], { x: x + 30, y: 662, w: colW - 60, h: 96 });
    text(s, [
      { text: "Por qué: ", options: { bold: true } },
      { text: a.por },
    ], { x: x + 20, y: 774, w: colW - 40, h: 70, fontSize: 10, color: HEX.navy, valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1 });
  }
  sectionHead(s, "¿POR QUÉ SEMIGLOBAL?", { x: 128, y: 882, w: 345, fontSize: 12 });
  const pill = (label, x, w) => text(s, label, { x, y: 882, w, h: 50, bold: true, fontSize: 10, color: HEX.navy, align: "center", valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: label });
  pill("Capacidades: su cultura sí viaja", 480, 420);
  pill("Mercado: gusto local distinto", 912, 430);
  pill("Recursos: menos inversión", 1354, 438);
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO) · Centros de distribución: diagnóstico interno", { placeholder: "source" });
  s.addNotes("Nuestra estrategia es semiglobal: lo que hace única a Merco se queda igual —su gente, sus garantías y su súper de cercanía— y solo adaptamos dos áreas. Primero el producto: en Guatemala la gente busca frijol negro, maíz y café local, así que el surtido mezcla lo chapín con lo mejor de México. Segundo la distribución: hoy todo sale de Nuevo León, pero a más de 2,100 km lo fresco no llega bien, así que abrimos un almacén en Ciudad de Guatemala con proveedores locales.");

  // 5 · Oportunidades y desafíos
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Análisis" });
  s.addImage({ path: A("chichicastenango-estrecha.png"), x: P(1412), y: P(132), w: P(380), h: P(736), altText: "Mercado de Chichicastenango" });
  photoCaption(s, "Foto: Toby Argüelles,\nCC BY-SA 2.0", 1412, 380);
  s.addText("PASO 4 · OPORTUNIDADES Y DESAFÍOS", { placeholder: "kicker" });
  s.addText("MUCHO POR GANAR, RETOS\nQUE SÍ SE PUEDEN MANEJAR", { placeholder: "title" });
  const od = [
    { head: "OPORTUNIDADES", y: 416, fill: HEX.card, items: [
      { v: "+18.7%", c: "remesas 2025", t: "O1 · Más dinero en casa", b: "Remesas por US$25,530 M y economía +4.1%.", r: "Lo aprovechamos con combos familiares y precios claros." },
      { v: "48%", c: "aún compra en tiendas pequeñas", t: "O2 · Del barrio al súper", b: "Hay espacio para un súper de barrio moderno.", r: "Lo aprovechamos con tiendas chicas, cerca de casa." },
    ] },
    { head: "DESAFÍOS", y: 706, fill: HEX.pale, items: [
      { v: "276", c: "tiendas de Walmart", t: "D1 · Walmart ya está", b: "Es mucho más grande y compite por precio.", r: "Lo manejamos ganando con trato y frescura, no con precio." },
      { v: "2,100 km", c: "desde Nuevo León", t: "D2 · Lejos e inseguro", b: "Distancia y extorsiones a comercios.", r: "Lo manejamos con almacén local y un socio guatemalteco." },
    ] },
  ];
  for (const row of od) {
    sectionHead(s, row.head, { x: 130, y: row.y - 46, w: 600, color: "FFFFFF", fontSize: 18, h: 42 });
    for (const [i, it] of row.items.entries()) {
      const x = 128 + i * 642;
      box(s, { x, y: row.y, w: 612, h: 234, fill: row.fill, line: row.fill === HEX.card ? HEX.mid : undefined, lineW: 2, name: it.t });
      bigStat(s, it.v, it.c, { x: x + 22, y: row.y + 14, w: 220, h: 132, size: 22, color: HEX.blue, captionColor: HEX.muted, name: `${it.t} cifra` });
      text(s, [
        { text: it.t, options: { bold: true, fontSize: 12.5, color: HEX.navy, breakLine: true } },
        { text: it.b, options: { fontSize: 10, color: HEX.muted } },
      ], { x: x + 252, y: row.y + 14, w: 342, h: 132, valign: "middle" });
      text(s, [{ text: "→ " }, { text: it.r }], { x: x + 18, y: row.y + 152, w: 576, h: 68, fontSize: 10, bold: true, color: "FFFFFF", valign: "middle", fill: { color: HEX.blue }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: `${it.t} respuesta` });
    }
  }
  s.addText("Fuentes: Banguat; Kantar; Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Dos oportunidades y dos desafíos, cada uno con su respuesta. O1: hay más dinero en los hogares por las remesas, y lo aprovechamos con combos familiares. O2: casi la mitad aún compra en tiendas pequeñas; ahí entra el súper de barrio de Merco. D1: Walmart ya está, así que no peleamos por precio sino por trato y frescura. D2: la distancia y la inseguridad, que resolvemos con un almacén local y un socio guatemalteco.");

  // 6 · Acción estratégica
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Celeste", sectionTitle: "Propuesta" });
  s.addText("PASO 4 · ACCIÓN ESTRATÉGICA", { placeholder: "kicker" });
  s.addText("PASO A PASO: UN SOCIO LOCAL\nY 2 TIENDAS PILOTO", { placeholder: "title" });
  const phases = [
    { n: "1", head: "PREPARAR", when: "Meses 0 a 6", fill: HEX.mid, num: HEX.navy, what: "Firmar con un socio guatemalteco, abrir el almacén en la capital y elegir proveedores locales.", gate: "Para avanzar: almacén y proveedores listos" },
    { n: "2", head: "PROBAR", when: "Meses 6 a 24", fill: HEX.blue, num: "FFFFFF", what: "Abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura y medir resultados.", gate: "Para avanzar: cumplir las 3 metas del piloto" },
    { n: "3", head: "CRECER", when: "Desde el mes 24", fill: HEX.navy, num: "FFFFFF", what: "Llevar el modelo a Quetzaltenango y Escuintla.", gate: "Solo si el piloto funciona" },
  ];
  for (const [i, ph] of phases.entries()) {
    const x = cols[i];
    box(s, { x, y: 400, w: colW, h: 360, fill: "FFFFFF", name: `Fase ${ph.n}` });
    text(s, ph.n, { x: x + 26, y: 420, w: 84, h: 84, fontFace: HEAD, bold: true, fontSize: 30, color: ph.num, align: "center", valign: "middle", fill: { color: ph.fill }, shape: pres.shapes.OVAL, name: `Número fase ${ph.n}` });
    text(s, [
      { text: `FASE ${ph.n} · ${ph.when.toUpperCase()}`, options: { bold: true, fontSize: 10.5, color: HEX.muted, breakLine: true } },
      { text: ph.head, options: { fontFace: HEAD, bold: true, fontSize: 24, color: HEX.blue } },
    ], { x: x + 130, y: 420, w: colW - 150, h: 84, valign: "middle" });
    text(s, ph.what, { x: x + 26, y: 524, w: colW - 52, h: 130, fontSize: 11.5, color: HEX.navy });
    text(s, ph.gate, { x: x + 20, y: 676, w: colW - 40, h: 62, bold: true, fontSize: 10.5, color: HEX.navy, align: "center", valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1 });
    if (i < 2) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x: P(x + colW - 4), y: P(560), w: P(37), h: P(40), fill: { color: HEX.orange }, line: { type: "none" }, objectName: "Flecha" });
    }
  }
  text(s, "¡OFERTAS DE\nAPERTURA!", { x: cols[1] + colW - 210, y: 318, w: 230, h: 80, fontFace: HEAD, bold: true, fontSize: 13, color: "FFFFFF", align: "center", valign: "middle", fill: { color: "102E50" }, rotate: 7, lineSpacingMultiple: 0.95, name: "Etiqueta ofertas" });
  box(s, { x: 128, y: 790, w: 816, h: 150, fill: HEX.blue, name: "Metas" });
  text(s, "METAS DEL PILOTO", { x: 160, y: 802, w: 400, h: 40, fontFace: HEAD, bold: true, fontSize: 15, color: "FFFFFF", valign: "middle" });
  const goals = ["Más ventas por m²", "Clientes que regresan", "Rotación menor al 41%"];
  for (const [i, g] of goals.entries()) {
    text(s, g, { x: 160 + i * 255, y: 852, w: 240, h: 66, bold: true, fontSize: 10, color: HEX.navy, align: "center", valign: "middle", fill: { color: "FFFFFF" }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1 });
  }
  await card(s, { x: 976, y: 790, w: 816, h: 150, fill: HEX.blue, icon: "LuLightbulb", iconColor: "FFFFFF", title: "Por qué esta acción", body: "Aprovecha O1 y O2, resuelve D2 con el socio y el almacén, y compite con Walmart (D1) con trato, no con precio. Arriesga poco: 2 tiendas son el 5% del tamaño de Merco.", titleColor: "FFFFFF", bodyColor: HEX.pale, titleSize: 12.5, name: "Por qué" });
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });
  s.addNotes("La acción tiene tres fases y no pasamos a la siguiente sin cumplir la anterior. Fase 1, preparar: socio, almacén y proveedores en los primeros 6 meses. Fase 2, probar: 2 tiendas en la capital con ofertas de apertura, midiendo ventas por metro cuadrado, clientes que regresan y rotación menor al 41%. Fase 3, crecer a Quetzaltenango y Escuintla, solo si el piloto funciona. Así arriesgamos poco: 2 tiendas son apenas el 5% del tamaño de Merco.");

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Propuesta" });
  s.addText("RECOMENDACIÓN FINAL", { placeholder: "kicker" });
  s.addText("RECOMENDAMOS LA OPCIÓN B:\nSÍ A GUATEMALA, PASO A PASO", { placeholder: "title" });
  stamp(s, "APROBADO CON CONDICIONES", { x: 1250, y: 196, w: 540, h: 90, fontSize: 13.5, color: "FFFFFF", rotate: -5, name: "Sello aprobado" });
  const options = [
    { k: "A", t: "Quedarse en México", b: "Sin riesgo, pero no aprovecha un mercado que crece.", chosen: false },
    { k: "B", t: "Guatemala paso a paso", b: "Socio local y 2 tiendas piloto: crecer con poco riesgo.", chosen: true },
    { k: "C", t: "Entrar en grande", b: "Muchas tiendas de golpe: demasiado riesgo con Walmart y la distancia.", chosen: false },
  ];
  for (const [i, o] of options.entries()) {
    const x = cols[i];
    box(s, { x, y: 400, w: colW, h: 190, fill: o.chosen ? "FFFFFF" : HEX.pale, line: o.chosen ? HEX.orange : undefined, lineW: 4, name: `Opción ${o.k}` });
    text(s, o.k, { x: x + 26, y: 440, w: 80, h: 80, fontFace: HEAD, bold: true, fontSize: 30, color: o.chosen ? "FFFFFF" : HEX.blue, align: "center", valign: "middle", fill: { color: o.chosen ? HEX.orange : "FFFFFF" }, shape: pres.shapes.OVAL });
    text(s, [
      { text: o.t, options: { bold: true, fontSize: 13, color: HEX.navy, breakLine: true } },
      { text: o.b, options: { fontSize: 10.5, color: HEX.muted } },
    ], { x: x + 126, y: 418, w: colW - 150, h: 154, valign: "middle", name: `Opción ${o.k} texto` });
    if (o.chosen) {
      text(s, "ELEGIDA", { x: x + colW - 170, y: 380, w: 150, h: 40, fontFace: HEAD, bold: true, fontSize: 12, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.2 });
    }
  }
  sectionHead(s, "POR QUÉ B", { x: 130, y: 614, w: 600, color: "FFFFFF", fontSize: 20 });
  const why = [
    { n: "1", t: "El mercado crece", b: "economía +4.1% y remesas +18.7% (Banguat)" },
    { n: "2", t: "Merco sabe ganar", b: "rotación de 98% a 41% y $118,000 por m²" },
    { n: "3", t: "El riesgo es pequeño", b: "2 tiendas son el 5% de su tamaño" },
  ];
  for (const [i, w] of why.entries()) {
    const y = 672 + i * 92;
    text(s, w.n, { x: 128, y, w: 70, h: 76, fontFace: HEAD, bold: true, fontSize: 26, color: HEX.blue, align: "center", valign: "middle", fill: { color: HEX.pale }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15 });
    text(s, [
      { text: `${w.t}: `, options: { bold: true, color: "FFFFFF" } },
      { text: w.b, options: { color: HEX.pale } },
    ], { x: 218, y, w: 830, h: 76, fontSize: 12, valign: "middle" });
  }
  box(s, { x: 1100, y: 614, w: 692, h: 326, fill: HEX.card, name: "Siguientes pasos" });
  sectionHead(s, "SIGUIENTES PASOS", { x: 1134, y: 630, w: 600, fontSize: 18 });
  const steps = [
    { d: "Ene–jun 2027", t: "Socio, almacén y proveedores" },
    { d: "Jul 2027–dic 2028", t: "2 tiendas con ofertas de apertura" },
    { d: "2029", t: "Crecer a Quetzaltenango y Escuintla" },
  ];
  for (const [i, st] of steps.entries()) {
    const y = 690 + i * 80;
    text(s, String(i + 1), { x: 1134, y: y + 6, w: 52, h: 52, fontFace: HEAD, bold: true, fontSize: 20, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.blue }, shape: pres.shapes.OVAL });
    text(s, [
      { text: st.d, options: { bold: true, fontSize: 12.5, color: HEX.navy, breakLine: true } },
      { text: st.t, options: { fontSize: 10.5, color: HEX.muted } },
    ], { x: 1206, y, w: 560, h: 64, valign: "middle" });
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });
  s.addNotes("Comparamos tres opciones. A, quedarse en México: no arriesga, pero deja pasar un mercado que crece. C, entrar en grande: demasiado riesgo frente a Walmart y la distancia. Recomendamos B: entrar a Guatemala paso a paso, con un socio local y 2 tiendas piloto. Por tres razones: el mercado crece, Merco sabe ganar con su gente y el riesgo es pequeño. La aprobamos con condiciones: solo crecemos si el piloto cumple sus metas. Siguientes pasos: socio en 2027, tiendas en 2028 y crecer en 2029. Gracias.");
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
    ref("Equipo 2 (2026). Diagnóstico externo, diagnóstico interno y matriz FODA de Supermercados Merco."),
    { text: "Fotos (Wikimedia Commons): ", options: { bold: true } },
    ref("Chad Davis, CC BY-SA 2.0 (Antigua); Juan Francisco, CC BY-SA 2.0 (Ciudad de Guatemala); Toby Argüelles, CC BY-SA 2.0 (Chichicastenango). Mapa: Natural Earth."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 984, y: 255, w: 805, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
