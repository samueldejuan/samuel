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
function sectionHead(slide, label, o) {
  text(slide, label, { fontFace: HEAD, bold: true, fontSize: o.fontSize || 20, color: o.color || HEX.blue, valign: "middle", x: o.x, y: o.y, w: o.w, h: o.h || 50, charSpacing: o.charSpacing });
}
// Círculo de color con un icono dentro
async function iconCircle(slide, name, o) {
  box(slide, { x: o.x, y: o.y, w: o.d, h: o.d, fill: o.fill, shape: pres.shapes.OVAL, name: `Círculo ${name}` });
  const pad = o.d * 0.24;
  slide.addImage({ data: await icon(name, o.color), x: P(o.x + pad), y: P(o.y + pad), w: P(o.d - 2 * pad), h: P(o.d - 2 * pad), altText: o.alt || name });
}
// Cifra enorme con una etiqueta corta
function bigNumber(slide, value, label, o) {
  o.nh = o.nh || Math.round((o.size || 44) * 3.1);
  text(slide, value, { x: o.x, y: o.y, w: o.w, h: o.nh, fontFace: HEAD, bold: true, fontSize: o.size || 44, color: o.color, valign: "bottom", align: o.align, name: `${label} cifra` });
  text(slide, label, { x: o.x, y: o.y + o.nh, w: o.lw || o.w, h: o.lh || 50, bold: true, fontSize: o.labelSize || 12.5, color: o.labelColor || o.color, align: o.align, name: label });
}

(async () => {
  const colW = 535;
  const cols = [128, 128 + colW + 29, 128 + 2 * (colW + 29)];

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
  ], { placeholder: "title", fontSize: 50, y: P(200), h: P(220) });
  text(s, " TODAVÍA NO TIENE ", { x: 128, y: 430, w: 780, h: 118, fontFace: HEAD, bold: true, fontSize: 36, color: HEX.blue, fill: { color: "F5F8FC" }, valign: "middle", name: "Gancho destacado" });
  box(s, { x: 128, y: 640, w: 598, h: 240, fill: HEX.card, line: HEX.mid, lineW: 3, rotate: -3, name: "Dato tienda de barrio" });
  text(s, [
    { text: "1 de cada 2", options: { fontFace: HEAD, bold: true, fontSize: 38, color: HEX.blue, breakLine: true } },
    { text: "compras se hacen en la tienda de barrio", options: { bold: true, fontSize: 13, color: HEX.navy } },
  ], { x: 175, y: 665, w: 520, h: 190, rotate: -3, valign: "middle", name: "Dato texto" });
  s.addImage({ path: A("mapa.png"), x: P(762), y: P(640), w: P(310), h: P(251), altText: "Mapa: de Monterrey a Guatemala" });
  s.addText("Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504)", { placeholder: "source" });
  text(s, "Fuente: Kantar · Foto: Chad Davis, CC BY-SA 2.0 · Mapa: Natural Earth", { x: 1020, y: 982, w: 770, h: 40, fontSize: 9, color: "FFFFFF", valign: "middle", align: "right", name: "Créditos portada" });
  s.addNotes("En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio (Kantar). Merco lleva 80 años siendo justo eso en el norte de México: un súper cercano para familias que cuidan su dinero. Hoy les proponemos llevar ese modelo a Ciudad de Guatemala. Vamos a ver si Merco está lista, por qué Guatemala, cómo entraríamos y qué recomendamos.");

  // 2 · Fortalezas y debilidades
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 1 · DIAGNÓSTICO", { placeholder: "kicker" });
  s.addText("¿QUÉ TAN LISTA ESTÁ MERCO?", { placeholder: "title" });
  stamp(s, "PREPARACIÓN MEDIA:\nLISTA PARA EMPEZAR", { x: 1270, y: 120, w: 530, h: 170, fontSize: 19, color: HEX.blue, rotate: -6, name: "Sello preparación" });
  sectionHead(s, "HACE BIEN", { x: 130, y: 330, w: 800, fontSize: 24 });
  sectionHead(s, "LE FALTA", { x: 980, y: 330, w: 800, fontSize: 24 });
  const good = [
    { i: "LuUsers", v: "41%", l: "de rotación (antes 98%)" },
    { i: "LuStore", v: "$118K", l: "por m², más que Walmart" },
    { i: "LuHeartHandshake", v: "80", l: "años de súper de barrio" },
  ];
  const bad = [
    { i: "LuMapPin", v: "39", l: "tiendas en solo 2 estados" },
    { i: "LuRoute", v: "0", l: "países fuera de México" },
    { i: "LuSmartphone", v: "0", l: "programas de lealtad" },
  ];
  for (const [side, list] of [[0, good], [1, bad]]) {
    for (const [i, t] of list.entries()) {
      const x = side ? 977 : 128;
      const y = 392 + i * 148;
      const dark = side === 0;
      box(s, { x, y, w: 815, h: 132, fill: dark ? HEX.blue : HEX.pale, name: `${dark ? "Fortaleza" : "Debilidad"} ${i + 1}` });
      await iconCircle(s, t.i, { x: x + 28, y: y + 31, d: 70, fill: dark ? "FFFFFF" : HEX.blue, color: dark ? HEX.blue : "FFFFFF" });
      text(s, t.v, { x: x + 118, y, w: 310, h: 132, fontFace: HEAD, bold: true, fontSize: 42, color: dark ? "FFFFFF" : HEX.blue, valign: "middle", name: `${t.l} cifra` });
      text(s, t.l, { x: x + 440, y, w: 355, h: 132, bold: true, fontSize: 15, color: dark ? "FFFFFF" : HEX.navy, valign: "middle", name: t.l });
    }
  }
  text(s, "→  ENTRAR EN PEQUEÑO Y CON UN SOCIO LOCAL", { x: 410, y: 850, w: 1100, h: 72, fontFace: HEAD, bold: true, fontSize: 17, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.2, name: "Conclusión" });
  s.addText("Fuentes: diagnóstico interno del Equipo 2; caso Merco (Carreón et al., 2026); Great Place to Work (2025)", { placeholder: "source" });
  s.addNotes("Lo que Merco hace bien: su gente se queda —la rotación bajó de 98% a 41% gracias al ingreso digno y la Universidad Merco—; sus tiendas chicas venden $118,000 por metro cuadrado, más que Soriana y Walmart; y lleva 80 años como súper de barrio para familias que cuidan su dinero, el mismo cliente que hay en Guatemala. Lo que le falta: es chica, con 39 tiendas en 2 estados frente a las 3,316 de Walmart; nunca ha operado fuera de México; y no tiene programa de lealtad. Por eso la preparación es media: está lista para empezar, pero en pequeño y con un socio local.");

  // 3 · El mercado
  s = pres.addSlide({ masterName: "Oscuro", sectionTitle: "Análisis" });
  s.addImage({ path: A("ciudad.jpg"), x: 0, y: P(735), w: P(1920), h: P(250), altText: "Vista de Ciudad de Guatemala" });
  s.addImage({ path: A("fundido-v.png"), x: 0, y: P(733), w: P(1920), h: P(140), altText: "Fundido decorativo" });
  s.addText("PASO 2 · EL MERCADO", { placeholder: "kicker" });
  s.addText("¿POR QUÉ GUATEMALA?", { placeholder: "title" });
  text(s, "MERCADO ELEGIDO: GUATEMALA", { x: 1212, y: 188, w: 580, h: 56, fontFace: HEAD, bold: true, fontSize: 13, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, name: "Mercado elegido" });
  const stats = [
    { x: 128, rot: -2.5, big: "18.7 M", desc: "personas" },
    { x: 695, rot: 2, big: "+4.1%", desc: "crecerá la economía" },
    { x: 1262, rot: -2, big: "+18.7%", desc: "más remesas en 2025" },
  ];
  for (const [i, st] of stats.entries()) {
    box(s, { x: st.x, y: 320, w: 530, h: 300, fill: HEX.card, line: HEX.mid, lineW: 3, rotate: st.rot, name: `Dato ${i + 1}` });
    text(s, [
      { text: st.big, options: { fontFace: HEAD, bold: true, fontSize: 56, color: HEX.blue, breakLine: true } },
      { text: st.desc, options: { bold: true, fontSize: 16, color: HEX.navy } },
    ], { x: st.x + 40, y: 335, w: 470, h: 270, rotate: st.rot, valign: "middle", name: `Dato ${i + 1} texto` });
  }
  sectionHead(s, "EL RETO", { x: 130, y: 676, w: 260, color: HEX.sky, fontSize: 20, charSpacing: 2 });
  const reto = [
    { i: "LuStore", t: "Walmart: 276 tiendas", x: 400 },
    { i: "LuRoute", t: "2,100 km de distancia", x: 900 },
  ];
  for (const r of reto) {
    box(s, { x: r.x, y: 668, w: 470, h: 66, fill: "FFFFFF", radius: 0.3, name: r.t });
    s.addImage({ data: await icon(r.i, HEX.blue), x: P(r.x + 22), y: P(682), w: P(38), h: P(38), altText: r.t });
    text(s, r.t, { x: r.x + 76, y: 668, w: 380, h: 66, bold: true, fontSize: 13, color: HEX.navy, valign: "middle" });
  }
  s.addText("Fuentes: Banco Mundial (2025); Banguat (2025, 2026); Walmart (2025) · Foto: Juan Francisco, CC BY-SA 2.0", { placeholder: "source" });
  s.addNotes("Elegimos Guatemala por tres datos. Es el país más poblado de Centroamérica, con 18.7 millones de personas (Banco Mundial). Su economía crecerá 4.1%, más que América Latina (Banguat). Y las remesas subieron 18.7% en 2025, hasta US$25,530 millones: las familias tienen más dinero para el súper. Además el cliente cuida su dinero (PIB por persona de US$6,506), igual que el de Merco. El reto: Walmart ya tiene 276 tiendas, hay extorsiones a comercios y estamos a más de 2,100 km de Nuevo León.");

  // 4 · Estrategia semiglobal
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addImage({ path: A("mercado.png"), x: P(1256), y: P(132), w: P(532), h: P(736), altText: "Mercado de Guatemala con frutas y especias" });
  photoCaption(s, "Foto: Francisco Anzola, CC BY 2.0", 1256, 532);
  s.addText("PASO 3 · LA ESTRATEGIA", { placeholder: "kicker" });
  s.addText("SEMIGLOBAL: EL CORAZÓN DE\nMERCO CON SABOR CHAPÍN", { placeholder: "title" });
  sectionHead(s, "SE QUEDA IGUAL", { x: 130, y: 382, w: 600, fontSize: 20 });
  const same = [
    { i: "LuUsers", t: "Su gente" },
    { i: "LuShieldCheck", t: "Sus garantías" },
    { i: "LuStore", t: "Súper de barrio" },
  ];
  for (const [i, it] of same.entries()) {
    const x = 128 + i * 355;
    box(s, { x, y: 438, w: 340, h: 170, fill: HEX.blue, name: it.t });
    await iconCircle(s, it.i, { x: x + 24, y: 491, d: 64, fill: "FFFFFF", color: HEX.blue });
    text(s, it.t, { x: x + 102, y: 438, w: 228, h: 170, fontFace: HEAD, bold: true, fontSize: 17, color: "FFFFFF", valign: "middle" });
  }
  sectionHead(s, "SE ADAPTA", { x: 130, y: 632, w: 600, fontSize: 20, color: HEX.orange });
  const adapt = [
    { i: "LuWheat", t: "Producto chapín", b: "frijol, maíz y café local" },
    { i: "LuWarehouse", t: "Almacén local", b: "en Ciudad de Guatemala" },
  ];
  for (const [i, it] of adapt.entries()) {
    const x = 128 + i * 533;
    box(s, { x, y: 688, w: 518, h: 170, fill: HEX.orange, name: it.t });
    await iconCircle(s, it.i, { x: x + 28, y: 735, d: 76, fill: "FFFFFF", color: HEX.orange });
    text(s, [
      { text: it.t, options: { fontFace: HEAD, bold: true, fontSize: 19, color: "FFFFFF", breakLine: true } },
      { text: it.b, options: { fontSize: 11.5, color: "FFFFFF" } },
    ], { x: x + 124, y: 688, w: 375, h: 170, valign: "middle" });
  }
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO)", { placeholder: "source" });
  s.addNotes("Nuestra estrategia es semiglobal: lo que hace única a Merco se queda igual y solo adaptamos dos áreas. Se queda igual su gente (ingreso digno y Universidad Merco), sus garantías (frescura garantizada y 30 días para devolver) y su formato de súper de barrio. Adaptamos, primero, el producto: en Guatemala la gente busca frijol negro, maíz y café local, así que el surtido mezcla lo chapín con lo mejor de México. Y segundo, la distribución: hoy todo sale de 3 centros en Nuevo León, pero a más de 2,100 km lo fresco no llega bien, así que abrimos un almacén en Ciudad de Guatemala con proveedores locales. ¿Por qué semiglobal? Porque su cultura sí viaja, el gusto local es distinto y adaptar solo 2 áreas cuesta menos.");

  // 5 · Oportunidades y desafíos
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Análisis" });
  s.addImage({ path: A("chichicastenango-estrecha.png"), x: P(1412), y: P(132), w: P(380), h: P(736), altText: "Mercado de Chichicastenango" });
  photoCaption(s, "Foto: Toby Argüelles,\nCC BY-SA 2.0", 1412, 380);
  s.addText("PASO 4 · OPORTUNIDADES Y DESAFÍOS", { placeholder: "kicker" });
  s.addText("MUCHO POR GANAR, RETOS\nQUE SÍ SE PUEDEN MANEJAR", { placeholder: "title" });
  const od = [
    { head: "OPORTUNIDADES", y: 420, fill: HEX.card, items: [
      { v: "+18.7%", l: "más remesas", r: "combos familiares" },
      { v: "48%", l: "aún compra en tiendas pequeñas", r: "súper de barrio" },
    ] },
    { head: "DESAFÍOS", y: 700, fill: HEX.pale, items: [
      { v: "276", l: "tiendas de Walmart", r: "trato y frescura" },
      { v: "2,100 km", l: "de distancia + extorsiones", r: "socio + almacén" },
    ] },
  ];
  for (const row of od) {
    sectionHead(s, row.head, { x: 130, y: row.y - 50, w: 600, color: "FFFFFF", fontSize: 18, h: 44, charSpacing: 2 });
    for (const [i, it] of row.items.entries()) {
      const x = 128 + i * 642;
      box(s, { x, y: row.y, w: 612, h: 220, fill: row.fill, name: it.l });
      bigNumber(s, it.v, it.l, { x: x + 34, y: row.y + 4, w: 550, nh: 124, size: 40, color: HEX.blue, labelColor: HEX.navy, labelSize: 13, lh: 40 });
      text(s, `→ ${it.r}`, { x: x + 34, y: row.y + 168, w: 300, h: 42, bold: true, fontSize: 10.5, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.blue }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.25, name: `${it.l} respuesta` });
    }
  }
  s.addText("Fuentes: Banguat (2026); Kantar (2022); Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Dos oportunidades y dos desafíos, cada uno con su respuesta. O1: las remesas subieron 18.7% (US$25,530 millones) y la economía crece 4.1%: hay más dinero en los hogares, y lo aprovechamos con combos familiares y precios claros. O2: el 48% aún compra en tiendas pequeñas: hay espacio para un súper de barrio moderno, que es justo el formato de Merco. D1: Walmart ya tiene 276 tiendas y compite por precio; nosotros ganamos con trato y frescura. D2: estamos a 2,100 km y hay extorsiones a comercios; lo resolvemos con un almacén local y un socio guatemalteco que conoce el terreno.");

  // 6 · Acción estratégica
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Celeste", sectionTitle: "Propuesta" });
  s.addText("PASO 4 · ACCIÓN ESTRATÉGICA", { placeholder: "kicker" });
  s.addText("PASO A PASO: UN SOCIO LOCAL\nY 2 TIENDAS PILOTO", { placeholder: "title" });
  const phases = [
    { cx: 395, i: "LuHandshake", fill: HEX.mid, head: "PREPARAR", when: "Meses 0–6", what: "Socio + almacén" },
    { cx: 960, i: "LuStore", fill: HEX.blue, head: "PROBAR", when: "Meses 6–24", what: "2 tiendas piloto" },
    { cx: 1525, i: "LuTrendingUp", fill: HEX.navy, head: "CRECER", when: "Desde el mes 24", what: "Quetzaltenango y Escuintla" },
  ];
  box(s, { x: 395, y: 496, w: 1130, h: 8, fill: HEX.mid, shape: pres.shapes.RECTANGLE, name: "Línea de tiempo" });
  for (const [i, ph] of phases.entries()) {
    await iconCircle(s, ph.i, { x: ph.cx - 95, y: 405, d: 190, fill: ph.fill, color: "FFFFFF", alt: ph.head });
    text(s, ph.head, { x: ph.cx - 260, y: 605, w: 520, h: 100, fontFace: HEAD, bold: true, fontSize: 32, color: HEX.blue, align: "center", valign: "middle", name: `Fase ${i + 1}` });
    text(s, ph.when, { x: ph.cx - 260, y: 704, w: 520, h: 34, bold: true, fontSize: 12, color: HEX.muted, align: "center", valign: "middle" });
    text(s, ph.what, { x: ph.cx - 260, y: 740, w: 520, h: 70, bold: true, fontSize: 15, color: HEX.navy, align: "center", valign: "middle" });
  }
  text(s, "¡OFERTAS DE\nAPERTURA!", { x: 1030, y: 362, w: 250, h: 84, fontFace: HEAD, bold: true, fontSize: 14, color: "FFFFFF", align: "center", valign: "middle", fill: { color: HEX.orange }, rotate: 7, lineSpacingMultiple: 0.95, name: "Etiqueta ofertas" });
  sectionHead(s, "METAS DEL PILOTO", { x: 128, y: 838, w: 360, fontSize: 16 });
  const goals = ["Más ventas por m²", "Clientes que regresan", "Rotación < 41%"];
  for (const [i, g] of goals.entries()) {
    text(s, g, { x: 500 + i * 435, y: 832, w: 420, h: 62, bold: true, fontSize: 13, color: HEX.navy, align: "center", valign: "middle", fill: { color: "FFFFFF" }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.3 });
  }
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });
  s.addNotes("La acción tiene tres fases, y no pasamos a la siguiente sin cumplir la anterior. Fase 1, preparar (meses 0 a 6): firmar con un socio guatemalteco, abrir el almacén en la capital y elegir proveedores locales. Fase 2, probar (meses 6 a 24): abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura y medir tres metas: ventas por metro cuadrado, clientes que regresan y rotación menor al 41%, como en México. Fase 3, crecer (desde el mes 24): llevar el modelo a Quetzaltenango y Escuintla, solo si el piloto funciona. Esta acción aprovecha O1 y O2, resuelve D2 con el socio y el almacén, y compite con Walmart con trato, no con precio. Y arriesga poco: 2 tiendas son el 5% del tamaño de Merco.");

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Propuesta" });
  s.addText("RECOMENDACIÓN FINAL", { placeholder: "kicker" });
  text(s, "SÍ A GUATEMALA.", { x: 126, y: 178, w: 1210, h: 150, fontFace: HEAD, bold: true, fontSize: 56, color: "FFFFFF", valign: "middle", name: "Título" });
  text(s, "Paso a paso, con un socio local.", { x: 130, y: 335, w: 1000, h: 60, fontSize: 24, color: HEX.pale, name: "Subtítulo" });
  stamp(s, "APROBADO CON\nCONDICIONES", { x: 1350, y: 190, w: 440, h: 150, fontSize: 19, color: "FFFFFF", rotate: -6, name: "Sello aprobado" });
  const why = [
    { i: "LuTrendingUp", v: "+4.1%", l: "un mercado que crece" },
    { i: "LuTrophy", v: "41%", l: "rotación: su modelo gana" },
    { i: "LuShieldCheck", v: "5%", l: "de su tamaño: riesgo pequeño" },
  ];
  for (const [i, w] of why.entries()) {
    const x = cols[i];
    box(s, { x, y: 450, w: colW, h: 270, fill: HEX.card, name: w.l });
    await iconCircle(s, w.i, { x: x + 34, y: 500, d: 84, fill: HEX.blue, color: "FFFFFF" });
    text(s, w.v, { x: x + 140, y: 462, w: colW - 160, h: 160, fontFace: HEAD, bold: true, fontSize: 52, color: HEX.blue, valign: "middle", name: `${w.l} cifra` });
    text(s, w.l, { x: x + 34, y: 630, w: colW - 60, h: 70, bold: true, fontSize: 14, color: HEX.navy, valign: "middle", name: w.l });
  }
  const steps = ["2027 · Socio y almacén", "2028 · 2 tiendas piloto", "2029 · Crecer"];
  for (const [i, st] of steps.entries()) {
    text(s, st, { x: cols[i], y: 780, w: colW, h: 70, fontFace: HEAD, bold: true, fontSize: 19, color: "FFFFFF", align: "center", valign: "middle", line: { color: "FFFFFF", width: 2 }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.3 });
    if (i < 2) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x: P(cols[i] + colW + 2), y: P(803), w: P(25), h: P(24), fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: "Flecha" });
    }
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });
  s.addNotes("Nuestra recomendación: sí a Guatemala, paso a paso y con un socio local. Comparamos tres caminos: quedarse en México no arriesga pero deja pasar un mercado que crece; entrar en grande es demasiado riesgo frente a Walmart y la distancia; entrar paso a paso es el equilibrio. Tres razones: el mercado crece (economía +4.1% y remesas +18.7%); el modelo de Merco gana (rotación de 98% a 41% y $118,000 por m²); y el riesgo es pequeño, porque 2 tiendas son el 5% de su tamaño. La aprobamos con condiciones: solo crecemos si el piloto cumple sus metas. Siguientes pasos: socio y almacén en 2027, 2 tiendas en 2028 y crecer en 2029. Gracias.");
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
    ref("Chad Davis, CC BY-SA 2.0 (Antigua); Juan Francisco, CC BY-SA 2.0 (Ciudad de Guatemala); Francisco Anzola, CC BY 2.0 (mercado); Toby Argüelles, CC BY-SA 2.0 (Chichicastenango). Mapa: Natural Earth."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 984, y: 255, w: 805, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
