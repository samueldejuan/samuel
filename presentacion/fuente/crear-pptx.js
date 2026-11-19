// Genera Pitch_Merco_Guatemala.pptx (editable) con estilo póster de folleto de súper.
// Uso: node crear-pptx.js <carpeta-de-recursos> <salida.pptx> <apply_theme.js>
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");

const [ASSETS, OUT, APPLY_THEME] = process.argv.slice(2);
const { applyTheme } = require(APPLY_THEME);
const A = (f) => path.join(ASSETS, f);
const G = JSON.parse(fs.readFileSync(A("graficos.json")));

const THEME = {
  name: "Merco Guatemala · Póster",
  headFontFace: "Impact",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1D1A33", lt1: "FFFFFF", dk2: "0E6E73", lt2: "FFF4E0",
    accent1: "0E6E73", accent2: "C2185B", accent3: "FFC845", accent4: "4B2C82",
    accent5: "2A9D6F", accent6: "F28C28", hlink: "0E6E73", folHlink: "1D1A33",
  },
};
// Paleta inspirada en los textiles de Guatemala
const HEX = {
  blue: "0E6E73", navy: "1D1A33", cream: "FFF4E0", paper: "FFFDF7", yellow: "FFC845",
  red: "C2185B", brightRed: "FF6B9A", green: "2A9D6F", purple: "4B2C82", muted: "5E5A6E",
};

// Medidas en píxeles de una lámina de 1920×1080
const P = (v) => v / 192;
const HEAD = "+mj-lt";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Pitch Merco: Guatemala";
pres.author = "Samuel de Juan Ortiz";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

async function icon(name, hex) {
  const Comp = lu[name];
  if (!Comp) throw new Error(`Icono no encontrado: ${name}`);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: 256, color: `#${hex}`, strokeWidth: 2 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Diseños (layouts) ----------
function layout(name, bg, { kicker, title, source }) {
  pres.defineSlideMaster({
    title: name,
    background: { path: A(`fondo-${bg}.jpg`) },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: P(120), y: P(96), w: P(1200), h: P(46), fontSize: 13, bold: true, charSpacing: 3, color: kicker, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: P(116), y: P(146), w: P(1300), h: P(200), fontSize: 40, color: title, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.9 }, text: "" } },
      { placeholder: { options: { name: "source", type: "body", x: P(120), y: P(1000), w: P(1520), h: P(40), fontSize: 9.5, color: source, margin: 0, valign: "middle" }, text: "" } },
    ],
    slideNumber: { x: P(1700), y: P(1000), w: P(100), h: P(40), fontSize: 10, color: source, align: "right", margin: 0 },
  });
}
layout("Turquesa", "turquesa", { kicker: HEX.yellow, title: HEX.cream, source: "CDE8E6" });
layout("Magenta", "magenta", { kicker: HEX.yellow, title: HEX.cream, source: "F6D3E0" });
layout("Morado", "morado", { kicker: HEX.yellow, title: HEX.cream, source: "DDD3EE" });
layout("Crema", "crema", { kicker: HEX.red, title: HEX.blue, source: HEX.muted });

// ---------- Ayudantes ----------
const SHADOW = () => ({ type: "outer", color: "140F26", blur: 0, offset: 3, angle: 45, opacity: 0.55 });
function text(slide, runs, o) {
  slide.addText(runs, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), margin: o.margin ?? 0, isTextBox: true,
    valign: o.valign || "top", align: o.align || "left", fontSize: o.fontSize || 12,
    color: o.color || HEX.navy, bold: o.bold, fontFace: o.fontFace, rotate: o.rotate,
    charSpacing: o.charSpacing, lineSpacingMultiple: o.lineSpacingMultiple, objectName: o.name,
    fill: o.fill, line: o.line, shape: o.shape, rectRadius: o.rectRadius, shadow: o.shadow,
  });
}
function box(slide, o) {
  slide.addShape(o.shape || pres.shapes.ROUNDED_RECTANGLE, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), rectRadius: o.radius ?? 0.1,
    fill: { color: o.fill }, line: o.line ? { color: o.line, width: o.lineW || 2, dashType: o.dash } : { type: "none" },
    rotate: o.rotate, objectName: o.name, shadow: o.shadow,
  });
}
// Título grande en Impact con palabras resaltadas: [["TEXTO ", color], ...] por línea
function poster(slide, lines, o) {
  const runs = [];
  lines.forEach((line, li) => line.forEach(([t, c, size], i) => runs.push({ text: t, options: { color: c || o.color, fontSize: size || o.fontSize, breakLine: i === line.length - 1 && li < lines.length - 1 } })));
  text(slide, runs, { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: HEAD, fontSize: o.fontSize, lineSpacingMultiple: 0.9, shadow: SHADOW(), name: o.name || "Título" });
}
// Etiqueta de precio (imagen) con texto editable encima
function priceTag(slide, x, y, w, rot, big, small, o = {}) {
  const k = w / G.etiqueta.W, h = G.etiqueta.H * k, ta = G.etiqueta.textArea;
  slide.addImage({ path: A("etiqueta.png"), x: P(x), y: P(y), w: P(w), h: P(h), rotate: rot, altText: "Etiqueta de precio" });
  text(slide, [
    { text: big, options: { fontFace: HEAD, fontSize: o.bigSize || 40, color: HEX.red, breakLine: true } },
    { text: small, options: { bold: true, fontSize: o.smallSize || 12, color: HEX.navy } },
  ], { x: x + ta.x * k, y: y + ta.y * k, w: ta.w * k, h: ta.h * k, valign: "middle", align: "center", rotate: rot, name: `Etiqueta ${big}` });
}
function stamp(slide, label, o) {
  text(slide, label, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: HEAD, fontSize: o.fontSize, color: o.color, align: "center", valign: "middle",
    rotate: o.rotate, line: { color: o.color, width: 4, dashType: "dash" }, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1,
    lineSpacingMultiple: 0.95, name: o.name, charSpacing: 1,
  });
}
function chip(slide, label, o) {
  text(slide, label, {
    x: o.x, y: o.y, w: o.w, h: o.h || 56, fontFace: o.head ? HEAD : undefined, bold: !o.head, fontSize: o.fontSize || 14,
    color: o.color || HEX.navy, align: "center", valign: "middle", fill: { color: o.fill }, shape: pres.shapes.ROUNDED_RECTANGLE,
    rectRadius: o.radius ?? 0.25, rotate: o.rotate, name: o.name || label, shadow: o.shadow ? SHADOW() : undefined, charSpacing: o.charSpacing,
  });
}
async function sticker(slide, name, o) {
  box(slide, { x: o.x, y: o.y, w: o.d, h: o.d, fill: o.fill, shape: pres.shapes.OVAL, line: o.ring, lineW: o.ringW || 5, shadow: SHADOW(), name: `Sticker ${name}` });
  const pad = o.d * 0.25;
  slide.addImage({ data: await icon(name, o.color), x: P(o.x + pad), y: P(o.y + pad), w: P(o.d - 2 * pad), h: P(o.d - 2 * pad), altText: o.alt || name });
}

(async () => {
  // 1 · Portada
  pres.addSection({ title: "Inicio" });
  let s = pres.addSlide({ masterName: "Turquesa", sectionTitle: "Inicio" });
  s.addImage({ path: A("polaroid-antigua.png"), x: P(1180), y: P(96), w: P(580), h: P(580 * 932 / 750), altText: "Polaroid: Arco de Santa Catalina, Antigua Guatemala" });
  s.addImage({ path: A("sticker-ruta.png"), x: P(980), y: P(560), w: P(380), h: P(380), altText: "Pegatina con la ruta de Monterrey a Guatemala" });
  stamp(s, "MONTERREY → GUATEMALA", { x: 1350, y: 860, w: 480, h: 84, fontSize: 14, color: HEX.yellow, rotate: -6, name: "Sello ruta" });
  s.addText("PROPUESTA DE EXPANSIÓN · MERCO", { placeholder: "kicker" });
  poster(s, [
    [["PRÓXIMA PARADA:"]],
    [["GUATEMALA", HEX.yellow, 78]],
  ], { x: 116, y: 150, w: 1000, h: 380, fontSize: 48, color: HEX.cream });
  text(s, "El súper de barrio del norte de México cruza la frontera.", { x: 122, y: 520, w: 860, h: 90, bold: true, fontSize: 20, color: HEX.cream, name: "Subtítulo" });
  priceTag(s, 70, 650, 640, -5, "1 DE CADA 2", "compras en Guatemala se hacen en la tienda de barrio", { bigSize: 31, smallSize: 11 });
  s.addText("Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504) · Fuente: Kantar · Mapa: Natural Earth", { placeholder: "source" });
  s.addNotes("En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio (Kantar). Merco lleva 80 años siendo justo eso en el norte de México: un súper cercano para familias que cuidan su dinero. Hoy les proponemos llevar ese modelo de Monterrey a Ciudad de Guatemala. Vamos a ver si Merco está lista, por qué Guatemala, cómo entraríamos y qué recomendamos.");

  // 2 · Diagnóstico
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Crema", sectionTitle: "Análisis" });
  s.addText("PASO 1 · LA MALETA DE MERCO", { placeholder: "kicker" });
  poster(s, [[["MAPA DE FORTALEZAS Y DEBILIDADES"]]], { x: 116, y: 150, w: 1270, h: 120, fontSize: 33, color: HEX.blue });
  stamp(s, "PREPARACIÓN MEDIA:\nLISTA PARA EMPEZAR", { x: 1390, y: 96, w: 460, h: 130, fontSize: 17, color: HEX.red, rotate: -7, name: "Sello preparación" });
  const cols2 = [
    { head: "FORTALEZAS", color: HEX.green, x: 120, items: [
      { i: "LuUsers", v: "41%", l: "de rotación (antes 98%)" },
      { i: "LuStore", v: "$118K", l: "por m², más que Walmart" },
      { i: "LuHeartHandshake", v: "80", l: "años de súper de barrio" },
    ] },
    { head: "DEBILIDADES", color: HEX.red, x: 980, items: [
      { i: "LuMapPin", v: "39", l: "tiendas en solo 2 estados" },
      { i: "LuRoute", v: "0", l: "países fuera de México" },
      { i: "LuSmartphone", v: "0", l: "programas de lealtad" },
    ] },
  ];
  for (const c of cols2) {
    chip(s, c.head, { x: c.x, y: 318, w: 490, h: 60, head: true, fontSize: 24, color: HEX.cream, fill: c.color, rotate: -2, radius: 0.15, shadow: true, charSpacing: 2 });
    for (const [i, it] of c.items.entries()) {
      const y = 404 + i * 160;
      box(s, { x: c.x, y, w: 820, h: 140, fill: HEX.paper, line: c.color, lineW: 2.5, dash: "dash", radius: 0.12, name: it.l });
      await sticker(s, it.i, { x: c.x + 26, y: y + 25, d: 90, fill: c.color, color: "FFFFFF" });
      text(s, it.v, { x: c.x + 140, y, w: 320, h: 140, fontFace: HEAD, fontSize: 48, color: c.color, valign: "middle", name: `${it.l} cifra` });
      text(s, it.l, { x: c.x + 470, y, w: 330, h: 140, bold: true, fontSize: 16, color: HEX.navy, valign: "middle" });
    }
  }
  chip(s, "→  ENTRAR EN PEQUEÑO Y CON UN SOCIO LOCAL", { x: 385, y: 896, w: 1150, h: 76, head: true, fontSize: 20, color: HEX.cream, fill: HEX.blue, radius: 0.2, shadow: true, charSpacing: 1 });
  s.addText("Fuentes: diagnóstico interno del Equipo 2; caso Merco (Carreón et al., 2026); Great Place to Work (2025)", { placeholder: "source" });
  s.addNotes("Lo que Merco hace bien: su gente se queda —la rotación bajó de 98% a 41% gracias al ingreso digno y la Universidad Merco—; sus tiendas chicas venden $118,000 por metro cuadrado, más que Soriana y Walmart; y lleva 80 años como súper de barrio para familias que cuidan su dinero. Lo que le falta: es chica, con 39 tiendas en 2 estados frente a las 3,316 de Walmart; nunca ha operado fuera de México; y no tiene programa de lealtad. Por eso la preparación es media: está lista para empezar, pero en pequeño y con un socio local.");

  // 3 · El mercado
  s = pres.addSlide({ masterName: "Magenta", sectionTitle: "Análisis" });
  s.addText("PASO 2 · UN PAÍS CON HAMBRE DE SÚPER", { placeholder: "kicker" });
  poster(s, [[["MERCADO SELECCIONADO"]], [["Y DATOS CLAVE", HEX.yellow]]], { x: 116, y: 136, w: 1260, h: 300, fontSize: 46, color: HEX.cream });
  chip(s, "MERCADO ELEGIDO", { x: 1390, y: 166, w: 440, h: 64, head: true, fontSize: 18, color: HEX.navy, fill: HEX.yellow, rotate: 4, radius: 0.15, shadow: true, charSpacing: 1 });
  priceTag(s, 60, 440, 620, -5, "18.7 M", "personas: la mayor población de Centroamérica", { bigSize: 40, smallSize: 11 });
  priceTag(s, 660, 465, 620, 3, "+4.1%", "crecerá su economía, más que América Latina", { bigSize: 40, smallSize: 11 });
  priceTag(s, 1260, 440, 620, -4, "+18.7%", "subieron las remesas en 2025", { bigSize: 40, smallSize: 11 });
  chip(s, "EL RETO", { x: 120, y: 800, w: 260, h: 66, head: true, fontSize: 24, color: HEX.navy, fill: HEX.yellow, rotate: -3, radius: 0.15, shadow: true, charSpacing: 2 });
  const retos = [
    { i: "LuStore", t: "Walmart: 276 tiendas", x: 420 },
    { i: "LuRoute", t: "A 2,100 km", x: 880 },
    { i: "LuShieldAlert", t: "Extorsiones", x: 1340 },
  ];
  for (const r of retos) {
    box(s, { x: r.x, y: 800, w: 440, h: 66, fill: HEX.cream, radius: 0.3, shadow: SHADOW(), name: r.t });
    s.addImage({ data: await icon(r.i, HEX.red), x: P(r.x + 22), y: P(814), w: P(38), h: P(38), altText: r.t });
    text(s, r.t, { x: r.x + 76, y: 800, w: 350, h: 66, bold: true, fontSize: 13, color: HEX.navy, valign: "middle" });
  }
  s.addText("Fuentes: Banco Mundial (2025); Banguat (2025, 2026); Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Elegimos Guatemala por tres datos. Es el país más poblado de Centroamérica, con 18.7 millones de personas (Banco Mundial). Su economía crecerá 4.1%, más que América Latina (Banguat). Y las remesas subieron 18.7% en 2025, hasta US$25,530 millones: las familias tienen más dinero para el súper. Además el cliente cuida su dinero (PIB por persona de US$6,506), igual que el de Merco. El reto: Walmart ya tiene 276 tiendas, hay extorsiones a comercios y estamos a más de 2,100 km de Nuevo León.");

  // 4 · Estrategia semiglobal
  s = pres.addSlide({ masterName: "Morado", sectionTitle: "Análisis" });
  s.addImage({ path: A("polaroid-mercado.png"), x: P(1330), y: P(170), w: P(530), h: P(530 * 922 / 736), altText: "Polaroid: mercado de Guatemala" });
  s.addText("PASO 3 · MISMO CORAZÓN, NUEVO SABOR", { placeholder: "kicker" });
  poster(s, [[["ESTRATEGIA SEMIGLOBAL"]], [["Y 2 ÁREAS DE ADAPTACIÓN", HEX.yellow]]], { x: 116, y: 136, w: 1200, h: 300, fontSize: 43, color: HEX.cream });
  chip(s, "SE QUEDA IGUAL", { x: 120, y: 508, w: 460, h: 58, head: true, fontSize: 22, color: HEX.navy, fill: HEX.cream, rotate: -2, radius: 0.15, shadow: true, charSpacing: 1 });
  const same = [
    { i: "LuUsers", t: "SU GENTE" },
    { i: "LuShieldCheck", t: "SUS GARANTÍAS" },
    { i: "LuStore", t: "SÚPER DE BARRIO" },
  ];
  for (const [i, it] of same.entries()) {
    const cx = 230 + i * 360;
    await sticker(s, it.i, { x: cx - 62, y: 578, d: 124, fill: HEX.cream, color: HEX.purple, ring: HEX.yellow, ringW: 6 });
    text(s, it.t, { x: cx - 170, y: 708, w: 340, h: 44, fontFace: HEAD, fontSize: 18, color: HEX.cream, align: "center", valign: "middle", shadow: SHADOW() });
  }
  chip(s, "SE ADAPTA", { x: 120, y: 770, w: 330, h: 58, head: true, fontSize: 22, color: HEX.navy, fill: HEX.yellow, rotate: -2, radius: 0.15, shadow: true, charSpacing: 1 });
  const adapt = [
    { i: "LuWheat", t: "PRODUCTO CHAPÍN", b: "frijol, maíz y café local" },
    { i: "LuWarehouse", t: "ALMACÉN LOCAL", b: "en Ciudad de Guatemala" },
  ];
  for (const [i, it] of adapt.entries()) {
    const x = 120 + i * 580;
    box(s, { x, y: 846, w: 555, h: 130, fill: HEX.yellow, radius: 0.15, shadow: SHADOW(), name: it.t });
    await sticker(s, it.i, { x: x + 22, y: 862, d: 98, fill: HEX.paper, color: HEX.purple });
    text(s, [
      { text: it.t, options: { fontFace: HEAD, fontSize: 20, color: HEX.navy, breakLine: true } },
      { text: it.b, options: { bold: true, fontSize: 12, color: HEX.navy } },
    ], { x: x + 140, y: 846, w: 400, h: 130, valign: "middle" });
  }
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO)", { placeholder: "source" });
  s.addNotes("Nuestra estrategia es semiglobal: lo que hace única a Merco se queda igual y solo adaptamos dos áreas. Se queda igual su gente (ingreso digno y Universidad Merco), sus garantías (frescura garantizada y 30 días para devolver) y su formato de súper de barrio. Adaptamos, primero, el producto: en Guatemala la gente busca frijol negro, maíz y café local, así que el surtido mezcla lo chapín con lo mejor de México. Y segundo, la distribución: hoy todo sale de 3 centros en Nuevo León, pero a más de 2,100 km lo fresco no llega bien, así que abrimos un almacén en Ciudad de Guatemala con proveedores locales. ¿Por qué semiglobal? Porque su cultura sí viaja, el gusto local es distinto y adaptar solo 2 áreas cuesta menos.");

  // 5 · Oportunidades y desafíos
  s = pres.addSlide({ masterName: "Turquesa", sectionTitle: "Análisis" });
  s.addImage({ path: A("polaroid-chichi.png"), x: P(1400), y: P(290), w: P(470), h: P(470 * 922 / 736), altText: "Polaroid: mercado de Chichicastenango" });
  s.addText("PASO 4 · LA OPORTUNIDAD ES REAL", { placeholder: "kicker" });
  poster(s, [[["2 OPORTUNIDADES"]], [["Y 2 DESAFÍOS", HEX.yellow]]], { x: 116, y: 136, w: 1280, h: 280, fontSize: 46, color: HEX.cream });
  const od = [
    { tag: "OPORTUNIDAD", tagFill: HEX.green, v: "+18.7%", l: "más remesas", r: "combos familiares", x: 120, y: 452 },
    { tag: "OPORTUNIDAD", tagFill: HEX.green, v: "48%", l: "compra en tiendas pequeñas", r: "súper de barrio", x: 760, y: 452 },
    { tag: "DESAFÍO", tagFill: HEX.navy, v: "276", l: "tiendas de Walmart", r: "trato y frescura", x: 120, y: 728 },
    { tag: "DESAFÍO", tagFill: HEX.navy, v: "2,100 km", l: "de distancia + extorsiones", r: "socio + almacén", x: 760, y: 728 },
  ];
  for (const it of od) {
    box(s, { x: it.x, y: it.y, w: 610, h: 240, fill: HEX.paper, line: HEX.red, lineW: 2.5, dash: "dash", radius: 0.1, shadow: SHADOW(), name: `${it.tag} ${it.v}` });
    chip(s, it.tag, { x: it.x + 24, y: it.y - 22, w: 290, h: 46, head: true, fontSize: 15, color: HEX.cream, fill: it.tagFill, radius: 0.2, charSpacing: 1 });
    text(s, it.v, { x: it.x + 30, y: it.y + 24, w: 560, h: 116, fontFace: HEAD, fontSize: 50, color: HEX.red, valign: "middle" });
    text(s, it.l, { x: it.x + 30, y: it.y + 138, w: 560, h: 36, bold: true, fontSize: 13, color: HEX.navy, valign: "middle" });
    chip(s, `→ ${it.r}`, { x: it.x + 30, y: it.y + 180, w: 380, h: 44, fontSize: 12, color: HEX.cream, fill: HEX.blue });
  }
  s.addText("Fuentes: Banguat (2026); Kantar (2022); Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Dos oportunidades y dos desafíos, cada uno con su respuesta. O1: las remesas subieron 18.7% (US$25,530 millones) y la economía crece 4.1%: hay más dinero en los hogares, y lo aprovechamos con combos familiares y precios claros. O2: el 48% aún compra en tiendas pequeñas: hay espacio para un súper de barrio moderno, que es justo el formato de Merco. D1: Walmart ya tiene 276 tiendas y compite por precio; nosotros ganamos con trato y frescura. D2: estamos a 2,100 km y hay extorsiones a comercios; lo resolvemos con un almacén local y un socio guatemalteco que conoce el terreno.");

  // 6 · Acción estratégica
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Crema", sectionTitle: "Propuesta" });
  s.addText("PASO 5 · ASÍ ABRIMOS LA PRIMERA TIENDA", { placeholder: "kicker" });
  poster(s, [[["ACCIÓN ESTRATÉGICA"]], [["PASO A PASO", HEX.red]]], { x: 116, y: 152, w: 1320, h: 290, fontSize: 50, color: HEX.blue });
  s.addShape(pres.shapes.LINE, { x: P(380), y: P(540), w: P(1160), h: 0, line: { color: HEX.blue, width: 5, dashType: "dash" }, objectName: "Ruta" });
  const phases = [
    { cx: 380, i: "LuHandshake", fill: HEX.yellow, color: HEX.navy, head: "PREPARAR", when: "MESES 0–6", what: "Socio + almacén" },
    { cx: 960, i: "LuStore", fill: HEX.red, color: "FFFFFF", head: "PROBAR", when: "MESES 6–24", what: "2 tiendas piloto" },
    { cx: 1540, i: "LuTrendingUp", fill: HEX.green, color: "FFFFFF", head: "CRECER", when: "DESDE EL MES 24", what: "Quetzaltenango y Escuintla" },
  ];
  for (const [i, ph] of phases.entries()) {
    await sticker(s, ph.i, { x: ph.cx - 100, y: 440, d: 200, fill: ph.fill, color: ph.color, ring: "FFFFFF", ringW: 8 });
    text(s, `${i + 1}`, { x: ph.cx + 56, y: 428, w: 64, h: 64, fontFace: HEAD, fontSize: 26, color: HEX.cream, align: "center", valign: "middle", fill: { color: HEX.blue }, shape: pres.shapes.OVAL });
    text(s, ph.head, { x: ph.cx - 260, y: 650, w: 520, h: 84, fontFace: HEAD, fontSize: 40, color: HEX.blue, align: "center", valign: "middle" });
    text(s, ph.when, { x: ph.cx - 260, y: 734, w: 520, h: 34, bold: true, fontSize: 12, color: HEX.red, align: "center", valign: "middle", charSpacing: 2 });
    text(s, ph.what, { x: ph.cx - 260, y: 770, w: 520, h: 84, bold: true, fontSize: 15, color: HEX.navy, align: "center", valign: "middle" });
  }
  text(s, "¡OFERTAS DE\nAPERTURA!", { x: 1200, y: 300, w: 380, h: 230, fontFace: HEAD, fontSize: 18, color: HEX.red, align: "center", valign: "middle", fill: { color: HEX.yellow }, shape: "irregularSeal1", rotate: 8, lineSpacingMultiple: 0.9, name: "Estrella ofertas" });
  chip(s, "METAS DEL PILOTO", { x: 120, y: 880, w: 430, h: 64, head: true, fontSize: 18, color: HEX.cream, fill: HEX.blue, rotate: -2, radius: 0.15, shadow: true, charSpacing: 1 });
  const goals = ["Más ventas por m²", "Clientes que vuelven", "Rotación < 41%"];
  for (const [i, g] of goals.entries()) {
    box(s, { x: 575 + i * 415, y: 878, w: 405, h: 68, fill: HEX.paper, line: HEX.blue, lineW: 2.5, dash: "dash", radius: 0.3, name: g });
    text(s, g, { x: 575 + i * 415, y: 878, w: 405, h: 68, bold: true, fontSize: 12.5, color: HEX.navy, align: "center", valign: "middle" });
  }
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });
  s.addNotes("La acción tiene tres fases, y no pasamos a la siguiente sin cumplir la anterior. Fase 1, preparar (meses 0 a 6): firmar con un socio guatemalteco, abrir el almacén en la capital y elegir proveedores locales. Fase 2, probar (meses 6 a 24): abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura y medir tres metas: ventas por metro cuadrado, clientes que regresan y rotación menor al 41%, como en México. Fase 3, crecer (desde el mes 24): llevar el modelo a Quetzaltenango y Escuintla, solo si el piloto funciona. Esta acción aprovecha O1 y O2, resuelve D2 con el socio y el almacén, y compite con Walmart con trato, no con precio. Y arriesga poco: 2 tiendas son el 5% del tamaño de Merco.");

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Magenta", sectionTitle: "Propuesta" });
  s.addText("PASO 6 · ¡VAMOS A GUATEMALA!", { placeholder: "kicker" });
  poster(s, [[["RECOMENDACIÓN FINAL"]], [["¡SÍ A GUATEMALA!", HEX.yellow]]], { x: 116, y: 150, w: 1270, h: 300, fontSize: 50, color: HEX.cream });
  stamp(s, "APROBADO CON\nCONDICIONES", { x: 1400, y: 140, w: 440, h: 160, fontSize: 22, color: HEX.yellow, rotate: -8, name: "Sello aprobado" });
  priceTag(s, 60, 470, 620, -4, "+4.1%", "un mercado que crece", { bigSize: 48, smallSize: 13 });
  priceTag(s, 660, 495, 620, 3, "41%", "de rotación: su modelo gana", { bigSize: 48, smallSize: 13 });
  priceTag(s, 1260, 470, 620, -3, "5%", "de su tamaño: riesgo pequeño", { bigSize: 48, smallSize: 13 });
  const steps = ["2027 · SOCIO Y ALMACÉN", "2028 · 2 TIENDAS PILOTO", "2029 · CRECER"];
  for (const [i, st] of steps.entries()) {
    chip(s, st, { x: 120 + i * 570, y: 860, w: 520, h: 72, head: true, fontSize: 17, color: i === 2 ? HEX.navy : HEX.cream, fill: i === 2 ? HEX.yellow : HEX.purple, radius: 0.2, shadow: true });
    if (i < 2) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x: P(120 + i * 570 + 528), y: P(880), w: P(36), h: P(32), fill: { color: HEX.cream }, line: { type: "none" }, objectName: "Flecha" });
    }
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });
  s.addNotes("Nuestra recomendación: sí a Guatemala, paso a paso y con un socio local. Comparamos tres caminos: quedarse en México no arriesga pero deja pasar un mercado que crece; entrar en grande es demasiado riesgo frente a Walmart y la distancia; entrar paso a paso es el equilibrio. Tres razones: el mercado crece (economía +4.1% y remesas +18.7%); el modelo de Merco gana (rotación de 98% a 41% y $118,000 por m²); y el riesgo es pequeño, porque 2 tiendas son el 5% de su tamaño. La aprobamos con condiciones: solo crecemos si el piloto cumple sus metas. Siguientes pasos: socio y almacén en 2027, 2 tiendas en 2028 y crecer en 2029. Gracias.");

  // 8 · Fuentes
  pres.addSection({ title: "Anexo" });
  s = pres.addSlide({ masterName: "Crema", sectionTitle: "Anexo" });
  s.addText("NO SE EXPONE", { placeholder: "kicker" });
  s.addText("FUENTES", { placeholder: "title", h: P(90) });
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
  text(s, left.map((t, i) => ref(t, i === left.length - 1)), { x: 120, y: 260, w: 810, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 1" });
  text(s, [
    ref("Carreón Flores, H. et al. (2026). Supermercados Merco: Crecimiento centrado en las personas (RC32-11-270). Tec de Monterrey."),
    ref("Great Place to Work México (2025). Merco."),
    ref("Emisoras Unidas (2025). Súper del Barrio celebra su tienda 200 en Guatemala."),
    ref("Dess, G. G. et al. (2021). Strategic management (10.ª ed.). McGraw-Hill."),
    ref("Equipo 2 (2026). Diagnóstico externo, diagnóstico interno y matriz FODA de Supermercados Merco."),
    ref("[Passport, Statista y EMIS: agregar]"),
    { text: "Fotos (Wikimedia Commons): ", options: { bold: true } },
    ref("Francisco Anzola, CC BY 2.0 (mercado); Toby Argüelles, CC BY-SA 2.0 (Chichicastenango). Mapa: Natural Earth. Iconos: Lucide (ISC)."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 990, y: 260, w: 810, h: 700, fontSize: 10.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
