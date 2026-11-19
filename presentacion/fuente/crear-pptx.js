// Genera Pitch_Merco_Guatemala.pptx: estilo editorial (serio y atractivo) y compatible:
// sin giros, sombras ni formas especiales, para que se vea igual en cualquier programa.
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

const HEX = {
  navy: "16223D", navy2: "24345A", gold: "C9A04A", goldTint: "F3E9D2", ivory: "F7F4EE", white: "FFFFFF",
  teal: "1E6B66", tealTint: "DDECEA", terra: "B5523B", terraTint: "F5E3DD", gray: "5E6472", line: "E3DED3",
  soft: "C8CFDC", map: "E8EDF3",
};
const SERIF = "Georgia";
const THEME = {
  name: "Merco Guatemala",
  headFontFace: SERIF,
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.navy, lt1: HEX.white, dk2: HEX.navy2, lt2: HEX.ivory,
    accent1: HEX.navy, accent2: HEX.gold, accent3: HEX.teal, accent4: HEX.terra,
    accent5: HEX.gray, accent6: HEX.soft, hlink: HEX.teal, folHlink: HEX.navy,
  },
};

// Medidas en píxeles de una lámina de 1920×1080
const P = (v) => v / 192;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Pitch Merco: Guatemala";
pres.author = "Samuel de Juan Ortiz";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

async function icon(name, hex) {
  const Comp = lu[name];
  if (!Comp) throw new Error(`Icono no encontrado: ${name}`);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: 256, color: `#${hex}`, strokeWidth: 1.8 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Diseños (layouts) ----------
function layout(name, bg, { kicker, title, source, x = 120, w = 1680, number = true }) {
  const def = {
    title: name,
    background: { color: bg },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: P(x), y: P(84), w: P(w), h: P(36), fontSize: 11, bold: true, charSpacing: 1, color: kicker, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: P(x), y: P(124), w: P(w), h: P(150), fontFace: SERIF, fontSize: 28, color: title, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.95 }, text: "" } },
      { placeholder: { options: { name: "source", type: "body", x: P(x), y: P(1010), w: P(w - 120), h: P(36), fontSize: 9, color: source, margin: 0, valign: "middle" }, text: "" } },
    ],
  };
  if (number) def.slideNumber = { x: P(1720), y: P(1010), w: P(80), h: P(36), fontSize: 9, color: source, align: "right", margin: 0 };
  pres.defineSlideMaster(def);
}
layout("Marfil", HEX.ivory, { kicker: HEX.gold, title: HEX.navy, source: HEX.gray });
layout("Marfil con foto", HEX.ivory, { kicker: HEX.gold, title: HEX.navy, source: HEX.gray, x: 740, w: 1060 });
layout("Azul", HEX.navy, { kicker: HEX.gold, title: HEX.white, source: HEX.soft, w: 1100 });

// ---------- Ayudantes ----------
function text(slide, runs, o) {
  slide.addText(runs, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), margin: o.margin ?? 0, isTextBox: true,
    valign: o.valign || "top", align: o.align || "left", fontSize: o.fontSize || 13,
    fontFace: o.serif ? SERIF : undefined, charSpacing: o.charSpacing,
    color: o.color || HEX.navy, bold: o.bold, lineSpacingMultiple: o.lineSpacingMultiple, objectName: o.name,
    fill: o.fill ? { color: o.fill } : undefined, shape: o.fill ? pres.shapes.RECTANGLE : undefined, paraSpaceAfter: o.paraSpaceAfter,
  });
}
function rect(slide, o) {
  slide.addShape(o.oval ? pres.shapes.OVAL : pres.shapes.RECTANGLE, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h),
    fill: { color: o.fill }, line: o.line ? { color: o.line, width: o.lineW || 1 } : { type: "none" }, objectName: o.name,
  });
}
function hline(slide, x, y, w, color, width = 1.5) {
  slide.addShape(pres.shapes.LINE, { x: P(x), y: P(y), w: P(w), h: 0, line: { color, width }, objectName: "Línea" });
}
async function iconCircle(slide, name, o) {
  rect(slide, { x: o.x, y: o.y, w: o.d, h: o.d, fill: o.fill, oval: true, name: `Icono ${name}` });
  const pad = o.d * 0.27;
  slide.addImage({ data: await icon(name, o.color), x: P(o.x + pad), y: P(o.y + pad), w: P(o.d - 2 * pad), h: P(o.d - 2 * pad), altText: o.alt || name });
}
const kicker = (slide, t, x, y, w, color = HEX.gold) => text(slide, t.toUpperCase(), { x, y, w, h: 32, fontSize: 10.5, bold: true, charSpacing: 1, color, valign: "middle" });

(async () => {
  // 1 · Portada
  pres.addSection({ title: "Inicio" });
  let s = pres.addSlide({ masterName: "Azul", sectionTitle: "Inicio" });
  s.addImage({ path: A("foto-portada.jpg"), x: P(820), y: 0, w: P(1100), h: P(1080), altText: "Arco de Santa Catalina, Antigua Guatemala" });
  s.addImage({ path: A("degradado-azul.png"), x: P(818), y: 0, w: P(420), h: P(1080), altText: "Degradado decorativo" });
  text(s, "PROPUESTA DE EXPANSIÓN · MERCO", { x: 120, y: 320, w: 880, h: 32, fontSize: 11, bold: true, charSpacing: 1, color: HEX.gold, valign: "middle" });
  text(s, [
    { text: "Próxima parada:", options: { breakLine: true } },
    { text: "Guatemala", options: { color: HEX.gold, italic: true } },
  ], { x: 120, y: 366, w: 920, h: 250, serif: true, fontSize: 42, color: HEX.white, lineSpacingMultiple: 0.95, name: "Título" });
  hline(s, 124, 668, 110, HEX.gold, 2.5);
  text(s, "En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio. Merco lleva 80 años siendo ese súper en el norte de México.", { x: 120, y: 700, w: 760, h: 110, fontSize: 15, color: HEX.soft, lineSpacingMultiple: 1.1 });
  text(s, "Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504)", { x: 120, y: 990, w: 900, h: 36, fontSize: 10, color: HEX.soft, valign: "middle" });
  text(s, "Dato: Kantar (2022) · Foto: Chad Davis, CC BY-SA 2.0", { x: 1300, y: 1036, w: 600, h: 30, fontSize: 9, color: HEX.white, align: "right", valign: "middle" });
  s.addNotes("En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio (Kantar). Merco lleva 80 años siendo justo eso en el norte de México: un súper cercano para familias que cuidan su dinero. Hoy proponemos llevar ese modelo a Ciudad de Guatemala, paso a paso. Vamos a ver si Merco está lista, por qué Guatemala, cómo entraríamos y qué recomendamos.");

  // 2 · Mapa de fortalezas y debilidades
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Marfil", sectionTitle: "Análisis" });
  s.addText("01 · DIAGNÓSTICO", { placeholder: "kicker" });
  s.addText("Mapa de fortalezas y debilidades", { placeholder: "title" });
  kicker(s, "Fortaleza", 120, 262, 400, HEX.teal);
  text(s, "CÓMO AYUDA", { x: 826, y: 262, w: 268, h: 32, fontSize: 10.5, bold: true, charSpacing: 1, color: HEX.gray, align: "center", valign: "middle" });
  kicker(s, "Debilidad que cubre", 1100, 262, 500, HEX.terra);
  const map = [
    { f: ["$118K", "Tiendas que rinden", "Vende más por m² que Soriana y Walmart"], m: "Con tiendas chicas, abrir cuesta poco", d: ["39", "Es pequeña", "Tiendas en solo 2 estados"] },
    { f: ["80", "Experta en cercanía", "Años vendiendo a familias que cuidan su dinero"], m: "El trato cercano vale más que una app", d: ["0", "Atrás en lo digital", "Sin programa de lealtad"] },
    { f: ["41%", "Cuida a su gente", "Empleados que se van al año (antes 98%)"], m: "Forma a su gente y se apoya en un socio", d: ["0", "Nunca ha salido", "Sin experiencia fuera de México"] },
  ];
  for (const [i, r] of map.entries()) {
    const y = 306 + i * 178;
    rect(s, { x: 120, y, w: 700, h: 158, fill: HEX.navy, name: `Fortaleza ${i + 1}` });
    text(s, r.f[0], { x: 150, y, w: 220, h: 158, bold: true, fontSize: 28, color: HEX.gold, valign: "middle" });
    text(s, [{ text: r.f[1], options: { bold: true, fontSize: 15, color: HEX.white, breakLine: true } }, { text: r.f[2], options: { fontSize: 12, color: HEX.soft } }], { x: 380, y, w: 420, h: 158, valign: "middle" });
    text(s, r.m, { x: 836, y: y + 14, w: 248, h: 80, fontSize: 11.5, bold: true, color: HEX.navy, align: "center", valign: "bottom" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: P(880), y: P(y + 104), w: P(160), h: P(22), fill: { color: HEX.gold }, line: { type: "none" }, objectName: `Flecha ${i + 1}` });
    rect(s, { x: 1100, y, w: 700, h: 158, fill: HEX.white, line: HEX.line, name: `Debilidad ${i + 1}` });
    text(s, r.d[0], { x: 1130, y, w: 160, h: 158, bold: true, fontSize: 30, color: HEX.terra, valign: "middle" });
    text(s, [{ text: r.d[1], options: { bold: true, fontSize: 15, breakLine: true } }, { text: r.d[2], options: { fontSize: 12, color: HEX.gray } }], { x: 1300, y, w: 480, h: 158, valign: "middle" });
  }
  rect(s, { x: 120, y: 862, w: 1680, h: 112, fill: HEX.goldTint, name: "Conclusión" });
  text(s, [{ text: "PREPARACIÓN", options: { fontSize: 10, bold: true, charSpacing: 1, color: HEX.gray, breakLine: true } }, { text: "Media", options: { fontFace: SERIF, fontSize: 24, color: HEX.navy } }], { x: 150, y: 862, w: 260, h: 112, valign: "middle" });
  rect(s, { x: 420, y: 884, w: 2, h: 68, fill: HEX.gold, name: "Separador" });
  text(s, "Su forma de trabajar funciona, pero es pequeña y nunca ha salido de México. Por eso debe empezar con 2 tiendas y un socio local.", { x: 450, y: 862, w: 1320, h: 112, fontSize: 14, color: HEX.navy, valign: "middle" });
  s.addText("Fuentes: diagnóstico interno del Equipo 2; caso Merco (Carreón et al., 2026); Great Place to Work (2025)", { placeholder: "source" });
  s.addNotes("Este mapa conecta cada fortaleza con la debilidad que ayuda a cubrir. Sus tiendas chicas venden $118,000 por metro cuadrado, más que Soriana y Walmart: como no necesita tiendas grandes, ser pequeña (39 tiendas en 2 estados) pesa menos. Lleva 80 años como súper de barrio: su trato cercano compensa que no tenga programa de lealtad. Y cuida a su gente: antes se le iba el 98% de los empleados al año y ahora solo el 41%; formando a la gente nueva y con un socio local puede trabajar en un país donde nunca ha estado. Conclusión: preparación media. Está lista para empezar, pero en pequeño y con un socio local.");

  // 3 · Mercado seleccionado y datos clave
  s = pres.addSlide({ masterName: "Marfil", sectionTitle: "Análisis" });
  rect(s, { x: 1180, y: 0, w: 740, h: 1080, fill: HEX.map, name: "Panel mapa" });
  s.addText("02 · EL MERCADO", { placeholder: "kicker" });
  text(s, "Mercado seleccionado\ny datos clave", { x: 120, y: 124, w: 1000, h: 160, serif: true, fontSize: 28, color: HEX.navy, name: "Título" });
  const stats = [
    { v: "18.7 M", l: "personas: el país con más gente de Centroamérica", src: "Banco Mundial, 2025" },
    { v: "+4.1%", l: "crecerá su economía, más que América Latina", src: "Banguat, 2025" },
    { v: "+18.7%", l: "de remesas: dinero de familiares en el extranjero", src: "Banguat, 2026" },
  ];
  for (const [i, st] of stats.entries()) {
    const x = 120 + i * 335;
    hline(s, x, 318, 300, HEX.gold, 2);
    text(s, st.v, { x, y: 330, w: 310, h: 80, bold: true, fontSize: 28, color: HEX.navy, valign: "middle" });
    text(s, [{ text: st.l, options: { breakLine: true } }, { text: st.src, options: { fontSize: 9.5, color: HEX.gray, bold: false } }], { x, y: 414, w: 300, h: 130, fontSize: 11.5, bold: true, color: HEX.navy });
  }
  const lists = [
    { x: 120, head: "Encaja con Merco", color: HEX.teal, tint: HEX.tealTint, rows: [
      ["LuPiggyBank", "Cuida su dinero", "Ingreso medio: US$6,506 por persona"],
      ["LuMapPin", "Entrada por la capital", "Ciudad de Guatemala"],
      ["LuHandshake", "Cerca de México", "México ya es su 3.er proveedor"],
    ] },
    { x: 640, head: "Competencia (amenaza)", color: HEX.terra, tint: HEX.terraTint, rows: [
      ["LuStore", "Walmart", "276 tiendas"],
      ["LuStore", "Súper del Barrio", "200 tiendas de cercanía"],
      ["LuStore", "La Torre", "Cadena local de supermercados"],
    ] },
  ];
  for (const ls of lists) {
    rect(s, { x: ls.x, y: 580, w: 490, h: 400, fill: HEX.white, line: HEX.line, name: ls.head });
    kicker(s, ls.head, ls.x + 30, 598, 440, ls.color);
    for (const [i, [ic, t, b]] of ls.rows.entries()) {
      const y = 642 + i * 108;
      await iconCircle(s, ic, { x: ls.x + 30, y: y + 14, d: 62, fill: ls.tint, color: ls.color });
      text(s, [{ text: t, options: { bold: true, fontSize: 13, breakLine: true } }, { text: b, options: { fontSize: 11, color: HEX.gray } }], { x: ls.x + 112, y, w: 360, h: 90, valign: "middle" });
    }
  }
  const mw = 740, mk = mw / G.mapa.W, mh = G.mapa.H * mk, mx = 1180, my = 250;
  s.addImage({ path: A("mapa-ruta.png"), x: P(mx), y: P(my), w: P(mw), h: P(mh), altText: "Mapa de la ruta de Monterrey a Ciudad de Guatemala" });
  const [ax, ay] = [mx + G.mapa.mty[0] * mk, my + G.mapa.mty[1] * mk];
  const [gx, gy] = [mx + G.mapa.gt[0] * mk, my + G.mapa.gt[1] * mk];
  text(s, "Monterrey", { x: ax - 180, y: ay - 44, w: 170, h: 32, fontSize: 12.5, bold: true, color: HEX.navy, align: "right", valign: "middle" });
  text(s, "Ciudad de Guatemala", { x: gx - 300, y: gy + 16, w: 280, h: 32, fontSize: 12.5, bold: true, color: HEX.navy, align: "right", valign: "middle" });
  text(s, [{ text: "≈ 2,100 km", options: { bold: true, fontSize: 26, color: HEX.navy, breakLine: true } }, { text: "por carretera desde Nuevo León", options: { fontSize: 11.5, color: HEX.gray } }], { x: 1240, y: 820, w: 620, h: 100, align: "center", valign: "middle" });
  s.addText("Fuentes: Banco Mundial (2025); Banguat (2025, 2026); FocusEconomics (2026); MINECO (2025); Walmart (2025); Emisoras Unidas (2025) · Mapa: Natural Earth", { placeholder: "source", w: P(1020) });
  s.addNotes("Elegimos Guatemala y entramos por su capital. Tres datos clave: es el país con más gente de Centroamérica, 18.7 millones de personas; su economía crecerá 4.1%, más que América Latina; y las remesas, el dinero que mandan los familiares desde fuera, subieron 18.7% en 2025, hasta US$25,530 millones. Encaja con Merco porque el cliente cuida su dinero (ingreso medio de US$6,506 por persona) y México ya es su tercer proveedor. La competencia es la principal amenaza: Walmart tiene 276 tiendas, Súper del Barrio 200 y está La Torre. Y estamos a unos 2,100 km de Nuevo León.");

  // 4 · Estrategia semiglobal y 2 áreas de adaptación
  s = pres.addSlide({ masterName: "Marfil con foto", sectionTitle: "Análisis" });
  s.addImage({ path: A("foto-mercado-alta.jpg"), x: 0, y: 0, w: P(640), h: P(1080), altText: "Mercado de Guatemala con frutas y especias" });
  text(s, "Foto: Francisco Anzola, CC BY 2.0", { x: 20, y: 1036, w: 600, h: 30, fontSize: 9, color: HEX.white, valign: "middle" });
  s.addText("03 · LA ESTRATEGIA", { placeholder: "kicker" });
  s.addText("Estrategia semiglobal y 2 áreas de adaptación", { placeholder: "title" });
  text(s, [{ text: "Semiglobal: ", options: { bold: true, color: HEX.gold } }, { text: "lo que hace especial a Merco se queda igual; solo cambiamos 2 cosas para el cliente de Guatemala." }],
    { x: 740, y: 300, w: 1060, h: 84, fontSize: 13.5, color: HEX.white, fill: HEX.navy, valign: "middle", margin: 12 });
  kicker(s, "Se queda igual", 740, 400, 600, HEX.teal);
  const keep = [
    ["LuUsers", "Su gente", "Buen sueldo y formación propia"],
    ["LuShieldCheck", "Sus garantías", "Frescura y 30 días para devolver"],
    ["LuStore", "Súper de barrio", "Tiendas chicas y cercanas"],
  ];
  for (const [i, [ic, t, b]] of keep.entries()) {
    const x = 740 + i * 360;
    await iconCircle(s, ic, { x, y: 466, d: 60, fill: HEX.tealTint, color: HEX.teal });
    text(s, [{ text: t, options: { bold: true, fontSize: 13, breakLine: true } }, { text: b, options: { fontSize: 11, color: HEX.gray } }], { x: x + 74, y: 448, w: 270, h: 96, valign: "middle" });
  }
  kicker(s, "Se adapta: 2 áreas", 740, 562, 600, HEX.gold);
  const adapt = [
    { n: "1", t: "Producto", hoy: "Surtido del norte", gt: "Frijol negro, maíz y café local + lo mejor de México", por: "El cliente quiere lo de siempre" },
    { n: "2", t: "Distribución", hoy: "Todo sale de Nuevo León", gt: "Almacén en la capital y proveedores locales", por: "Lo fresco no aguanta 2,100 km" },
  ];
  for (const [i, a] of adapt.entries()) {
    const x = 740 + i * 545;
    rect(s, { x, y: 602, w: 515, h: 384, fill: HEX.white, line: HEX.line, name: a.t });
    text(s, [{ text: `${a.n} `, options: { color: HEX.gold } }, { text: a.t }], { x: x + 30, y: 614, w: 460, h: 56, serif: true, fontSize: 21, color: HEX.navy, valign: "middle" });
    hline(s, x + 30, 678, 455, HEX.line, 1);
    text(s, [
      { text: "Hoy: ", options: { bold: true, color: HEX.gray } }, { text: a.hoy, options: { color: HEX.gray, breakLine: true } },
      { text: "En Guatemala: ", options: { bold: true, color: HEX.teal } }, { text: a.gt, options: { bold: true, breakLine: true } },
      { text: "Por qué: ", options: { bold: true, color: HEX.gold } }, { text: a.por },
    ], { x: x + 30, y: 692, w: 460, h: 280, fontSize: 11.5, paraSpaceAfter: 7 });
  }
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO, 2025)", { placeholder: "source" });
  s.addNotes("Nuestra estrategia es semiglobal: mantenemos lo que hace especial a Merco y cambiamos solo dos cosas. Se queda igual su gente (buen sueldo y su propia universidad), sus garantías (frescura garantizada y 30 días para devolver) y su formato de súper de barrio. Cambiamos, primero, el producto: en Guatemala la gente busca frijol negro, maíz y café local, así que mezclamos lo chapín con lo mejor de México. Y segundo, la distribución: hoy todo sale de 3 centros en Nuevo León, pero lo fresco no aguanta 2,100 km, así que abrimos un almacén en Ciudad de Guatemala con proveedores locales. ¿Por qué semiglobal? Porque su cultura sí viaja, pero el gusto local es distinto, y cambiar solo 2 cosas cuesta menos.");

  // 5 · 2 oportunidades y 2 desafíos
  s = pres.addSlide({ masterName: "Marfil", sectionTitle: "Análisis" });
  s.addText("04 · OPORTUNIDADES Y AMENAZAS", { placeholder: "kicker" });
  s.addText("2 oportunidades y 2 desafíos", { placeholder: "title" });
  const od = [
    { x: 120, y: 280, head: "OPORTUNIDAD 1 · PRIORIDAD", fill: HEX.teal, c: HEX.teal, v: "+18.7%", t: "Más dinero en casa", b: "Llega más dinero de familiares en el extranjero", r: "Combos familiares y precios claros" },
    { x: 120, y: 636, head: "OPORTUNIDAD 2", fill: HEX.teal, c: HEX.teal, v: "48%", t: "Del barrio al súper", b: "Aún compra en tiendas pequeñas", r: "Súper de barrio moderno, cerca de casa" },
    { x: 980, y: 280, head: "DESAFÍO 1 · PRIORIDAD", fill: HEX.terra, c: HEX.terra, v: "276", t: "Walmart ya está", b: "Tiendas en Guatemala con precios bajos", r: "Ganar con trato y frescura, no con precio" },
    { x: 980, y: 636, head: "DESAFÍO 2", fill: HEX.terra, c: HEX.terra, v: "2,100 km", t: "Lejos y con inseguridad", b: "Lejos de México y extorsiones a tiendas", r: "Almacén local y socio guatemalteco" },
  ];
  for (const it of od) {
    rect(s, { x: it.x, y: it.y, w: 820, h: 330, fill: HEX.white, line: HEX.line, name: `${it.head} ${it.t}` });
    text(s, it.head, { x: it.x, y: it.y, w: 820, h: 50, fontSize: 10.5, bold: true, charSpacing: 1, color: HEX.white, fill: it.fill, valign: "middle", margin: 12 });
    text(s, it.v, { x: it.x + 32, y: it.y + 66, w: 350, h: 130, bold: true, fontSize: 28, color: it.c, valign: "middle" });
    text(s, [{ text: it.t, options: { bold: true, fontSize: 15, breakLine: true } }, { text: it.b, options: { fontSize: 12, color: HEX.gray } }], { x: it.x + 395, y: it.y + 66, w: 400, h: 130, valign: "middle" });
    hline(s, it.x + 32, it.y + 220, 756, HEX.line, 1);
    text(s, [{ text: "QUÉ HACEMOS  ", options: { fontSize: 10, bold: true, charSpacing: 1, color: it.c } }, { text: it.r, options: { fontSize: 14, bold: true, color: HEX.navy } }], { x: it.x + 32, y: it.y + 232, w: 760, h: 76, valign: "middle" });
  }
  s.addText("Fuentes: Banguat (2025, 2026); Kantar (2022); Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Dos oportunidades y dos desafíos o amenazas, cada uno con lo que haríamos. Oportunidad 1, la principal: las remesas subieron 18.7%; hay más dinero en los hogares y lo aprovechamos con combos familiares y precios claros. Oportunidad 2: el 48% aún compra en tiendas pequeñas; hay espacio para un súper de barrio moderno, justo el formato de Merco. Desafío 1, el principal: Walmart ya tiene 276 tiendas y compite con precios bajos; nosotros ganamos con trato y frescura. Desafío 2: estamos a 2,100 km y hay extorsiones a comercios; lo resolvemos con un almacén local y un socio guatemalteco que conoce el terreno.");

  // 6 · Acción estratégica
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Marfil", sectionTitle: "Propuesta" });
  s.addText("05 · ACCIÓN ESTRATÉGICA", { placeholder: "kicker" });
  s.addText("Un socio local y 2 tiendas piloto, paso a paso", { placeholder: "title" });
  const phases = [
    { ic: "LuHandshake", n: "01", head: "Preparar", when: "Meses 0–6", what: "Firmar con un socio, abrir el almacén y elegir proveedores locales", gate: "Almacén y proveedores listos" },
    { ic: "LuStore", n: "02", head: "Probar", when: "Meses 6–24", what: "Abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura", gate: "Cumplir las 3 metas del piloto" },
    { ic: "LuTrendingUp", n: "03", head: "Crecer", when: "Desde el mes 24", what: "Llevar el modelo a Quetzaltenango y Escuintla", gate: "Solo si el piloto funciona" },
  ];
  hline(s, 210, 340, 1500, HEX.gold, 2);
  for (const [i, ph] of phases.entries()) {
    const x = 120 + i * 568;
    await iconCircle(s, ph.ic, { x: x + 40, y: 290, d: 100, fill: i === 1 ? HEX.gold : HEX.navy, color: HEX.white, alt: ph.head });
    text(s, [{ text: ph.when.toUpperCase(), options: { fontSize: 10.5, bold: true, charSpacing: 1, color: HEX.gold, breakLine: true } }, { text: ph.head, options: { fontFace: SERIF, fontSize: 26, color: HEX.navy } }], { x: x + 40, y: 412, w: 480, h: 96, valign: "middle" });
    text(s, ph.what, { x: x + 40, y: 514, w: 470, h: 130, fontSize: 13, color: HEX.navy });
    text(s, [{ text: "PARA AVANZAR  ", options: { fontSize: 9.5, bold: true, charSpacing: 1, color: HEX.teal } }, { text: ph.gate, options: { fontSize: 12.5, bold: true } }], { x: x + 40, y: 660, w: 480, h: 74, fill: HEX.tealTint, valign: "middle", margin: 12 });
  }
  rect(s, { x: 120, y: 770, w: 1160, h: 200, fill: HEX.navy, name: "Metas del piloto" });
  kicker(s, "Metas del piloto", 160, 794, 600, HEX.gold);
  const goals = [["LuChartColumn", "Vender bien"], ["LuRepeat", "Clientes que vuelven"], ["LuUsers", "El equipo se queda"]];
  for (const [i, [ic, g]] of goals.entries()) {
    const x = 160 + i * 370;
    await iconCircle(s, ic, { x, y: 852, d: 64, fill: HEX.navy2, color: HEX.gold });
    text(s, g, { x: x + 80, y: 852, w: 270, h: 64, fontSize: 14.5, bold: true, color: HEX.white, valign: "middle" });
  }
  rect(s, { x: 1310, y: 770, w: 490, h: 200, fill: HEX.goldTint, name: "Riesgo pequeño" });
  text(s, "5%", { x: 1336, y: 770, w: 130, h: 200, bold: true, fontSize: 32, color: HEX.gold, valign: "middle" });
  text(s, [{ text: "Riesgo pequeño", options: { bold: true, fontSize: 15, breakLine: true } }, { text: "2 tiendas = 5% del tamaño de Merco", options: { fontSize: 12, color: HEX.gray } }], { x: 1470, y: 770, w: 315, h: 200, valign: "middle" });
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });
  s.addNotes("La acción tiene tres fases, y no pasamos a la siguiente sin cumplir la anterior. Fase 1, preparar (meses 0 a 6): firmar con un socio guatemalteco, abrir el almacén en la capital y elegir proveedores locales. Fase 2, probar (meses 6 a 24): abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura y medir tres metas: que vendan bien por metro cuadrado, que los clientes vuelvan y que el equipo se quede (menos del 41% se va al año, como en México). Fase 3, crecer (desde el mes 24): llevar el modelo a Quetzaltenango y Escuintla, solo si el piloto funciona. Esta acción aprovecha las dos oportunidades, resuelve la distancia con el socio y el almacén, y compite con Walmart con trato. Y el riesgo es pequeño: 2 tiendas son el 5% del tamaño de Merco.");

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Azul", sectionTitle: "Propuesta" });
  s.addImage({ path: A("foto-chichi-alta.jpg"), x: P(1280), y: 0, w: P(640), h: P(1080), altText: "Mercado de Chichicastenango" });
  s.addImage({ path: A("degradado-azul.png"), x: P(1278), y: 0, w: P(300), h: P(1080), altText: "Degradado decorativo" });
  text(s, "Foto: Toby Argüelles, CC BY-SA 2.0", { x: 1500, y: 1036, w: 400, h: 30, fontSize: 9, color: HEX.white, align: "right", valign: "middle" });
  s.addText("06 · RECOMENDACIÓN FINAL", { placeholder: "kicker" });
  text(s, [{ text: "Sí a Guatemala:", options: { breakLine: true } }, { text: "paso a paso y con un socio local", options: { color: HEX.gold, italic: true } }], { x: 120, y: 124, w: 1140, h: 160, serif: true, fontSize: 28, color: HEX.white, lineSpacingMultiple: 0.95, name: "Título" });
  text(s, "APROBADO CON CONDICIONES", { x: 120, y: 300, w: 520, h: 46, fontSize: 11, bold: true, color: HEX.navy, fill: HEX.gold, align: "center", valign: "middle" });
  const why = [
    { v: "+4.1%", t: "El mercado crece", b: "Y llega más dinero a casa" },
    { v: "41%", t: "Merco sabe ganar", b: "Su gente se queda (antes 98%)" },
    { v: "5%", t: "Riesgo pequeño", b: "2 tiendas = 5% de su tamaño" },
  ];
  for (const [i, w] of why.entries()) {
    const x = 120 + i * 372;
    hline(s, x, 386, 340, HEX.gold, 2);
    text(s, w.v, { x, y: 396, w: 340, h: 84, bold: true, fontSize: 30, color: HEX.gold, valign: "middle" });
    text(s, [{ text: w.t, options: { bold: true, fontSize: 15, color: HEX.white, breakLine: true } }, { text: w.b, options: { fontSize: 12, color: HEX.soft } }], { x, y: 484, w: 340, h: 110 });
  }
  rect(s, { x: 120, y: 640, w: 520, h: 330, fill: HEX.navy2, name: "Condición" });
  kicker(s, "La condición", 152, 662, 460);
  text(s, "Solo abrimos más tiendas si las 2 primeras funcionan: venden bien, los clientes vuelven y el equipo se queda.", { x: 152, y: 706, w: 460, h: 240, fontSize: 14.5, bold: true, color: HEX.white, lineSpacingMultiple: 1.1 });
  kicker(s, "Siguientes pasos", 680, 640, 500);
  const steps = [["2027", "Socio, almacén y proveedores"], ["2028", "2 tiendas piloto con ofertas de apertura"], ["2029", "Crecer a Quetzaltenango y Escuintla"]];
  for (const [i, [yr, t]] of steps.entries()) {
    const y = 690 + i * 94;
    text(s, yr, { x: 680, y, w: 130, h: 72, bold: true, fontSize: 19, color: HEX.gold, valign: "middle" });
    text(s, t, { x: 820, y, w: 430, h: 72, fontSize: 13.5, color: HEX.white, valign: "middle" });
    if (i < 2) hline(s, 680, y + 82, 540, HEX.navy2, 1);
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });
  s.addNotes("Nuestra recomendación: sí a Guatemala, paso a paso y con un socio local. Comparamos tres caminos: quedarse en México no arriesga pero deja pasar un mercado que crece; entrar en grande es demasiado riesgo frente a Walmart y la distancia; entrar paso a paso es el equilibrio. Tres razones: el mercado crece (economía +4.1% y más remesas); Merco sabe ganar (su gente se queda y sus tiendas rinden); y el riesgo es pequeño, porque 2 tiendas son el 5% de su tamaño. La aprobamos con una condición: solo crecemos si las 2 primeras tiendas funcionan. Siguientes pasos: socio y almacén en 2027, 2 tiendas en 2028 y crecer en 2029. Gracias.");

  // 8 · Fuentes
  pres.addSection({ title: "Anexo" });
  s = pres.addSlide({ masterName: "Marfil", sectionTitle: "Anexo" });
  s.addText("NO SE EXPONE", { placeholder: "kicker" });
  s.addText("Fuentes", { placeholder: "title" });
  const ref = (t, last) => ({ text: t, options: { breakLine: !last, paraSpaceAfter: 6 } });
  const left = [
    "Banco de Guatemala (2026). Ingreso de divisas por remesas familiares 2025. Vía AGN y Emisoras Unidas.",
    "Banco de Guatemala (2025, dic.). Proyección de crecimiento económico 2025–2026 e inflación. banguat.gob.gt",
    "Banco Mundial (2025). Población total, Guatemala. data.worldbank.org",
    "Kantar (2022). Canales de consumo masivo en Guatemala.",
    "FocusEconomics (2026). Guatemala: PIB per cápita en US$.",
    "Prensa Latina (2025, 27 ago.). Cámaras empresariales de Guatemala señalaron aumento de extorsiones.",
    "Walmart (2025). Central America: retail units. corporate.walmart.com",
    "MINECO Guatemala (2025). Informe de comercio exterior.",
    "Emisoras Unidas (2025). Súper del Barrio celebra su tienda 200 en Guatemala.",
  ];
  text(s, left.map((t, i) => ref(t, i === left.length - 1)), { x: 120, y: 250, w: 810, h: 720, fontSize: 11.5, color: HEX.navy, name: "Fuentes columna 1" });
  text(s, [
    ref("Carreón Flores, H. et al. (2026). Supermercados Merco: Crecimiento centrado en las personas (RC32-11-270). Tec de Monterrey."),
    ref("Great Place to Work México (2025). Merco."),
    ref("Dess, G. G. et al. (2021). Strategic management (10.ª ed.). McGraw-Hill."),
    ref("Equipo 2 (2026). Diagnóstico externo, diagnóstico interno y matriz FODA de Supermercados Merco."),
    ref("[Passport, Statista y EMIS: agregar]"),
    { text: "Fotos (Wikimedia Commons): ", options: { bold: true } },
    ref("Chad Davis, CC BY-SA 2.0 (Antigua); Francisco Anzola, CC BY 2.0 (mercado); Toby Argüelles, CC BY-SA 2.0 (Chichicastenango). Mapa: Natural Earth. Iconos: Lucide (ISC)."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 990, y: 250, w: 810, h: 720, fontSize: 11.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
