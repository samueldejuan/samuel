// Genera Pitch_Merco_Guatemala.pptx: estilo sobrio y compatible (sin giros, sombras ni formas especiales).
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
  navy: "1B2A4A", navy2: "2A3B5F", gold: "C99A2E", goldTint: "F6EEDB", teal: "1F6F6B", tealTint: "E3EFEE",
  red: "A63D40", redTint: "F6E7E7", light: "F4F5F7", white: "FFFFFF", gray: "5B6475", line: "DDE1E8", soft: "C9D2E0",
};
const THEME = {
  name: "Merco Guatemala",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.navy, lt1: HEX.white, dk2: HEX.navy2, lt2: HEX.light,
    accent1: HEX.navy, accent2: HEX.gold, accent3: HEX.teal, accent4: HEX.red,
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
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: 256, color: `#${hex}`, strokeWidth: 2 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Diseños (layouts) ----------
function layout(name, bg, { kicker, title, source, number = true }) {
  const def = {
    title: name,
    background: { color: bg },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: P(120), y: P(70), w: P(1300), h: P(40), fontSize: 9.5, bold: true, color: kicker, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: P(120), y: P(116), w: P(1680), h: P(160), fontSize: 26, bold: true, color: title, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.95 }, text: "" } },
      { placeholder: { options: { name: "source", type: "body", x: P(120), y: P(1000), w: P(1500), h: P(40), fontSize: 9, color: source, margin: 0, valign: "middle" }, text: "" } },
    ],
  };
  if (number) def.slideNumber = { x: P(1700), y: P(1000), w: P(100), h: P(40), fontSize: 9, color: source, align: "right", margin: 0 };
  pres.defineSlideMaster(def);
}
layout("Claro", HEX.light, { kicker: HEX.teal, title: HEX.navy, source: HEX.gray });
layout("Oscuro", HEX.navy, { kicker: HEX.gold, title: HEX.white, source: HEX.soft });

// ---------- Ayudantes ----------
function text(slide, runs, o) {
  slide.addText(runs, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), margin: o.margin ?? 0, isTextBox: true,
    valign: o.valign || "top", align: o.align || "left", fontSize: o.fontSize || 14,
    color: o.color || HEX.navy, bold: o.bold, lineSpacingMultiple: o.lineSpacingMultiple, objectName: o.name,
    fill: o.fill ? { color: o.fill } : undefined, shape: o.fill ? pres.shapes.ROUNDED_RECTANGLE : undefined, rectRadius: 0.06,
    paraSpaceAfter: o.paraSpaceAfter,
  });
}
function box(slide, o) {
  slide.addShape(o.oval ? pres.shapes.OVAL : pres.shapes.ROUNDED_RECTANGLE, {
    x: P(o.x), y: P(o.y), w: P(o.w), h: P(o.h), rectRadius: 0.06,
    fill: { color: o.fill }, line: o.line ? { color: o.line, width: o.lineW || 1 } : { type: "none" }, objectName: o.name,
  });
}
async function iconCircle(slide, name, o) {
  box(slide, { x: o.x, y: o.y, w: o.d, h: o.d, fill: o.fill, oval: true, name: `Icono ${name}` });
  const pad = o.d * 0.26;
  slide.addImage({ data: await icon(name, o.color), x: P(o.x + pad), y: P(o.y + pad), w: P(o.d - 2 * pad), h: P(o.d - 2 * pad), altText: o.alt || name });
}
const label = (slide, t, x, y, w, color = HEX.teal, align = "left") => text(slide, t.toUpperCase(), { x, y, w, h: 36, fontSize: 9.5, bold: true, color, valign: "middle", align });

(async () => {
  // 1 · Portada
  pres.addSection({ title: "Inicio" });
  let s = pres.addSlide({ masterName: "Oscuro", sectionTitle: "Inicio" });
  s.addImage({ path: A("foto-antigua.jpg"), x: P(1100), y: 0, w: P(820), h: P(1080), altText: "Arco de Santa Catalina, Antigua Guatemala" });
  text(s, "PROPUESTA DE EXPANSIÓN · SUPERMERCADOS MERCO", { x: 120, y: 300, w: 940, h: 40, fontSize: 11, bold: true, color: HEX.gold, valign: "middle" });
  text(s, [
    { text: "Próxima parada:", options: { breakLine: true } },
    { text: "Guatemala", options: { color: HEX.gold } },
  ], { x: 120, y: 350, w: 940, h: 250, fontSize: 44, bold: true, color: HEX.white, lineSpacingMultiple: 0.95, name: "Título" });
  s.addShape(pres.shapes.LINE, { x: P(124), y: P(640), w: P(120), h: 0, line: { color: HEX.gold, width: 3 }, objectName: "Separador" });
  text(s, "En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio. Merco lleva 80 años siendo ese súper en el norte de México.", { x: 120, y: 670, w: 880, h: 110, fontSize: 15, color: HEX.soft });
  text(s, "Samuel de Juan Ortiz · A01833027 · Estrategia y globalización (Gpo 504)", { x: 120, y: 990, w: 940, h: 40, fontSize: 9, color: HEX.soft, valign: "middle" });
  text(s, "Dato: Kantar (2022) · Foto: Chad Davis, CC BY-SA 2.0", { x: 1300, y: 1030, w: 600, h: 30, fontSize: 9, color: HEX.white, align: "right", valign: "middle" });
  s.addNotes("En Guatemala, 1 de cada 2 compras se hace en la tienda de barrio (Kantar). Merco lleva 80 años siendo justo eso en el norte de México: un súper cercano para familias que cuidan su dinero. Hoy proponemos llevar ese modelo a Ciudad de Guatemala, paso a paso. Vamos a ver si Merco está lista, por qué Guatemala, cómo entraríamos y qué recomendamos.");

  // 2 · Mapa de fortalezas y debilidades
  pres.addSection({ title: "Análisis" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 1 · DIAGNÓSTICO", { placeholder: "kicker" });
  s.addText("Mapa de fortalezas y debilidades", { placeholder: "title" });
  label(s, "Fortaleza", 120, 296, 700);
  label(s, "Cómo la usamos", 820, 296, 280, HEX.gray, "center");
  label(s, "Debilidad que cubre", 1100, 296, 700, HEX.red);
  const map = [
    { f: ["$118K", "Tiendas que rinden", "Vende más por metro cuadrado que Soriana y Walmart"], m: "Con tiendas chicas, abrir cuesta poco", d: ["39", "Es chica", "Tiendas en solo 2 estados"] },
    { f: ["80", "Experta en cercanía", "Años vendiendo a familias que cuidan su dinero"], m: "El trato cercano vale más que una app", d: ["0", "Atrás en lo digital", "Sin programa de lealtad"] },
    { f: ["41%", "Cuida a su gente", "Empleados que se van al año (antes 98%)"], m: "Forma a su gente y se apoya en un socio", d: ["0", "Nunca ha salido", "Sin experiencia fuera de México"] },
  ];
  for (const [i, r] of map.entries()) {
    const y = 344 + i * 170;
    box(s, { x: 120, y, w: 700, h: 150, fill: HEX.white, line: HEX.line, name: `Fortaleza ${i + 1}` });
    box(s, { x: 120, y, w: 230, h: 150, fill: HEX.tealTint, name: `Fortaleza ${i + 1} dato` });
    text(s, r.f[0], { x: 120, y, w: 230, h: 150, fontSize: 24, bold: true, color: HEX.teal, align: "center", valign: "middle" });
    text(s, [{ text: r.f[1], options: { bold: true, fontSize: 13, breakLine: true } }, { text: r.f[2], options: { fontSize: 10.5, color: HEX.gray } }], { x: 376, y, w: 420, h: 150, valign: "middle" });
    text(s, r.m, { x: 830, y: y, w: 260, h: 88, fontSize: 10, bold: true, color: HEX.navy, align: "center", valign: "bottom" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: P(870), y: P(y + 92), w: P(180), h: P(26), fill: { color: HEX.gold }, line: { type: "none" }, objectName: `Flecha ${i + 1}` });
    box(s, { x: 1100, y, w: 700, h: 150, fill: HEX.white, line: HEX.line, name: `Debilidad ${i + 1}` });
    box(s, { x: 1100, y, w: 180, h: 150, fill: HEX.redTint, name: `Debilidad ${i + 1} dato` });
    text(s, r.d[0], { x: 1100, y, w: 180, h: 150, fontSize: 24, bold: true, color: HEX.red, align: "center", valign: "middle" });
    text(s, [{ text: r.d[1], options: { bold: true, fontSize: 13, breakLine: true } }, { text: r.d[2], options: { fontSize: 10.5, color: HEX.gray } }], { x: 1306, y, w: 470, h: 150, valign: "middle" });
  }
  box(s, { x: 120, y: 870, w: 430, h: 96, fill: HEX.goldTint, name: "Nivel de preparación" });
  text(s, [{ text: "PREPARACIÓN: MEDIA", options: { bold: true, fontSize: 12, color: HEX.navy, breakLine: true } }, { text: "Lista para empezar en pequeño", options: { fontSize: 9.5, color: HEX.gray } }], { x: 146, y: 870, w: 390, h: 96, valign: "middle" });
  text(s, [{ text: "Conclusión: ", options: { bold: true, color: HEX.gold } }, { text: "su forma de trabajar funciona, pero es pequeña y nunca ha salido de México. Por eso debe empezar con 2 tiendas y un socio local." }],
    { x: 570, y: 870, w: 1230, h: 96, fontSize: 12.5, color: HEX.white, fill: HEX.navy, valign: "middle", margin: 0.2 });
  s.addText("Fuentes: diagnóstico interno del Equipo 2; caso Merco (Carreón et al., 2026); Great Place to Work (2025)", { placeholder: "source" });
  s.addNotes("Este mapa conecta cada fortaleza con la debilidad que ayuda a cubrir. Sus tiendas chicas venden $118,000 por metro cuadrado, más que Soriana y Walmart: como no necesita tiendas grandes, su tamaño pequeño (39 tiendas en 2 estados) pesa menos. Lleva 80 años como súper de barrio: su trato cercano compensa que no tenga programa de lealtad ni mucha tecnología. Y cuida a su gente: la rotación bajó de 98% a 41%; con la Universidad Merco y un socio local puede formar equipo en un país donde nunca ha operado. Conclusión: preparación media. Está lista para empezar, pero en pequeño y con un socio local.");

  // 3 · Mercado seleccionado y datos clave
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 2 · EL MERCADO", { placeholder: "kicker" });
  s.addText("Mercado seleccionado y datos clave", { placeholder: "title" });
  const stats = [
    { v: "18.7 M", l: "personas: el país con más gente de Centroamérica", src: "Banco Mundial, 2025" },
    { v: "+4.1%", l: "crecerá su economía, más que América Latina", src: "Banguat, 2025" },
    { v: "+18.7%", l: "de remesas: dinero de familiares en el extranjero", src: "Banguat, 2026" },
  ];
  for (const [i, st] of stats.entries()) {
    const x = 120 + i * 325;
    box(s, { x, y: 300, w: 305, h: 220, fill: HEX.white, line: HEX.line, name: `Dato ${i + 1}` });
    text(s, st.v, { x: x + 26, y: 316, w: 260, h: 74, fontSize: 27, bold: true, color: HEX.navy, valign: "middle" });
    text(s, [{ text: st.l, options: { breakLine: true } }, { text: st.src, options: { fontSize: 9, color: HEX.gray, bold: false } }], { x: x + 26, y: 394, w: 260, h: 116, fontSize: 10.5, bold: true, color: HEX.navy });
  }
  const panels = [
    { x: 120, head: "Por qué encaja con Merco", color: HEX.teal, tint: HEX.tealTint, rows: [
      ["LuPiggyBank", "Cuida su dinero", "Ingreso medio: US$6,506 por persona"],
      ["LuMapPin", "Entrada por la capital", "Ciudad de Guatemala"],
      ["LuHandshake", "Cerca de México", "México ya es su 3.er proveedor"],
    ] },
    { x: 610, head: "Competencia (amenaza)", color: HEX.red, tint: HEX.redTint, rows: [
      ["LuStore", "Walmart", "276 tiendas"],
      ["LuStore", "Súper del Barrio", "200 tiendas de cercanía"],
      ["LuStore", "La Torre", "Cadena local de supermercados"],
    ] },
  ];
  for (const pn of panels) {
    box(s, { x: pn.x, y: 548, w: 470, h: 390, fill: HEX.white, line: HEX.line, name: pn.head });
    label(s, pn.head, pn.x + 28, 562, 420, pn.color);
    for (const [i, [ic, t, b]] of pn.rows.entries()) {
      const y = 616 + i * 104;
      await iconCircle(s, ic, { x: pn.x + 28, y: y + 12, d: 64, fill: pn.tint, color: pn.color });
      text(s, [{ text: t, options: { bold: true, fontSize: 11.5, breakLine: true } }, { text: b, options: { fontSize: 10, color: HEX.gray } }], { x: pn.x + 110, y, w: 340, h: 88, valign: "middle" });
    }
  }
  const mw = 680, mk = mw / G.mapa.W, mh = G.mapa.H * mk, mx = 1120, my = 300;
  s.addImage({ path: A("mapa-ruta.png"), x: P(mx), y: P(my), w: P(mw), h: P(mh), altText: "Mapa de la ruta de Monterrey a Ciudad de Guatemala" });
  const [ax, ay] = [mx + G.mapa.mty[0] * mk, my + G.mapa.mty[1] * mk];
  const [gx, gy] = [mx + G.mapa.gt[0] * mk, my + G.mapa.gt[1] * mk];
  text(s, "Monterrey", { x: ax - 170, y: ay - 46, w: 160, h: 34, fontSize: 10.5, bold: true, color: HEX.navy, align: "right", valign: "middle" });
  text(s, "Ciudad de Guatemala", { x: gx - 300, y: gy + 18, w: 280, h: 34, fontSize: 10.5, bold: true, color: HEX.navy, align: "right", valign: "middle" });
  text(s, [{ text: "≈ 2,100 km ", options: { bold: true, color: HEX.navy } }, { text: "por carretera desde Nuevo León", options: { color: HEX.gray } }], { x: mx, y: my + mh + 20, w: mw, h: 50, fontSize: 10.5, align: "center", valign: "middle", fill: HEX.white });
  s.addText("Fuentes: Banco Mundial (2025); Banguat (2025, 2026); FocusEconomics (2026); MINECO (2025); Walmart (2025); Emisoras Unidas (2025) · Mapa: Natural Earth", { placeholder: "source" });
  s.addNotes("Elegimos Guatemala y entramos por su capital. Tres datos clave: es el país más poblado de Centroamérica, con 18.7 millones de personas; su economía crecerá 4.1%, más que América Latina; y las remesas subieron 18.7% en 2025, hasta US$25,530 millones, así que las familias tienen más dinero para el súper. Encaja con Merco porque el cliente cuida su dinero (PIB por persona de US$6,506) y México ya es su tercer proveedor. La competencia es real: Walmart tiene 276 tiendas, Súper del Barrio 200 y está La Torre. Y estamos a unos 2,100 km de Nuevo León.");

  // 4 · Estrategia semiglobal y 2 áreas de adaptación
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 3 · LA ESTRATEGIA", { placeholder: "kicker" });
  s.addText("Estrategia semiglobal y 2 áreas de adaptación", { placeholder: "title" });
  s.addImage({ path: A("foto-mercado.jpg"), x: P(1380), y: P(300), w: P(420), h: P(632), altText: "Mercado de Guatemala con frutas y especias" });
  text(s, "Foto: Francisco Anzola, CC BY 2.0", { x: 1380, y: 940, w: 420, h: 30, fontSize: 9, color: HEX.gray, align: "right", valign: "middle" });
  text(s, [{ text: "Semiglobal: ", options: { bold: true, color: HEX.gold } }, { text: "lo que hace especial a Merco se queda igual; solo cambiamos 2 cosas para el cliente de Guatemala." }],
    { x: 120, y: 296, w: 1230, h: 84, fontSize: 12.5, color: HEX.white, fill: HEX.navy, valign: "middle", margin: 0.2 });
  label(s, "Se mantiene", 120, 392, 600);
  const keep = [
    ["LuUsers", "Su gente", "Buen sueldo y formación propia"],
    ["LuShieldCheck", "Sus garantías", "Frescura y 30 días para devolver"],
    ["LuStore", "Súper de barrio", "Tiendas chicas y cercanas"],
  ];
  for (const [i, [ic, t, b]] of keep.entries()) {
    const x = 120 + i * 418;
    box(s, { x, y: 436, w: 394, h: 130, fill: HEX.white, line: HEX.line, name: t });
    await iconCircle(s, ic, { x: x + 22, y: 469, d: 64, fill: HEX.tealTint, color: HEX.teal });
    text(s, [{ text: t, options: { bold: true, fontSize: 12, breakLine: true } }, { text: b, options: { fontSize: 9.5, color: HEX.gray } }], { x: x + 102, y: 436, w: 276, h: 130, valign: "middle" });
  }
  label(s, "Se adapta: 2 áreas", 120, 592, 600, HEX.gold);
  const adapt = [
    { ic: "LuWheat", t: "1 · Producto", hoy: "Surtido para el norte de México", gt: "Frijol negro, maíz y café local + lo mejor de México", por: "El cliente busca sus productos de siempre" },
    { ic: "LuWarehouse", t: "2 · Distribución", hoy: "Todo sale de 3 centros en Nuevo León", gt: "Almacén en Ciudad de Guatemala y proveedores locales", por: "Lo fresco no aguanta 2,100 km" },
  ];
  for (const [i, a] of adapt.entries()) {
    const x = 120 + i * 627;
    box(s, { x, y: 636, w: 603, h: 296, fill: HEX.white, line: HEX.line, name: a.t });
    await iconCircle(s, a.ic, { x: x + 24, y: 656, d: 60, fill: HEX.goldTint, color: HEX.gold });
    text(s, a.t, { x: x + 100, y: 656, w: 480, h: 60, fontSize: 13.5, bold: true, valign: "middle" });
    text(s, [
      { text: "Hoy: ", options: { bold: true, color: HEX.gray } }, { text: a.hoy, options: { color: HEX.gray, breakLine: true } },
      { text: "En Guatemala: ", options: { bold: true, color: HEX.teal } }, { text: a.gt, options: { bold: true, breakLine: true } },
      { text: "Por qué: ", options: { bold: true, color: HEX.gold } }, { text: a.por },
    ], { x: x + 24, y: 724, w: 555, h: 196, fontSize: 11, paraSpaceAfter: 5 });
  }
  s.addText("Marco: Dess et al. (2021), cap. 1 · México, 3.er proveedor de Guatemala (MINECO, 2025)", { placeholder: "source" });
  s.addNotes("Nuestra estrategia es semiglobal: mantenemos lo que hace única a Merco y adaptamos solo dos áreas. Se mantiene su gente (ingreso digno y Universidad Merco), sus garantías (frescura garantizada y 30 días para devolver) y su formato de súper de barrio. Adaptamos, primero, el producto: en Guatemala la gente busca frijol negro, maíz y café local, así que el surtido mezcla lo chapín con lo mejor de México. Y segundo, la distribución: hoy todo sale de 3 centros en Nuevo León, pero a más de 2,100 km lo fresco no llega bien, así que abrimos un almacén en Ciudad de Guatemala con proveedores locales. ¿Por qué semiglobal y no global o local? Porque su cultura sí viaja, el gusto local es distinto y adaptar solo 2 áreas cuesta menos.");

  // 5 · 2 oportunidades y 2 desafíos
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Análisis" });
  s.addText("PASO 4 · OPORTUNIDADES Y DESAFÍOS", { placeholder: "kicker" });
  s.addText("2 oportunidades y 2 desafíos", { placeholder: "title" });
  label(s, "Oportunidades", 120, 296, 700, HEX.teal);
  label(s, "Desafíos (amenazas)", 980, 296, 700, HEX.red);
  const od = [
    { x: 120, y: 340, c: HEX.teal, tint: HEX.tealTint, tag: "O1 · Prioridad 1", v: "+18.7%", t: "Más dinero en casa", b: "Llega más dinero de familiares en el extranjero", r: "Combos familiares y precios claros" },
    { x: 120, y: 650, c: HEX.teal, tint: HEX.tealTint, tag: "O2", v: "48%", t: "Del barrio al súper", b: "Aún compra en tiendas pequeñas", r: "Súper de barrio moderno, cerca de casa" },
    { x: 980, y: 340, c: HEX.red, tint: HEX.redTint, tag: "D1 · Prioridad 1", v: "276", t: "Walmart ya está", b: "Tiendas en Guatemala con precios bajos", r: "Ganar con trato y frescura, no con precio" },
    { x: 980, y: 650, c: HEX.red, tint: HEX.redTint, tag: "D2", v: "2,100 km", t: "Lejos y con inseguridad", b: "Lejos de México y extorsiones a tiendas", r: "Almacén local y socio guatemalteco" },
  ];
  for (const it of od) {
    box(s, { x: it.x, y: it.y, w: 820, h: 284, fill: HEX.white, line: HEX.line, name: `${it.tag} ${it.t}` });
    text(s, it.tag.toUpperCase(), { x: it.x + 28, y: it.y + 18, w: 400, h: 32, fontSize: 9, bold: true, color: it.c, valign: "middle" });
    text(s, it.v, { x: it.x + 28, y: it.y + 54, w: 340, h: 110, fontSize: 27, bold: true, color: it.c, valign: "middle" });
    text(s, [{ text: it.t, options: { bold: true, fontSize: 13.5, breakLine: true } }, { text: it.b, options: { fontSize: 10.5, color: HEX.gray } }], { x: it.x + 380, y: it.y + 54, w: 416, h: 110, valign: "middle" });
    text(s, [{ text: "Qué hacemos: ", options: { bold: true, color: it.c } }, { text: it.r }], { x: it.x + 20, y: it.y + 186, w: 780, h: 76, fontSize: 11, color: HEX.navy, fill: it.tint, valign: "middle", margin: 0.2 });
  }
  s.addText("Fuentes: Banguat (2025, 2026); Kantar (2022); Walmart (2025); Prensa Latina (2025)", { placeholder: "source" });
  s.addNotes("Dos oportunidades y dos desafíos, cada uno con su respuesta. O1, la prioridad: las remesas subieron 18.7% y la economía crece 4.1%; hay más dinero en los hogares y lo aprovechamos con combos familiares y precios claros. O2: el 48% aún compra en tiendas pequeñas; hay espacio para un súper de barrio moderno, que es justo el formato de Merco. D1, el desafío principal: Walmart ya tiene 276 tiendas y compite por precio; nosotros ganamos con trato y frescura. D2: estamos a 2,100 km y hay extorsiones a comercios; lo resolvemos con un almacén local y un socio guatemalteco que conoce el terreno.");

  // 6 · Acción estratégica
  pres.addSection({ title: "Propuesta" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Propuesta" });
  s.addText("PASO 5 · ACCIÓN ESTRATÉGICA", { placeholder: "kicker" });
  s.addText("Acción estratégica: un socio local y 2 tiendas piloto", { placeholder: "title" });
  const phases = [
    { ic: "LuHandshake", n: "1", head: "Preparar", when: "Meses 0–6", what: "Firmar con un socio guatemalteco, abrir el almacén y elegir proveedores locales", gate: "Almacén y proveedores listos" },
    { ic: "LuStore", n: "2", head: "Probar", when: "Meses 6–24", what: "Abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura", gate: "Cumplir las 3 metas del piloto" },
    { ic: "LuTrendingUp", n: "3", head: "Crecer", when: "Desde el mes 24", what: "Llevar el modelo a Quetzaltenango y Escuintla", gate: "Solo si el piloto funciona" },
  ];
  for (const [i, ph] of phases.entries()) {
    const x = 120 + i * 568;
    box(s, { x, y: 300, w: 540, h: 440, fill: HEX.white, line: HEX.line, name: `Fase ${ph.n}` });
    await iconCircle(s, ph.ic, { x: x + 30, y: 330, d: 80, fill: i === 1 ? HEX.gold : HEX.navy, color: HEX.white, alt: ph.head });
    text(s, [{ text: `FASE ${ph.n} · ${ph.when.toUpperCase()}`, options: { fontSize: 9, bold: true, color: HEX.teal, breakLine: true } }, { text: ph.head, options: { fontSize: 21, bold: true } }], { x: x + 132, y: 330, w: 390, h: 80, valign: "middle" });
    text(s, ph.what, { x: x + 30, y: 440, w: 480, h: 170, fontSize: 12, color: HEX.navy });
    text(s, [{ text: "Para avanzar: ", options: { bold: true, color: HEX.teal } }, { text: ph.gate }], { x: x + 20, y: 640, w: 500, h: 80, fontSize: 10.5, fill: HEX.tealTint, valign: "middle", margin: 0.2 });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: P(x + 542), y: P(500), w: P(24), h: P(40), fill: { color: HEX.gold }, line: { type: "none" }, objectName: "Flecha" });
  }
  box(s, { x: 120, y: 770, w: 1100, h: 170, fill: HEX.navy, name: "Metas del piloto" });
  label(s, "Metas del piloto", 150, 786, 600, HEX.gold);
  const goals = ["Vender bien", "Clientes que vuelven", "El equipo se queda"];
  for (const [i, g] of goals.entries()) text(s, g, { x: 150 + i * 350, y: 840, w: 330, h: 74, fontSize: 11.5, bold: true, color: HEX.navy, align: "center", valign: "middle", fill: HEX.white });
  box(s, { x: 1250, y: 770, w: 550, h: 170, fill: HEX.goldTint, name: "Riesgo acotado" });
  text(s, "5%", { x: 1276, y: 770, w: 170, h: 170, fontSize: 35, bold: true, color: HEX.gold, valign: "middle" });
  text(s, [{ text: "Riesgo pequeño", options: { bold: true, fontSize: 13, breakLine: true } }, { text: "2 tiendas son el 5% del tamaño de Merco", options: { fontSize: 10.5, color: HEX.gray } }], { x: 1450, y: 770, w: 330, h: 170, valign: "middle" });
  s.addText("Propuesta propia a partir del diagnóstico · Inversión: se define en el estudio de factibilidad [validar con EMIS]", { placeholder: "source" });
  s.addNotes("La acción tiene tres fases, y no pasamos a la siguiente sin cumplir la anterior. Fase 1, preparar (meses 0 a 6): firmar con un socio guatemalteco, abrir el almacén en la capital y elegir proveedores locales. Fase 2, probar (meses 6 a 24): abrir 2 tiendas de cercanía en Ciudad de Guatemala con ofertas de apertura y medir tres metas: ventas por metro cuadrado, clientes que regresan y rotación menor al 41%, como en México. Fase 3, crecer (desde el mes 24): llevar el modelo a Quetzaltenango y Escuintla, solo si el piloto funciona. Esta acción aprovecha O1 y O2, resuelve D2 con el socio y el almacén, y compite con Walmart con trato. Y el riesgo es acotado: 2 tiendas son el 5% del tamaño de Merco.");

  // 7 · Recomendación final
  s = pres.addSlide({ masterName: "Oscuro", sectionTitle: "Propuesta" });
  s.addText("PASO 6 · RECOMENDACIÓN FINAL", { placeholder: "kicker" });
  s.addText("Recomendación final: sí a Guatemala, paso a paso y con un socio local", { placeholder: "title", w: P(1230) });
  text(s, "APROBADO CON CONDICIONES", { x: 1300, y: 50, w: 500, h: 56, fontSize: 11, bold: true, color: HEX.navy, align: "center", valign: "middle", fill: HEX.gold });
  const why = [
    { v: "+4.1%", t: "El mercado crece", b: "Y llega más dinero a las familias" },
    { v: "41%", t: "Merco sabe ganar", b: "Su gente se queda (antes se iba el 98%)" },
    { v: "5%", t: "El riesgo es pequeño", b: "2 tiendas son el 5% de su tamaño" },
  ];
  for (const [i, w] of why.entries()) {
    const x = 120 + i * 568;
    box(s, { x, y: 300, w: 540, h: 260, fill: HEX.navy2, name: w.t });
    text(s, w.v, { x: x + 34, y: 316, w: 480, h: 110, fontSize: 35, bold: true, color: HEX.gold, valign: "middle" });
    text(s, [{ text: w.t, options: { bold: true, fontSize: 14.5, color: HEX.white, breakLine: true } }, { text: w.b, options: { fontSize: 11, color: HEX.soft } }], { x: x + 34, y: 432, w: 480, h: 110, valign: "top" });
  }
  box(s, { x: 120, y: 600, w: 640, h: 330, fill: HEX.white, name: "Condición" });
  label(s, "La condición", 150, 620, 560, HEX.gold);
  text(s, "Solo abrimos más tiendas si las 2 primeras funcionan: venden bien, los clientes vuelven y el equipo se queda.", { x: 150, y: 668, w: 580, h: 240, fontSize: 13.5, bold: true, color: HEX.navy });
  label(s, "Siguientes pasos", 810, 600, 600, HEX.gold);
  const steps = [["2027", "Socio, almacén y proveedores"], ["2028", "2 tiendas piloto con ofertas de apertura"], ["2029", "Crecer a Quetzaltenango y Escuintla"]];
  for (const [i, [yr, t]] of steps.entries()) {
    const y = 650 + i * 96;
    text(s, yr, { x: 810, y, w: 150, h: 80, fontSize: 16, bold: true, color: HEX.navy, align: "center", valign: "middle", fill: HEX.gold });
    text(s, t, { x: 984, y, w: 816, h: 80, fontSize: 13, color: HEX.white, valign: "middle" });
  }
  s.addText("Evidencia: láminas 2 a 6 · Fuentes completas en la lámina 8", { placeholder: "source" });
  s.addNotes("Nuestra recomendación: sí a Guatemala, paso a paso y con un socio local. Comparamos tres caminos: quedarse en México no arriesga pero deja pasar un mercado que crece; entrar en grande es demasiado riesgo frente a Walmart y la distancia; entrar paso a paso es el equilibrio. Tres razones: el mercado crece (economía +4.1% y remesas +18.7%); el modelo de Merco gana (rotación de 98% a 41% y $118,000 por m²); y el riesgo es acotado, porque 2 tiendas son el 5% de su tamaño. La aprobamos con una condición: solo crecemos si el piloto cumple sus 3 metas. Siguientes pasos: socio y almacén en 2027, 2 tiendas en 2028 y crecer en 2029. Gracias.");

  // 8 · Fuentes
  pres.addSection({ title: "Anexo" });
  s = pres.addSlide({ masterName: "Claro", sectionTitle: "Anexo" });
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
  text(s, left.map((t, i) => ref(t, i === left.length - 1)), { x: 120, y: 230, w: 810, h: 740, fontSize: 9.5, color: HEX.navy, name: "Fuentes columna 1" });
  text(s, [
    ref("Carreón Flores, H. et al. (2026). Supermercados Merco: Crecimiento centrado en las personas (RC32-11-270). Tec de Monterrey."),
    ref("Great Place to Work México (2025). Merco."),
    ref("Dess, G. G. et al. (2021). Strategic management (10.ª ed.). McGraw-Hill."),
    ref("Equipo 2 (2026). Diagnóstico externo, diagnóstico interno y matriz FODA de Supermercados Merco."),
    ref("[Passport, Statista y EMIS: agregar]"),
    { text: "Fotos (Wikimedia Commons): ", options: { bold: true } },
    ref("Chad Davis, CC BY-SA 2.0 (Antigua); Francisco Anzola, CC BY 2.0 (mercado). Mapa: Natural Earth. Iconos: Lucide (ISC)."),
    { text: "Uso de IA: ", options: { bold: true } },
    ref("Claude (Anthropic, 2026) para investigar, buscar fotos libres y diseñar; cada dato se revisó en su fuente.", true),
  ], { x: 990, y: 230, w: 810, h: 740, fontSize: 9.5, color: HEX.navy, name: "Fuentes columna 2" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
