// Neubewertung Kreuzchenplan P800 ohne Schuss 309/310
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require("/root/.claude/skills/synced/10f6f57f-31e2-469d-b54e-5e73bdc38b45_cb9ce55c-1b2b-43d1-a3b0-bcd2798483a7/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Werft Stahl",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B232C", lt1: "FFFFFF", dk2: "22303E", lt2: "EEF1F4",
    accent1: "E8641B", // Schweiß-Orange
    accent2: "2E9BD6", // Schuss 301
    accent3: "5FA84A", // Schuss 302 / positiv
    accent4: "8E63B5", // nicht verwendet (früher 309)
    accent5: "8A96A3", // Testschuss 1 / neutral
    accent6: "C8102E", // kritisch
    hlink: "2E9BD6", folHlink: "8E63B5",
  },
};
const HEX = THEME.colors;
const OUT = "/tmp/build/Kreuzchenplan_P800_ohne_309_Neubewertung.pptx";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Kreuzchenplan P800 ohne Schuss 309 – Neubewertung";
pres.author = "Projektmanagement";
const C = pres.SchemeColor;

const FOOT = "Kreuzchenplan Schweißproben PHPL Wismar · Neubewertung ohne Schuss 309/310 · Stand 31.10.2026";

pres.defineSlideMaster({
  title: "TITEL",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.9, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.35, w: 11.5, h: 1.0, fontSize: 46, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 3.55, w: 10.5, h: 0.9, fontSize: 20, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "INHALT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.32, w: 9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.3, h: 0.62, fontSize: 26, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 7.05, w: 10, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.33, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});
pres.defineSlideMaster({
  title: "INHALT_DUNKEL",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.32, w: 9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.3, h: 0.62, fontSize: 26, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 7.05, w: 10, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.33, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
});

let nObj = 0;
const nm = (p) => `${p}_${++nObj}`;
function tb(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, objectName: nm("txt"), fontSize: 14, color: C.text1, margin: 0, valign: "top" }, o));
}
function box(slide, o) {
  const r = o.rounded; const opts = Object.assign({}, o); delete opts.rounded;
  slide.addShape(r ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE,
    Object.assign({ objectName: nm("form"), line: { type: "none" } }, r ? { rectRadius: 0.06 } : {}, opts));
}
function shadow() { return { type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 }; }
function card(slide, x, y, w, h, fill) { box(slide, { x, y, w, h, fill: { color: fill || C.background1 }, rounded: true, shadow: shadow(), line: { color: "DDE2E7", width: 0.75 } }); }

// Icons
async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
async function iconCircle(slide, Comp, x, y, d, bg, fg) {
  slide.addShape(pres.shapes.OVAL, { objectName: nm("iconkreis"), x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  const pad = d * 0.24;
  slide.addImage({ data: await icon(Comp, fg || "FFFFFF"), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: "Symbol", objectName: nm("icon") });
}

// Datum
const D = (s) => { const [d, m, y] = s.split("."); return new Date(Date.UTC(2000 + +y, +m - 1, +d)); };
const DAY = 86400000;
function gantt(slide, cfg) {
  const { x, y, w, labelW, start, end, rowH, rows } = cfg;
  const gx = x + labelW, gw = w - labelW;
  const t0 = D(start).getTime(), t1 = D(end).getTime() + DAY;
  const X = (s) => gx + ((D(s).getTime() - t0) / (t1 - t0)) * gw;
  const Xe = (s) => gx + ((D(s).getTime() + DAY - t0) / (t1 - t0)) * gw;
  const headH = 0.34, totalH = headH + rows.length * rowH;
  const months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
  rows.forEach((r, i) => { if (i % 2 === 0) box(slide, { x, y: y + headH + i * rowH, w, h: rowH, fill: { color: C.background2 } }); });
  let cur = new Date(t0); cur = new Date(Date.UTC(cur.getUTCFullYear(), cur.getUTCMonth(), 1));
  while (cur.getTime() < t1) {
    const nxt = new Date(Date.UTC(cur.getUTCFullYear(), cur.getUTCMonth() + 1, 1));
    const a = Math.max(cur.getTime(), t0), b = Math.min(nxt.getTime(), t1);
    const xa = gx + ((a - t0) / (t1 - t0)) * gw, xb = gx + ((b - t0) / (t1 - t0)) * gw;
    if (xb - xa > 0.6) tb(slide, `${months[cur.getUTCMonth()]} ${String(cur.getUTCFullYear()).slice(2)}`, { x: xa, y, w: xb - xa, h: headH, fontSize: 11, bold: true, color: C.text2, align: "center", valign: "middle" });
    if (a === cur.getTime()) slide.addShape(pres.shapes.LINE, { objectName: nm("monat"), x: xa, y: y + headH, w: 0, h: totalH - headH, line: { color: "C9D1D9", width: 0.75 } });
    cur = nxt;
  }
  // Zonen (Hervorhebung)
  (cfg.zones || []).forEach((z) => {
    const zy = y + headH + z.row * rowH;
    box(slide, { x: X(z.s), y: zy + 0.03, w: Xe(z.e) - X(z.s), h: rowH * (z.rows || 1) - 0.06, fill: { color: z.color, transparency: z.tr || 80 }, line: { color: z.color, width: 1, dashType: "dash" } });
    if (z.text) tb(slide, z.text, { x: X(z.s), y: zy, w: Xe(z.e) - X(z.s), h: rowH, fontSize: z.fs || 11, bold: true, color: z.tcolor || C.text2, align: "center", valign: "middle" });
  });
  (cfg.vlines || []).forEach((v) => {
    const vx = X(v.d);
    slide.addShape(pres.shapes.LINE, { objectName: nm("stichtag"), x: vx, y: y + headH - 0.05, w: 0, h: totalH - headH + 0.1, line: { color: v.color || HEX.accent6, width: 1.75, dashType: v.dash || "solid" } });
    if (v.text) tb(slide, v.text, { x: vx - 1.0, y: y + totalH + 0.07, w: 2.0, h: 0.25, fontSize: 10, bold: true, color: v.tcolor || C.accent6, align: "center" });
  });
  rows.forEach((r, i) => {
    const ry = y + headH + i * rowH;
    tb(slide, r.label, { x: x + 0.1, y: ry, w: labelW - 0.15, h: rowH, fontSize: cfg.labelSize || 12, color: C.text1, valign: "middle", bold: !!r.bold });
    (r.bars || []).forEach((b) => {
      const n = b.lanes || 1, lane = b.lane || 0, pad = 0.06, lh = (rowH - 2 * pad) / n;
      const bx = X(b.s), bw = Math.max(0.06, Xe(b.e) - X(b.s));
      if (b.arrow) slide.addShape(pres.shapes.RIGHT_ARROW, { objectName: nm("pfeil"), x: bx, y: ry + pad + lane * lh + lh * 0.15, w: bw, h: lh * 0.7, fill: { color: b.color, transparency: 30 }, line: { type: "none" } });
      else box(slide, { x: bx, y: ry + pad + lane * lh + lh * 0.08, w: bw, h: lh * 0.84, fill: { color: b.color } });
      if (b.text) tb(slide, b.text, { x: b.tout ? Xe(b.e) + 0.06 : bx, y: ry + pad + lane * lh, w: b.tout ? (b.tw || 1.6) : bw, h: lh, fontSize: b.fs || 9.5, bold: true, color: b.tout ? (b.tcolor || C.text1) : C.background1, align: b.tout ? "left" : "center", valign: "middle" });
    });
    (r.ms || []).forEach((m) => {
      const mx = X(m.d) + (Xe(m.d) - X(m.d)) / 2, s = 0.2;
      slide.addShape(pres.shapes.DIAMOND, { objectName: nm("meilenstein"), x: mx - s / 2, y: ry + rowH / 2 - s / 2, w: s, h: s, fill: { color: m.color || HEX.accent6 }, line: { color: "FFFFFF", width: 0.75 } });
      if (m.text) tb(slide, m.text, { x: m.left ? mx - s / 2 - (m.tw || 1.4) - 0.05 : mx + s / 2 + 0.05, y: ry, w: m.tw || 1.4, h: rowH, fontSize: 10, bold: true, color: m.tcolor || C.accent6, valign: "middle", align: m.left ? "right" : "left" });
    });
  });
  return { X, Xe, bottom: y + totalH };
}
function legend(slide, x, y, items) {
  let cx = x;
  items.forEach(([label, color, w]) => {
    box(slide, { x: cx, y: y + 0.05, w: 0.22, h: 0.18, fill: { color } });
    tb(slide, label, { x: cx + 0.3, y, w, h: 0.28, fontSize: 11, valign: "middle" });
    cx += 0.3 + w + 0.25;
  });
}

(async () => {
  let S, s;
  // ================= 1 Titel
  S = "Einführung"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "TITEL", sectionTitle: S });
  s.addText("NEUBEWERTUNG KREUZCHENPLAN", { placeholder: "kicker" });
  s.addText("Ohne Schuss 309 reicht eine Schicht", { placeholder: "title" });
  s.addText("P800 mit Schüssen 301/302 im Vergleich ein- und zweischichtig: Ablauf, Engpass, Risiko, Empfehlung", { placeholder: "sub" });
  const kpiT = [["2 AT", "Zeitgewinn durch 2. Schicht"], ["20 : 25", "Risikopunkte 1- zu 2-schichtig"], ["2×", "Personaleinsatz bei 2 Schichten"]];
  kpiT.forEach(([n, l], i) => {
    const kx = 0.8 + i * 3.4;
    tb(s, n, { x: kx, y: 5.45, w: 3.1, h: 0.7, fontSize: 34, bold: true, color: C.background1 });
    tb(s, l, { x: kx, y: 6.15, w: 3.1, h: 0.4, fontSize: 13, color: C.background2 });
  });
  tb(s, "Projektmanagement · 31.10.2026", { x: 9.8, y: 6.9, w: 3.0, h: 0.3, fontSize: 11, color: C.accent5, align: "right" });
  s.addNotes("Neubewertung des Kreuzchenplans unter der Vorgabe: Schuss 309 wird vollständig aus dem Plan genommen. Da auf Station 18 immer zwei Schüsse paarweise verarbeitet werden, entfällt mit 309 auch der Partnerschuss 310. Betrachtet wird nur noch das Projekt P800 mit den Schüssen 301 und 302 sowie der laufende Testschuss 1.");

  // ================= 2 Executive Summary
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("KERNAUSSAGE", { placeholder: "kicker" });
  s.addText("Die zweite Schicht verliert ohne 309 ihren Zweck", { placeholder: "title" });
  const kpis = [
    [fa.FaClock, "2 AT", "früherer Start der Rundnaht 301+302", "20.01. statt 22.01.27", HEX.accent5],
    [fa.FaShieldHalved, "9 → 2 AT", "Puffer nach SLV-Ergebnis bis Spantschweißen St. 16", "einschichtig → zweischichtig", HEX.accent6],
    [fa.FaChartColumn, "20 : 25", "Risikopunkte nach Neubewertung", "vorher 28 : 30", HEX.accent3],
    [fa.FaUsers, "20 → 40", "Mitarbeiter in Halle 20+40 pro Tag", "ES, SB, QS, CNC je Schicht", HEX.accent1],
  ];
  for (let i = 0; i < kpis.length; i++) {
    const [ic, n, l, sub, col] = kpis[i];
    const ky = 1.55 + i * 1.32;
    card(s, 0.5, ky, 7.4, 1.15);
    await iconCircle(s, ic, 0.75, ky + 0.25, 0.65, col);
    tb(s, n, { x: 1.65, y: ky + 0.12, w: 2.2, h: 0.9, fontSize: 30, bold: true, color: C.text2, valign: "middle" });
    tb(s, [{ text: l, options: { bold: true, breakLine: true } }, { text: sub, options: { color: C.accent5, fontSize: 12 } }], { x: 3.9, y: ky + 0.12, w: 3.85, h: 0.9, fontSize: 14, valign: "middle" });
  }
  box(s, { x: 8.3, y: 1.55, w: 4.5, h: 5.11, fill: { color: C.text2 }, rounded: true, shadow: shadow() });
  await iconCircle(s, fa.FaCircleCheck, 8.6, 1.85, 0.7, HEX.accent3);
  tb(s, "Empfehlung", { x: 9.45, y: 1.85, w: 3.2, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
  tb(s, "Einschichtig fertigen", { x: 9.45, y: 2.18, w: 3.3, h: 0.45, fontSize: 22, bold: true, color: C.background1 });
  tb(s, [
    { text: "Schuss 302 bestimmt den Takt. Er läuft in beiden Szenarien fast gleich.", options: { bullet: true, breakLine: true } },
    { text: "Eine zweite Schicht macht nur 301 früher fertig. 301 wartet dann 11 AT auf 302.", options: { bullet: true, breakLine: true } },
    { text: "Einschichtig bleiben 2 Wochen Reserve für die WPS-Freigabe St. 16.", options: { bullet: true, breakLine: true } },
    { text: "Station 18 ist 8 Wochen frei und kann als Puffer dienen.", options: { bullet: true } },
  ], { x: 8.6, y: 3.0, w: 3.95, h: 3.5, fontSize: 14, color: C.background1, paraSpaceAfter: 9 });
  s.addNotes("Die Kernaussage auf einen Blick. Ohne 309/310 bleibt als einziger Vorteil der zweiten Schicht, dass die Rundnaht 301+302 zwei Arbeitstage früher starten kann (20.01. statt 22.01.2027). Dafür verdoppelt sich der tägliche Personaleinsatz (10 Schweißer, 7 Schiffbauer, 2 QS, 1 CNC je Schicht), die Unterstützung aus Kiel muss zwei Schichten abdecken, und der Puffer zwischen dem SLV-Ergebnis von Testschuss 1 (ca. 20.11.) und dem ersten Spantschweißen auf Station 16 schrumpft von 9 auf 2 Arbeitstage. Die Risikosumme nach Neubewertung: einschichtig 20, zweischichtig 25.");

  // ================= 3 Was ändert sich
  S = "Ausgangslage"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("AUSGANGSLAGE", { placeholder: "kicker" });
  s.addText("Mit 309 entfallen 310, ein Umbau und drei Risikopunkte", { placeholder: "title" });
  const colsCh = [
    ["Bisher: mit 309/310", HEX.accent5, fa.FaCircleXmark, [
      ["Vier Schüsse im zweischichtigen Plan", "309, 310, 301, 302"],
      ["Umbau Station 12 ab 05.10.26", "für Sektion 475/309"],
      ["Spante 309/310 aus Kiel", "Bedarf 12.10., Seetransport"],
      ["Station 18 doppelt belegt", "einschichtig blockiert 475 die Station für P800"],
      ["WPS St. 16 ab 16.11. nötig", "vor dem SLV-Ergebnis von Testschuss 1"],
    ]],
    ["Neu: nur P800", HEX.accent3, fa.FaCircleCheck, [
      ["Zwei Schüsse in beiden Szenarien", "301 und 302 plus Testschuss 1"],
      ["Umbau Station 12 ab 19.10.26", "direkt für P800/301"],
      ["Keine Materialbeistellung 475", "Risikopunkte Kiel-Material entfallen"],
      ["Station 18 frei für P800", "Konflikt mit 475 entfällt"],
      ["WPS St. 16 ab 23.11. bzw. 04.12.", "nach dem SLV-Ergebnis"],
    ]],
  ];
  for (let c = 0; c < 2; c++) {
    const [h, col, ic, items] = colsCh[c];
    const cx = 0.5 + c * 6.25;
    card(s, cx, 1.5, 6.05, 5.25, c ? C.background1 : C.background2);
    await iconCircle(s, ic, cx + 0.3, 1.72, 0.55, col);
    tb(s, h, { x: cx + 1.0, y: 1.72, w: 4.8, h: 0.55, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
    items.forEach(([a, b], i) => {
      const iy = 2.55 + i * 0.83;
      s.addShape(pres.shapes.LINE, { objectName: nm("trenner"), x: cx + 0.3, y: iy - 0.08, w: 5.45, h: 0, line: { color: "DDE2E7", width: 0.75 } });
      tb(s, [{ text: a, options: { bold: true, breakLine: true, color: c ? C.text1 : C.accent5 } }, { text: b, options: { fontSize: 12, color: C.accent5 } }], { x: cx + 0.3, y: iy, w: 5.45, h: 0.7, fontSize: 14.5 });
    });
  }
  s.addNotes("Was durch das Herausnehmen von 309 wegfällt: Schuss 310 als Partner auf Station 18, der vorgezogene Umbau von Station 12 (05.–09.10.), die Beistellung der Spante 475/309 und des Schusses 310 aus Kiel (Risikobewertung 3 und 1), das MPM-Risiko 'Station 18 durch Sektion 475 belegt' (einschichtig 3). Die P800-Termine für 301/302 werden unverändert aus den Blättern '1 Schichtig' und '2 Schichtig' übernommen.");

  // ================= 4 Gantt einschichtig
  S = "Ablauf"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("SZENARIO A · EINSCHICHTIG", { placeholder: "kicker" });
  s.addText("Einschichtig ist Station 18 acht Wochen frei", { placeholder: "title" });
  const L = 3;
  const rowsA = [
    { label: "St. 12 Montage", bars: [{ s: "19.10.26", e: "30.10.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "02.11.26", e: "05.11.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Vermessen", bars: [{ s: "02.11.26", e: "06.11.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "06.11.26", e: "12.11.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "St. 15 Längsnähte", bars: [{ s: "09.11.26", e: "01.12.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "16.11.26", e: "07.12.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "St. 16 Spante", bars: [{ s: "30.09.26", e: "08.10.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "03.12.26", e: "10.12.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "11.12.26", e: "16.12.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Putzen, VT, NDT", bars: [{ s: "12.10.26", e: "27.10.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "11.12.26", e: "05.01.27", color: HEX.accent2, lanes: L, lane: 1 }, { s: "18.12.26", e: "07.01.27", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Messen, Abgabe", bars: [{ s: "06.01.27", e: "18.01.27", color: HEX.accent2, lanes: L, lane: 1 }, { s: "12.01.27", e: "21.01.27", color: HEX.accent3, lanes: L, lane: 2 }], ms: [] },
    { label: "St. 18 Rundnaht", bars: [{ s: "29.10.26", e: "20.11.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "22.01.27", e: "02.02.27", color: HEX.dk2, lanes: 1, lane: 0, text: "301+302", fs: 10 }] },
  ];
  gantt(s, {
    x: 0.5, y: 1.45, w: 12.3, labelW: 1.95, start: "28.09.26", end: "05.02.27", rowH: 0.64, rows: rowsA,
    zones: [{ s: "23.11.26", e: "21.01.27", row: 6, color: HEX.accent3, tr: 82, text: "frei: 37 AT Puffer für Reparaturen oder Wiederholungsproben", tcolor: C.text2 }],
    vlines: [{ d: "20.11.26", color: HEX.accent6, dash: "dash", text: "SLV-Ergebnis ≈ 20.11." }, { d: "14.12.26", color: HEX.accent1, dash: "dash", text: "Spant 302 ab 14.12.", tcolor: C.accent1 }],
  });
  legend(s, 0.5, 6.62, [["Testschuss 1", HEX.accent5, 1.1], ["Schuss 301", HEX.accent2, 0.95], ["Schuss 302", HEX.accent3, 0.95], ["Paar 301+302", HEX.dk2, 1.1]]);
  tb(s, [{ text: "Abgabe 301 ", options: { bold: true } }, { text: "18.01.   " }, { text: "Abgabe 302 ", options: { bold: true } }, { text: "21.01.   " }, { text: "Rundnaht ab ", options: { bold: true } }, { text: "22.01.27" }], { x: 7.0, y: 6.62, w: 5.8, h: 0.28, fontSize: 13, color: C.text2, align: "right" });
  s.addNotes("Daten aus Blatt '1 Schichtig', Zeile 475/309 ist dort bereits leer. Station 18 ist nach der Probe-Rundnaht von Testschuss 1 (bis 20.11.) erst wieder ab 22.01.2027 durch das Paar 301+302 belegt. Das sind 37 Arbeitstage ohne Belegung (ohne Betriebsruhe 24.12.–01.01.). Diese Zeit ist Reserve, falls das SLV-Ergebnis von Testschuss 1 Nacharbeit oder Wiederholungsproben auf Station 18 erfordert.");

  // ================= 5 Gantt zweischichtig
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("SZENARIO B · ZWEISCHICHTIG OHNE 309", { placeholder: "kicker" });
  s.addText("Zweischichtig ist 301 früher fertig und wartet dann auf 302", { placeholder: "title" });
  const rowsB = [
    { label: "St. 12 Montage", bars: [{ s: "19.10.26", e: "28.10.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "02.11.26", e: "05.11.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Vermessen", bars: [{ s: "29.10.26", e: "04.11.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "06.11.26", e: "12.11.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "St. 15 Längsnähte", bars: [{ s: "05.11.26", e: "23.11.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "16.11.26", e: "07.12.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "St. 16 Spante", bars: [{ s: "30.09.26", e: "08.10.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "23.11.26", e: "27.11.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "11.12.26", e: "16.12.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Putzen, VT, NDT", bars: [{ s: "12.10.26", e: "27.10.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "30.11.26", e: "15.12.26", color: HEX.accent2, lanes: L, lane: 1 }, { s: "17.12.26", e: "23.12.26", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "Messen, Abgabe", bars: [{ s: "16.12.26", e: "04.01.27", color: HEX.accent2, lanes: L, lane: 1 }, { s: "18.01.27", e: "18.01.27", color: HEX.accent3, lanes: L, lane: 2 }] },
    { label: "St. 18 Rundnaht", bars: [{ s: "29.10.26", e: "20.11.26", color: HEX.accent5, lanes: L, lane: 0 }, { s: "20.01.27", e: "02.02.27", color: HEX.dk2, lanes: 1, lane: 0, text: "301+302", fs: 10 }] },
  ];
  gantt(s, {
    x: 0.5, y: 1.45, w: 12.3, labelW: 1.95, start: "28.09.26", end: "05.02.27", rowH: 0.64, rows: rowsB,
    zones: [{ s: "05.01.27", e: "19.01.27", row: 5, color: HEX.accent6, tr: 85, text: "11 AT", fs: 11, tcolor: C.accent6 }],
    vlines: [{ d: "20.11.26", color: HEX.accent6, dash: "dash", text: "SLV-Ergebnis ≈ 20.11." }, { d: "14.12.26", color: HEX.accent1, dash: "dash", text: "Spant 302 ab 14.12.", tcolor: C.accent1 }],
  });
  // Callout Liegezeit
  legend(s, 0.5, 6.62, [["Testschuss 1", HEX.accent5, 1.1], ["Schuss 301", HEX.accent2, 0.95], ["Schuss 302", HEX.accent3, 0.95], ["Liegezeit 301", HEX.accent6, 1.1]]);
  tb(s, [{ text: "Abgabe 301 ", options: { bold: true } }, { text: "04.01.   " }, { text: "Abgabe 302 ", options: { bold: true } }, { text: "18.01.   " }, { text: "Rundnaht ab ", options: { bold: true } }, { text: "20.01.27" }], { x: 7.0, y: 6.62, w: 5.8, h: 0.28, fontSize: 13, color: C.text2, align: "right" });
  s.addNotes("Daten aus Blatt '2 Schichtig' für 301/302, die Balken für 309/310 sind entfernt. 301 wird zwei Wochen früher abgegeben (04.01. statt 18.01.), kann aber erst mit 302 auf Station 18 – Abgabe 302 am 18.01. Zwischen Abgabe 301 und Rundnahtstart liegen 11 Arbeitstage Liegezeit. Annahme: Die Termine für 301/302 bleiben gegenüber dem Blatt unverändert. Ohne 309 könnte 301 theoretisch noch früher starten – das ändert den Paartermin aber nicht, weil 302 den Takt vorgibt.");

  // ================= 6 Engpass 302
  S = "Bewertung"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("ENGPASS", { placeholder: "kicker" });
  s.addText("Schuss 302 gibt den Takt vor, nicht die Schichtzahl", { placeholder: "title" });
  const ms302 = [["St. 12", "02.11.", "02.11."], ["St. 15", "16.11.", "16.11."], ["St. 16", "11.12.", "11.12."], ["Abgabe", "21.01.", "18.01."], ["Rundnaht", "22.01.", "20.01."]];
  const tx0 = 2.6, tx1 = 12.3, tl = (tx1 - tx0) / (ms302.length - 1);
  [["Einschichtig", 2.25], ["Zweischichtig", 3.65]].forEach(([lab, ly], r) => {
    tb(s, lab, { x: 0.5, y: ly - 0.2, w: 1.9, h: 0.4, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    s.addShape(pres.shapes.LINE, { objectName: nm("zeitstrahl"), x: tx0, y: ly, w: tx1 - tx0, h: 0, line: { color: "C9D1D9", width: 3 } });
    ms302.forEach((m, i) => {
      const mx = tx0 + i * tl, same = m[1] === m[2], col = same ? HEX.accent3 : HEX.accent1;
      s.addShape(pres.shapes.OVAL, { objectName: nm("punkt302"), x: mx - 0.16, y: ly - 0.16, w: 0.32, h: 0.32, fill: { color: col }, line: { color: "FFFFFF", width: 2 } });
      tb(s, m[1 + r], { x: mx - 0.8, y: ly + 0.22, w: 1.6, h: 0.35, fontSize: 16, bold: true, color: C.text2, align: "center" });
      if (r === 0) tb(s, m[0], { x: mx - 0.8, y: ly - 0.62, w: 1.6, h: 0.32, fontSize: 13, color: C.accent5, align: "center", bold: true });
    });
  });
  tb(s, "grün = gleicher Termin in beiden Szenarien · orange = Unterschied", { x: 2.6, y: 4.45, w: 9.7, h: 0.3, fontSize: 11, color: C.accent5, align: "center" });
  const why = [
    [fa.FaLock, "Fester Starttermin", "Spante für 302 dürfen laut Plan nicht vor dem 14.12. geschweißt werden. Das gilt in beiden Szenarien."],
    [fa.FaHourglassHalf, "Feste Wartezeiten", "Wartezeit vor NDT und die Betriebsruhe 24.12.–01.01. verkürzen sich durch eine zweite Schicht nicht."],
    [fa.FaLink, "Paarweise Rundnaht", "Station 18 verarbeitet 301 und 302 nur zusammen. Der spätere Schuss bestimmt den Start."],
  ];
  for (let i = 0; i < why.length; i++) {
    const [ic, h, t] = why[i];
    const wx = 0.5 + i * 4.15;
    card(s, wx, 4.95, 3.95, 1.8);
    await iconCircle(s, ic, wx + 0.25, 5.15, 0.5, HEX.dk2);
    tb(s, h, { x: wx + 0.9, y: 5.15, w: 2.9, h: 0.5, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    tb(s, t, { x: wx + 0.25, y: 5.75, w: 3.5, h: 0.95, fontSize: 12.5 });
  }
  s.addNotes("Schuss 302 hat in beiden Blättern bis einschließlich Station 16 identische Termine: Montage ab 02.11., Längsnähte ab 16.11., Spante ab 11.12. (einrichten) mit dem Vermerk 'Start nicht vor dem 14.12.'. Erst danach unterscheiden sich die Szenarien um drei Arbeitstage (Abgabe 18.01. statt 21.01.), der Rundnahtstart um zwei Arbeitstage. Die zweite Schicht beschleunigt also vor allem 301 – und der wartet dann auf 302.");

  // ================= 7 Szenarienvergleich
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("SZENARIENVERGLEICH", { placeholder: "kicker" });
  s.addText("Einschichtig gewinnt in fünf von sechs Kriterien", { placeholder: "title" });
  const cmp = [
    ["Start Rundnaht 301+302", "22.01.27", "20.01.27", 2, "2 AT früher"],
    ["Liegezeit 301 vor Rundnaht", "3 AT", "11 AT", 1, "weniger Zwischenlager"],
    ["Puffer WPS-Freigabe St. 16", "9 AT", "2 AT", 1, "Reserve bei Nacharbeit"],
    ["Personal Halle pro Tag", "20 MA", "40 MA", 1, "halber Einsatz"],
    ["Anlagenfahrer Kiel", "1 Schicht", "2 Schichten", 1, "Risiko 1 statt 3"],
    ["Risikosumme neu", "20", "25", 1, "5 Punkte weniger"],
  ];
  const hdrY = 1.5;
  [["Kriterium", 0.5, 3.6], ["Einschichtig", 4.2, 2.6], ["Zweischichtig", 6.9, 2.6], ["Vorteil", 9.6, 3.2]].forEach(([t, x, w]) => {
    tb(s, t, { x: x + 0.15, y: hdrY, w, h: 0.4, fontSize: 13, bold: true, color: C.accent5, valign: "middle" });
  });
  for (let i = 0; i < cmp.length; i++) {
    const [k, a, b, win, why] = cmp[i];
    const ry = 1.95 + i * 0.79;
    box(s, { x: 0.5, y: ry, w: 12.3, h: 0.68, fill: { color: i % 2 ? C.background1 : C.background2 }, rounded: true });
    tb(s, k, { x: 0.65, y: ry, w: 3.5, h: 0.68, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    [[a, 4.2, win === 1], [b, 6.9, win === 2]].forEach(([v, x, good]) => {
      if (good) box(s, { x: x + 0.08, y: ry + 0.1, w: 2.45, h: 0.48, fill: { color: HEX.accent3, transparency: 80 }, rounded: true, line: { color: HEX.accent3, width: 1 } });
      tb(s, v, { x: x + 0.2, y: ry, w: 2.3, h: 0.68, fontSize: 16, bold: good, color: good ? C.text2 : C.accent5, valign: "middle" });
    });
    await iconCircle(s, win === 1 ? fa.FaCheck : fa.FaArrowRight, 9.75, ry + 0.16, 0.36, win === 1 ? HEX.accent3 : HEX.accent5);
    tb(s, [{ text: win === 1 ? "Einschichtig: " : "Zweischichtig: ", options: { bold: true } }, { text: why }], { x: 10.25, y: ry, w: 2.5, h: 0.68, fontSize: 12, valign: "middle" });
  }
  s.addNotes("Gegenüberstellung der beiden Szenarien ohne 309/310. MA = Mitarbeiter: laut Risikoblatt je Schicht 10 Schweißer (ES), 7 Schiffbauer (SB), 2 QS und 1 CNC-Bediener, also 20 pro Schicht, plus NDT. Einziger Vorteil der zweiten Schicht ist der um zwei Arbeitstage frühere Start der Rundnaht 301+302.");

  // ================= 8 WPS-Puffer
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("KRITISCHER PFAD", { placeholder: "kicker" });
  s.addText("Einschichtig bleiben 9 AT Reserve nach dem SLV-Ergebnis", { placeholder: "title" });
  gantt(s, {
    x: 0.5, y: 1.45, w: 12.3, labelW: 2.6, start: "26.10.26", end: "11.12.26", rowH: 0.72, labelSize: 13,
    rows: [
      { label: "Proben Testschuss 1", bars: [{ s: "28.10.26", e: "29.10.26", color: HEX.accent5 }], ms: [{ d: "30.10.26", text: "Versand SLV 30.10.", tw: 1.6 }] },
      { label: "SLV Rostock", bars: [{ s: "31.10.26", e: "19.11.26", color: HEX.accent6, arrow: true, text: "Prüfung ≈ 3 Wochen", fs: 10 }], ms: [{ d: "20.11.26", text: "Ergebnis ≈ 20.11.", tw: 1.6 }] },
      { label: "Zweischichtig: 301 St. 16", bars: [{ s: "23.11.26", e: "27.11.26", color: HEX.accent2, text: "Spante", fs: 10 }] },
      { label: "Einschichtig: 301 St. 16", bars: [{ s: "03.12.26", e: "10.12.26", color: HEX.accent2, text: "Spante", fs: 10 }] },
    ],
    zones: [{ s: "23.11.26", e: "24.11.26", row: 2, color: HEX.accent6, tr: 75 }, { s: "23.11.26", e: "03.12.26", row: 3, color: HEX.accent3, tr: 80, text: "9 AT Reserve", tcolor: C.text2, fs: 12 }],
    vlines: [{ d: "20.11.26", color: HEX.accent6, dash: "dash", text: "SLV-Ergebnis" }],
  });
  const wp = [
    [fa.FaTriangleExclamation, HEX.accent6, "Zweischichtig: 2 AT", "Einrichten St. 16 am Montag nach dem Ergebnis. Fällt eine Probe durch, steht Station 16 still, und die gesamte Kette verschiebt sich."],
    [fa.FaShieldHalved, HEX.accent3, "Einschichtig: 9 AT", "Zeit, um Befunde zu bewerten, die WPS zu erstellen und freizugeben, ohne den Ablauf von 301 zu verschieben."],
  ];
  for (let i = 0; i < wp.length; i++) {
    const [ic, col, h, t] = wp[i];
    const bx = 0.5 + i * 6.25;
    card(s, bx, 5.0, 6.05, 1.75);
    await iconCircle(s, ic, bx + 0.25, 5.2, 0.55, col);
    tb(s, h, { x: bx + 1.0, y: 5.2, w: 4.8, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    tb(s, t, { x: bx + 0.25, y: 5.85, w: 5.6, h: 0.85, fontSize: 13 });
  }
  s.addNotes("Mit 309 wurde die WPS 121 für Station 16 schon ab 16.11. gebraucht, also vor dem SLV-Ergebnis. Ohne 309 ist der erste Serienschuss auf Station 16 der Schuss 301: zweischichtig Einrichten ab 23.11., Spante ab 25.11.; einschichtig Einrichten 03.12., Spante ab 04.12. Gerechnet ab dem erwarteten Ergebnis am Freitag, 20.11., bleiben zweischichtig 2 und einschichtig 9 Arbeitstage bis zum Spantschweißen. Eine Wiederholung der Verfahrensprüfung würde bei 5 Wochen SLV-Laufzeit in beiden Fällen nicht in den Puffer passen – einschichtig gibt es aber Zeit für Bewertung, Klärung und Reparaturfreigabe.");

  // ================= 9 Risiko neu
  S = "Risiko"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("RISIKO-NEUBEWERTUNG", { placeholder: "kicker" });
  s.addText("Ohne 309 sinkt das Risiko einschichtig am stärksten", { placeholder: "title" });
  // Hantel-Diagramm alt -> neu
  card(s, 0.5, 1.5, 5.4, 5.25);
  tb(s, "Risikosumme alt → neu", { x: 0.8, y: 1.7, w: 4.8, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  const sx0 = 1.0, sx1 = 5.4, smin = 15, smax = 32;
  const SX = (v) => sx0 + ((v - smin) / (smax - smin)) * (sx1 - sx0);
  [15, 20, 25, 30].forEach((v) => {
    s.addShape(pres.shapes.LINE, { objectName: nm("raster"), x: SX(v), y: 2.35, w: 0, h: 2.6, line: { color: "E3E7EB", width: 0.75 } });
    tb(s, String(v), { x: SX(v) - 0.3, y: 5.0, w: 0.6, h: 0.25, fontSize: 10, color: C.accent5, align: "center" });
  });
  [["Einschichtig", 28, 20, 2.95], ["Zweischichtig", 30, 25, 4.15]].forEach(([lab, a, b, yy]) => {
    tb(s, lab, { x: 1.0, y: yy - 0.62, w: 2.5, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
    s.addShape(pres.shapes.LINE, { objectName: nm("hantel"), x: SX(b), y: yy, w: SX(a) - SX(b), h: 0, line: { color: "C9D1D9", width: 4 } });
    s.addShape(pres.shapes.OVAL, { objectName: nm("alt"), x: SX(a) - 0.17, y: yy - 0.17, w: 0.34, h: 0.34, fill: { color: HEX.accent5 }, line: { color: "FFFFFF", width: 1.5 } });
    s.addShape(pres.shapes.OVAL, { objectName: nm("neu"), x: SX(b) - 0.2, y: yy - 0.2, w: 0.4, h: 0.4, fill: { color: lab === "Einschichtig" ? HEX.accent3 : HEX.accent1 }, line: { color: "FFFFFF", width: 1.5 } });
    tb(s, String(a), { x: SX(a) - 0.3, y: yy + 0.22, w: 0.6, h: 0.3, fontSize: 12, color: C.accent5, align: "center" });
    tb(s, String(b), { x: SX(b) - 0.35, y: yy + 0.24, w: 0.7, h: 0.35, fontSize: 18, bold: true, color: C.text2, align: "center" });
  });
  tb(s, "Grau = alte Bewertung mit 309/310, farbig = neue Bewertung nur P800. Skala je Punkt 0 bis 3.", { x: 0.8, y: 5.45, w: 4.85, h: 0.9, fontSize: 12, color: C.accent5 });
  // Änderungstabelle
  tb(s, "Geänderte Bewertungen", { x: 6.3, y: 1.5, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  const hdr = ["Risiko", "1-schichtig", "2-schichtig", "Begründung"];
  const chg = [
    ["Material Spante 475/309", "3 → –", "3 → –", "entfällt mit 309"],
    ["Material Schuss 310", "1 → –", "1 → –", "entfällt mit 309"],
    ["MPM: Station 18 belegt", "3 → 0", "1 → 0", "kein Konflikt mit 475"],
    ["WPS 121 Station 16", "3 → 2", "3 → 3", "9 AT bzw. 2 AT Puffer"],
  ];
  const trows = [hdr.map((h) => ({ text: h, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 12 } }))].concat(
    chg.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 12.5, bold: j === 0 || ((j === 1 || j === 2) && c.includes("0")), color: j === 3 ? C.accent5 : C.text1, fill: { color: i % 2 ? C.background1 : C.background2 }, align: j === 1 || j === 2 ? "center" : "left" } }))));
  s.addTable(trows, { x: 6.3, y: 1.95, w: 6.5, colW: [2.15, 1.1, 1.1, 2.15], rowH: 0.48, valign: "middle", border: { type: "solid", pt: 0.5, color: "DDE2E7" }, objectName: nm("tab") });
  tb(s, "Unverändert", { x: 6.3, y: 4.55, w: 6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
  tb(s, [
    { text: "Personal und Kiel-Unterstützung: zweischichtig höher (Anlage 1 zu 3, MSG 1 zu 2, UP 2 zu 3).", options: { bullet: true, breakLine: true } },
    { text: "Gleich in beiden: WPS St. 15 (2), Messtrupp (2), Zusatzwerkstoffe (3 × 2), Wendeschneidplatten (2).", options: { bullet: true } },
  ], { x: 6.3, y: 4.95, w: 6.5, h: 1.6, fontSize: 13, paraSpaceAfter: 6 });
  s.addNotes("Grundlage ist das Blatt 'Auswertung Risiko' (Summen 28 und 30). Entfernt wurden die beiden Positionen der Materialbeistellung aus Kiel für 475/309 (3/3) und Schuss 310 (1/1). Das MPM-Risiko bezog sich auf die Belegung von Station 18 durch Sektion 475 und entfällt (0/0). Die WPS 121 für Station 16 wurde nach dem Puffer bis zum ersten Spantschweißen neu bewertet: einschichtig 2 (9 AT Reserve), zweischichtig weiterhin 3 (2 AT Reserve). Ergebnis: einschichtig 28 → 20, zweischichtig 30 → 25. Die Neubewertung der WPS St. 16 ist eine Einschätzung des Projektmanagements und sollte mit Schweißaufsicht und QS bestätigt werden.");

  // ================= 10 Risikoprofil neu (Chart)
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("RISIKOPROFIL NEU", { placeholder: "kicker" });
  s.addText("Die zweite Schicht verlagert das Risiko auf Personal und Kiel", { placeholder: "title" });
  const rk = [
    ["WPS 138P/121 (St. 15)", 2, 2], ["WPS 121 (St. 16)", 2, 3], ["WPS 121 (St. 18)", 1, 1],
    ["Anlage einfahren, Hilfe Kiel", 1, 3], ["Schweißer MSG", 1, 2], ["Messtrupp", 2, 2],
    ["Schweißer UP", 2, 3], ["NDT", 1, 1], ["Zusatzwerkstoffe (3 Sorten)", 6, 6], ["Wendeschneidplatten", 2, 2],
  ];
  const lr = rk.map((r) => r[0]).reverse();
  s.addChart(pres.charts.BAR, [
    { name: "Einschichtig", labels: lr, values: rk.map((r) => r[1]).reverse() },
    { name: "Zweischichtig", labels: lr, values: rk.map((r) => r[2]).reverse() },
  ], {
    x: 0.5, y: 1.45, w: 8.0, h: 5.4, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 55,
    chartColors: [HEX.accent3, HEX.accent1],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 12, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
    valAxisHidden: true, valAxisMaxVal: 7, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.dk1,
    objectName: nm("risikochart"),
  });
  const hot = [
    [fa.FaUserGear, "Anlagenfahrer Kiel", "1 → 3", "Unterstützung muss zwei Schichten abdecken."],
    [fa.FaUsers, "Geprüfte Schweißer", "+2", "MSG 1 → 2 und UP 2 → 3: doppelte Zahl mit gültiger Prüfung."],
    [fa.FaShieldHalved, "WPS Station 16", "2 → 3", "Nur 2 AT Puffer nach SLV-Ergebnis."],
  ];
  for (let i = 0; i < hot.length; i++) {
    const [ic, h, n, t] = hot[i];
    const hy = 1.5 + i * 1.78;
    card(s, 8.85, hy, 3.95, 1.6);
    await iconCircle(s, ic, 9.05, hy + 0.2, 0.5, HEX.accent1);
    tb(s, h, { x: 9.7, y: hy + 0.2, w: 2.0, h: 0.5, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    tb(s, n, { x: 11.6, y: hy + 0.2, w: 1.05, h: 0.5, fontSize: 18, bold: true, color: C.accent1, align: "right", valign: "middle" });
    tb(s, t, { x: 9.05, y: hy + 0.82, w: 3.6, h: 0.7, fontSize: 12.5 });
  }
  s.addNotes("Neues Risikoprofil ohne 309/310 je Position (0 kein, 1 wenig, 2 mittel, 3 hoch). Die drei Zusatzwerkstoffe Megafil 940 M, S3 NiMo 3 mm und 421 TT sind mit je 2 bewertet und hier zusammengefasst. Gas (Erdgas, M21, Acetylen/Sauerstoff) ist mit 0 bewertet und nicht dargestellt. Rechts die drei Positionen, in denen die zweite Schicht schlechter abschneidet.");

  // ================= 11 Bleibende Risiken
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("UNABHÄNGIG VON 309", { placeholder: "kicker" });
  s.addText("Vier Fertigungsstarts liegen weiterhin vor dem SLV-Ergebnis", { placeholder: "title" });
  const gaps = [["Station 2", "Start P488 Gurte", "25.02.27", "26.02.27", "1 AT"], ["Station 7", "Start P488 Stege", "25.02.27", "19.03.27", "3 Wo"], ["Station 15", "Start Schuss 601", "21.05.27", "28.05.27", "1 Wo"], ["Station 16", "Start 201", "01.06.27", "25.06.27", "3,5 Wo"]];
  for (let i = 0; i < gaps.length; i++) {
    const [st, what, a, b, g] = gaps[i];
    const gx = 0.5 + i * 3.1;
    card(s, gx, 1.5, 2.9, 3.4);
    await iconCircle(s, fa.FaFlask, gx + 0.25, 1.72, 0.55, HEX.dk2);
    tb(s, st, { x: gx + 0.95, y: 1.72, w: 1.85, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    tb(s, g, { x: gx + 0.25, y: 2.45, w: 2.5, h: 0.85, fontSize: 38, bold: true, color: C.accent6 });
    tb(s, "Lücke", { x: gx + 0.25, y: 3.25, w: 2.5, h: 0.3, fontSize: 11, color: C.accent5 });
    tb(s, [{ text: what + ": ", options: { bold: true } }, { text: a, options: { breakLine: true } }, { text: "SLV-Ergebnis AP: ", options: { bold: true } }, { text: b }], { x: gx + 0.25, y: 3.65, w: 2.55, h: 1.1, fontSize: 12.5 });
  }
  box(s, { x: 0.5, y: 5.15, w: 12.3, h: 1.6, fill: { color: C.background2 }, rounded: true });
  await iconCircle(s, fa.FaLightbulb, 0.75, 5.45, 0.55, HEX.accent1);
  tb(s, "Hebel", { x: 1.5, y: 5.35, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  tb(s, "Die SLV braucht für jede Prüfung 5 Wochen. Die Arbeitsproben müssen deshalb mindestens 5 Wochen vor Fertigungsstart geschweißt sein, sonst braucht es eine Vorabfreigabe auf Basis der Verfahrensprüfung. Diese vier Lücken stehen bisher nicht in der Risikoliste.", { x: 1.5, y: 5.72, w: 11.1, h: 0.95, fontSize: 13.5 });
  s.addNotes("Diese Befunde betreffen die Qualifizierung der Stationen für die Folgeprojekte (P488, 201, 601) und sind vom Herausnehmen von 309 nicht berührt. Quelle: Kommentare 'Start P488 Gurtfertigung/Stegfertigung' (25.02.27), Blatt 'Übersicht Schweißproben' Spalte 'Fertigungsstart Projekt', sowie die Pfeile vom Probenversand bis zum Ergebnis der SLV. Durch einschichtiges Fertigen werden keine Kapazitäten für die Schweißproben gebunden – ein weiteres Argument für eine Schicht.");

  // ================= 12 Empfehlung
  S = "Empfehlung"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT_DUNKEL", sectionTitle: S });
  s.addText("EMPFEHLUNG", { placeholder: "kicker" });
  s.addText("P800 einschichtig fertigen und die Reserve gezielt nutzen", { placeholder: "title" });
  const steps = [
    [fa.FaCalendarCheck, "Bis KW 45", "Entscheidung einschichtig im Projektteam bestätigen und 309/310 aus dem Kreuzchenplan entfernen."],
    [fa.FaFlask, "Bis 20.11.", "SLV-Ergebnis Testschuss 1 verbindlich terminieren, WPS 121 St. 16 bis 03.12. freigeben."],
    [fa.FaWarehouse, "Ab 23.11.", "Freie Station 18 als Reserve für Nacharbeit, Wiederholungsproben oder vorgezogene Proben nutzen."],
    [fa.FaClipboardList, "Bis Dez 26", "Arbeitsproben St. 2, 7, 15, 16 so vorziehen, dass die SLV-Ergebnisse vor Fertigungsstart vorliegen."],
  ];
  for (let i = 0; i < steps.length; i++) {
    const [ic, w, t] = steps[i];
    const sy = 1.55 + i * 1.12;
    await iconCircle(s, ic, 0.5, sy + 0.08, 0.7, HEX.accent1);
    tb(s, w, { x: 1.45, y: sy, w: 1.9, h: 0.85, fontSize: 17, bold: true, color: C.accent1, valign: "middle" });
    tb(s, t, { x: 3.3, y: sy, w: 5.0, h: 0.85, fontSize: 14.5, color: C.background1, valign: "middle" });
  }
  box(s, { x: 8.75, y: 1.55, w: 4.05, h: 5.2, fill: { color: "2E3E4F" }, rounded: true });
  tb(s, "Wann doch zwei Schichten?", { x: 9.05, y: 1.8, w: 3.5, h: 0.4, fontSize: 16, bold: true, color: C.background1 });
  tb(s, [
    { text: "Der Kunde fordert die Rundnaht 301+302 vor dem 22.01.27.", options: { bullet: true, breakLine: true } },
    { text: "Ein neuer Auftrag soll die Stationen ab Mitte Januar belegen.", options: { bullet: true, breakLine: true } },
    { text: "Die Freigabe des Spantstarts 302 kommt deutlich vor dem 14.12.", options: { bullet: true } },
  ], { x: 9.05, y: 2.35, w: 3.5, h: 2.6, fontSize: 14, color: C.background2, paraSpaceAfter: 10 });
  tb(s, "Dann neu rechnen. Ohne diese Auslöser bringt die 2. Schicht nur 2 AT.", { x: 9.05, y: 5.4, w: 3.5, h: 1.1, fontSize: 14, bold: true, color: C.accent1 });
  s.addNotes("Empfehlung des Projektmanagements auf Basis der Neubewertung. Die Bedingungen rechts nennen die Fälle, in denen die Bewertung kippen würde. Hinweis: Heute ist der 31.10.2026 – Montage 301 läuft laut Plan bereits seit 19.10./26.10.; der Ist-Stand ist vor der Entscheidung abzugleichen.");

  // ================= 13 Anhang Annahmen
  S = "Anhang"; pres.addSection({ title: S });
  s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
  s.addText("ANHANG", { placeholder: "kicker" });
  s.addText("Annahmen und Datengrundlage der Neubewertung", { placeholder: "title" });
  const ann = [
    ["309 und 310 entfallen gemeinsam", "Station 18 verarbeitet immer zwei Schüsse. Ohne 309 hat 310 keinen Partner im Plan."],
    ["Termine 301/302 unverändert", "Übernommen aus den Blättern „1 Schichtig“ und „2 Schichtig“. Ein früherer Start von 301 ändert den Paartermin nicht."],
    ["SLV-Ergebnis Testschuss 1 ≈ 20.11.26", "Laut Pfeil und Kommentar im Blatt „1 Schichtig“. Versand der Proben am 30.10."],
    ["Arbeitstage ohne Betriebsruhe", "Wochenenden und 24.12.–01.01. sind nicht mitgezählt (AT = Arbeitstag)."],
    ["Risikoskala unverändert", "0 kein, 1 wenig, 2 mittel, 3 hoch. Nur vier Positionen wurden geändert (Folie 9)."],
    ["Personal je Schicht", "10 ES, 7 SB, 2 QS, 1 CNC laut Risikoblatt, NDT separat."],
  ];
  ann.forEach(([h, t], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const ax = 0.5 + col * 6.25, ay = 1.55 + row * 1.75;
    card(s, ax, ay, 6.05, 1.55);
    tb(s, String(i + 1).padStart(2, "0"), { x: ax + 0.25, y: ay + 0.2, w: 0.8, h: 0.6, fontSize: 26, bold: true, color: C.accent1 });
    tb(s, h, { x: ax + 1.1, y: ay + 0.2, w: 4.75, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    tb(s, t, { x: ax + 1.1, y: ay + 0.65, w: 4.75, h: 0.8, fontSize: 12.5 });
  });
  s.addNotes("Transparenz über die Annahmen. Quelle ist ausschließlich die Excel-Datei 'Kreutzchenplan_Schweißproben_PHPL-Wismar_Ablauf.xlsx'.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok");
})();
