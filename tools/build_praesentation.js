const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/10f6f57f-31e2-469d-b54e-5e73bdc38b45_cb9ce55c-1b2b-43d1-a3b0-bcd2798483a7/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Werft Stahl",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B232C", lt1: "FFFFFF", dk2: "2B3A4A", lt2: "EEF1F4",
    accent1: "E8641B", // Schweiß-Orange (Hervorhebung)
    accent2: "2E9BD6", // Schuss 301
    accent3: "5FA84A", // Schuss 302
    accent4: "8E63B5", // Schuss 309/310
    accent5: "8A96A3", // Testschuss 1
    accent6: "C8102E", // kritisch / Termin
    hlink: "2E9BD6", folHlink: "8E63B5",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Kreuzchenplan Schweißproben – Ablauf & Abhängigkeiten";
pres.author = "Projektmanagement";
const C = pres.SchemeColor;
const W = 13.333, H = 7.5;

const FOOT = "Kreuzchenplan Schweißproben PHPL Wismar · Auswertung Projektmanagement · Stand 31.10.2026";

pres.defineSlideMaster({
  title: "TITEL_DUNKEL",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 44, bold: true, color: C.background1, valign: "bottom", fontFace: THEME.headFontFace }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.95, w: 11.7, h: 1.2, fontSize: 20, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "KAPITEL",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 2.2, y: 2.6, w: 10.3, h: 1.2, fontSize: 40, bold: true, color: C.background1, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 2.2, y: 3.9, w: 10.3, h: 1.0, fontSize: 18, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "INHALT",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.3, h: 0.8, fontSize: 25, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
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
  slide.addShape(o.rounded ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE,
    Object.assign({ objectName: nm("shape"), line: { type: "none" } }, o.rounded ? { rectRadius: 0.08 } : {}, o, { rounded: undefined }));
}
function circleNum(slide, x, y, d, label, color) {
  slide.addShape(pres.shapes.OVAL, { objectName: nm("kreis"), x, y, w: d, h: d, fill: { color: color || C.accent1 }, line: { type: "none" } });
  tb(slide, String(label), { x, y, w: d, h: d, fontSize: d > 0.5 ? 18 : 12, bold: true, color: C.background1, align: "center", valign: "middle" });
}
function sec(title) { pres.addSection({ title }); return title; }

// ---------- Datumshilfen ----------
const D = (s) => { const [d, m, y] = s.split("."); return new Date(Date.UTC(2000 + +y, +m - 1, +d)); };
const DAY = 86400000;

// Gantt-Zeichner: rows = [{label, sub?, bars:[{s,e,color,text?}]}], milestones = [{d,label,color}]
function gantt(slide, cfg) {
  const { x, y, w, labelW, start, end, rowH, rows } = cfg;
  const gx = x + labelW, gw = w - labelW;
  const t0 = D(start).getTime(), t1 = D(end).getTime() + DAY;
  const X = (s) => gx + ((D(s).getTime() - t0) / (t1 - t0)) * gw;
  const Xe = (s) => gx + ((D(s).getTime() + DAY - t0) / (t1 - t0)) * gw;
  const headH = 0.32;
  // Monatskopf
  let cur = new Date(t0);
  cur = new Date(Date.UTC(cur.getUTCFullYear(), cur.getUTCMonth(), 1));
  const months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
  const totalH = headH + rows.length * rowH;
  while (cur.getTime() < t1) {
    const nxt = new Date(Date.UTC(cur.getUTCFullYear(), cur.getUTCMonth() + 1, 1));
    const a = Math.max(cur.getTime(), t0), b = Math.min(nxt.getTime(), t1);
    const xa = gx + ((a - t0) / (t1 - t0)) * gw, xb = gx + ((b - t0) / (t1 - t0)) * gw;
    const idx = cur.getUTCMonth();
    if (xb - xa > 0.6) tb(slide, `${months[idx]} ${String(cur.getUTCFullYear()).slice(2)}`, { x: xa, y, w: xb - xa, h: headH, fontSize: 11, bold: true, color: C.text2, align: "center", valign: "middle" });
    if (a === cur.getTime()) slide.addShape(pres.shapes.LINE, { objectName: nm("monatslinie"), x: xa, y: y + headH, w: 0, h: totalH - headH, line: { color: "C9D1D9", width: 0.75 } });
    cur = nxt;
  }
  // Zeilen: Hintergrund + Beschriftung
  rows.forEach((r, i) => {
    const ry = y + headH + i * rowH;
    if (i % 2 === 0) box(slide, { x, y: ry, w, h: rowH, fill: { color: C.background2 } });
  });
  // vertikale Stichtage (hinter den Balken)
  (cfg.vlines || []).forEach((v) => {
    const vx = X(v.d);
    slide.addShape(pres.shapes.LINE, { objectName: nm("stichtag"), x: vx, y: y + headH - 0.05, w: 0, h: totalH - headH + 0.1, line: { color: v.color || HEX.accent6, width: 1.75, dashType: v.dash || "solid" } });
    if (v.text) tb(slide, v.text, { x: vx - 0.9, y: y + totalH + 0.06, w: 1.8, h: 0.25, fontSize: 10, bold: true, color: v.tcolor || C.accent6, align: "center" });
  });
  rows.forEach((r, i) => {
    const ry = y + headH + i * rowH;
    tb(slide, r.sub ? [{ text: r.label, options: { bold: true, breakLine: true } }, { text: r.sub, options: { fontSize: 10, color: C.accent5 } }] : r.label,
      { x: x + 0.08, y: ry, w: labelW - 0.12, h: rowH, fontSize: cfg.labelSize || 12, color: C.text1, valign: "middle", bold: !r.sub && !!r.bold });
    (r.bars || []).forEach((b) => {
      const n = b.lanes || 1, lane = b.lane || 0;
      const pad = 0.05, lh = (rowH - 2 * pad) / n;
      const bx = X(b.s), bw = Math.max(0.05, Xe(b.e) - X(b.s));
      if (b.arrow) {
        slide.addShape(pres.shapes.RIGHT_ARROW, { objectName: nm("pfeil"), x: bx, y: ry + pad + lane * lh + lh * 0.18, w: bw, h: lh * 0.64, fill: { color: b.color, transparency: 35 }, line: { type: "none" } });
      } else {
        box(slide, { x: bx, y: ry + pad + lane * lh, w: bw, h: lh, fill: { color: b.color } });
      }
      if (b.text) {
        const tw = Math.max(bw, b.tw || 0);
        tb(slide, b.text, { x: b.tout ? Xe(b.e) + 0.05 : bx, y: ry + pad + lane * lh, w: b.tout ? (b.tw || 1.5) : tw, h: lh, fontSize: b.fs || 9, color: b.tout ? (b.tcolor || C.text1) : (b.tcolor || C.background1), bold: true, align: b.tout ? "left" : "center", valign: "middle" });
      }
    });
    (r.ms || []).forEach((m) => {
      const mx = X(m.d) + (Xe(m.d) - X(m.d)) / 2, s = 0.2;
      slide.addShape(pres.shapes.DIAMOND, { objectName: nm("meilenstein"), x: mx - s / 2, y: ry + rowH / 2 - s / 2, w: s, h: s, fill: { color: m.color || C.accent6 }, line: { color: "FFFFFF", width: 0.75 } });
      if (m.text) tb(slide, m.text, { x: m.left ? mx - s / 2 - (m.tw || 1.4) - 0.04 : mx + s / 2 + 0.04, y: ry, w: m.tw || 1.4, h: rowH, fontSize: 9, bold: true, color: m.tcolor || C.accent6, valign: "middle", align: m.left ? "right" : "left" });
    });
  });
  return { X, Xe, bottom: y + totalH };
}

function legend(slide, x, y, items, gap) {
  let cx = x;
  items.forEach(([label, color, wdt]) => {
    box(slide, { x: cx, y: y + 0.04, w: 0.22, h: 0.18, fill: { color } });
    tb(slide, label, { x: cx + 0.3, y, w: wdt || 1.5, h: 0.26, fontSize: 11, color: C.text1, valign: "middle" });
    cx += 0.3 + (wdt || 1.5) + (gap || 0.2);
  });
}

// =====================================================================
// 1 Titel
let S = sec("Einführung");
let s = pres.addSlide({ masterName: "TITEL_DUNKEL", sectionTitle: S });
s.addText("Kreuzchenplan Schweißproben", { placeholder: "title" });
s.addText("Ablauf Testschuss 1, Schüsse 301/302 (P800) und 309/310 (475) – Szenarien, Abhängigkeiten, Risiken", { placeholder: "body" });
// Motiv: Kette aus Stationen
const chain = ["St. 12", "St. 15", "St. 16", "St. 18", "SLV"];
chain.forEach((t, i) => {
  const cx = 0.8 + i * 1.55;
  s.addShape(pres.shapes.OVAL, { objectName: nm("titelkreis"), x: cx, y: 5.55, w: 0.9, h: 0.9, fill: { color: i === 4 ? C.accent6 : C.accent1 }, line: { type: "none" } });
  tb(s, t, { x: cx, y: 5.55, w: 0.9, h: 0.9, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
  if (i < chain.length - 1) s.addShape(pres.shapes.LINE, { objectName: nm("titellinie"), x: cx + 0.95, y: 6.0, w: 0.55, h: 0, line: { color: HEX.lt2, width: 1.5, endArrowType: "triangle" } });
});
tb(s, "Projektmanagement · Stand 31.10.2026", { x: 8.8, y: 6.75, w: 3.7, h: 0.3, fontSize: 12, color: C.background2, align: "right" });
s.addNotes("Diese Präsentation erklärt den Kreuzchenplan 'Schweißproben PHPL Wismar – Ablauf'. Die Excel-Datei besteht aus vier Blättern: '1 Schichtig', '2 Schichtig', 'Übersicht Schweißproben 1.3964' und 'Auswertung Risiko'. Ziel des Plans: zeigen, wie Testschuss 1 und die Schweißverfahrensprüfungen so laufen, dass die Sektion 475/309 und die P800-Schüsse 301/302 gefertigt werden können.");

// 2 Worum geht es
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Der Plan beantwortet eine Frage: Reicht eine Schicht?", { placeholder: "title" });
tb(s, "Der Kreuzchenplan legt die Fertigung mehrerer Rohrschüsse aus Werkstoff 1.3964 tageweise über die Stationen der Halle. Gleichzeitig müssen die Schweißverfahren (WPS) über Proben bei der SLV Rostock qualifiziert werden.",
  { x: 0.5, y: 1.25, w: 12.3, h: 0.75, fontSize: 16, color: C.text2 });
const cards = [
  ["Testschuss 1", "Läuft bereits (Sep–Nov 26). Liefert die Proben für die Verfahrensprüfung von Station 16 und 18.", C.accent5],
  ["Schüsse 301 / 302", "Projekt P800. Laufen in beiden Szenarien über St. 12 → 15 → 16 → 18.", C.accent2],
  ["Schüsse 309 / 310", "Sektion 475 (Boot 1). Nur im 2-schichtigen Szenario im Plan eingetaktet.", C.accent4],
  ["Schweißproben 1.3964", "VP und AP je Station (2, 5, 7, 9, 15, 16, Rahmenspant) mit SLV-Prüfung, Start P488 ab Feb 27.", C.accent1],
];
cards.forEach(([t, d, col], i) => {
  const cx = 0.5 + i * 3.13, cy = 2.3;
  box(s, { x: cx, y: cy, w: 2.88, h: 2.55, fill: { color: C.background2 }, rounded: true });
  s.addShape(pres.shapes.OVAL, { objectName: nm("kartenkreis"), x: cx + 0.25, y: cy + 0.25, w: 0.5, h: 0.5, fill: { color: col }, line: { type: "none" } });
  tb(s, String(i + 1), { x: cx + 0.25, y: cy + 0.25, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
  tb(s, t, { x: cx + 0.25, y: cy + 0.9, w: 2.4, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  tb(s, d, { x: cx + 0.25, y: cy + 1.35, w: 2.45, h: 1.1, fontSize: 13, color: C.text1 });
});
box(s, { x: 0.5, y: 5.2, w: 12.3, h: 1.55, fill: { color: C.text2 }, rounded: true });
tb(s, "Die vier Excel-Blätter", { x: 0.8, y: 5.35, w: 4, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
const sheets = [["1 Schichtig", "Ablauf mit einer Schicht"], ["2 Schichtig", "Ablauf mit zwei Schichten"], ["Übersicht Schweißproben", "VP/AP-Termine je Station"], ["Auswertung Risiko", "Bewertung 0–3 je Szenario"]];
sheets.forEach(([a, b], i) => {
  tb(s, [{ text: a, options: { bold: true, breakLine: true, color: C.background1 } }, { text: b, options: { color: C.background2 } }],
    { x: 0.8 + i * 3.0, y: 5.8, w: 2.8, h: 0.8, fontSize: 14 });
});
s.addNotes("Kernfrage des Plans: Kann die Fertigung einschichtig laufen, oder braucht es zwei Schichten, damit 475/309 rechtzeitig gefertigt werden kann? Der Kopf der Excel-Blätter lautet 'möglicher Ablauf Testschuss 1 um 475/309 fertigen zu können'. Rechts daneben steht 'Empfehlung: Damit …' – der Satz ist im Original nicht zu Ende geschrieben. Die Empfehlung fehlt also noch und wird am Ende dieser Präsentation aus den Daten abgeleitet.");

// 3 So liest man den Plan
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("So liest man den Kreuzchenplan", { placeholder: "title" });
// Mock-Raster
const gx0 = 0.5, gy0 = 1.35, cw = 0.42, rh = 0.36;
const days = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So", "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"];
tb(s, "KW 45", { x: gx0 + 2.6, y: gy0, w: cw * 7, h: 0.3, fontSize: 11, bold: true, color: C.text2, align: "center" });
tb(s, "KW 46", { x: gx0 + 2.6 + cw * 7, y: gy0, w: cw * 7, h: 0.3, fontSize: 11, bold: true, color: C.text2, align: "center" });
days.forEach((d, i) => tb(s, d, { x: gx0 + 2.6 + i * cw, y: gy0 + 0.3, w: cw, h: 0.28, fontSize: 9, color: C.accent5, align: "center" }));
const mock = [
  ["Schussvermessen", "", [[0, 3, HEX.accent2], [7, 10, HEX.accent3]]],
  ["St. 15 innen LN", "Schicht 1", [[8, 11, HEX.accent2]]],
  ["St. 15 innen LN", "Schicht 2", []],
  ["St. 15 innen LN", "Schicht 3", []],
  ["Wartezeit", "", [[4, 4, HEX.accent2]]],
];
mock.forEach(([l, sh, bars], r) => {
  const ry = gy0 + 0.62 + r * rh;
  tb(s, l, { x: gx0, y: ry, w: 1.7, h: rh, fontSize: 11, valign: "middle" });
  tb(s, sh, { x: gx0 + 1.7, y: ry, w: 0.85, h: rh, fontSize: 9, color: C.accent5, valign: "middle" });
  days.forEach((d, i) => {
    const we = i % 7 >= 5;
    let col = we ? "D5DAE0" : "FFFFFF";
    bars.forEach(([a, b, c]) => { if (i >= a && i <= b && !we) col = c; });
    box(s, { x: gx0 + 2.6 + i * cw, y: ry, w: cw, h: rh, fill: { color: col }, line: { color: "C9D1D9", width: 0.5 } });
  });
});
// Erklärungen
const expl = [
  ["Zeilen = Arbeitsschritte", "von Station 12 (Schussmontage) bis Station 18 (Rundnaht), darunter die Schweißproben je Station."],
  ["Spalten = Kalendertage", "21.09.2026 bis 28.06.2027 (KW 39 bis KW 26), grau = Wochenende und Betriebsruhe 24.12.–01.01."],
  ["Schichtzeilen 1–3", "Schweißschritte haben je Schicht eine eigene Zeile. Belegt = in dieser Schicht wird gearbeitet."],
  ["Farbe = Schuss", "Jede Farbe steht für einen Schuss. Die Zellbeschriftung nennt Detail: LN 1–3, Spt 1–6, NDT-Verfahren."],
];
expl.forEach(([h, t], i) => {
  const ey = 4.0 + i * 0.72;
  circleNum(s, 0.5, ey + 0.05, 0.42, i + 1);
  tb(s, [{ text: h + "  ", options: { bold: true, color: C.text2 } }, { text: t }], { x: 1.1, y: ey, w: 7.0, h: 0.62, fontSize: 13, valign: "middle" });
});
// Legende rechts
box(s, { x: 8.6, y: 1.35, w: 4.2, h: 5.4, fill: { color: C.background2 }, rounded: true });
tb(s, "Farbcode in dieser Präsentation", { x: 8.85, y: 1.5, w: 3.8, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
const leg = [
  [HEX.accent5, "Testschuss 1", "Excel: dunkelgrau"],
  [HEX.accent2, "Schuss 301 (P800)", "Excel: hellblau"],
  [HEX.accent3, "Schuss 302 (P800)", "Excel: hellgrün"],
  [HEX.accent4, "Schüsse 309/310 (475)", "Excel: orange"],
  [HEX.accent1, "Schweißprobe VP/AP", "Excel: diverse"],
  [HEX.accent6, "Stichtag, SAT, Fertigungsstart", "Excel: rot / Pfeile"],
];
leg.forEach(([c, a, b], i) => {
  const ly = 2.05 + i * 0.6;
  box(s, { x: 8.85, y: ly + 0.07, w: 0.4, h: 0.3, fill: { color: c } });
  tb(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: b, options: { fontSize: 10, color: C.accent5 } }], { x: 9.4, y: ly, w: 3.3, h: 0.5, fontSize: 12 });
});
tb(s, "Pfeile im Excel markieren Wartezeiten auf SLV-Ergebnisse, Kommentare (Autor: Klein, Bjoern) nennen Bedarfstermine.", { x: 8.85, y: 5.75, w: 3.75, h: 0.85, fontSize: 11, color: C.text1 });
s.addNotes("Aufbau: Spalte A = Arbeitsschritt, Spalte B = Schichtnummer (1, 2, 3), ab Spalte C ein Kalendertag pro Spalte, Zeile 3 enthält die Kalenderwoche, Zeile 4 das Datum. Ein 'Kreuzchen' ist eine farbig gefüllte Zelle: an diesem Tag, in dieser Schicht, wird dieser Schritt für den Schuss mit dieser Farbe ausgeführt. Wochenenden und die Betriebsruhe zwischen Weihnachten und Neujahr sind grau hinterlegt. Zusätzliche Informationen stecken in Zellkommentaren (Bedarfstermine, Probenversand, SLV-Ergebnisse) und in eingezeichneten Pfeilen (Wartezeit bis zum SLV-Ergebnis) sowie einer roten senkrechten Linie am 29.10.2026.");

// 4 Prozesskette
S = sec("Ablauf eines Schusses");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Jeder Schuss durchläuft dieselbe Kette aus sieben Stationen", { placeholder: "title" });
const steps = [
  ["St. 12", "Schussmontage", "Umbau, Heften", "≈ 5–10 AT"],
  ["Halle 40", "Vermessen", "Transport + Messen", "≈ 4–5 AT"],
  ["St. 15", "Längsnähte", "innen, ausfugen, außen, DL schleifen, VT, NDT", "≈ 15 AT"],
  ["St. 16", "Spante", "Einrichten, 6 Spante schweißen", "≈ 3–7 AT"],
  ["Putzplatz", "Putzen + Prüfen", "Putzen, VT, Wartezeit, NDT UT/MT", "≈ 10 AT"],
  ["Messplatz", "Messen + Abgabe", "Endmaß, Übergabe", "≈ 5–10 AT"],
  ["St. 18", "Rundnaht", "2 Schüsse verbinden, RN innen/außen", "≈ 12–15 AT"],
];
const sw = 1.66, sg = 0.12;
steps.forEach(([st, t, d, dur], i) => {
  const sx = 0.5 + i * (sw + sg), sy = 1.5;
  s.addShape(pres.shapes.CHEVRON, { objectName: nm("prozess"), x: sx, y: sy, w: sw + 0.1, h: 0.9, fill: { color: i === 2 || i === 3 || i === 6 ? C.accent1 : C.text2 }, line: { type: "none" } });
  tb(s, st, { x: sx + 0.25, y: sy, w: sw - 0.35, h: 0.9, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
  tb(s, t, { x: sx, y: sy + 1.05, w: sw, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center" });
  tb(s, d, { x: sx + 0.05, y: sy + 1.42, w: sw - 0.1, h: 0.8, fontSize: 11, color: C.text1, align: "center" });
  tb(s, dur, { x: sx, y: sy + 2.25, w: sw, h: 0.3, fontSize: 12, bold: true, color: C.accent1, align: "center" });
});
tb(s, "AT = Arbeitstage, abgelesen aus dem Plan (Schuss 301, einschichtig). Orange = Schweißstationen, an denen eine qualifizierte WPS vorliegen muss.", { x: 0.5, y: 4.15, w: 12.3, h: 0.3, fontSize: 11, color: C.accent5 });
// Drei Merkmale
const feat = [
  ["Wartezeit vor jeder NDT", "Nach VT folgt immer ein Block „Wartezeit“ von 5 AT, erst dann Ultraschall, Röntgen und Magnetpulver. Diese Tage sind fest und lassen sich nicht durch mehr Personal verkürzen."],
  ["Paarweise Weiterverarbeitung", "Auf Station 18 werden immer zwei Schüsse zusammen verarbeitet (301+302, 309+310). Der spätere Schuss eines Paares bestimmt den Start der Rundnaht."],
  ["Station = Engpass", "Jede Station kann nur einen Schuss gleichzeitig tragen. Ein Schuss rückt erst nach, wenn der vorherige die Station verlassen hat."],
];
feat.forEach(([h, t], i) => {
  const fx = 0.5 + i * 4.15;
  box(s, { x: fx, y: 4.65, w: 3.95, h: 2.15, fill: { color: C.background2 }, rounded: true });
  circleNum(s, fx + 0.2, 4.82, 0.42, i + 1);
  tb(s, h, { x: fx + 0.75, y: 4.82, w: 3.05, h: 0.42, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
  tb(s, t, { x: fx + 0.2, y: 5.38, w: 3.6, h: 1.35, fontSize: 12.5, color: C.text1 });
});
s.addNotes("Die Kette im Detail (Zeilen im Excel): Station 12 Schussmontage (inkl. Umbau der Vorrichtung zwischen den Projekten) – Transport in Halle 40 zum Vermessen – Transport auf Station 15 – Längsnähte LN 1 bis 3 innen schweißen – Ausfugen inkl. MT – Längsnähte außen – Decklage schleifen – VT – Wartezeit – NDT der Längsnähte (UT/RT/MT) – Restanbindung – Filmausschnitte – Transport auf Station 16 – Station einrichten – Spante 1–6 schweißen (paarweise je Schicht) – Transport Putzplatz – Spante putzen – VT Schuss – Wartezeit – NDT UT/MT – Transport Messplatz – Messen – Abgabe – Transport auf Station 18 – Einrichten – Rundnaht verarbeiten – Spantabstand messen – Rundnaht innen schweißen – außen fräsen inkl. MT – Rundnaht außen – Reserve – Decklage außen schleifen. Abkürzungen: LN = Längsnaht, RN = Rundnaht, Spt = Spant, DL = Decklage, SL = Schleifen, VT/UT/RT/MT = Sicht-, Ultraschall-, Durchstrahlungs-, Magnetpulverprüfung.");

// 5 Testschuss 1
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Testschuss 1 liefert die Proben, auf die alles wartet", { placeholder: "title" });
const g1 = gantt(s, {
  x: 0.5, y: 1.3, w: 8.2, labelW: 2.3, start: "28.09.26", end: "27.11.26", rowH: 0.48, labelSize: 12,
  rows: [
    { label: "St. 16 Spante schweißen", bars: [{ s: "30.09.26", e: "08.10.26", color: HEX.accent5, text: "Spt 1–6" }] },
    { label: "Putzen, VT, Wartezeit", bars: [{ s: "12.10.26", e: "23.10.26", color: HEX.accent5, text: "VP Reparatur Spt 4 (22.10.)", tout: true, tw: 2.2, tcolor: C.accent6 }], ms: [{ d: "22.10.26" }] },
    { label: "NDT UT/MT", bars: [{ s: "26.10.26", e: "27.10.26", color: HEX.accent5 }] },
    { label: "Probenentnahme Spant/DK", bars: [{ s: "28.10.26", e: "29.10.26", color: HEX.accent1 }], ms: [{ d: "30.10.26", text: "Versand SLV", tw: 1.0 }] },
    { label: "SLV Rostock prüft", bars: [{ s: "31.10.26", e: "19.11.26", color: HEX.accent6, arrow: true, text: "≈ 3 Wochen", fs: 10 }], ms: [{ d: "20.11.26", text: "Ergebnis", tw: 0.8 }] },
    { label: "St. 18 Rundnaht (Probe)", bars: [{ s: "29.10.26", e: "20.11.26", color: HEX.accent5, text: "einrichten → RN innen/außen → schleifen" }] },
  ],
  vlines: [{ d: "29.10.26", text: "Stichtag 29.10." }],
});
box(s, { x: 9.0, y: 1.3, w: 3.8, h: 5.45, fill: { color: C.text2 }, rounded: true });
tb(s, "Warum er kritisch ist", { x: 9.25, y: 1.5, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
tb(s, [
  { text: "Proben aus Spant 4 und 5 gehen am 30.10. an die SLV Rostock.", options: { bullet: true, breakLine: true } },
  { text: "Erst mit dem SLV-Ergebnis ist die WPS 121 für Station 16 freigegeben.", options: { bullet: true, breakLine: true } },
  { text: "Ohne WPS darf kein Schuss 301, 302 oder 309 auf St. 16 Spante schweißen.", options: { bullet: true, breakLine: true } },
  { text: "Parallel qualifiziert der Testschuss die Rundnaht auf St. 18.", options: { bullet: true } },
], { x: 9.25, y: 2.0, w: 3.35, h: 3.2, fontSize: 14, color: C.background1, paraSpaceAfter: 10 });
tb(s, "20.11.", { x: 9.25, y: 5.2, w: 3.35, h: 0.8, fontSize: 40, bold: true, color: C.accent1 });
tb(s, "erwartetes SLV-Ergebnis Testschuss 1", { x: 9.25, y: 6.0, w: 3.35, h: 0.5, fontSize: 13, color: C.background2 });
s.addNotes("Testschuss 1 ist in beiden Blättern identisch (dunkelgrau). Spante schweißen auf Station 16 vom 30.09. bis 08.10.2026, danach Putzen (14.–21.10.), VT, Wartezeit, NDT UT/MT am 26./27.10. Kommentar 22.10.: Schweißverfahrensprüfung Spant 4 und Reparatur einer Stelle im Spant als Verfahrensprüfung Reparatur. Am 28.10. Probenentnahme Spant/Deck, am 29.10. werden die Entnahmelöcher geschlossen (gelb markiert). Kommentar 30.10.: Versand der Proben Spant 4 und 5 an die SLV Rostock. Ein Pfeil führt bis zum 19./20.11. – dort steht der Kommentar 'Probenergebnisse der SLV'. Ab 29.10. wird der Testschuss auf Station 18 gebracht und die Rundnaht probeweise verarbeitet (bis 20.11.). Die rote senkrechte Linie am 29.10. zieht sich über das gesamte Blatt und markiert diesen Stichtag. Heute (31.10.) liegt also genau zwischen Probenversand und SLV-Ergebnis.");

// 6 Szenario 1-schichtig
S = sec("Szenarien im Vergleich");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Szenario A: einschichtig – nur 301/302, kein Platz für 309", { placeholder: "title" });
const L4 = 4;
const gA = gantt(s, {
  x: 0.5, y: 1.25, w: 12.3, labelW: 1.9, start: "28.09.26", end: "05.02.27", rowH: 0.58, labelSize: 12,
  rows: [
    { label: "St. 12 Montage", bars: [{ s: "19.10.26", e: "30.10.26", color: HEX.accent2, lanes: L4, lane: 2, }, { s: "02.11.26", e: "05.11.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Vermessen", bars: [{ s: "02.11.26", e: "06.11.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "06.11.26", e: "12.11.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 15 Längsnähte", bars: [{ s: "09.11.26", e: "01.12.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "16.11.26", e: "07.12.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 16 Spante", bars: [{ s: "30.09.26", e: "08.10.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "03.12.26", e: "10.12.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "11.12.26", e: "16.12.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Putzen, VT, NDT", bars: [{ s: "12.10.26", e: "27.10.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "11.12.26", e: "05.01.27", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "18.12.26", e: "07.01.27", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Messen, Abgabe", bars: [{ s: "06.01.27", e: "18.01.27", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "12.01.27", e: "21.01.27", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 18 Rundnaht", bars: [{ s: "29.10.26", e: "20.11.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "22.01.27", e: "02.02.27", color: HEX.accent2, lanes: 2, lane: 1, text: "301+302", fs: 9 }] },
    { label: "475/309, 310", bars: [{ s: "28.09.26", e: "05.02.27", color: HEX.accent4, lanes: 1, text: "Nicht eingeplant – Station 12 bis 18 durchgehend belegt", tcolor: C.background1, fs: 12 }] },
  ],
  vlines: [{ d: "29.10.26", text: "Stichtag 29.10." }, { d: "14.12.26", color: HEX.accent1, dash: "dash", text: "Spant 302 ≥ 14.12.", tcolor: C.accent1 }],
});
legend(s, 0.5, 6.62, [["Testschuss 1", HEX.accent5, 1.15], ["Schuss 301", HEX.accent2, 0.95], ["Schuss 302", HEX.accent3, 0.95], ["309/310", HEX.accent4, 0.8]]);
tb(s, [{ text: "Abgabe 301: ", options: { bold: true } }, { text: "18.01.27   " }, { text: "Abgabe 302: ", options: { bold: true } }, { text: "21.01.27" }], { x: 8.3, y: 6.62, w: 4.5, h: 0.28, fontSize: 13, color: C.text2, align: "right" });
s.addNotes("Blatt '1 Schichtig': Station 12 wird ab 19.10. umgebaut, ab 26.10. wird 800/301 montiert, ab 02.11. folgt 302. 301 ist auf Station 15 vom 10.11. bis Anfang Dezember, Spante auf Station 16 vom 04.12. bis 10.12. Für 302 steht der Kommentar 'Start nicht vor dem 14.12.' (roter Pfeil im Excel). Abgabe 301 am 18.01.2027, Abgabe 302 am 21.01.2027, danach Rundnaht 301+302 auf Station 18 ab 22.01.2027. Die Zeile 'Station 12 475/309' ist leer: im einschichtigen Ablauf ist für 475/309 und 310 kein Platz. Laut Risikoblatt belegt im einschichtigen Fall die Sektion 475/31x die Station 18, wenn P800 sie braucht (MPM-Risiko 3). Der Basisterminplan für Schuss 309 (Bild im Risikoblatt) sieht 11.11.2026 bis FAT 22.01.2027 vor – das kollidiert direkt mit 301/302.");

// 7 Szenario 2-schichtig
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Szenario B: zweischichtig – erst 309/310, dann P800", { placeholder: "title" });
const gB = gantt(s, {
  x: 0.5, y: 1.25, w: 12.3, labelW: 1.9, start: "28.09.26", end: "05.02.27", rowH: 0.58, labelSize: 12,
  rows: [
    { label: "St. 12 Montage", bars: [{ s: "05.10.26", e: "14.10.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "19.10.26", e: "28.10.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "02.11.26", e: "05.11.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Vermessen", bars: [{ s: "15.10.26", e: "21.10.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "29.10.26", e: "04.11.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "06.11.26", e: "12.11.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 15 Längsnähte", bars: [{ s: "22.10.26", e: "12.11.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "05.11.26", e: "23.11.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "16.11.26", e: "07.12.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 16 Spante", bars: [{ s: "30.09.26", e: "08.10.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "13.11.26", e: "18.11.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "23.11.26", e: "27.11.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "11.12.26", e: "16.12.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Putzen, VT, NDT", bars: [{ s: "12.10.26", e: "27.10.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "19.11.26", e: "01.12.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "30.11.26", e: "15.12.26", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "17.12.26", e: "23.12.26", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "Messen, Abgabe", bars: [{ s: "02.12.26", e: "14.12.26", color: HEX.accent4, lanes: L4, lane: 1 }, { s: "16.12.26", e: "04.01.27", color: HEX.accent2, lanes: L4, lane: 2 }, { s: "18.01.27", e: "18.01.27", color: HEX.accent3, lanes: L4, lane: 3 }] },
    { label: "St. 18 Rundnaht", bars: [{ s: "29.10.26", e: "20.11.26", color: HEX.accent5, lanes: L4, lane: 0 }, { s: "14.12.26", e: "15.01.27", color: HEX.accent4, lanes: L4, lane: 1, text: "309+310", fs: 8 }, { s: "20.01.27", e: "02.02.27", color: HEX.accent2, lanes: 2, lane: 1, text: "301+302", fs: 9 }] },
  ],
  vlines: [{ d: "29.10.26", text: "Stichtag 29.10." }, { d: "16.11.26", color: HEX.accent1, dash: "dash", text: "WPS St. 16 nötig 16.11.", tcolor: C.accent1 }],
});
legend(s, 0.5, 6.62, [["Testschuss 1", HEX.accent5, 1.15], ["Schuss 301", HEX.accent2, 0.95], ["Schuss 302", HEX.accent3, 0.95], ["309/310", HEX.accent4, 0.8]]);
tb(s, [{ text: "Abgabe 309: ", options: { bold: true } }, { text: "14.12.26   " }, { text: "301: ", options: { bold: true } }, { text: "04.01.27   " }, { text: "302: ", options: { bold: true } }, { text: "18.01.27" }], { x: 7.3, y: 6.62, w: 5.5, h: 0.28, fontSize: 13, color: C.text2, align: "right" });
s.addNotes("Blatt '2 Schichtig': Station 12 wird schon vom 05. bis 09.10. für 475/309 umgebaut, 309 wird vom 12. bis 14.10. in zwei Schichten montiert (Kommentar: 'Spätester Bedarf Spante für 475/309 inkl. Schuss 475/310' am 12.10.). Die Längsnähte auf Station 15 laufen in zwei Schichten parallel (LN 1 und LN 2 gleichzeitig), dadurch halbieren sich die Durchlaufzeiten. 309 ist am 14.12. abgegeben, die Rundnaht 309+310 auf Station 18 läuft vom 14.12. bis 15.01.2027 – eine Woche vor dem FAT-Termin 22.01.2027 aus dem Basisplan. 301 wird am 04.01.2027 abgegeben (zwei Wochen früher als einschichtig), 302 am 18.01.2027. Rundnaht 301+302 ab 20.01.2027. Kommentar: WPS Station 15 (138P/121) ab 23.10.26 notwendig, WPS Station 16 (121) ab 16.11.2026.");

// 8 Vergleich Kennzahlen
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Zwei Schichten: 309/310 im Plan, 301 zwei Wochen früher", { placeholder: "title" });
const cols = [
  ["Kriterium", "Einschichtig", "Zweischichtig"],
  ["Schüsse im Plan", "301, 302", "309, 310, 301, 302"],
  ["Start Station 12", "19.10.26 (Umbau)", "05.10.26 (Umbau 309)"],
  ["Abgabe 309", "nicht eingeplant", "14.12.26"],
  ["Rundnaht 309+310 fertig", "–", "15.01.27 (FAT-Plan 22.01.27)"],
  ["Abgabe 301", "18.01.27", "04.01.27"],
  ["Abgabe 302", "21.01.27", "18.01.27"],
  ["Start Rundnaht 301+302", "22.01.27", "20.01.27"],
  ["Zeit für WPS-Freigabe St. 16", "3,5 Wochen", "2 Wochen"],
  ["Risikosumme (0–3 je Punkt)", "28", "30"],
];
const rows = cols.map((r, i) => r.map((c, j) => ({
  text: c,
  options: {
    bold: i === 0 || j === 0, fontSize: 13,
    color: i === 0 ? C.background1 : (j === 2 && [3, 4, 5].includes(i) ? C.accent3 : (j === 1 && i === 3 ? C.accent6 : C.text1)),
    fill: { color: i === 0 ? C.text2 : (i % 2 ? C.background1 : C.background2) },
    valign: "middle",
  },
})));
s.addTable(rows, { x: 0.5, y: 1.3, w: 8.0, colW: [3.1, 2.3, 2.6], rowH: 0.45, border: { type: "solid", pt: 0.5, color: "C9D1D9" }, objectName: nm("tabelle") });
// Stat-Callouts
const stats = [["+2", "Schüsse (309/310) zusätzlich", C.accent4], ["–14", "Tage früher Abgabe 301", C.accent2], ["–1,5", "Wochen weniger Puffer für WPS St. 16", C.accent6]];
stats.forEach(([n, l, c], i) => {
  const sy = 1.3 + i * 1.62;
  box(s, { x: 8.9, y: sy, w: 3.9, h: 1.45, fill: { color: C.background2 }, rounded: true });
  tb(s, n, { x: 9.1, y: sy + 0.12, w: 1.6, h: 1.2, fontSize: 44, bold: true, color: c, valign: "middle" });
  tb(s, l, { x: 10.75, y: sy + 0.12, w: 1.9, h: 1.2, fontSize: 14, color: C.text1, valign: "middle" });
});
s.addNotes("Gegenüberstellung der Termine aus den beiden Blättern. Der wesentliche Gewinn des zweischichtigen Szenarios ist, dass 309/310 überhaupt eingeplant werden kann – und zwar mit einer Woche Puffer zum FAT-Termin 22.01.2027 aus dem Basisplan der Sektion 33. Der Preis: die Freigabe der Schweißanweisung für Station 16 muss in 2 statt 3,5 Wochen kommen, und die Personalrisiken steigen. Die Risikosummen (28 gegen 30) sind fast gleich – die Art des Risikos unterscheidet sich aber deutlich (siehe Risikoauswertung).");

// 9 Abhängigkeitsnetz
S = sec("Abhängigkeiten");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Die Schweißfreigabe steuert den kritischen Pfad", { placeholder: "title" });
function node(slide, x, y, w, h, title, sub, fill, txt) {
  box(slide, { x, y, w, h, fill: { color: fill }, rounded: true });
  tb(slide, [{ text: title, options: { bold: true, breakLine: !!sub } }].concat(sub ? [{ text: sub, options: { fontSize: 10.5 } }] : []),
    { x: x + 0.1, y, w: w - 0.2, h, fontSize: 12.5, color: txt || C.background1, align: "center", valign: "middle" });
}
function arrow(slide, x1, y1, x2, y2, color, dash) {
  slide.addShape(pres.shapes.LINE, { objectName: nm("abh"), x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1, flipV: y2 < y1, line: { color: color || HEX.dk2, width: 1.75, endArrowType: "triangle", dashType: dash || "solid" } });
}
// Spalte 1: Eingänge
node(s, 0.5, 1.35, 2.4, 0.85, "Testschuss 1", "Proben Spant/DK 28.10.", HEX.accent5);
node(s, 0.5, 2.95, 2.4, 0.85, "Anlage einfahren", "Unterstützung Kiel", HEX.dk2);
node(s, 0.5, 4.15, 2.4, 0.85, "Personal", "Schweißer mit Prüfung MSG/UP", HEX.dk2);
node(s, 0.5, 5.55, 2.4, 0.85, "Material Kiel", "Spante 309/310 bis 12.10.", HEX.dk2);
// Spalte 2
node(s, 3.55, 1.35, 2.4, 0.85, "SLV Rostock", "Prüfung ≈ 3 Wochen", HEX.accent6);
node(s, 3.55, 2.35, 2.4, 0.85, "WPS St. 15", "138P/121, ab 23./27.10.", HEX.accent1);
node(s, 3.55, 3.6, 2.4, 0.85, "Kapazität je Schicht", "1 oder 2 Schichten", HEX.dk2);
// Spalte 3/4
node(s, 6.6, 1.35, 2.4, 0.85, "WPS St. 16", "121, Bedarf 16.11. / 04.12.", HEX.accent1);
node(s, 6.6, 2.95, 2.4, 0.85, "St. 15 Längsnähte", "", HEX.dk2);
node(s, 9.65, 2.25, 2.4, 0.85, "St. 16 Spante", "301 · 302 · 309", HEX.dk2);
node(s, 9.65, 3.95, 2.4, 0.85, "Messen + Abgabe", "", HEX.dk2);
node(s, 9.65, 5.55, 3.15, 0.85, "St. 18 Rundnaht", "nur paarweise, Station einmal belegt", HEX.dk2);
node(s, 6.6, 5.55, 2.4, 0.85, "475/309 FAT", "22.01.2027", HEX.accent4);
arrow(s, 2.9, 1.78, 3.55, 1.78, HEX.accent6);
arrow(s, 5.95, 1.78, 6.6, 1.78, HEX.accent6);
arrow(s, 9.0, 1.78, 9.65, 2.5, HEX.accent6);
arrow(s, 2.9, 3.38, 3.55, 3.85);
arrow(s, 2.9, 4.58, 3.55, 4.2);
arrow(s, 5.95, 2.78, 6.6, 3.2);
arrow(s, 5.95, 4.02, 6.6, 3.55);
arrow(s, 9.0, 3.38, 9.65, 2.9);
arrow(s, 10.85, 3.1, 10.85, 3.95);
arrow(s, 10.85, 4.8, 10.85, 5.55);
arrow(s, 9.65, 5.98, 9.0, 5.98, HEX.accent4);
arrow(s, 2.9, 5.98, 6.6, 5.98, HEX.accent4, "dash");
tb(s, "Rot = kritischer Pfad über die Verfahrensprüfung. Verzögert sich die SLV, verschieben sich St. 16 und alle Abgaben.", { x: 9.65, y: 1.25, w: 3.15, h: 0.85, fontSize: 11.5, color: C.text1 });
s.addNotes("Die Abhängigkeiten, die der Plan über Kommentare und Pfeile ausdrückt: (1) Testschuss 1 → Probenentnahme 28.10. → Versand an die SLV Rostock 30.10. → Ergebnis ca. 20.11. → Freigabe WPS 121 für Station 16. Ohne diese Freigabe darf auf Station 16 kein Spant eines Serienschusses geschweißt werden. (2) WPS 138P/121 für Station 15 muss vor den Längsnähten vorliegen (Kommentar: ab 23.10. bei zwei Schichten, ab 27.10. bei einer Schicht). (3) Material aus Kiel: Spante für 475/309 und Schuss 310 spätestens 12.10. – laut Risikoblatt wegen Seetransport mit zwei Wochen Vorlauf. (4) Anlage einfahren mit Unterstützung aus Kiel und Schweißer mit gültiger Prüfung (MSG, UP) begrenzen die Kapazität. (5) Station 18 kann nur ein Schusspaar gleichzeitig aufnehmen – das ist das MPM-Risiko im einschichtigen Fall. (6) Der FAT von 475/309 am 22.01.2027 hängt an der Rundnaht 309+310.");

// 10 WPS-Fenster St. 16
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Zweischichtig wird die WPS St. 16 vor dem SLV-Ergebnis gebraucht", { placeholder: "title" });
const tl = gantt(s, {
  x: 0.5, y: 1.4, w: 12.3, labelW: 2.6, start: "26.10.26", end: "11.12.26", rowH: 0.7, labelSize: 13,
  rows: [
    { label: "Proben Testschuss 1", bars: [{ s: "28.10.26", e: "29.10.26", color: HEX.accent5 }], ms: [{ d: "30.10.26", text: "Versand 30.10.", tw: 1.2 }] },
    { label: "SLV Rostock", bars: [{ s: "31.10.26", e: "19.11.26", color: HEX.accent6, arrow: true, text: "Prüfung", fs: 10 }], ms: [{ d: "20.11.26", text: "Ergebnis ≈ 20.11.", tw: 1.5 }] },
    { label: "Zweischichtig: 309 St. 16", bars: [{ s: "16.11.26", e: "18.11.26", color: HEX.accent4, text: "Spant", fs: 9 }], ms: [{ d: "13.11.26", text: "16.11. vor Ergebnis", left: true, tw: 1.6, color: HEX.accent6 }] },
    { label: "Einschichtig: 301 St. 16", bars: [{ s: "04.12.26", e: "10.12.26", color: HEX.accent2, text: "Spante", fs: 9 }] },
  ],
  vlines: [{ d: "20.11.26", color: HEX.accent6, dash: "dash", text: "SLV-Ergebnis" }],
});
const cx2 = [["Zweischichtig", "Risikoblatt: 2 Wochen Freigabezeit", "Der Kommentar nennt 16.11. als Bedarf, der Pfeil zur SLV endet erst am 19./20.11. Ohne früheres Ergebnis oder Vorabfreigabe fehlen rund 4 Arbeitstage.", HEX.accent6],
  ["Einschichtig", "Risikoblatt: 3,5 Wochen Freigabezeit", "Zwischen SLV-Ergebnis und Start 301 auf St. 16 liegen 2 Wochen Puffer für Nacharbeit oder Wiederholungsproben.", HEX.accent3]];
cx2.forEach(([h, a, b, col], i) => {
  const bx = 0.5 + i * 6.25;
  box(s, { x: bx, y: 4.85, w: 6.05, h: 1.95, fill: { color: C.background2 }, rounded: true });
  s.addShape(pres.shapes.OVAL, { objectName: nm("ampel"), x: bx + 0.25, y: 5.05, w: 0.38, h: 0.38, fill: { color: col }, line: { type: "none" } });
  tb(s, h, { x: bx + 0.8, y: 5.03, w: 2.2, h: 0.42, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  tb(s, a, { x: bx + 3.0, y: 5.03, w: 2.85, h: 0.42, fontSize: 12, color: C.accent5, valign: "middle", align: "right" });
  tb(s, b, { x: bx + 0.25, y: 5.6, w: 5.6, h: 1.1, fontSize: 13, color: C.text1 });
});
s.addNotes("Das ist der wichtigste Befund der Auswertung. Im Blatt '2 Schichtig' steht ein roter Pfeil vom 29.10. bis 16.11. und der Kommentar 'WPS Station 16 121 wird ab 16.11.2026 benötigt'. Am 16.11. startet laut Plan auch das Schweißen der Spante von 309 auf Station 16. Im Blatt '1 Schichtig' endet der Pfeil der SLV-Prüfung aber erst am 19.11., der Kommentar 'Probenergebnisse der SLV' steht am 20.11. Das Risikoblatt bewertet die WPS 121 Station 16 in beiden Szenarien mit 3 (hoch) und schreibt: bei 1-schichtig 3,5 Wochen Zeit für die Genehmigung, bei 2-schichtig nur 2 Wochen. Konsequenz: Für das zweischichtige Szenario muss mit der SLV ein früheres Ergebnis vereinbart werden, oder die Reihenfolge auf Station 16 wird angepasst (z. B. 309-Spante erst ab ca. 23.11.).");

// 11 Schweißproben-Roadmap
S = sec("Schweißproben 1.3964");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Schweißproben 1.3964: VP und AP je Station, jeweils mit SLV", { placeholder: "title" });
const P = HEX.accent1, A = HEX.dk2, R = HEX.accent6;
const gP = gantt(s, {
  x: 0.5, y: 1.2, w: 12.3, labelW: 2.15, start: "02.11.26", end: "31.07.27", rowH: 0.52, labelSize: 11,
  rows: [
    { label: "Rahmenspant Gurt+Steg", bars: [{ s: "23.11.26", e: "04.12.26", color: P }, { s: "05.12.26", e: "15.01.27", color: R, arrow: true }] },
    { label: "Rahmenspant Gurt", bars: [{ s: "07.12.26", e: "18.12.26", color: P }, { s: "19.12.26", e: "29.01.27", color: R, arrow: true }] },
    { label: "Station 2 (Gurt)", bars: [{ s: "09.11.26", e: "20.11.26", color: P }, { s: "21.11.26", e: "22.12.26", color: R, arrow: true }, { s: "11.01.27", e: "22.01.27", color: A }, { s: "23.01.27", e: "26.02.27", color: R, arrow: true, text: "Start P488 Gurte 25.02.", tout: true, tw: 1.8, tcolor: C.accent6 }], ms: [{ d: "06.01.27", color: HEX.accent5 }, { d: "25.02.27" }] },
    { label: "Station 5 (Gurt)", bars: [{ s: "01.02.27", e: "05.02.27", color: A }, { s: "06.02.27", e: "12.03.27", color: R, arrow: true }], ms: [{ d: "20.01.27", color: HEX.accent5 }] },
    { label: "Station 7 (Stege)", bars: [{ s: "11.01.27", e: "12.02.27", color: P }, { s: "15.02.27", e: "19.02.27", color: A }, { s: "20.02.27", e: "19.03.27", color: R, arrow: true, text: "Start P488 Stege 25.02.", tout: true, tw: 1.8, tcolor: C.accent6 }], ms: [{ d: "08.01.27", color: HEX.accent5 }, { d: "25.02.27" }] },
    { label: "St. 3 / 6 / 8 Vorstufen", bars: [{ s: "25.01.27", e: "29.01.27", color: HEX.accent5, text: "Walzen", tout: true, tw: 0.6, fs: 9 }, { s: "22.02.27", e: "26.02.27", color: HEX.accent5 }, { s: "01.03.27", e: "05.03.27", color: HEX.accent5, text: "Fräsen, Zusammenbau", tout: true, tw: 1.6, fs: 9 }] },
    { label: "Station 9 (Spant FW)", bars: [{ s: "08.03.27", e: "26.03.27", color: P }, { s: "27.03.27", e: "30.04.27", color: R, arrow: true }] },
    { label: "Station 15 (Längsnaht)", bars: [{ s: "16.04.27", e: "23.04.27", color: A }, { s: "24.04.27", e: "28.05.27", color: R, arrow: true, text: "Start 601 21.05.", tout: true, tw: 1.4, tcolor: C.accent6 }], ms: [{ d: "21.05.27" }] },
    { label: "Station 16 (Spant DK)", bars: [{ s: "25.01.27", e: "05.02.27", color: P }, { s: "06.02.27", e: "12.03.27", color: R, arrow: true }, { s: "18.05.27", e: "20.05.27", color: A }, { s: "21.05.27", e: "25.06.27", color: R, arrow: true, text: "Start 201 01.06.", tout: true, tw: 1.3, tcolor: C.accent6 }], ms: [{ d: "01.06.27" }] },
  ],
});
legend(s, 0.5, 6.5, [["Verfahrensprüfung (VP)", P, 1.85], ["Arbeitsprobe (AP)", A, 1.5], ["SLV-Prüfung ≈ 5 Wochen", R, 1.95], ["SAT der Anlage", HEX.accent5, 1.3]]);
tb(s, "Rote Raute = Fertigungsstart", { x: 10.2, y: 6.5, w: 2.6, h: 0.26, fontSize: 11, color: C.accent6, bold: true, align: "right" });
s.addNotes("Unterer Teil der Kreuzchenpläne (ab Zeile 80) und Blatt 'Übersicht Schweißproben 1.3964'. Reihenfolge je Station: Anlage abnehmen (SAT) → Verfahrensprüfung (VP) schweißen → Proben an die SLV → Ergebnis → Arbeitsprobe (AP) unter Fertigungsbedingungen → Proben an die SLV → Ergebnis → Fertigungsstart. Details: Rahmenspant Gurt+Steg: VP 23.11.–04.12.26 (Steg 30 mm, Gurt min. 45 mm, 55°, halb/halb, 9 mm Nase, zwei Traktoren nötig – ggf. in Kiel leihen), Ergebnis 15.01.27. Rahmenspant Gurt: VP 07.–18.12.26 (500x200x90, Tulpe 20°), Ergebnis 29.01.27. Station 2 und 5: gemeinsame VP mit Traktor in der Kursstätte (35°, halb/halb, 2 mm Nase), Ergebnis ca. 22.12.26; SAT St. 2 am 06.01.27, AP St. 2 11.–22.01.27, Ergebnis 26.02.27; SAT St. 5 20.01.27, AP 01.–05.02.27, Ergebnis 12.03.27. Station 7: SAT 08.01.27, VP Stege 11.01.–12.02.27, AP 5 Teststege mit 100 % RT und PT 15.–19.02.27, Ergebnis 19.03.27. Station 3 walzt die Testgurte (25.–29.01.27), Station 6 fräst CrNi-Stege (ab 22.02.27), Station 8 baut die Testspante zusammen (ab 01.03.27). Station 9: VP und AP am ganzen Spant (Kehlnaht umlaufend) 08.–26.03.27, Ergebnis 30.04.27. Station 15: AP an vorgezogenem Schuss 16.–23.04.27, Ergebnis 28.05.27, Fertigungsstart Schuss 601 am 21.05.27. Station 16: VP als Blech mit zwei Traktoren 25.01.–05.02.27, Ergebnis 12.03.27; AP 18.–20.05. (im Excel steht 2026, gemeint ist sehr wahrscheinlich 2027), Ergebnis 25.06.27, Fertigungsstart 01.06.27 (201) und 08.06.27 (601).");

// 12 Befunde Schweißproben
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Vier Fertigungsstarts liegen vor dem SLV-Ergebnis", { placeholder: "title" });
// Stat links
box(s, { x: 0.5, y: 1.3, w: 3.7, h: 5.45, fill: { color: C.text2 }, rounded: true });
tb(s, "5", { x: 0.75, y: 1.55, w: 3.2, h: 1.3, fontSize: 80, bold: true, color: C.accent1, valign: "middle" });
tb(s, "Wochen dauert jede SLV-Prüfung", { x: 0.75, y: 2.9, w: 3.2, h: 0.75, fontSize: 18, bold: true, color: C.background1 });
tb(s, "Alle Pfeile im Plan, vom Probenversand bis zum Ergebnis, sind 5 Wochen lang, über Weihnachten 6. Das ist die feste Taktzeit jeder Qualifizierung.", { x: 0.75, y: 3.75, w: 3.2, h: 1.6, fontSize: 13, color: C.background2 });
tb(s, "Testschuss 1 ist mit ≈ 3 Wochen eingeplant, also deutlich knapper.", { x: 0.75, y: 5.45, w: 3.2, h: 1.0, fontSize: 13, bold: true, color: C.accent1 });
const fnd = [
  ["Station 2: Start P488 Gurte 25.02.27", "SLV-Ergebnis AP: 26.02.27", "1 Tag"],
  ["Station 7: Start P488 Stege 25.02.27", "SLV-Ergebnis AP: 19.03.27", "3 Wochen"],
  ["Station 15: Start Schuss 601 21.05.27", "SLV-Ergebnis AP: 28.05.27", "1 Woche"],
  ["Station 16: Start 201 am 01.06.27", "SLV-Ergebnis AP: 25.06.27", "3,5 Wochen"],
];
tb(s, "Fertigungsstart vor Ergebnis", { x: 4.6, y: 1.3, w: 5, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
tb(s, "Lücke", { x: 11.0, y: 1.3, w: 1.8, h: 0.4, fontSize: 13, bold: true, color: C.accent5, align: "right" });
fnd.forEach(([a, b, gap], i) => {
  const fy = 1.85 + i * 0.95;
  box(s, { x: 4.6, y: fy, w: 8.2, h: 0.8, fill: { color: C.background2 }, rounded: true });
  s.addShape(pres.shapes.OVAL, { objectName: nm("warn"), x: 4.8, y: fy + 0.2, w: 0.4, h: 0.4, fill: { color: C.accent6 }, line: { type: "none" } });
  tb(s, "!", { x: 4.8, y: fy + 0.2, w: 0.4, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
  tb(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: b, options: { color: C.accent5 } }], { x: 5.4, y: fy + 0.05, w: 5.4, h: 0.7, fontSize: 13, valign: "middle" });
  tb(s, gap, { x: 10.8, y: fy, w: 1.8, h: 0.8, fontSize: 20, bold: true, color: C.accent6, align: "right", valign: "middle" });
});
tb(s, "Entweder wird für diese Starts eine Vorabfreigabe auf Basis der VP vereinbart, oder die Arbeitsproben müssen mindestens 5 Wochen vor Fertigungsstart geschweißt sein.", { x: 4.6, y: 5.75, w: 8.2, h: 0.9, fontSize: 14, color: C.text1 });
s.addNotes("Vergleich der Fertigungsstarts (Kommentare 'Start P488 Gurtfertigung', 'Start P488 Stegfertigung', Übersicht Spalte 'Fertigungsstart Projekt') mit den Ergebnisterminen der SLV. Station 2: Start 25.02., Ergebnis 26.02. – praktisch zeitgleich, kein Puffer. Station 7: Start 25.02., Ergebnis der AP erst 19.03. Station 15: Start Schuss 601 am 21.05., Ergebnis 28.05. Station 16: Start 01.06. (201) bzw. 08.06. (601), Ergebnis 25.06. Zusätzlich vermerkt die Übersicht bei Station 16: Proben aus dem Akkuluk erst nach dem Bau der Sektion 20. Diese Punkte sind im Plan nicht als Risiko bewertet und sollten in die Risikoliste aufgenommen werden.");

// 13 Risiko-Chart
S = sec("Risiken und Ressourcen");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Risikobewertung: fast gleiche Summe, verschiedene Treiber", { placeholder: "title" });
const rk = [
  ["WPS 138P/121 (St. 15)", 2, 2], ["WPS 121 (St. 16)", 3, 3], ["WPS 121 (St. 18)", 1, 1],
  ["Anlage einfahren, Hilfe Kiel", 1, 3], ["Kapa Schweißer MSG", 1, 2], ["Kapa Messtrupp", 2, 2],
  ["Kapa Schweißer UP", 2, 3], ["Kapa NDT", 1, 1], ["MPM (Station 18 belegt)", 3, 1],
  ["Material 475/309 Spante", 3, 3], ["Material Schuss 310", 1, 1], ["Zusatzwerkstoffe (3 Sorten)", 6, 6],
  ["Wendeschneidplatten RN-Fräse", 2, 2],
];
const labelsR = rk.map((r) => r[0]).reverse();
s.addChart(pres.charts.BAR, [
  { name: "Einschichtig", labels: labelsR, values: rk.map((r) => r[1]).reverse() },
  { name: "Zweischichtig", labels: labelsR, values: rk.map((r) => r[2]).reverse() },
], {
  x: 0.5, y: 1.2, w: 8.2, h: 5.65, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 60,
  chartColors: [HEX.accent5, HEX.accent1],
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 11, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt",
  valAxisHidden: true, valAxisMaxVal: 7, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.dk1,
  objectName: nm("risikochart"),
});
const tot = [["28", "Einschichtig", C.accent5], ["30", "Zweischichtig", C.accent1]];
tot.forEach(([n, l, c], i) => {
  const tx = 9.0 + i * 1.95;
  box(s, { x: tx, y: 1.3, w: 1.8, h: 1.45, fill: { color: C.background2 }, rounded: true });
  tb(s, n, { x: tx, y: 1.35, w: 1.8, h: 0.85, fontSize: 40, bold: true, color: c, align: "center", valign: "middle" });
  tb(s, l, { x: tx, y: 2.2, w: 1.8, h: 0.4, fontSize: 12, color: C.text1, align: "center" });
});
tb(s, "Skala je Punkt: 0 kein · 1 wenig · 2 mittel · 3 hoch. Zusatzwerkstoffe hier zusammengefasst (Megafil 940 M, S3 NiMo, 421 TT je 2).", { x: 9.0, y: 2.9, w: 3.8, h: 0.75, fontSize: 10.5, color: C.accent5 });
tb(s, "Wo sich die Szenarien unterscheiden", { x: 9.0, y: 3.75, w: 3.8, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
tb(s, [
  { text: "Einschichtig höher: MPM, Station 18 ist durch 475 belegt, wenn P800 sie braucht (3 zu 1).", options: { bullet: true, breakLine: true } },
  { text: "Zweischichtig höher: Anlagenfahrer aus Kiel (3 zu 1), Schweißer MSG (2 zu 1) und UP (3 zu 2).", options: { bullet: true, breakLine: true } },
  { text: "Hoch in beiden: WPS St. 16 und Spante 475/309 aus Kiel.", options: { bullet: true } },
], { x: 9.0, y: 4.15, w: 3.8, h: 2.7, fontSize: 12.5, color: C.text1, paraSpaceAfter: 6 });
s.addNotes("Blatt 'Auswertung Risiko'. Jeder Punkt wird je Szenario mit 0 bis 3 bewertet, die Summe steht in Zeile 34 (Formel). Ergebnis: einschichtig 28, zweischichtig 30. Die Summe allein ist kaum aussagekräftig – entscheidend ist die Art des Risikos. Einschichtig ist das Risiko strukturell (Station 18 ist belegt, 475/309 passt nicht in den Ablauf, das lässt sich nicht durch Personal lösen). Zweischichtig liegt das Risiko bei Personal (Schweißer mit gültiger Prüfung MSG und UP, Anlagenfahrer aus Kiel) – das lässt sich durch frühzeitige Personalplanung und Absprachen mit Kiel steuern. Gleich hoch in beiden: WPS 121 Station 16 (3) und die Beistellung der Spante 475/309 aus Kiel (3), Grund laut Kommentar: früher Bedarfstermin wegen Seetransport mit zwei Wochen Vorlauf. Gas (Erdgas, M21, Acetylen/Sauerstoff) ist mit 0 bewertet.");

// 14 Ressourcen
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Personal je Schicht begrenzt, was parallel laufen kann", { placeholder: "title" });
const res = [["10", "Schweißer (ES)"], ["7", "Schlosser (SB)"], ["2", "Qualitätssicherung"], ["1", "CNC-Bediener"], ["2 / 3", "NDT-Prüfer (1. / 3. Schicht)"]];
res.forEach(([n, l], i) => {
  const rx = 0.5 + i * 2.5;
  box(s, { x: rx, y: 1.4, w: 2.3, h: 2.0, fill: { color: i === 0 ? C.text2 : C.background2 }, rounded: true });
  tb(s, n, { x: rx, y: 1.55, w: 2.3, h: 1.0, fontSize: 44, bold: true, color: i === 0 ? C.accent1 : C.text2, align: "center", valign: "middle" });
  tb(s, l, { x: rx + 0.1, y: 2.6, w: 2.1, h: 0.65, fontSize: 13, color: i === 0 ? C.background1 : C.text1, align: "center" });
});
tb(s, "Laut Risikoblatt: Halle 20 + 40, „alle können alles“ – Personal ist flexibel zwischen den Stationen einsetzbar.", { x: 0.5, y: 3.55, w: 12.3, h: 0.35, fontSize: 12, color: C.accent5 });
const imp = [
  ["Was eine zweite Schicht bringt", ["Längsnähte LN 1 und LN 2 laufen auf St. 15 gleichzeitig statt nacheinander.", "Spante werden paarweise geschweißt (Spt 1+2, 3+4, 5+6).", "Durchlauf eines Schusses sinkt von ≈ 13 auf ≈ 9 Wochen."], C.accent3],
  ["Was sie kostet", ["Doppelt so viele Schweißer mit gültiger Prüfung (MSG und UP).", "Anlagenfahrer aus Kiel muss beide Schichten abdecken.", "Wartezeiten vor NDT und SLV-Laufzeiten verkürzen sich nicht."], C.accent6],
];
imp.forEach(([h, items, col], i) => {
  const ix = 0.5 + i * 6.25;
  box(s, { x: ix, y: 4.1, w: 6.05, h: 2.7, fill: { color: C.background2 }, rounded: true });
  s.addShape(pres.shapes.OVAL, { objectName: nm("punkt"), x: ix + 0.25, y: 4.3, w: 0.38, h: 0.38, fill: { color: col }, line: { type: "none" } });
  tb(s, h, { x: ix + 0.8, y: 4.28, w: 5, h: 0.42, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  tb(s, items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })), { x: ix + 0.3, y: 4.85, w: 5.55, h: 1.85, fontSize: 13.5, color: C.text1, paraSpaceAfter: 6 });
});
s.addNotes("Personaltabelle im Blatt 'Auswertung Risiko' (Spalten M–O): pro Schicht 10 Schweißer (ES), 7 Schlosser/Schiffbauer (SB), 2 QS, 1 CNC, NDT 2 Personen in der 1. Schicht und 3 in der 3. Schicht. Durchlaufzeit Schuss 301 einschichtig: Start St. 12 am 19.10. bis Abgabe 18.01. ≈ 13 Wochen. Schuss 309 zweischichtig: Start Umbau 05.10. bis Abgabe 14.12. ≈ 10 Wochen, ab Montage 12.10. ≈ 9 Wochen.");

// 15 Empfehlung
S = sec("Empfehlung");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Empfehlung: zweischichtig, wenn drei Bedingungen erfüllt sind", { placeholder: "title" });
box(s, { x: 0.5, y: 1.3, w: 4.3, h: 5.45, fill: { color: C.text2 }, rounded: true });
tb(s, "Warum zweischichtig", { x: 0.8, y: 1.5, w: 3.8, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
tb(s, [
  { text: "Nur so ist 475/309 überhaupt im Plan, mit einer Woche Puffer zum FAT 22.01.27.", options: { bullet: true, breakLine: true } },
  { text: "Das Risiko einschichtig ist strukturell (Station 18), zweischichtig ist es steuerbar (Personal).", options: { bullet: true, breakLine: true } },
  { text: "P800 wird nicht später, 301 ist sogar 2 Wochen früher fertig.", options: { bullet: true } },
], { x: 0.8, y: 2.05, w: 3.8, h: 3.2, fontSize: 14.5, color: C.background1, paraSpaceAfter: 12 });
tb(s, "1 Woche", { x: 0.8, y: 5.35, w: 3.8, h: 0.75, fontSize: 36, bold: true, color: C.accent1 });
tb(s, "Puffer zwischen Rundnaht 309+310 und FAT", { x: 0.8, y: 6.1, w: 3.8, h: 0.5, fontSize: 13, color: C.background2 });
const conds = [
  ["WPS Station 16 bis 16.11. klären", "Mit der SLV Rostock ein Vorabergebnis vereinbaren oder 309-Spante auf St. 16 ca. eine Woche nach hinten schieben."],
  ["Personal für die 2. Schicht sofort sichern", "Zweite Schicht mit geprüften MSG- und UP-Schweißern besetzen, Anlagenfahrer Kiel für beide Schichten anfragen."],
  ["Material Kiel absichern", "Bedarfstermin Spante 475/309 und Schuss 310 war der 12.10. (Seetransport, 2 Wochen Vorlauf): Ist-Stand prüfen, Plan nachtakten."],
];
conds.forEach(([h, t], i) => {
  const cy = 1.3 + i * 1.3;
  circleNum(s, 5.15, cy + 0.15, 0.55, i + 1);
  tb(s, h, { x: 5.95, y: cy + 0.05, w: 6.85, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
  tb(s, t, { x: 5.95, y: cy + 0.47, w: 6.85, h: 0.7, fontSize: 13.5, color: C.text1 });
});
box(s, { x: 5.15, y: 5.3, w: 7.65, h: 1.45, fill: { color: C.background2 }, rounded: true });
tb(s, "Zusätzlich für alle Szenarien", { x: 5.4, y: 5.4, w: 7.2, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
tb(s, "Arbeitsproben der Stationen 2, 7, 15 und 16 so vorziehen, dass das SLV-Ergebnis (5 Wochen) vor dem Fertigungsstart vorliegt, und diese Lücken in die Risikoliste aufnehmen.", { x: 5.4, y: 5.78, w: 7.2, h: 0.9, fontSize: 13, color: C.text1 });
s.addNotes("Im Excel ist die Empfehlung ('Empfehlung: Damit …') noch nicht ausformuliert. Aus den Daten ergibt sich: Wenn 475/309 zum FAT 22.01.2027 fertig sein soll, gibt es nur den zweischichtigen Ablauf. Dieser ist an drei Bedingungen geknüpft. Achtung zum Datum: Einige Termine des zweischichtigen Plans (Umbau St. 12 ab 05.10., Spante 309 bis 12.10.) liegen heute (31.10.2026) bereits in der Vergangenheit. Vor der Entscheidung muss der Ist-Stand geprüft und der Plan neu getaktet werden. Diese Empfehlung ist eine Ableitung des Projektmanagements aus dem Plan und muss mit Fertigung, Schweißaufsicht und QS abgestimmt werden.");

// 16 Nächste Schritte + Datenqualität
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Nächste Schritte und offene Punkte im Plan", { placeholder: "title" });
const next = [
  ["Sofort", "Ist-Stand Testschuss 1 und 309 gegen den Plan abgleichen; Termin SLV-Ergebnis Testschuss 1 verbindlich erfragen."],
  ["KW 45", "Entscheidung 1- oder 2-schichtig im Projektteam treffen und die fehlende Empfehlung im Plan ergänzen."],
  ["KW 46", "Personal 2. Schicht und Anlagenfahrer Kiel bestätigen; Material Kiel nachverfolgen."],
  ["Bis Dez 26", "Terminplan Schweißproben so anpassen, dass alle AP-Ergebnisse vor Fertigungsstart liegen."],
];
next.forEach(([w, t], i) => {
  const ny = 1.35 + i * 0.95;
  box(s, { x: 0.5, y: ny, w: 1.5, h: 0.75, fill: { color: C.accent1 }, rounded: true });
  tb(s, w, { x: 0.5, y: ny, w: 1.5, h: 0.75, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  tb(s, t, { x: 2.2, y: ny, w: 5.2, h: 0.75, fontSize: 13.5, color: C.text1, valign: "middle" });
});
box(s, { x: 7.8, y: 1.35, w: 5.0, h: 5.4, fill: { color: C.background2 }, rounded: true });
tb(s, "Im Excel zu bereinigen", { x: 8.05, y: 1.5, w: 4.5, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
tb(s, [
  { text: "Empfehlung in Zelle J1 bricht nach „Damit“ ab.", options: { bullet: true, breakLine: true } },
  { text: "Übersicht St. 15: AP-Datum „16.04.207“.", options: { bullet: true, breakLine: true } },
  { text: "Übersicht St. 16: AP „18.05.2026“, vermutlich 2027.", options: { bullet: true, breakLine: true } },
  { text: "WPS St. 15 einschichtig: Kommentar nennt 27.10., Längsnähte starten aber erst 10.11.", options: { bullet: true, breakLine: true } },
  { text: "Einschichtig: Rundnaht 301+302 nur als Text, ohne Farbe.", options: { bullet: true, breakLine: true } },
  { text: "Kommentar Station 9 und Rahmenspant teilweise kopiert.", options: { bullet: true } },
], { x: 8.05, y: 2.0, w: 4.55, h: 4.6, fontSize: 13, color: C.text1, paraSpaceAfter: 6 });
s.addNotes("Hinweise zur Datenqualität, damit der Plan als Entscheidungsgrundlage taugt. Die genannten Unstimmigkeiten sind Schreib- oder Übertragungsfehler und ändern die Kernaussage nicht. Sie sollten vor der Weitergabe an Kunden oder die SLV korrigiert werden.");

// 17 Glossar
S = sec("Anhang");
s = pres.addSlide({ masterName: "INHALT", sectionTitle: S });
s.addText("Anhang: Abkürzungen im Kreuzchenplan", { placeholder: "title" });
const gl = [
  ["WPS", "Schweißanweisung (Welding Procedure Specification)"], ["VP", "Verfahrensprüfung, qualifiziert die WPS"],
  ["AP", "Arbeitsprobe unter Fertigungsbedingungen"], ["SLV", "Schweißtechnische Lehr- und Versuchsanstalt Rostock"],
  ["SAT", "Abnahme der Anlage vor Ort"], ["121 / 138P", "Unterpulverschweißen / Metall-Fülldraht"],
  ["MSG / UP", "Metall-Schutzgas / Unterpulver"], ["LN / RN", "Längsnaht / Rundnaht"],
  ["Spt / DK", "Spant / Deck"], ["DL / SL", "Decklage / Schleifen"],
  ["VT · UT · RT · MT · PT", "Sicht-, Ultraschall-, Röntgen-, Magnetpulver-, Eindringprüfung"], ["NDT", "Zerstörungsfreie Prüfung"],
  ["FW", "Kehlnaht (Fillet Weld)"], ["MPM", "Projekt- und Fertigungsplanung, hier: Stationsbelegung"],
  ["1.3964 / 1.6782", "Werkstoffnummern der zu schweißenden Stähle"], ["P800 · 475 · P488", "Projekte bzw. Bauvorhaben im Plan"],
];
gl.forEach(([k, v], i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const gx = 0.5 + col * 6.25, gy = 1.3 + row * 0.68;
  box(s, { x: gx, y: gy, w: 6.05, h: 0.58, fill: { color: row % 2 ? C.background1 : C.background2 } });
  tb(s, k, { x: gx + 0.15, y: gy, w: 1.85, h: 0.58, fontSize: 13, bold: true, color: C.accent1, valign: "middle" });
  tb(s, v, { x: gx + 2.05, y: gy, w: 3.9, h: 0.58, fontSize: 12.5, color: C.text1, valign: "middle" });
});
s.addNotes("Die Abkürzungen sind aus dem Excel übernommen. Die Bedeutung von 'ES' (z. B. 'innen ES LN', 'Kapa MA mit ES Prüfung') ist im Plan nicht erklärt; hier als Schweißen/Schweißer interpretiert. MPM wird im Risikoblatt im Zusammenhang mit der Belegung von Station 18 verwendet.");

(async () => {
  await pres.writeFile({ fileName: "/tmp/build/Kreuzchenplan_Schweissproben_Praesentation.pptx" });
  await applyTheme("/tmp/build/Kreuzchenplan_Schweissproben_Praesentation.pptx", THEME);
  console.log("done");
})();
