// Genera la presentazione "Cuore, respiro e primo soccorso".
// Ogni riquadro è UN solo oggetto (forma con il testo dentro): niente testi
// appoggiati sopra le forme e niente icone sovrapposte, così la disposizione
// resta identica in PowerPoint desktop, web e mobile.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const JSZip = require("jszip");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Scienze Motorie",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E1E24", lt1: "FFFFFF", dk2: "6E0F1E", lt2: "F3F4F6",
    accent1: "C8102E", accent2: "1F6F8B", accent3: "B8730F", accent4: "2E7D4F",
    accent5: "5A6472", accent6: "3A3F8F", hlink: "1F6F8B", folHlink: "6E0F1E",
  },
};
const H = THEME.colors;
const PINK = "F2C9CF";
const T = { red: "FBE9EC", teal: "E4F1F5", amber: "FCF1E1", green: "E6F3EB", gray: "F3F4F6", indigo: "E9EAF6" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Cuore, respiro e primo soccorso";
pres.subject = "Scienze motorie";
pres.author = "Scienze motorie";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const W = 13.333, M = 0.6, CW = W - 2 * M;
const PAD = [12, 12, 10, 10]; // margini interni dei riquadri (pt): sx, dx, basso, alto

// ---------- Layout ----------
pres.defineSlideMaster({
  title: "Titolo",
  background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.0, w: 11, h: 2.1, fontSize: 48, bold: true, color: "FFFFFF", valign: "bottom", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.35, w: 11, h: 1.2, fontSize: 22, color: PINK, valign: "top", align: "left", margin: 0 }, text: "Sottotitolo" } },
  ],
});
pres.defineSlideMaster({
  title: "Sezione",
  background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "body", type: "body", x: M, y: 1.9, w: 11, h: 0.5, fontSize: 18, bold: true, color: PINK, charSpacing: 3, align: "left", valign: "bottom", margin: 0 }, text: "Parte" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.5, w: 12, h: 1.6, fontSize: 44, bold: true, color: "FFFFFF", valign: "top", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo sezione" } },
  ],
  slideNumber: { x: W - 1.1, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: PINK, align: "right" },
});
pres.defineSlideMaster({
  title: "Contenuto",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: CW, h: 0.8, fontSize: 32, bold: true, color: H.dk2, valign: "middle", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo" } },
    { text: { text: "Scienze motorie", options: { x: M, y: 6.95, w: 4, h: 0.35, fontSize: 11, color: H.accent5, margin: 0 } } },
  ],
  slideNumber: { x: W - 1.1, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: H.accent5, align: "right" },
});

// ---------- Helper ----------
const S = { intro: "Introduzione", cardio: "Apparato cardiocircolatorio", resp: "Apparato respiratorio", ps: "Il primo soccorso", loco: "Traumi dell'apparato locomotore", em: "Le emergenze", fine: "Conclusione" };
let curSection = null;
function section(name) { pres.addSection({ title: name }); curSection = name; }

function contentSlide(title) {
  const s = pres.addSlide({ masterName: "Contenuto", sectionTitle: curSection });
  s.addText(title, { placeholder: "title" });
  return s;
}

function sectionSlide(part, title, topics, notes) {
  const s = pres.addSlide({ masterName: "Sezione", sectionTitle: curSection });
  s.addText(part, { placeholder: "body" });
  s.addText(title, { placeholder: "title" });
  s.addText(topics.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i < topics.length - 1, paraSpaceAfter: 6 } })),
    { x: M, y: 4.3, w: 11, h: 2.4, fontSize: 18, color: "FFFFFF", valign: "top", margin: 0, isTextBox: true, objectName: "Argomenti" });
  s.addNotes(notes);
  return s;
}

function lead(s, text, h = 0.75) {
  s.addText(text, { x: M, y: 1.2, w: CW, h, fontSize: 17, italic: true, color: H.dk1, margin: 0, valign: "top", isTextBox: true, objectName: "Introduzione" });
}

// "**Etichetta:** resto" -> etichetta in grassetto
function runs(t, base) {
  const m = t.match(/^\*\*(.+?)\*\*(.*)$/);
  if (!m) return [{ text: t, options: { ...base } }];
  return [{ text: m[1], options: { ...base, bold: true } }, { text: m[2], options: { ...base, bullet: undefined } }];
}

// Riquadro: un'unica forma arrotondata con intestazione e punti elenco dentro
function box(s, { x, y, w, h, fill, head, headColor, headSize = 18, items = [], fontSize = 16, color = H.dk1, numbered = false, align = "left", space = 8, valign = "top", name }) {
  const paras = [];
  if (head) paras.push({ text: head, options: { bold: true, color: headColor, fontSize: headSize, paraSpaceAfter: items.length ? 8 : 0, breakLine: items.length > 0, align } });
  items.forEach((t, i) => {
    const bullet = numbered ? { type: "number", indent: 20 } : { indent: 16 };
    const r = runs(t, { color, fontSize, paraSpaceAfter: space, align });
    r[0].options.bullet = align === "center" ? undefined : bullet;
    r[r.length - 1].options.breakLine = i < items.length - 1;
    paras.push(...r);
  });
  s.addText(paras, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x, y, w, h, fill: { color: fill }, line: { color: fill }, margin: PAD, valign, objectName: name || head || "Riquadro" });
}

// Riquadro "dato": numero grande + etichetta
function stat(s, { x, y, w, h, fill, num, label, numColor, numSize = 34, labelColor = H.dk1 }) {
  s.addText([
    { text: num, options: { bold: true, fontSize: numSize, color: numColor, fontFace: THEME.headFontFace, breakLine: true, paraSpaceAfter: 4 } },
    { text: label, options: { fontSize: 15, color: labelColor } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x, y, w, h, fill: { color: fill }, line: { color: fill }, align: "center", valign: "middle", margin: PAD, objectName: "Dato " + num });
}

// Cerchio con numero (un solo oggetto)
function badge(s, n, x, y, d, fill, size = 18) {
  s.addText(String(n), { shape: pres.shapes.OVAL, x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, fontSize: size, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, objectName: "Numero " + n });
}

// Passo numerato: cerchio + testo (titolo in grassetto e descrizione) a fianco
function step(s, n, title, desc, { x, y, w, h, color, d = 0.55, descSize = 15 }) {
  badge(s, n, x, y + 0.05, d, color);
  s.addText([
    { text: title, options: { bold: true, color: H.dk2, fontSize: 17, breakLine: true } },
    { text: desc, options: { color: H.dk1, fontSize: descSize } },
  ], { x: x + d + 0.25, y, w: w - d - 0.25, h, valign: "top", margin: 0, isTextBox: true, objectName: "Passo " + n });
}

function cols3(top = 2.1, h = 4.6, gap = 0.3) {
  const w = (CW - 2 * gap) / 3;
  return [0, 1, 2].map((i) => ({ x: M + i * (w + gap), y: top, w, h }));
}

const SEGNI = (items) => ({ fill: T.gray, head: "Come si riconosce", headColor: H.accent5, items });
const FARE = (items) => ({ fill: T.green, head: "Cosa fare", headColor: H.accent4, items });
const NONFARE = (items) => ({ fill: T.red, head: "Cosa NON fare", headColor: H.accent1, items });
const CAUSE = (items) => ({ fill: T.amber, head: "Cause", headColor: H.accent3, items });
const ATT = (head, items) => ({ fill: T.amber, head, headColor: H.accent3, items });

function scheda(title, def, cols, notes) {
  const s = contentSlide(title);
  lead(s, def, 0.8);
  const pos = cols3(2.15, 4.6);
  cols.forEach((c, i) => box(s, { ...pos[i], fontSize: 17, space: 10, headSize: 20, ...c }));
  s.addNotes(notes);
  return s;
}

// =========================================================
// INTRODUZIONE
// =========================================================
section(S.intro);
{
  const s = pres.addSlide({ masterName: "Titolo", sectionTitle: curSection });
  s.addText("Cuore, respiro e primo soccorso", { placeholder: "title" });
  s.addText("Apparato cardiocircolatorio e respiratorio, esercizio fisico e gestione dei traumi più comuni", { placeholder: "body" });
  s.addText("Scienze motorie", { x: M, y: 6.6, w: 6, h: 0.4, fontSize: 14, color: PINK, margin: 0, isTextBox: true, objectName: "Materia" });
  s.addNotes("Presentazione di scienze motorie. Tre grandi temi: apparato cardiocircolatorio, apparato respiratorio e primo soccorso, con particolare attenzione ai traumi sportivi e alle emergenze.");
}
{
  const s = contentSlide("Di cosa parliamo");
  const rows = [
    [H.accent1, T.red, "Apparato cardiocircolatorio", "Il cuore, la circolazione, il sangue e l'adattamento all'esercizio fisico"],
    [H.accent2, T.teal, "Apparato respiratorio", "Gli organi, la meccanica del respiro e la respirazione durante lo sforzo"],
    [H.accent4, T.green, "Il primo soccorso", "Come intervenire: P.A.S., valutazione, contusioni, ferite ed emorragie"],
    [H.accent3, T.amber, "Traumi dell'apparato locomotore", "Muscoli, tendini, articolazioni e ossa: dalla contrattura alla frattura"],
    [H.accent6, T.indigo, "Le emergenze", "Lipotimia, arresto cardiaco, shock, trauma cranico, soffocamento, ipertermia, ipotermia"],
  ];
  const y0 = 1.4, rh = 0.92, gap = 0.12;
  rows.forEach(([col, tint, t, d], i) => {
    const y = y0 + i * (rh + gap);
    badge(s, i + 1, M, y + 0.16, 0.6, col, 20);
    s.addText([
      { text: t, options: { bold: true, fontSize: 19, color: col, breakLine: true } },
      { text: d, options: { fontSize: 15, color: H.dk1 } },
    ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: M + 0.85, y, w: CW - 0.85, h: rh, fill: { color: tint }, line: { color: tint }, valign: "middle", margin: [14, 12, 6, 6], objectName: "Parte " + (i + 1) });
  });
  s.addNotes("Le parti 3, 4 e 5 sono il cuore della lezione: sapere cosa fare, e cosa non fare, davanti a un infortunio può fare la differenza.");
}

// =========================================================
// PARTE 1 — CARDIOCIRCOLATORIO
// =========================================================
section(S.cardio);
sectionSlide("PARTE 1", "Apparato cardiocircolatorio", ["Il cuore", "La circolazione sanguigna", "Il sangue", "Apparato cardiocircolatorio ed esercizio fisico"],
  "L'apparato cardiocircolatorio è formato da cuore, vasi sanguigni e sangue. Il suo compito è portare ossigeno e nutrienti a tutte le cellule e rimuovere le sostanze di scarto.");

{
  const s = contentSlide("Il cuore");
  lead(s, "Un muscolo cavo grande quanto un pugno, posto nel torace tra i polmoni, leggermente spostato a sinistra. Funziona come una pompa che spinge il sangue in tutto il corpo.", 0.8);
  box(s, { x: M, y: 2.15, w: 7.0, h: 4.6, fill: T.gray, head: "Com'è fatto", headColor: H.accent1, headSize: 20, fontSize: 17, space: 10, items: [
    "**4 cavità:** 2 atri (in alto) che ricevono il sangue e 2 ventricoli (in basso) che lo spingono fuori",
    "**Setto:** divide il cuore destro (sangue povero di ossigeno) dal sinistro (ricco di ossigeno)",
    "**4 valvole** (tricuspide, mitrale, polmonare, aortica): il sangue scorre in una sola direzione",
    "**3 strati:** pericardio (esterno), miocardio (muscolo), endocardio (interno)",
    "**Ciclo cardiaco:** sistole (contrazione) e diastole (rilassamento)",
  ] });
  const gx = 7.9, gw = (W - M - gx - 0.3) / 2, gh = 2.15;
  [["≈ 300 g", "peso del cuore di un adulto"], ["60–80", "battiti al minuto a riposo"], ["≈ 5 litri", "di sangue pompati al minuto"], ["100.000", "battiti circa al giorno"]].forEach(([n, l], i) => {
    stat(s, { x: gx + (i % 2) * (gw + 0.3), y: 2.15 + Math.floor(i / 2) * (gh + 0.3), w: gw, h: gh, fill: T.red, num: n, label: l, numColor: H.accent1, numSize: 30 });
  });
  s.addNotes("Il cuore lavora senza sosta per tutta la vita. Il battito nasce da un pacemaker naturale, il nodo seno-atriale, che genera l'impulso elettrico senza bisogno di comandi volontari. Le valvole sono come porte a senso unico: il rumore che sentiamo con lo stetoscopio è la loro chiusura.");
}

{
  const s = contentSlide("La circolazione sanguigna");
  lead(s, "Nell'uomo la circolazione è doppia (due circuiti) e completa: il sangue ricco di ossigeno non si mescola mai con quello povero di ossigeno.", 0.8);
  const pw = (CW - 0.3) / 2;
  box(s, { x: M, y: 2.1, w: pw, h: 2.75, fill: T.teal, head: "Piccola circolazione (polmonare)", headColor: H.accent2, numbered: true, space: 4, fontSize: 17, items: ["Ventricolo destro", "Arterie polmonari", "**Polmoni:** il sangue prende O₂ e lascia CO₂", "Vene polmonari", "Atrio sinistro"] });
  box(s, { x: M + pw + 0.3, y: 2.1, w: pw, h: 2.75, fill: T.red, head: "Grande circolazione (sistemica)", headColor: H.accent1, numbered: true, space: 4, fontSize: 17, items: ["Ventricolo sinistro", "Aorta e arterie", "**Organi e muscoli:** il sangue cede O₂ e raccoglie CO₂", "Vene cave", "Atrio destro"] });
  const pos = cols3(5.05, 1.7);
  [["Arterie", "Portano il sangue dal cuore agli organi. Pareti spesse ed elastiche."], ["Vene", "Riportano il sangue al cuore. Hanno valvole che impediscono il reflusso."], ["Capillari", "Vasi sottilissimi dove avvengono gli scambi con le cellule."]].forEach(([t, d], i) => {
    s.addText([{ text: t, options: { bold: true, color: H.dk2, fontSize: 17, breakLine: true, paraSpaceAfter: 4 } }, { text: d, options: { fontSize: 16, color: H.dk1 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, ...pos[i], fill: { color: T.gray }, line: { color: T.gray }, valign: "top", margin: PAD, objectName: t });
  });
  s.addNotes("Attenzione: le arterie si chiamano così perché partono dal cuore, non perché portano sangue ossigenato. Le arterie polmonari portano infatti sangue povero di ossigeno ai polmoni, e le vene polmonari riportano al cuore sangue ricco di ossigeno.");
}

{
  const s = contentSlide("Il sangue");
  lead(s, "Un tessuto liquido: in un adulto ce ne sono circa 5 litri (7–8% del peso corporeo). Trasporta ossigeno, nutrienti, ormoni e scarti, difende l'organismo e regola la temperatura.", 0.8);
  const by = 2.1, bh = 0.7;
  s.addText([{ text: "Plasma 55%", options: { bold: true } }, { text: " · acqua, proteine, sali, zuccheri" }], { shape: pres.shapes.RECTANGLE, x: M, y: by, w: CW * 0.55, h: bh, fill: { color: "E9C46A" }, line: { color: "E9C46A" }, fontSize: 16, color: H.dk1, valign: "middle", margin: [12, 8, 0, 0], objectName: "Plasma" });
  s.addText([{ text: "Cellule 45%", options: { bold: true } }, { text: " · globuli e piastrine" }], { shape: pres.shapes.RECTANGLE, x: M + CW * 0.55, y: by, w: CW * 0.45, h: bh, fill: { color: H.accent1 }, line: { color: H.accent1 }, fontSize: 16, color: "FFFFFF", valign: "middle", margin: [12, 8, 0, 0], objectName: "Parte corpuscolata" });
  const pos = cols3(3.0, 2.65);
  box(s, { ...pos[0], fill: T.red, head: "Globuli rossi (eritrociti)", headColor: H.accent1, fontSize: 17, space: 8, items: ["Contengono emoglobina, che trasporta l'ossigeno", "4,5–5,5 milioni per mm³", "Vivono circa 120 giorni"] });
  box(s, { ...pos[1], fill: T.indigo, head: "Globuli bianchi (leucociti)", headColor: H.accent6, fontSize: 17, space: 8, items: ["Difendono da virus e batteri", "4.000–10.000 per mm³", "Aumentano durante un'infezione"] });
  box(s, { ...pos[2], fill: T.amber, head: "Piastrine", headColor: H.accent3, fontSize: 17, space: 8, items: ["Frammenti di cellule che fermano le emorragie", "150.000–400.000 per mm³", "Formano il coagulo"] });
  const fw = (CW - 3 * 0.25) / 4;
  [["Trasporto", "ossigeno, nutrienti, ormoni"], ["Difesa", "contro le infezioni"], ["Coagulazione", "chiude le ferite"], ["Termoregolazione", "distribuisce il calore"]].forEach(([t, d], i) => {
    s.addText([{ text: t, options: { bold: true, color: "FFFFFF", fontSize: 15, breakLine: true } }, { text: d, options: { color: PINK, fontSize: 13 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + i * (fw + 0.25), y: 5.9, w: fw, h: 0.85, fill: { color: H.dk2 }, line: { color: H.dk2 }, align: "center", valign: "middle", margin: [6, 6, 4, 4], objectName: "Funzione " + t });
  });
  s.addNotes("Con l'allenamento, soprattutto di resistenza, aumentano il volume di sangue e la quantità di emoglobina: il corpo trasporta più ossigeno ai muscoli. Lo stesso accade vivendo in alta quota.");
}

{
  const s = contentSlide("Il cuore durante l'esercizio fisico");
  lead(s, "Sotto sforzo i muscoli chiedono molto più ossigeno: il cuore risponde battendo più spesso e più forte.", 0.5);
  const pos = cols3(1.85, 2.35);
  [["Frequenza cardiaca", "70", "180–200", "battiti al minuto"], ["Gittata sistolica", "70", "120", "ml di sangue per battito (atleti fino a 200)"], ["Portata cardiaca", "5", "20–25", "litri di sangue al minuto (atleti fino a 35)"]].forEach(([t, a, b, u], i) => {
    s.addText([
      { text: t, options: { bold: true, fontSize: 17, color: H.dk2, breakLine: true, paraSpaceAfter: 6 } },
      { text: "Riposo: ", options: { fontSize: 16, color: H.accent5 } }, { text: a, options: { fontSize: 24, bold: true, color: H.accent5, breakLine: true } },
      { text: "Sforzo: ", options: { fontSize: 16, color: H.accent1 } }, { text: b, options: { fontSize: 24, bold: true, color: H.accent1, breakLine: true, paraSpaceAfter: 4 } },
      { text: u, options: { fontSize: 14, color: H.accent5 } },
    ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, ...pos[i], fill: { color: T.red }, line: { color: T.red }, valign: "middle", margin: PAD, objectName: t });
  });
  box(s, { x: M, y: 4.4, w: 7.8, h: 2.35, fill: T.gray, head: "Con l'allenamento il cuore si adatta", headColor: H.accent1, fontSize: 16, space: 5, items: [
    "**Bradicardia da sport:** a riposo batte più lentamente (anche 40–60 al minuto)",
    "**Cuore d'atleta:** più grande e forte, spinge più sangue a ogni battito",
    "**Più capillari** nei muscoli e più sangue in circolo",
    "**Pressione** a riposo più bassa e recupero più rapido",
  ] });
  s.addText([
    { text: "Frequenza cardiaca massima", options: { bold: true, fontSize: 16, color: PINK, breakLine: true } },
    { text: "220 − età", options: { bold: true, fontSize: 34, color: "FFFFFF", fontFace: THEME.headFontFace, breakLine: true } },
    { text: "A 16 anni ≈ 204 battiti al minuto. Il polso si misura al polso o al collo.", options: { fontSize: 14, color: "FFFFFF" } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + 8.1, y: 4.4, w: CW - 8.1, h: 2.35, fill: { color: H.dk2 }, line: { color: H.dk2 }, valign: "top", margin: PAD, objectName: "Frequenza massima" });
  s.addNotes("Portata cardiaca = frequenza cardiaca × gittata sistolica. Il sedentario aumenta la portata soprattutto aumentando la frequenza; l'atleta, avendo una gittata maggiore, ottiene la stessa portata con meno battiti. La formula 220 meno età è una stima: serve per dosare l'intensità dell'allenamento.");
}

// =========================================================
// PARTE 2 — RESPIRATORIO
// =========================================================
section(S.resp);
sectionSlide("PARTE 2", "Apparato respiratorio", ["Gli organi della respirazione", "La respirazione", "La respirazione durante l'esercizio"],
  "L'apparato respiratorio porta l'ossigeno dall'aria al sangue ed elimina l'anidride carbonica prodotta dalle cellule.");

{
  const s = contentSlide("Gli organi della respirazione");
  box(s, { x: M, y: 1.35, w: 7.4, h: 5.4, fill: T.teal, head: "Il percorso dell'aria", headColor: H.accent2, headSize: 20, numbered: true, fontSize: 18, space: 14, items: [
    "**Naso e cavità nasali:** filtrano, scaldano e umidificano l'aria",
    "**Faringe:** incrocio tra la via dell'aria e quella del cibo",
    "**Laringe:** contiene le corde vocali; l'epiglottide la chiude quando deglutiamo",
    "**Trachea:** tubo di anelli di cartilagine lungo circa 12 cm",
    "**Bronchi e bronchioli:** si ramificano come un albero dentro i polmoni",
    "**Alveoli:** circa 300 milioni di sacchetti dove avviene lo scambio dei gas",
  ] });
  const rx = M + 7.7, rw = W - M - rx;
  box(s, { x: rx, y: 1.35, w: rw, h: 2.55, fill: T.gray, head: "I polmoni", headColor: H.dk2, headSize: 20, fontSize: 17, space: 8, items: ["Destro con 3 lobi, sinistro con 2", "Avvolti dalla pleura", "Superficie degli alveoli: 70–100 m²"] });
  box(s, { x: rx, y: 4.15, w: rw, h: 2.6, fill: T.gray, head: "I muscoli respiratori", headColor: H.dk2, headSize: 20, fontSize: 17, space: 8, items: ["**Diaframma:** il principale", "**Intercostali:** muovono le costole", "**Accessori** (collo, addome): sotto sforzo"] });
  s.addNotes("Naso, faringe e laringe sono le vie aeree superiori; trachea, bronchi e polmoni quelle inferiori. Respirare col naso è importante perché l'aria arriva ai polmoni filtrata, calda e umida. Nello sforzo intenso si respira anche con la bocca per far entrare più aria.");
}

{
  const s = contentSlide("La respirazione");
  lead(s, "Un atto respiratorio ha due fasi. A riposo respiriamo 12–16 volte al minuto, circa mezzo litro d'aria per volta.", 0.5);
  const pw = (CW - 0.3) / 2;
  box(s, { x: M, y: 1.85, w: pw, h: 2.55, fill: T.teal, head: "Inspirazione: fase attiva", headColor: H.accent2, headSize: 20, fontSize: 18, space: 10, items: ["Il diaframma si contrae e si abbassa", "Gli intercostali sollevano le costole", "Il torace si allarga e l'aria entra"] });
  box(s, { x: M + pw + 0.3, y: 1.85, w: pw, h: 2.55, fill: T.gray, head: "Espirazione: passiva a riposo", headColor: H.accent5, headSize: 20, fontSize: 18, space: 10, items: ["I muscoli si rilassano, il diaframma risale", "Il torace si restringe e l'aria esce", "Sotto sforzo diventa attiva (addominali)"] });
  box(s, { x: M, y: 4.6, w: CW, h: 2.15, fill: H.dk2, head: "Gli scambi dei gas avvengono per diffusione", headColor: "FFFFFF", color: "FFFFFF", fontSize: 16, space: 6, items: [
    "**Negli alveoli:** l'ossigeno passa nel sangue, l'anidride carbonica passa nell'aria da espirare",
    "**Nei tessuti:** il sangue cede ossigeno alle cellule e raccoglie anidride carbonica",
    "**Ossigeno nell'aria:** 21% in quella inspirata, circa 16% in quella espirata",
  ] });
  s.addNotes("La respirazione è controllata automaticamente dai centri respiratori nel tronco encefalico, che misurano soprattutto la quantità di anidride carbonica nel sangue. Possiamo però controllarla anche volontariamente, ad esempio in apnea.");
}

{
  const s = contentSlide("La respirazione durante l'esercizio");
  const chartW = 6.3;
  s.addChart(pres.charts.BAR, [{ name: "Aria respirata al minuto", labels: ["Riposo", "Corsa leggera", "Sforzo intenso", "Atleta al massimo"], values: [6, 40, 100, 150] }], {
    x: M, y: 1.3, w: chartW, h: 5.45, barDir: "col", chartColors: [H.accent2],
    showTitle: true, title: "Litri d'aria respirati al minuto (valori indicativi)", titleFontSize: 14, titleColor: H.dk1, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: H.dk1, dataLabelFontFace: "+mn-lt", dataLabelFontBold: true,
    catAxisLabelColor: H.accent5, valAxisLabelColor: H.accent5, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 175, valAxisMajorUnit: 25, barGapWidthPct: 60, showLegend: false,
  });
  const rx = M + chartW + 0.35, rw = W - M - rx;
  box(s, { x: rx, y: 1.35, w: rw, h: 3.95, fill: T.teal, head: "Cosa cambia sotto sforzo", headColor: H.accent2, headSize: 20, fontSize: 17, space: 10, items: [
    "**Frequenza:** da 12–16 fino a 40–50 respiri al minuto",
    "**Volume per respiro:** da 0,5 fino a 2–3 litri",
    "**Debito di ossigeno:** finito lo sforzo continuiamo ad ansimare per recuperare",
    "**Con l'allenamento:** muscoli respiratori più efficienti e più ossigeno usato (VO₂max)",
  ] });
  box(s, { x: rx, y: 5.5, w: rw, h: 1.25, fill: T.gray, head: "Consiglio", headColor: H.dk2, headSize: 17, fontSize: 16, items: ["Respira con ritmo regolare ed espira nella fase di sforzo"] });
  s.addNotes("La ventilazione polmonare è il prodotto di frequenza respiratoria per volume corrente. Trattenere il fiato durante uno sforzo massimale fa salire molto la pressione: meglio espirare nella fase di spinta, ad esempio quando si solleva un peso.");
}

// =========================================================
// PARTE 3 — PRIMO SOCCORSO
// =========================================================
section(S.ps);
sectionSlide("PARTE 3", "Il primo soccorso", ["Come si presta il primo soccorso: P.A.S.", "Valutare la vittima", "Contusioni, ferite, emorragie", "Emorragia nasale"],
  "Il primo soccorso è l'aiuto che chiunque può dare a una persona infortunata o malata in attesa dei soccorsi. Prestarlo è un dovere: in Italia l'omissione di soccorso è un reato (art. 593 del Codice Penale). Anche solo chiamare i soccorsi è già prestare aiuto.");

{
  const s = contentSlide("Come si presta il primo soccorso: P.A.S.");
  lead(s, "Tre azioni, sempre in questo ordine. Il soccorritore non deve mai diventare una seconda vittima.", 0.5);
  const pos = cols3(1.85, 4.9);
  box(s, { ...pos[0], fill: T.amber, head: "P · Proteggere", headColor: H.accent3, headSize: 24, fontSize: 18, space: 12, items: ["Valuta i pericoli: traffico, fuoco, corrente elettrica, gas, crolli", "Metti in sicurezza prima te stesso, poi la vittima", "Usa guanti monouso se c'è sangue", "Sposta la vittima solo se c'è un pericolo immediato"] });
  box(s, { ...pos[1], fill: T.red, head: "A · Avvertire", headColor: H.accent1, headSize: 24, fontSize: 18, space: 12, items: ["Chiama il 112, Numero Unico di Emergenza", "Di' dove sei, cosa è successo, quante persone sono coinvolte e come stanno", "Rispondi alle domande e non riattaccare per primo"] });
  box(s, { ...pos[2], fill: T.green, head: "S · Soccorrere", headColor: H.accent4, headSize: 24, fontSize: 18, space: 12, items: ["Valuta coscienza e respiro", "Esegui solo le manovre che conosci", "Rassicura la vittima, coprila e non lasciarla sola", "Non darle da bere né da mangiare"] });
  s.addNotes("Proteggere, Avvertire, Soccorrere. Prima di tutto la propria sicurezza: un soccorritore ferito non aiuta nessuno. Il 112 è gratuito e si può chiamare anche senza credito. Al telefono restare calmi e lasciare che sia l'operatore a chiudere la chiamata.");
}

{
  const s = contentSlide("Valutare la vittima");
  const steps = [
    ["Coscienza", "Chiamala ad alta voce e scuotila delicatamente per le spalle. Risponde?"],
    ["Respiro", "Testa indietro e mento in su; per 10 secondi guarda il torace, ascolta e senti il respiro (G.A.S.)"],
    ["Decidi", "In base a ciò che hai visto scegli cosa fare, e chiama il 112 se non l'hai già fatto"],
  ];
  steps.forEach(([t, d], i) => step(s, i + 1, t, d, { x: M, y: 1.4 + i * 1.0, w: CW, h: 0.9, color: H.dk2 }));
  const pos = cols3(4.55, 2.2);
  box(s, { ...pos[0], fill: T.green, head: "Cosciente", headColor: H.accent4, headSize: 20, fontSize: 17, items: ["Lasciala nella posizione in cui sta meglio, chiedi cosa è successo e tienila d'occhio"] });
  box(s, { ...pos[1], fill: T.amber, head: "Incosciente ma respira", headColor: H.accent3, headSize: 20, fontSize: 17, items: ["Posizione laterale di sicurezza (se non c'è un trauma), 112 e controlla il respiro"] });
  box(s, { ...pos[2], fill: T.red, head: "Incosciente, non respira", headColor: H.accent1, headSize: 20, fontSize: 17, items: ["112, fai portare un defibrillatore (DAE) e inizia subito la RCP"] });
  s.addNotes("Posizione laterale di sicurezza: la vittima è su un fianco, con la testa estesa e la bocca rivolta verso il basso, così la lingua non chiude le vie aeree e un eventuale vomito può uscire. Non si usa se si sospetta un trauma alla colonna vertebrale. Il respiro va ricontrollato di continuo.");
}

scheda("Le contusioni",
  "Lesione dei tessuti molli (pelle, muscoli) causata da un urto, senza rottura della pelle. È il trauma più comune nello sport.",
  [
    SEGNI(["Dolore nel punto colpito", "Gonfiore", "Livido o ematoma (sangue uscito dai piccoli vasi)", "Arrossamento e calore", "Difficoltà a muovere la parte"]),
    FARE(["Interrompi l'attività", "Ghiaccio avvolto in un panno per 15–20 minuti, più volte al giorno", "Tieni la parte a riposo e sollevata", "Se il dolore è forte o non passa, vai dal medico"]),
    NONFARE(["Non mettere il ghiaccio direttamente sulla pelle", "Non applicare calore nei primi 2–3 giorni", "Non massaggiare la zona", "Non sottovalutare un colpo forte a testa, addome o torace"]),
  ],
  "Il freddo restringe i vasi sanguigni: limita il gonfiore e l'ematoma e riduce il dolore. Il calore invece dilata i vasi e nei primi giorni peggiora il gonfiore.");

scheda("Le ferite",
  "Interruzione della pelle. Tipi: abrasioni (sbucciature), ferite da taglio, da punta, lacero-contuse (pelle strappata da un urto).",
  [
    ATT("Quando è grave", ["Sanguina molto o a fiotti", "È profonda, lunga o con bordi aperti", "È sul viso, sulle mani o vicino agli occhi", "C'è un oggetto conficcato", "È un morso o è molto sporca"]),
    FARE(["Lavati le mani e metti i guanti", "Lava la ferita con acqua corrente o soluzione fisiologica", "Disinfetta dal centro verso l'esterno", "Copri con garza sterile e cerotto", "Ferita profonda o sporca: medico e antitetanica"]),
    NONFARE(["Non usare cotone: lascia fili nella ferita", "Non versare alcol nella ferita", "Non togliere oggetti conficcati: fissali e chiama il 112", "Non toccare la ferita a mani nude"]),
  ],
  "La vaccinazione antitetanica va richiamata ogni 10 anni. Il tetano è causato da un batterio presente nella terra e nella ruggine. Un oggetto conficcato può fare da tappo: toglierlo può scatenare un'emorragia. Per disinfettare meglio un disinfettante non alcolico.");

{
  const s = contentSlide("Le emorragie: i tipi");
  lead(s, "L'emorragia è l'uscita di sangue dai vasi. Perdere molto sangue in poco tempo è pericoloso e può portare allo shock.", 0.5);
  const a = cols3(1.85, 2.0);
  box(s, { ...a[0], fill: T.red, head: "Arteriosa", headColor: H.accent1, headSize: 20, fontSize: 17, items: ["Rosso vivo, esce a fiotti al ritmo del battito. È la più grave"] });
  box(s, { ...a[1], fill: T.red, head: "Venosa", headColor: H.dk2, headSize: 20, fontSize: 17, items: ["Rosso scuro, esce in modo continuo e lento"] });
  box(s, { ...a[2], fill: T.amber, head: "Capillare", headColor: H.accent3, headSize: 20, fontSize: 17, items: ["Esce a gocce, come in una sbucciatura. Si ferma facilmente"] });
  s.addText("In base a dove finisce il sangue", { x: M, y: 4.05, w: CW, h: 0.45, fontSize: 18, bold: true, color: H.dk2, margin: 0, isTextBox: true, objectName: "Sottotitolo" });
  const b = cols3(4.6, 2.15);
  box(s, { ...b[0], fill: T.gray, head: "Esterna", headColor: H.dk2, headSize: 20, fontSize: 17, items: ["Il sangue esce da una ferita ed è visibile"] });
  box(s, { ...b[1], fill: T.gray, head: "Interna", headColor: H.dk2, headSize: 20, fontSize: 17, items: ["Resta dentro il corpo: non si vede, si riconosce dai segni dello shock"] });
  box(s, { ...b[2], fill: T.gray, head: "Esteriorizzata", headColor: H.dk2, headSize: 20, fontSize: 17, items: ["È interna ma esce da naso, orecchio o bocca"] });
  s.addNotes("Un'emorragia interna è insidiosa perché non si vede: dopo un trauma forte all'addome o al torace bisogna sospettarla se compaiono pallore, sudore freddo, polso rapido e debole. In quel caso: 112 subito.");
}

{
  const s = contentSlide("Come fermare un'emorragia esterna");
  const steps = [
    ["Proteggiti", "Metti i guanti e fai sdraiare la vittima"],
    ["Comprimi", "Premi forte sulla ferita con garza o panno pulito per almeno 10 minuti, senza sollevarla"],
    ["Solleva", "Se è un arto e non c'è frattura, tienilo più in alto del cuore"],
    ["Aggiungi", "Se il sangue passa, metti altre garze sopra senza togliere le prime e fai un bendaggio stretto"],
    ["Comprimi l'arteria", "Se non basta, premi l'arteria a monte contro l'osso: braccio o inguine"],
  ];
  steps.forEach(([t, d], i) => step(s, i + 1, t, d, { x: M, y: 1.35 + i * 1.08, w: 7.6, h: 1.0, color: H.accent1 }));
  const rx = M + 8.0, rw = W - M - rx;
  box(s, { x: rx, y: 1.35, w: rw, h: 3.1, fill: H.dk2, head: "Il laccio emostatico", headColor: "FFFFFF", color: "FFFFFF", fontSize: 16, space: 8, items: ["Solo come ultima risorsa, per un'emorragia grave di un arto", "Si stringe a monte della ferita e si annota l'ora", "Non va mai allentato: lo toglie il medico"] });
  box(s, { x: rx, y: 4.65, w: rw, h: 2.1, fill: T.red, head: "Chiama il 112 se", headColor: H.accent1, fontSize: 16, space: 6, items: ["il sangue esce a fiotti o non si ferma", "la vittima è pallida, sudata e confusa"] });
  s.addNotes("La compressione diretta ferma la grande maggioranza delle emorragie esterne. L'errore più comune è sollevare la garza per guardare: così si stacca il coagulo che si stava formando.");
}

scheda("L'emorragia nasale (epistassi)",
  "Uscita di sangue dal naso, frequente nello sport. Può dipendere da un colpo, dall'aria secca, dal soffiarsi il naso con forza o dalla pressione alta.",
  [
    FARE(["Fai sedere la persona con la testa leggermente in avanti", "Stringi la parte morbida del naso per 10 minuti", "Falla respirare con la bocca", "Ghiaccio avvolto in un panno sul naso", "Se non si ferma in 20 minuti: pronto soccorso"]),
    NONFARE(["Non piegare la testa all'indietro: il sangue scende in gola e può far vomitare", "Non far sdraiare la persona", "Non soffiare il naso per alcune ore", "Non infilare cotone nelle narici"]),
    ATT("Attenzione", ["Dopo un colpo alla testa, sangue o liquido chiaro da naso o orecchie può indicare una frattura del cranio: chiama il 112", "Se il naso è storto dopo un colpo, potrebbe essere rotto"]),
  ],
  "È un errore molto diffuso mettere la testa all'indietro: il sangue non si ferma, scende in gola, può andare nelle vie aeree o essere ingerito e far vomitare. La testa va in avanti.");

// =========================================================
// PARTE 4 — APPARATO LOCOMOTORE
// =========================================================
section(S.loco);
sectionSlide("PARTE 4", "Traumi dell'apparato locomotore", ["Il protocollo P.R.I.C.E.", "Contrattura, stiramento, strappo, tendinopatie", "Distorsione, lussazione, frattura"],
  "Traumi a muscoli, tendini, articolazioni e ossa: sono gli infortuni più frequenti durante l'attività sportiva.");

{
  const s = contentSlide("Il protocollo P.R.I.C.E.");
  lead(s, "Le prime cure per quasi tutti i traumi di muscoli, tendini e articolazioni, nelle prime 48–72 ore.", 0.5);
  const gap = 0.25, tw = (CW - 4 * gap) / 5;
  [["P", "Protezione", "Stop all'attività, proteggi la parte"], ["R", "Riposo", "Non caricare e non muovere la parte"], ["I", "Ice: ghiaccio", "15–20 minuti ogni 2–3 ore, in un panno"], ["C", "Compressione", "Benda elastica, non troppo stretta"], ["E", "Elevazione", "Parte sollevata sopra il cuore"]].forEach(([l, t, d], i) => {
    s.addText([
      { text: l, options: { bold: true, fontSize: 60, color: H.accent3, fontFace: THEME.headFontFace, breakLine: true } },
      { text: t, options: { bold: true, fontSize: 19, color: H.dk2, breakLine: true, paraSpaceAfter: 8 } },
      { text: d, options: { fontSize: 17, color: H.dk1 } },
    ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + i * (tw + gap), y: 1.85, w: tw, h: 3.5, fill: { color: T.amber }, line: { color: T.amber }, align: "center", valign: "top", margin: [8, 8, 10, 14], objectName: "P.R.I.C.E. " + l });
  });
  box(s, { x: M, y: 5.6, w: CW, h: 1.15, fill: T.red, head: "Nei primi giorni evita calore, massaggi e alcol", headColor: H.accent1, fontSize: 17, items: ["Aumentano il gonfiore e il sanguinamento interno"] });
  s.addNotes("P.R.I.C.E. è l'acronimo inglese di Protection, Rest, Ice, Compression, Elevation. Va bene per contusioni, stiramenti, strappi e distorsioni. Per fratture e lussazioni la priorità è immobilizzare e chiamare i soccorsi.");
}

{
  const s = contentSlide("Contrattura, stiramento e strappo");
  lead(s, "Tre lesioni del muscolo di gravità crescente: dalla contrattura (nessuna fibra rotta) allo strappo (fibre rotte).", 0.5);
  const pos = cols3(1.85, 4.9);
  box(s, { ...pos[0], fill: T.amber, head: "Contrattura · lieve", headColor: H.accent3, headSize: 20, fontSize: 17, space: 14, items: ["**Cos'è:** il muscolo resta contratto, duro e teso; le fibre non sono rotte", "**Segni:** dolore che cresce piano piano, si riesce a muoversi", "**Cosa fare:** riposo, calore, massaggio leggero, stretching dolce"] });
  box(s, { ...pos[1], fill: T.red, head: "Stiramento · medio", headColor: H.accent1, headSize: 20, fontSize: 17, space: 14, items: ["**Cos'è:** il muscolo è stato allungato troppo; le fibre si danneggiano ma non si rompono", "**Segni:** dolore acuto durante lo sforzo, bisogna rallentare", "**Cosa fare:** stop, P.R.I.C.E., riposo 1–2 settimane"] });
  box(s, { ...pos[2], fill: T.red, head: "Strappo · grave", headColor: H.dk2, headSize: 20, fontSize: 17, space: 14, items: ["**Cos'è:** rottura di una parte delle fibre muscolari", "**Segni:** fitta improvvisa come una coltellata, impossibile continuare, poi ematoma", "**Cosa fare:** stop, P.R.I.C.E., medico ed ecografia; recupero di settimane"] });
  s.addNotes("Le cause principali sono riscaldamento insufficiente, stanchezza, freddo e disidratazione. La contrattura può essere trattata col calore, lo strappo no: nel dubbio, nei primi giorni, sempre ghiaccio. Prevenzione: riscaldamento, gradualità, idratazione, stretching.");
}

scheda("Le tendinopatie",
  "Sofferenza del tendine (che collega il muscolo all'osso) dovuta al sovraccarico: movimenti ripetuti nel tempo, non un singolo trauma.",
  [
    CAUSE(["Allenamenti troppo intensi o aumentati in fretta", "Gesti ripetuti sempre uguali", "Scarpe inadatte, terreni duri", "Tecnica sbagliata", "Esempi: tendine d'Achille, rotuleo, gomito del tennista"]),
    SEGNI(["Dolore lungo il tendine che aumenta con il movimento", "Rigidità al mattino o a inizio attività", "A volte gonfiore", "Dolore se si preme sul tendine"]),
    FARE(["Riduci o sospendi l'attività che fa male", "Ghiaccio quando fa male", "Visita medica e fisioterapia", "Riprendi gradualmente", "Previeni: riscaldamento e carichi graduali"]),
  ],
  "Un tempo si chiamavano tendiniti. Tendono a diventare croniche se ci si allena col dolore. Il messaggio chiave: il dolore è un segnale da ascoltare.");

scheda("La distorsione",
  "Trauma di un'articolazione: le ossa si spostano per un attimo e poi tornano al loro posto. Si possono lesionare i legamenti. Le più frequenti: caviglia e ginocchio.",
  [
    SEGNI(["Dolore all'articolazione", "Gonfiore rapido", "Livido nelle ore successive", "Difficoltà a muoversi o a camminare"]),
    FARE(["Interrompi subito l'attività", "Applica il P.R.I.C.E.", "Immobilizza con una benda o un tutore", "Non caricare il peso sull'arto", "Vai dal medico: serve escludere una frattura"]),
    NONFARE(["Non continuare a giocare \"per scaldarla\"", "Non massaggiare e non scaldare", "Non sottovalutarla: se non guarisce bene si ripete"]),
  ],
  "La distorsione di caviglia avviene quasi sempre con il piede che si gira verso l'interno. La gravità va dalla lesione lieve dei legamenti (primo grado) alla rottura completa (terzo grado). Rinforzo muscolare ed esercizi di equilibrio aiutano a prevenirla.");

scheda("La lussazione",
  "Le ossa di un'articolazione perdono il contatto e restano fuori posto. È più grave della distorsione. Le più frequenti: spalla, dita, gomito.",
  [
    SEGNI(["Dolore molto forte", "Articolazione deformata", "Impossibile muoverla", "Gonfiore", "A volte formicolio o perdita di sensibilità"]),
    FARE(["Chiama il 112 o vai al pronto soccorso", "Immobilizza l'articolazione così com'è, con bende o un foulard", "Ghiaccio avvolto in un panno", "Tieni la persona tranquilla"]),
    NONFARE(["Non tentare mai di rimetterla a posto: rischi di rompere l'osso o lesionare nervi e vasi", "Non far muovere l'arto", "Non dare da bere né da mangiare"]),
  ],
  "Nella lussazione della spalla la vittima tiene di solito il braccio attaccato al corpo e la spalla appare squadrata. La riduzione deve farla solo un medico, spesso dopo una radiografia. Niente cibo né bevande: potrebbe servire un'anestesia.");

scheda("La frattura",
  "Rottura di un osso. Chiusa se la pelle è intatta, esposta se l'osso esce dalla pelle. Composta se i pezzi restano allineati, scomposta se si spostano.",
  [
    SEGNI(["Dolore forte che aumenta col movimento", "Arto deformato o in posizione strana", "Gonfiore e livido", "Impossibile usare l'arto", "Frattura esposta: osso visibile"]),
    FARE(["Chiama il 112", "Non muovere la vittima, immobilizza l'arto così com'è", "Blocca anche le articolazioni sopra e sotto", "Frattura esposta: copri con garza sterile", "Ghiaccio intorno, mai sull'osso esposto"]),
    NONFARE(["Non raddrizzare l'arto", "Non far rientrare l'osso", "Non spostare chi ha dolore a collo o schiena", "Non dare da bere né da mangiare"]),
  ],
  "In caso di sospetta lesione della colonna vertebrale (caduta dall'alto, forte dolore a collo o schiena, formicolio o perdita di forza alle gambe) la vittima non va mai mossa: si tiene la testa ferma in asse con il corpo fino all'arrivo dei soccorsi.");

{
  const s = contentSlide("Riepilogo: come distinguerli");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, fontSize: 15, valign: "middle" } });
  const c0 = (t) => ({ text: t, options: { bold: true, color: H.dk2, fill: { color: T.gray }, fontSize: 15, valign: "middle" } });
  const c = (t) => ({ text: t, options: { color: H.dk1, fontSize: 14, valign: "middle" } });
  const rows = [
    [hdr("Trauma"), hdr("Cosa si lesiona"), hdr("Segno tipico"), hdr("Primo intervento")],
    [c0("Contusione"), c("Pelle e muscoli, senza ferita"), c("Livido e gonfiore"), c("Ghiaccio e riposo")],
    [c0("Contrattura"), c("Muscolo teso, nessuna fibra rotta"), c("Muscolo duro, dolore graduale"), c("Riposo, calore, stretching dolce")],
    [c0("Strappo"), c("Fibre muscolari rotte"), c("Fitta improvvisa, si deve smettere"), c("P.R.I.C.E. e medico")],
    [c0("Tendinopatia"), c("Tendine sovraccaricato"), c("Dolore che cresce con l'uso"), c("Riposo, ghiaccio, medico")],
    [c0("Distorsione"), c("Legamenti; l'articolazione torna a posto"), c("Gonfiore rapido"), c("P.R.I.C.E., immobilizzare, medico")],
    [c0("Lussazione"), c("Articolazione fuori posto"), c("Articolazione deformata"), c("Immobilizzare così com'è, 112")],
    [c0("Frattura"), c("Osso rotto"), c("Arto deformato, dolore forte"), c("Non muovere, immobilizzare, 112")],
  ];
  s.addTable(rows, { x: M, y: 1.35, w: CW, colW: [2.2, 3.6, 3.2, 3.13], rowH: 0.64, border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.05, 0.12, 0.05, 0.12], fontFace: THEME.bodyFontFace, objectName: "Tabella riepilogo traumi" });
  s.addNotes("Tabella utile per lo studio. Nel dubbio tra distorsione, lussazione e frattura ci si comporta come per la frattura: si immobilizza e si chiede un controllo medico.");
}

// =========================================================
// PARTE 5 — EMERGENZE
// =========================================================
section(S.em);
sectionSlide("PARTE 5", "Le emergenze", ["Lipotimia, arresto cardiaco, shock", "Trauma cranico, soffocamento", "Ipertermia e ipotermia"],
  "Situazioni in cui sono in pericolo le funzioni vitali: coscienza, respiro, circolazione. Qui chiamare subito il 112 è fondamentale.");

scheda("La lipotimia (lo svenimento)",
  "Breve perdita di coscienza, o sensazione di svenire, perché arriva poco sangue al cervello. Di solito si risolve da sola in pochi minuti.",
  [
    CAUSE(["Caldo e ambienti affollati", "Stare a lungo in piedi fermi", "Emozioni forti, dolore, vista del sangue", "Digiuno", "Alzarsi in piedi di scatto"]),
    SEGNI(["Pallore e sudore freddo", "Vista annebbiata, nausea, ronzio nelle orecchie", "Debolezza, poi la persona si accascia", "Polso debole, ma respira"]),
    FARE(["Falla sdraiare", "Solleva le gambe di 30–40 cm", "Allenta gli abiti e fai circolare aria", "Quando si riprende, falla alzare piano", "Se non si riprende in 1–2 minuti: 112"]),
  ],
  "Sollevare le gambe fa tornare più sangue verso il cuore e il cervello. Non dare mai da bere a una persona non completamente cosciente: può soffocare. Se lo svenimento avviene durante uno sforzo fisico va sempre fatto un controllo medico, perché può nascondere un problema cardiaco.");

{
  const s = contentSlide("L'arresto cardiaco");
  lead(s, "Il cuore smette di pompare sangue: il cervello non riceve più ossigeno. La persona è incosciente e non respira, o respira in modo anomalo (boccheggia).", 0.8);
  const pos = cols3(2.15, 1.95);
  [["4–6 min", "dopo i quali il cervello senza ossigeno inizia a subire danni"], ["−10%", "di probabilità di sopravvivere per ogni minuto senza aiuto"], ["× 2–3", "la sopravvivenza se qualcuno inizia subito la RCP"]].forEach(([n, l], i) => stat(s, { ...pos[i], fill: T.red, num: n, label: l, numColor: H.accent1, numSize: 34 }));
  s.addText("La catena della sopravvivenza", { x: M, y: 4.3, w: CW, h: 0.45, fontSize: 18, bold: true, color: H.dk2, margin: 0, isTextBox: true, objectName: "Sottotitolo" });
  const gap = 0.25, cw = (CW - 3 * gap) / 4;
  [["1. Allarme", "chiama il 112"], ["2. RCP", "compressioni subito"], ["3. Defibrillazione", "con il DAE"], ["4. Soccorso avanzato", "medici e ospedale"]].forEach(([t, d], i) => {
    s.addText([{ text: t, options: { bold: true, fontSize: 17, color: "FFFFFF", breakLine: true, paraSpaceAfter: 4 } }, { text: d, options: { fontSize: 15, color: PINK } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + i * (cw + gap), y: 4.85, w: cw, h: 1.9, fill: { color: H.dk2 }, line: { color: H.dk2 }, valign: "middle", align: "center", margin: PAD, objectName: "Anello " + (i + 1) });
  });
  s.addNotes("Nello sport l'arresto cardiaco può colpire anche giovani apparentemente sani, per malattie cardiache non diagnosticate: per questo esistono la visita medico-sportiva e l'obbligo del defibrillatore negli impianti sportivi. I soccorsi impiegano diversi minuti per arrivare: chi è sul posto fa la differenza.");
}

{
  const s = contentSlide("La rianimazione cardiopolmonare (RCP)");
  const steps = [
    ["Sicurezza e coscienza", "La scena è sicura? Chiama e scuoti la vittima: non risponde"],
    ["Respiro", "Testa indietro e mento in su; per 10 secondi: non respira normalmente"],
    ["112 e DAE", "Chiama il 112 in vivavoce e fai portare un defibrillatore"],
    ["30 compressioni", "Al centro del torace, braccia tese: 5–6 cm, 100–120 al minuto"],
    ["2 ventilazioni", "Solo se sei addestrato; altrimenti continua con le sole compressioni"],
    ["Non fermarti", "Accendi il DAE, segui la sua voce e continua fino ai soccorsi"],
  ];
  steps.forEach(([t, d], i) => step(s, i + 1, t, d, { x: M, y: 1.35 + i * 0.9, w: 7.9, h: 0.85, color: i >= 3 ? H.accent1 : H.dk2, d: 0.5 }));
  const rx = M + 8.3, rw = W - M - rx;
  s.addText([
    { text: "Il ritmo da ricordare", options: { bold: true, fontSize: 16, color: PINK, breakLine: true } },
    { text: "30 : 2", options: { bold: true, fontSize: 54, color: "FFFFFF", fontFace: THEME.headFontFace, breakLine: true } },
    { text: "compressioni : ventilazioni", options: { fontSize: 15, color: "FFFFFF" } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: rx, y: 1.35, w: rw, h: 2.9, fill: { color: H.dk2 }, line: { color: H.dk2 }, valign: "middle", margin: PAD, objectName: "Ritmo RCP" });
  box(s, { x: rx, y: 4.45, w: rw, h: 2.3, fill: T.amber, head: "Il DAE (defibrillatore)", headColor: H.accent3, fontSize: 16, space: 6, items: ["Analizza il cuore e dà la scarica solo se serve", "Può usarlo chiunque"] });
  s.addNotes("Le compressioni devono permettere al torace di riespandersi completamente tra una e l'altra. Per tenere il ritmo si può pensare a una canzone a circa 100–120 battiti al minuto. Il DAE non può far male: dà la scarica solo se riconosce un ritmo defibrillabile. Durante la scarica nessuno deve toccare la vittima.");
}

scheda("Lo shock",
  "Grave insufficienza della circolazione: gli organi non ricevono abbastanza sangue e ossigeno. Cause: forte emorragia, reazione allergica grave, problema al cuore, gravi ustioni.",
  [
    SEGNI(["Pelle pallida, fredda e sudata", "Polso rapido e debole", "Respiro rapido e superficiale", "Sete, agitazione, poi confusione", "Se allergico: gonfiore di viso e gola, fatica a respirare"]),
    FARE(["Chiama subito il 112", "Ferma eventuali emorragie", "Sdraiala e solleva le gambe (se non ci sono traumi)", "Coprila per non farle perdere calore", "Rassicurala e controlla il respiro"]),
    NONFARE(["Non dare da bere né da mangiare", "Non lasciarla sola", "Non sollevare le gambe se ci sono fratture o trauma cranico", "Non perdere tempo: peggiora in fretta"]),
  ],
  "La posizione antishock (sdraiati con le gambe sollevate) aiuta il sangue a tornare verso gli organi vitali. Chi soffre di allergie gravi può avere con sé una penna di adrenalina autoiniettabile: si può aiutarlo a usarla.");

scheda("Il trauma cranico",
  "Un colpo alla testa che può danneggiare il cervello. Nello sport il più comune è la commozione cerebrale: un disturbo temporaneo, anche senza perdere i sensi.",
  [
    ATT("Segnali di allarme", ["Perdita di coscienza, anche breve", "Mal di testa forte, vomito", "Confusione, non ricorda cosa è successo", "Sonnolenza, fatica a parlare o camminare", "Sangue o liquido da naso o orecchie"]),
    FARE(["Stop: nel dubbio l'atleta esce dal campo", "Con un segnale di allarme: 112", "Se a terra con dolore al collo, tieni ferma la testa", "Ghiaccio sul bernoccolo", "Osservalo per almeno 24 ore"]),
    NONFARE(["Non farlo rientrare in gioco lo stesso giorno", "Non muovere testa e collo", "Non togliere il casco a un motociclista", "Non dare farmaci senza il medico"]),
  ],
  "I sintomi di un trauma cranico possono comparire anche ore dopo il colpo: per questo la persona va osservata. Il ritorno allo sport dopo una commozione cerebrale deve essere graduale e autorizzato da un medico.");

{
  const s = contentSlide("Il soffocamento");
  lead(s, "Un corpo estraneo, di solito cibo, blocca le vie aeree. Il segno tipico: la persona si porta le mani alla gola.", 0.5);
  const pw = (CW - 0.3) / 2;
  box(s, { x: M, y: 1.85, w: pw, h: 2.75, fill: T.amber, head: "Ostruzione parziale", headColor: H.accent3, headSize: 20, fontSize: 18, space: 10, items: ["**Segni:** tossisce, parla e respira, anche se con fatica", "**Cosa fare:** incoraggiala a tossire e resta con lei", "Non darle colpi sulla schiena e non farla bere"] });
  box(s, { x: M + pw + 0.3, y: 1.85, w: pw, h: 2.75, fill: T.red, head: "Ostruzione completa", headColor: H.accent1, headSize: 20, fontSize: 18, space: 10, items: ["**Segni:** non riesce a parlare, tossire né respirare", "**Cosa fare:** intervieni subito con le manovre qui sotto", "Fai chiamare il 112"] });
  const pos = cols3(4.8, 1.95);
  [["1. Cinque colpi sulla schiena", "Falla piegare in avanti e dai 5 colpi decisi tra le scapole"], ["2. Cinque compressioni addominali", "Manovra di Heimlich: da dietro, pugno sopra l'ombelico, spingi in dentro e in su"], ["3. Ripeti o RCP", "Alterna 5 e 5. Se perde coscienza: a terra, 112 e RCP"]].forEach(([t, d], i) => {
    s.addText([{ text: t, options: { bold: true, color: PINK, fontSize: 17, breakLine: true, paraSpaceAfter: 6 } }, { text: d, options: { color: "FFFFFF", fontSize: 16 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, ...pos[i], fill: { color: H.dk2 }, line: { color: H.dk2 }, valign: "top", margin: PAD, objectName: "Manovra " + (i + 1) });
  });
  s.addNotes("Le manovre per i lattanti (sotto 1 anno) sono diverse: colpi sulla schiena e compressioni sul torace, mai la manovra di Heimlich. Chi ha subito compressioni addominali va comunque controllato da un medico.");
}

scheda("L'ipertermia (colpo di calore)",
  "La temperatura del corpo sale oltre i 40 °C perché il corpo non riesce più a disperdere il calore. Il colpo di sole è causato dal sole diretto sulla testa.",
  [
    ATT("Cause e segni", ["Sforzo con caldo umido, poca acqua, abiti pesanti", "Pelle calda e arrossata, poi secca", "Mal di testa, nausea, vertigini, crampi", "Confusione o svenimento nei casi gravi"]),
    FARE(["Portala all'ombra o al fresco", "Togli gli indumenti in eccesso", "Panni bagnati su testa, collo, ascelle e inguine; fai aria", "Se è cosciente: acqua fresca a piccoli sorsi", "Se è confusa o sviene: 112"]),
    { fill: T.teal, head: "Prevenzione", headColor: H.accent2, items: ["Bevi prima, durante e dopo, anche senza sete", "Evita le ore più calde", "Abiti leggeri e chiari, cappello al sole", "Abituati al caldo gradualmente"] },
  ],
  "Il colpo di calore è un'emergenza: se la temperatura non viene abbassata rapidamente può danneggiare cervello e organi. Non dare da bere a chi non è completamente cosciente.");

scheda("L'ipotermia",
  "La temperatura del corpo scende sotto i 35 °C perché il corpo perde più calore di quanto ne produca: in montagna, in acqua fredda, con abiti bagnati e vento.",
  [
    SEGNI(["Brividi forti, che nei casi gravi smettono", "Pelle pallida e fredda, labbra bluastre", "Muscoli rigidi, movimenti impacciati", "Sonnolenza e confusione", "Respiro e polso lenti"]),
    FARE(["Portala in un luogo riparato e caldo", "Togli i vestiti bagnati e coprila (coperta termica)", "Riscaldala piano, partendo dal tronco", "Se è cosciente: bevande calde e zuccherate", "Nei casi gravi: 112"]),
    NONFARE(["Non dare alcolici", "Non strofinare né massaggiare braccia e gambe", "Non usare fonti di calore forti sulla pelle", "Non farla camminare o sforzare"]),
  ],
  "Strofinare gli arti fa tornare verso il cuore sangue molto freddo e può causare problemi al ritmo cardiaco. La coperta termica si usa con il lato dorato verso l'esterno per trattenere il calore. Bambini e anziani perdono calore più facilmente.");

// =========================================================
// CONCLUSIONE
// =========================================================
section(S.fine);
{
  const s = contentSlide("Le regole d'oro del soccorritore");
  const rules = [
    "Prima la tua sicurezza, poi quella della vittima",
    "Chiama il 112 e spiega dove sei e cosa è successo",
    "Fai solo ciò che sai fare: a volte non fare è meglio che fare male",
    "Non muovere chi ha un trauma, salvo pericolo immediato",
    "Non dare da bere a chi non è pienamente cosciente",
    "Resta con la vittima e rassicurala fino all'arrivo dei soccorsi",
  ];
  box(s, { x: M, y: 1.4, w: 8.3, h: 5.35, fill: T.gray, numbered: true, fontSize: 21, space: 20, valign: "middle", items: rules, name: "Regole d'oro" });
  s.addText([
    { text: "112", options: { bold: true, fontSize: 72, color: "FFFFFF", fontFace: THEME.headFontFace, breakLine: true } },
    { text: "Numero Unico di Emergenza", options: { bold: true, fontSize: 17, color: "FFFFFF", breakLine: true } },
    { text: "gratuito, anche senza credito", options: { fontSize: 15, color: PINK } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x: M + 8.6, y: 1.4, w: CW - 8.6, h: 5.35, fill: { color: H.accent1 }, line: { color: H.accent1 }, align: "center", valign: "middle", margin: PAD, objectName: "Numero emergenza" });
  s.addNotes("Riassunto finale. Il primo soccorso non richiede di essere medici: richiede calma, poche regole chiare e la prontezza di chiamare aiuto. Le manovre di RCP e di disostruzione vanno imparate con un corso pratico, ad esempio BLSD.");
}

// ---------- Scrittura ----------
(async () => {
  await pres.writeFile({ fileName: OUT });
  // pptxgenjs scrive un <a:pPr> in più davanti a ogni run successivo di un paragrafo: lo rimuoviamo
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  let fixed = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const out = xml.replace(/(<\/a:r>)<a:pPr\b(?:[^>]*[^\/>])?>.*?<\/a:pPr>/g, (m, r) => { fixed++; return r; });
    zip.file(name, out);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  await applyTheme(OUT, THEME);
  console.log("ok", OUT, "pPr rimossi:", fixed);
})().catch((e) => { console.error(e); process.exit(1); });
