const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const gi = require("react-icons/gi");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Scienze Motorie",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E1E24",
    lt1: "FFFFFF",
    dk2: "6E0F1E", // rosso sangue scuro
    lt2: "F3F4F6",
    accent1: "C8102E", // rosso
    accent2: "1F6F8B", // blu-verde (respirazione)
    accent3: "D98A1C", // ambra (attenzione)
    accent4: "2E7D4F", // verde (cosa fare)
    accent5: "5A6472", // grigio
    accent6: "3A3F8F", // indaco
    hlink: "1F6F8B",
    folHlink: "6E0F1E",
  },
};
const HEX = THEME.colors;

// Light tints (hex) for cards
const T = {
  red: "FBE9EC",
  teal: "E4F1F5",
  amber: "FCF1E1",
  green: "E6F3EB",
  gray: "F3F4F6",
  indigo: "E9EAF6",
};

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.title = "Apparato cardiocircolatorio, respirazione e primo soccorso";
  pres.subject = "Scienze motorie";
  pres.author = "Scienze motorie";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const W = 13.333;
  const M = 0.6;

  // ---------- Icons ----------
  const I = {};
  const iconSpecs = {
    heart: [gi.GiHeartOrgan],
    lungs: [gi.GiLungs],
    drop: [gi.GiDrop],
    blood: [gi.GiBlood],
    firstaid: [gi.GiFirstAidKit],
    run: [fa.FaRunning],
    bone: [gi.GiBrokenBone],
    brain: [fa.FaBrain],
    hot: [fa.FaTemperatureHigh],
    cold: [fa.FaTemperatureLow],
    snow: [fa.FaSnowflake],
    sun: [fa.FaSun],
    phone: [fa.FaPhoneAlt],
    check: [fa.FaCheck],
    times: [fa.FaTimes],
    eye: [fa.FaEye],
    shield: [fa.FaShieldAlt],
    ambulance: [fa.FaAmbulance],
    warn: [fa.FaExclamationTriangle],
    bandage: [fa.FaBandAid],
    wound: [gi.GiBleedingWound],
    nose: [gi.GiNoseSide],
    muscle: [gi.GiMuscleUp],
    knee: [gi.GiKneeCap],
    leg: [gi.GiLeg],
    arm: [gi.GiArmBandage],
    pulse: [gi.GiPulse],
    beats: [gi.GiHeartBeats],
    stopwatch: [fa.FaStopwatch],
    ice: [gi.GiIceCube],
    bed: [fa.FaBed],
    cough: [fa.FaHeadSideCough],
    hand: [fa.FaHandPaper],
    search: [fa.FaSearch],
    bolt: [fa.FaBolt],
    user: [fa.FaUserInjured],
    wind: [fa.FaWind],
    dumbbell: [fa.FaDumbbell],
    stetho: [gi.GiStethoscope],
    lightbulb: [fa.FaLightbulb],
    info: [fa.FaInfoCircle],
    clip: [fa.FaClipboardList],
    skull: [gi.GiSkullCrack],
    falling: [gi.GiFalling],
    heartplus: [gi.GiHeartPlus],
    siren: [gi.GiSiren],
    syringe: [gi.GiSyringe],
    arrowUp: [fa.FaArrowUp],
    arrowDown: [fa.FaArrowDown],
  };
  async function ic(name, color) {
    const key = name + "_" + color;
    if (!I[key]) I[key] = await icon(iconSpecs[name][0], color);
    return I[key];
  }

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "Titolo",
    background: { color: HEX.dk2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.3, w: 8.4, h: 2.0, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 4.55, w: 8.4, h: 1.2, fontSize: 20, color: "F2C9CF", valign: "top", align: "left", margin: 0 }, text: "Sottotitolo" } },
    ],
  });

  pres.defineSlideMaster({
    title: "Sezione",
    background: { color: HEX.dk2 },
    objects: [
      { placeholder: { options: { name: "body", type: "body", x: M, y: 2.2, w: 8, h: 0.6, fontSize: 18, bold: true, color: "F2C9CF", charSpacing: 3, align: "left", margin: 0 }, text: "Parte" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.85, w: 8.2, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo sezione" } },
    ],
    slideNumber: { x: W - 1.1, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: "F2C9CF", align: "right" },
  });

  pres.defineSlideMaster({
    title: "Contenuto",
    background: { color: HEX.lt1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "middle", align: "left", fontFace: THEME.headFontFace, margin: 0 }, text: "Titolo" } },
      { text: { text: "Scienze motorie", options: { x: M, y: 6.95, w: 4, h: 0.35, fontSize: 11, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: W - 1.1, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: C.accent5, align: "right" },
  });

  // ---------- Helpers ----------
  function contentSlide(section, title) {
    const s = pres.addSlide({ masterName: "Contenuto", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    return s;
  }

  function lead(s, text, y = 1.2, h = 0.7) {
    s.addText(text, { x: M, y, w: W - 2 * M, h, fontSize: 17, italic: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "Introduzione" });
  }

  function bullets(items, opts = {}) {
    return items.map((t, i) => {
      const o = { bullet: { indent: 16 }, paraSpaceAfter: opts.space ?? 6, breakLine: i < items.length - 1 };
      if (Array.isArray(t)) return t.map((r, j) => ({ text: r.text, options: { ...o, ...r.options, breakLine: j === t.length - 1 ? o.breakLine : false, bullet: j === 0 ? o.bullet : undefined } }));
      return [{ text: t, options: o }];
    }).flat();
  }

  // rich bullet: "**Label**: rest"
  function rb(items) {
    return items.map((t) => {
      const m = t.match(/^\*\*(.+?)\*\*(.*)$/);
      if (m) return [{ text: m[1], options: { bold: true } }, { text: m[2] }];
      return t;
    });
  }

  function richBullets(items, opts = {}) {
    const out = [];
    items.forEach((t, i) => {
      const last = i === items.length - 1;
      const base = { paraSpaceAfter: opts.space ?? 6 };
      const m = typeof t === "string" ? t.match(/^\*\*(.+?)\*\*(.*)$/) : null;
      if (m) {
        out.push({ text: m[1], options: { ...base, bullet: { indent: 16 }, bold: true } });
        out.push({ text: m[2], options: { ...base, breakLine: !last } });
      } else {
        out.push({ text: t, options: { ...base, bullet: { indent: 16 }, breakLine: !last } });
      }
    });
    return out;
  }

  async function iconCircle(s, name, x, y, d, fill, fg = "FFFFFF", objectName = "Icona") {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: objectName + " cerchio" });
    const p = d * 0.25;
    s.addImage({ data: await ic(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: objectName, objectName });
  }

  // Card with icon header + bullets
  async function card(s, { x, y, w, h, tint, color, iconName, label, items, fontSize = 15, space = 6 }) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: tint }, line: { color: tint }, objectName: "Scheda " + label });
    await iconCircle(s, iconName, x + 0.25, y + 0.22, 0.55, color, "FFFFFF", "Icona " + label);
    s.addText(label, { x: x + 0.95, y: y + 0.22, w: w - 1.15, h: 0.55, fontSize: 18, bold: true, color, valign: "middle", margin: 0, isTextBox: true, objectName: "Titolo scheda " + label });
    s.addText(richBullets(items, { space }), { x: x + 0.25, y: y + 0.95, w: w - 0.45, h: h - 1.1, fontSize, color: HEX.dk1, valign: "top", margin: 0, isTextBox: true, objectName: "Testo scheda " + label });
  }

  // Standard first-aid "scheda": lead + 3 columns
  async function scheda(section, title, { def, cols }) {
    const s = contentSlide(section, title);
    lead(s, def, 1.2, 0.75);
    const top = 2.1, h = 4.65, gap = 0.3;
    const w = (W - 2 * M - gap * (cols.length - 1)) / cols.length;
    for (let i = 0; i < cols.length; i++) {
      await card(s, { x: M + i * (w + gap), y: top, w, h, fontSize: 16, space: 9, ...cols[i] });
    }
    return s;
  }

  const SEGNI = (items) => ({ tint: T.gray, color: HEX.accent5, iconName: "eye", label: "Come si riconosce", items });
  const FARE = (items) => ({ tint: T.green, color: HEX.accent4, iconName: "check", label: "Cosa fare", items });
  const NONFARE = (items) => ({ tint: T.red, color: HEX.accent1, iconName: "times", label: "Cosa NON fare", items });
  const CAUSE = (items) => ({ tint: T.amber, color: HEX.accent3, iconName: "search", label: "Cause", items });

  const S1 = "Apparato cardiocircolatorio";
  const S2 = "Apparato respiratorio";
  const S3 = "Il primo soccorso";
  const S4 = "Traumi dell'apparato locomotore";
  const S5 = "Le emergenze";

  // =========================================================
  // 1. Titolo
  // =========================================================
  pres.addSection({ title: "Introduzione" });
  {
    const s = pres.addSlide({ masterName: "Titolo", sectionTitle: "Introduzione" });
    s.addText("Cuore, respiro e primo soccorso", { placeholder: "title" });
    s.addText("Apparato cardiocircolatorio e respiratorio, esercizio fisico e gestione dei traumi più comuni", { placeholder: "body" });
    // motif: three icon circles on right
    await iconCircle(s, "heart", 9.9, 1.35, 1.5, HEX.accent1, "FFFFFF", "Icona cuore");
    await iconCircle(s, "lungs", 10.9, 3.05, 1.5, HEX.accent2, "FFFFFF", "Icona polmoni");
    await iconCircle(s, "firstaid", 9.9, 4.75, 1.5, "FFFFFF", HEX.accent1, "Icona primo soccorso");
    s.addNotes("Presentazione di scienze motorie. Tre parti: apparato cardiocircolatorio, apparato respiratorio e primo soccorso, con particolare attenzione ai traumi sportivi e alle emergenze.");
  }

  // 2. Indice
  {
    const s = contentSlide("Introduzione", "Di cosa parliamo");
    const items = [
      ["heart", HEX.accent1, T.red, "1", "Apparato cardiocircolatorio", "Il cuore, la circolazione, il sangue e l'adattamento all'esercizio fisico"],
      ["lungs", HEX.accent2, T.teal, "2", "Apparato respiratorio", "Gli organi, la meccanica del respiro e la respirazione durante lo sforzo"],
      ["firstaid", HEX.accent4, T.green, "3", "Il primo soccorso", "Come intervenire: P.A.S., valutazione, contusioni, ferite ed emorragie"],
      ["bone", HEX.accent3, T.amber, "4", "Traumi dell'apparato locomotore", "Muscoli, tendini, articolazioni e ossa: dalla contrattura alla frattura"],
      ["siren", HEX.accent6, T.indigo, "5", "Le emergenze", "Lipotimia, arresto cardiaco, shock, trauma cranico, soffocamento, ipertermia e ipotermia"],
    ];
    const y0 = 1.4, rh = 1.0, gap = 0.08;
    for (let i = 0; i < items.length; i++) {
      const [ico, col, tint, n, t, d] = items[i];
      const y = y0 + i * (rh + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: W - 2 * M, h: rh, rectRadius: 0.1, fill: { color: tint }, line: { color: tint }, objectName: "Riga indice " + n });
      await iconCircle(s, ico, M + 0.2, y + 0.17, 0.66, col, "FFFFFF", "Icona " + t);
      s.addText(n, { x: M + 1.05, y, w: 0.5, h: rh, fontSize: 28, bold: true, color: col, valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
      s.addText(t, { x: M + 1.6, y, w: 4.4, h: rh, fontSize: 20, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: M + 6.1, y, w: W - 2 * M - 6.3, h: rh, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    }
    s.addNotes("Le parti 3, 4 e 5 sono il cuore della lezione: sapere cosa fare (e cosa non fare) davanti a un infortunio può fare la differenza.");
  }

  // =========================================================
  // PARTE 1 — CARDIOCIRCOLATORIO
  // =========================================================
  pres.addSection({ title: S1 });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: S1 });
    s.addText("PARTE 1", { placeholder: "body" });
    s.addText("Apparato cardiocircolatorio", { placeholder: "title" });
    await iconCircle(s, "heart", 9.6, 2.0, 2.8, HEX.accent1, "FFFFFF", "Icona cuore");
    s.addNotes("L'apparato cardiocircolatorio è formato da cuore, vasi sanguigni e sangue. Il suo compito è portare ossigeno e nutrienti a tutte le cellule e rimuovere le sostanze di scarto.");
  }

  // Il cuore
  {
    const s = contentSlide(S1, "Il cuore");
    lead(s, "Un muscolo cavo grande quanto un pugno, posto nel torace tra i polmoni, leggermente spostato a sinistra. Funziona come una pompa che spinge il sangue in tutto il corpo.", 1.2, 0.8);
    // Left: structure
    const lx = M, lw = 6.6, ty = 2.2;
    s.addText("Com'è fatto", { x: lx, y: ty, w: lw, h: 0.45, fontSize: 20, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(richBullets([
      "**4 cavità:** 2 atri (in alto) che ricevono il sangue e 2 ventricoli (in basso) che lo spingono fuori",
      "**Setto:** divide il cuore destro (sangue povero di ossigeno) dal sinistro (sangue ricco di ossigeno)",
      "**4 valvole** (tricuspide, mitrale, polmonare, aortica): fanno scorrere il sangue in una sola direzione",
      "**3 strati:** pericardio (involucro esterno), miocardio (muscolo), endocardio (rivestimento interno)",
      "**Ciclo cardiaco:** sistole (contrazione, il sangue esce) e diastole (rilassamento, il cuore si riempie)",
    ]), { x: lx, y: ty + 0.55, w: lw, h: 4.0, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    // Right: stats 2x2
    const stats = [
      ["≈ 300 g", "peso del cuore di un adulto"],
      ["60–80", "battiti al minuto a riposo"],
      ["≈ 5 litri", "di sangue pompati ogni minuto"],
      ["100.000", "battiti circa in un giorno"],
    ];
    const gx = 7.65, gw = 2.4, gh = 2.05, gg = 0.25;
    stats.forEach(([n, l], i) => {
      const x = gx + (i % 2) * (gw + gg), y = ty + Math.floor(i / 2) * (gh + gg);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: gw, h: gh, rectRadius: 0.12, fill: { color: T.red }, line: { color: T.red }, objectName: "Dato " + n });
      s.addText(n, { x: x + 0.15, y: y + 0.3, w: gw - 0.3, h: 0.8, fontSize: 32, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
      s.addText(l, { x: x + 0.15, y: y + 1.1, w: gw - 0.3, h: 0.75, fontSize: 14, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("Il cuore lavora senza sosta per tutta la vita. Il battito nasce da un pacemaker naturale, il nodo seno-atriale, che genera l'impulso elettrico senza bisogno di comandi volontari. Le valvole sono come porte a senso unico: il rumore che sentiamo con lo stetoscopio è la loro chiusura.");
  }

  // La circolazione sanguigna
  {
    const s = contentSlide(S1, "La circolazione sanguigna");
    lead(s, "Nell'uomo la circolazione è doppia (due circuiti) e completa (il sangue ossigenato non si mescola mai con quello povero di ossigeno).", 1.2, 0.7);
    const panels = [
      { title: "Piccola circolazione (polmonare)", color: HEX.accent2, tint: T.teal, steps: ["Ventricolo destro", "Arterie polmonari", "Polmoni: il sangue prende O₂ e lascia CO₂", "Vene polmonari", "Atrio sinistro"] },
      { title: "Grande circolazione (sistemica)", color: HEX.accent1, tint: T.red, steps: ["Ventricolo sinistro", "Aorta e arterie", "Organi e muscoli: il sangue cede O₂ e raccoglie CO₂", "Vene cave", "Atrio destro"] },
    ];
    const py = 2.0, ph = 2.9, pw = (W - 2 * M - 0.3) / 2;
    panels.forEach((p, i) => {
      const x = M + i * (pw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: py, w: pw, h: ph, rectRadius: 0.12, fill: { color: p.tint }, line: { color: p.tint }, objectName: p.title });
      s.addText(p.title, { x: x + 0.25, y: py + 0.15, w: pw - 0.5, h: 0.5, fontSize: 18, bold: true, color: p.color, margin: 0, isTextBox: true });
      // vertical chain of steps
      const sy = py + 0.75, sh = 0.36, sg = 0.06;
      p.steps.forEach((t, j) => {
        const y = sy + j * (sh + sg);
        s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.05, w: 0.26, h: 0.26, fill: { color: p.color }, line: { color: p.color }, objectName: "Punto " + t });
        s.addText(t, { x: x + 0.75, y, w: pw - 1.0, h: sh, fontSize: 15, color: HEX.dk1, bold: j === 2, valign: "middle", margin: 0, isTextBox: true });
      });
      s.addShape(pres.shapes.LINE, { x: x + 0.43, y: sy + 0.31, w: 0, h: 4 * (sh + sg) - 0.26, line: { color: p.color, width: 2 }, objectName: "Linea percorso" });
    });
    // vessels row
    const vessels = [
      ["Arterie", "Portano il sangue dal cuore agli organi. Pareti spesse ed elastiche, alta pressione."],
      ["Vene", "Riportano il sangue al cuore. Pareti sottili, con valvole a nido di rondine contro il reflusso."],
      ["Capillari", "Vasi sottilissimi dove avvengono gli scambi di gas e nutrienti con le cellule."],
    ];
    const vy = 5.1, vh = 1.6, vw = (W - 2 * M - 0.6) / 3;
    vessels.forEach(([t, d], i) => {
      const x = M + i * (vw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: vy, w: vw, h: vh, rectRadius: 0.1, fill: { color: T.gray }, line: { color: T.gray }, objectName: "Vaso " + t });
      s.addText([{ text: t, options: { bold: true, color: HEX.dk2, fontSize: 17, breakLine: true } }, { text: d, options: { fontSize: 14, color: HEX.dk1 } }], { x: x + 0.2, y: vy + 0.12, w: vw - 0.4, h: vh - 0.24, valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("Attenzione: le arterie si chiamano così perché partono dal cuore, non perché portano sangue ossigenato. Infatti le arterie polmonari portano sangue povero di ossigeno ai polmoni, e le vene polmonari riportano al cuore sangue ricco di ossigeno.");
  }

  // Il sangue
  {
    const s = contentSlide(S1, "Il sangue");
    lead(s, "Un tessuto liquido: in un adulto ce ne sono circa 5 litri (7–8% del peso corporeo). Trasporta ossigeno, nutrienti, ormoni e scarti, difende l'organismo e regola la temperatura.", 1.2, 0.8);
    // composition bar
    const by = 2.25, bh = 0.75, bw = W - 2 * M;
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: by, w: bw * 0.55, h: bh, fill: { color: "E9C46A" }, line: { color: "E9C46A" }, objectName: "Barra plasma" });
    s.addShape(pres.shapes.RECTANGLE, { x: M + bw * 0.55, y: by, w: bw * 0.45, h: bh, fill: { color: HEX.accent1 }, line: { color: HEX.accent1 }, objectName: "Barra parte corpuscolata" });
    s.addText([{ text: "Plasma 55%", options: { bold: true } }, { text: "  ·  acqua (90%), proteine, sali, zuccheri" }], { x: M + 0.2, y: by, w: bw * 0.55 - 0.3, h: bh, fontSize: 16, color: HEX.dk1, valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: "Parte corpuscolata 45%", options: { bold: true } }, { text: "  ·  le cellule" }], { x: M + bw * 0.55 + 0.2, y: by, w: bw * 0.45 - 0.3, h: bh, fontSize: 16, color: "FFFFFF", valign: "middle", margin: 0, isTextBox: true });
    const cells = [
      ["blood", HEX.accent1, T.red, "Globuli rossi (eritrociti)", ["Contengono emoglobina, che trasporta l'ossigeno", "4,5–5,5 milioni per mm³", "Vivono circa 120 giorni"]],
      ["shield", HEX.accent6, T.indigo, "Globuli bianchi (leucociti)", ["Difendono l'organismo da virus e batteri", "4.000–10.000 per mm³", "Aumentano durante un'infezione"]],
      ["bandage", HEX.accent3, T.amber, "Piastrine", ["Frammenti di cellule che fermano le emorragie", "150.000–400.000 per mm³", "Formano il coagulo"]],
    ];
    const cy = 3.25, ch = 2.45, cw = (W - 2 * M - 0.6) / 3;
    for (let i = 0; i < cells.length; i++) {
      const [ico, col, tint, t, it] = cells[i];
      await card(s, { x: M + i * (cw + 0.3), y: cy, w: cw, h: ch, tint, color: col, iconName: ico, label: t, items: it, fontSize: 15 });
    }
    const fy = 5.95, fh = 0.8, fgap = 0.25, fw = (W - 2 * M - 3 * fgap) / 4;
    [["Trasporto", "ossigeno, nutrienti, ormoni"], ["Difesa", "contro le infezioni"], ["Coagulazione", "chiude le ferite"], ["Termoregolazione", "distribuisce il calore"]].forEach(([t, d], i) => {
      const x = M + i * (fw + fgap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: fy, w: fw, h: fh, rectRadius: 0.4, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Funzione " + t });
      s.addText([{ text: t, options: { bold: true, color: "FFFFFF", fontSize: 15, breakLine: true } }, { text: d, options: { color: "F2C9CF", fontSize: 13 } }], { x: x + 0.2, y: fy, w: fw - 0.4, h: fh, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
    s.addNotes("Con l'allenamento, soprattutto di resistenza, aumenta il volume di sangue e la quantità di emoglobina: il corpo trasporta più ossigeno ai muscoli. Lo stesso accade vivendo in alta quota.");
  }

  // Cardio ed esercizio
  {
    const s = contentSlide(S1, "Apparato cardiocircolatorio ed esercizio fisico");
    lead(s, "Durante lo sforzo i muscoli chiedono molto più ossigeno: il cuore risponde battendo più spesso e più forte.", 1.2, 0.5);
    const stats = [
      ["Frequenza cardiaca", "70", "180–200", "battiti/min"],
      ["Gittata sistolica", "70", "120", "ml di sangue per battito (fino a 200 negli atleti)"],
      ["Portata cardiaca", "5", "20–25", "litri di sangue al minuto (fino a 35 negli atleti)"],
    ];
    const sy = 1.8, sh = 2.45, sw = (W - 2 * M - 0.6) / 3;
    for (let i = 0; i < stats.length; i++) {
      const [t, a, b, u] = stats[i];
      const x = M + i * (sw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: sy, w: sw, h: sh, rectRadius: 0.12, fill: { color: T.red }, line: { color: T.red }, objectName: "Dato " + t });
      s.addText(t, { x: x + 0.25, y: sy + 0.15, w: sw - 0.5, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true });
      s.addText([{ text: "riposo  ", options: { fontSize: 13, color: HEX.accent5 } }, { text: a, options: { fontSize: 30, bold: true, color: HEX.accent5 } }], { x: x + 0.25, y: sy + 0.65, w: sw - 0.5, h: 0.6, margin: 0, valign: "middle", isTextBox: true });
      s.addText([{ text: "sforzo  ", options: { fontSize: 13, color: HEX.accent1 } }, { text: b, options: { fontSize: 30, bold: true, color: HEX.accent1 } }], { x: x + 0.25, y: sy + 1.25, w: sw - 0.5, h: 0.6, margin: 0, valign: "middle", isTextBox: true });
      s.addText(u, { x: x + 0.25, y: sy + 1.82, w: sw - 0.5, h: 0.55, fontSize: 14, color: HEX.accent5, margin: 0, isTextBox: true });
    }
    // adaptations + formula
    const ay = 4.45;
    s.addText("Con l'allenamento regolare il cuore si adatta", { x: M, y: ay, w: 7.6, h: 0.45, fontSize: 18, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(richBullets([
      "**Bradicardia da sport:** a riposo il cuore batte più lentamente (anche 40–60 battiti/min)",
      "**Cuore d'atleta:** cavità più ampie e pareti più forti, più sangue a ogni battito",
      "**Più capillari** nei muscoli e più sangue in circolo",
      "**Pressione arteriosa** a riposo più bassa e recupero più veloce dopo lo sforzo",
    ], { space: 4 }), { x: M, y: ay + 0.5, w: 7.6, h: 2.0, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    const fx = 8.6, fw = W - M - fx;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: fx, y: ay, w: fw, h: 2.3, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Formula frequenza massima" });
    s.addText("Frequenza cardiaca massima", { x: fx + 0.25, y: ay + 0.2, w: fw - 0.5, h: 0.4, fontSize: 16, bold: true, color: "F2C9CF", margin: 0, isTextBox: true });
    s.addText("220 − età", { x: fx + 0.25, y: ay + 0.65, w: fw - 0.5, h: 0.8, fontSize: 36, bold: true, color: "FFFFFF", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
    s.addText("Esempio: a 16 anni ≈ 204 battiti/min. Il polso si misura al polso (arteria radiale) o al collo (carotide).", { x: fx + 0.25, y: ay + 1.45, w: fw - 0.5, h: 0.85, fontSize: 13, color: "FFFFFF", margin: 0, isTextBox: true });
    s.addNotes("Portata cardiaca = frequenza cardiaca × gittata sistolica. Il sedentario aumenta la portata soprattutto aumentando la frequenza; l'atleta, avendo una gittata maggiore, ottiene la stessa portata con meno battiti. La formula 220 meno età è una stima: serve per dosare l'intensità dell'allenamento.");
  }

  // =========================================================
  // PARTE 2 — RESPIRATORIO
  // =========================================================
  pres.addSection({ title: S2 });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: S2 });
    s.addText("PARTE 2", { placeholder: "body" });
    s.addText("Apparato respiratorio", { placeholder: "title" });
    await iconCircle(s, "lungs", 9.6, 2.0, 2.8, HEX.accent2, "FFFFFF", "Icona polmoni");
    s.addNotes("L'apparato respiratorio porta l'ossigeno dall'aria al sangue ed elimina l'anidride carbonica prodotta dalle cellule.");
  }

  // Organi
  {
    const s = contentSlide(S2, "Gli organi della respirazione");
    // left: path of air
    const path = [
      ["Naso e cavità nasali", "filtrano, scaldano e umidificano l'aria"],
      ["Faringe", "incrocio tra vie dell'aria e del cibo"],
      ["Laringe", "contiene le corde vocali; l'epiglottide chiude la via durante la deglutizione"],
      ["Trachea", "tubo di anelli di cartilagine, lungo circa 12 cm"],
      ["Bronchi e bronchioli", "si ramificano come un albero dentro i polmoni"],
      ["Alveoli polmonari", "circa 300 milioni di sacchetti dove avviene lo scambio di gas"],
    ];
    const px = M, py = 1.4, ph = 0.82, pw = 7.4;
    s.addText("Il percorso dell'aria", { x: px, y: py - 0.05, w: pw, h: 0.4, fontSize: 18, bold: true, color: C.accent2, margin: 0, isTextBox: true });
    path.forEach(([t, d], i) => {
      const y = py + 0.45 + i * ph;
      s.addShape(pres.shapes.OVAL, { x: px, y: y + 0.08, w: 0.52, h: 0.52, fill: { color: i < 3 ? HEX.accent2 : HEX.dk2 }, line: { color: i < 3 ? HEX.accent2 : HEX.dk2 }, objectName: "Passo " + (i + 1) });
      s.addText(String(i + 1), { x: px, y: y + 0.08, w: 0.52, h: 0.52, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: t, options: { bold: true, breakLine: true } }, { text: d, options: { fontSize: 14 } }], { x: px + 0.75, y, w: pw - 0.75, h: ph - 0.05, fontSize: 16, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    });
    s.addText("1–3: vie aeree superiori · 4–6: vie aeree inferiori", { x: px + 0.75, y: py + 0.45 + 6 * ph - 0.02, w: pw - 0.75, h: 0.3, fontSize: 12, italic: true, color: C.accent5, margin: 0, isTextBox: true });
    // right cards
    const rx = 8.45, rw = W - M - rx;
    await card(s, { x: rx, y: 1.4, w: rw, h: 2.55, tint: T.teal, color: HEX.accent2, iconName: "lungs", label: "I polmoni", items: ["Destro con 3 lobi, sinistro con 2 (fa spazio al cuore)", "Avvolti dalla pleura, una doppia membrana", "Superficie degli alveoli: circa 70–100 m²"], fontSize: 14, space: 4 });
    await card(s, { x: rx, y: 4.15, w: rw, h: 2.6, tint: T.gray, color: HEX.dk2, iconName: "muscle", label: "I muscoli respiratori", items: ["**Diaframma:** il muscolo principale, separa torace e addome", "**Intercostali:** muovono le costole", "**Accessori** (collo, addominali): entrano in gioco sotto sforzo"], fontSize: 14, space: 4 });
    s.addNotes("Respirare col naso è importante perché l'aria arriva ai polmoni filtrata, calda e umida. Nello sforzo intenso si respira anche con la bocca per far entrare più aria.");
  }

  // La respirazione
  {
    const s = contentSlide(S2, "La respirazione");
    lead(s, "Un atto respiratorio è fatto di due fasi. A riposo respiriamo 12–16 volte al minuto, circa mezzo litro d'aria per volta.", 1.2, 0.5);
    const cw = (W - 2 * M - 0.3) / 2, cy = 1.9, ch = 2.6;
    await card(s, { x: M, y: cy, w: cw, h: ch, tint: T.teal, color: HEX.accent2, iconName: "arrowDown", label: "Inspirazione · fase attiva", items: ["Il diaframma si contrae e si abbassa", "Gli intercostali sollevano le costole", "Il torace si allarga, la pressione nei polmoni scende e l'aria entra"], fontSize: 15, space: 4 });
    await card(s, { x: M + cw + 0.3, y: cy, w: cw, h: ch, tint: T.gray, color: HEX.accent5, iconName: "arrowUp", label: "Espirazione · fase passiva a riposo", items: ["I muscoli si rilassano e il diaframma risale", "Il torace si restringe e l'aria esce", "Sotto sforzo diventa attiva: lavorano addominali e intercostali interni"], fontSize: 15, space: 4 });
    // gas exchange band
    const gy = 4.75, gh = 2.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: gy, w: W - 2 * M, h: gh, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Scambi gassosi" });
    s.addText("Gli scambi gassosi avvengono per diffusione", { x: M + 0.3, y: gy + 0.15, w: 7, h: 0.45, fontSize: 18, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
    const ex = [
      ["Respirazione esterna", "Negli alveoli: l'O₂ passa nel sangue, la CO₂ passa nell'aria da espirare"],
      ["Respirazione interna", "Nei tessuti: il sangue cede O₂ alle cellule e raccoglie la CO₂"],
      ["Aria in entrata e in uscita", "Inspirata: 21% di O₂ · Espirata: circa 16% di O₂ e più CO₂"],
    ];
    const ew = (W - 2 * M - 0.6 - 0.6) / 3;
    ex.forEach(([t, d], i) => {
      s.addText([{ text: t, options: { bold: true, color: "F2C9CF", breakLine: true } }, { text: d, options: { color: "FFFFFF" } }], { x: M + 0.3 + i * (ew + 0.3), y: gy + 0.65, w: ew, h: 1.25, fontSize: 14, valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("La respirazione è controllata automaticamente dai centri respiratori nel tronco encefalico, che misurano soprattutto la quantità di CO₂ nel sangue. Possiamo però controllarla anche volontariamente, ad esempio in apnea.");
  }

  // Respirazione durante l'esercizio
  {
    const s = contentSlide(S2, "La respirazione durante l'esercizio");
    // chart left
    const chartW = 6.3;
    s.addChart(pres.charts.BAR, [{ name: "Ventilazione polmonare", labels: ["Riposo", "Corsa leggera", "Sforzo intenso", "Atleta al massimo"], values: [6, 40, 100, 150] }], {
      x: M, y: 1.3, w: chartW, h: 5.4,
      barDir: "col",
      chartColors: [HEX.accent2],
      showTitle: true, title: "Aria respirata al minuto (litri, valori indicativi)", titleFontSize: 14, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFontBold: true,
      catAxisLabelColor: HEX.accent5, valAxisLabelColor: HEX.accent5, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
      valGridLine: { color: "E5E7EB", size: 0.5 }, catGridLine: { style: "none" },
      valAxisMaxVal: 175, valAxisMinVal: 0, valAxisMajorUnit: 25, barGapWidthPct: 60,
      showLegend: false,
    });
    const rx = M + chartW + 0.4, rw = W - M - rx;
    s.addText("Cosa cambia sotto sforzo", { x: rx, y: 1.35, w: rw, h: 0.45, fontSize: 18, bold: true, color: C.accent2, margin: 0, isTextBox: true });
    s.addText(richBullets([
      "**Frequenza respiratoria:** da 12–16 fino a 40–50 atti al minuto",
      "**Volume corrente:** da 0,5 litri fino a 2–3 litri per atto",
      "**Prima reazione immediata** (nervosa), poi graduale in base a CO₂ e acido lattico",
      "**Debito di ossigeno:** finito lo sforzo continuiamo ad ansimare per recuperare",
      "**Allenamento:** muscoli respiratori più efficienti e maggiore massimo consumo di ossigeno (VO₂max)",
    ], { space: 5 }), { x: rx, y: 1.9, w: rw, h: 3.35, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: 5.35, w: rw, h: 1.4, rectRadius: 0.12, fill: { color: T.teal }, line: { color: T.teal }, objectName: "Consiglio respirazione" });
    await iconCircle(s, "lightbulb", rx + 0.2, 5.55, 0.5, HEX.accent2, "FFFFFF", "Icona consiglio");
    s.addText("Consiglio: respira con ritmo regolare, accordato al movimento; espira durante la fase di sforzo e non trattenere il fiato.", { x: rx + 0.85, y: 5.45, w: rw - 1.05, h: 1.2, fontSize: 14, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addNotes("La ventilazione polmonare è il prodotto di frequenza respiratoria per volume corrente. Trattenere il fiato durante uno sforzo massimale (manovra di Valsalva) fa salire molto la pressione: meglio espirare nella fase di spinta, ad esempio quando si solleva un peso.");
  }

  // =========================================================
  // PARTE 3 — PRIMO SOCCORSO
  // =========================================================
  pres.addSection({ title: S3 });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: S3 });
    s.addText("PARTE 3", { placeholder: "body" });
    s.addText("Il primo soccorso", { placeholder: "title" });
    await iconCircle(s, "firstaid", 9.6, 2.0, 2.8, "FFFFFF", HEX.accent1, "Icona primo soccorso");
    s.addNotes("Il primo soccorso è l'aiuto che chiunque può dare a una persona infortunata o malata in attesa dei soccorsi professionali. Prestarlo è un dovere: in Italia l'omissione di soccorso è un reato (art. 593 del Codice Penale). Basta anche solo chiamare i soccorsi.");
  }

  // P.A.S.
  {
    const s = contentSlide(S3, "Come si presta il primo soccorso: P.A.S.");
    lead(s, "Tre azioni, sempre in questo ordine. Il soccorritore non deve mai diventare una seconda vittima.", 1.2, 0.5);
    const steps = [
      ["P", "Proteggere", "shield", HEX.accent3, ["Valuta i pericoli: traffico, fuoco, corrente elettrica, gas, crolli", "Metti in sicurezza te stesso, poi la scena e infine la vittima", "Usa guanti monouso se c'è sangue", "Sposta la vittima solo se c'è un pericolo imminente"]],
      ["A", "Avvertire", "phone", HEX.accent1, ["Chiama il 112, il Numero Unico di Emergenza", "Di' dove sei, cosa è successo, quante persone sono coinvolte e come stanno", "Rispondi alle domande e non riattaccare per primo"]],
      ["S", "Soccorrere", "heartplus", HEX.accent4, ["Valuta coscienza e respiro", "Esegui solo le manovre che conosci", "Rassicura la vittima, coprila e non lasciarla sola", "Non dare da bere né da mangiare"]],
    ];
    const cy = 1.95, ch = 4.8, cw = (W - 2 * M - 0.6) / 3;
    steps.forEach(([l, t, ico, col, it], i) => {
      const x = M + i * (cw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: T.gray }, line: { color: T.gray }, objectName: "Passo " + t });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: cy + 0.3, w: 1.0, h: 1.0, fill: { color: col }, line: { color: col }, objectName: "Lettera " + l });
      s.addText(l, { x: x + 0.3, y: cy + 0.3, w: 1.0, h: 1.0, fontSize: 40, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
      s.addText(t, { x: x + 1.5, y: cy + 0.3, w: cw - 1.7, h: 1.0, fontSize: 24, bold: true, color: col, valign: "middle", margin: 0, isTextBox: true });
      s.addText(richBullets(it, { space: 10 }), { x: x + 0.3, y: cy + 1.55, w: cw - 0.55, h: ch - 1.75, fontSize: 17, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("Proteggere, Avvertire, Soccorrere. Prima di tutto la propria sicurezza: un soccorritore ferito non aiuta nessuno. Il 112 è gratuito e si può chiamare anche senza credito o senza SIM. Al telefono restare calmi e lasciare che sia l'operatore a chiudere la chiamata.");
  }

  // Valutazione della vittima
  {
    const s = contentSlide(S3, "Valutare la vittima");
    const steps = [
      ["1", "Coscienza", "Chiamala ad alta voce e scuotila delicatamente per le spalle. Risponde?"],
      ["2", "Respiro", "Apri le vie aeree (testa indietro, mento in su) e per 10 secondi: Guardo il torace, Ascolto e Sento il respiro (G.A.S.)"],
      ["3", "Decidi", "In base a ciò che hai visto, scegli cosa fare e chiama il 112 se non l'hai già fatto"],
    ];
    const sy = 1.35, sh = 1.05;
    steps.forEach(([n, t, d], i) => {
      const y = sy + i * (sh + 0.12);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: W - 2 * M, h: sh, rectRadius: 0.1, fill: { color: T.gray }, line: { color: T.gray }, objectName: "Fase " + t });
      s.addShape(pres.shapes.OVAL, { x: M + 0.25, y: y + 0.2, w: 0.65, h: 0.65, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Numero fase " + n });
      s.addText(n, { x: M + 0.25, y: y + 0.2, w: 0.65, h: 0.65, fontSize: 22, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(t, { x: M + 1.1, y, w: 2.0, h: sh, fontSize: 20, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: M + 3.1, y, w: W - 2 * M - 3.3, h: sh, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    });
    const oy = 4.95, oh = 1.8, ow = (W - 2 * M - 0.6) / 3;
    const outs = [
      [T.green, HEX.accent4, "Cosciente", "Lasciala nella posizione in cui sta meglio, chiedi cosa è successo e controllala"],
      [T.amber, HEX.accent3, "Incosciente ma respira", "Posizione laterale di sicurezza (se non c'è un trauma), chiama il 112 e controlla il respiro"],
      [T.red, HEX.accent1, "Incosciente e non respira", "Chiama il 112, fatti portare un defibrillatore (DAE) e inizia subito la RCP"],
    ];
    outs.forEach(([tint, col, t, d], i) => {
      const x = M + i * (ow + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: oy, w: ow, h: oh, rectRadius: 0.12, fill: { color: tint }, line: { color: tint }, objectName: "Esito " + t });
      s.addText([{ text: t, options: { bold: true, color: col, fontSize: 18, breakLine: true } }, { text: d, options: { fontSize: 16, color: HEX.dk1 } }], { x: x + 0.25, y: oy + 0.15, w: ow - 0.5, h: oh - 0.3, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 4 });
    });
    s.addNotes("Posizione laterale di sicurezza: la vittima è su un fianco, con la testa estesa e la bocca rivolta verso il basso, così la lingua non chiude le vie aeree e un eventuale vomito può uscire. Non si usa se si sospetta un trauma alla colonna vertebrale. Il respiro va ricontrollato di continuo.");
  }

  // Contusioni
  await (async () => {
    const s = await scheda(S3, "Le contusioni", {
      def: "Lesione dei tessuti molli (pelle, muscoli) causata da un urto contro un corpo contundente, senza rottura della pelle. È il trauma più comune nello sport.",
      cols: [
        SEGNI(["Dolore nel punto colpito", "Gonfiore (edema)", "Livido o ematoma: sangue uscito dai piccoli vasi sotto la pelle", "Arrossamento e calore", "Difficoltà a muovere la parte se la contusione è forte"]),
        FARE(["Interrompi l'attività", "Applica subito ghiaccio avvolto in un panno per 15–20 minuti, più volte al giorno", "Tieni la parte a riposo e sollevata", "Se il dolore è forte o non passa, vai dal medico"]),
        NONFARE(["Non mettere il ghiaccio direttamente sulla pelle", "Non applicare calore nei primi 2–3 giorni", "Non massaggiare la zona", "Non sottovalutare un colpo forte a testa, addome o torace"]),
      ],
    });
    s.addNotes("Il freddo restringe i vasi sanguigni (vasocostrizione): limita il gonfiore e l'ematoma e riduce il dolore. Il calore invece dilata i vasi e nei primi giorni peggiora il gonfiore.");
  })();

  // Ferite
  await (async () => {
    const s = await scheda(S3, "Le ferite", {
      def: "Interruzione della pelle con possibile lesione dei tessuti sotto. Tipi: abrasioni (sbucciature), ferite da taglio, da punta, lacero-contuse (pelle strappata da un urto).",
      cols: [
        { tint: T.amber, color: HEX.accent3, iconName: "warn", label: "Quando è grave", items: ["Sanguina molto o a fiotti", "È profonda, lunga o con bordi aperti", "È sul viso, sulle mani o vicino agli occhi", "C'è un oggetto conficcato", "È causata da un morso o è molto sporca"] },
        FARE(["Lavati le mani e indossa guanti monouso", "Lava la ferita con acqua corrente o soluzione fisiologica", "Disinfetta dal centro verso l'esterno con un disinfettante non alcolico", "Copri con garza sterile e cerotto", "Ferita profonda o sporca: medico e controllo dell'antitetanica"]),
        NONFARE(["Non usare cotone idrofilo: lascia fili nella ferita", "Non versare alcol nella ferita", "Non togliere oggetti conficcati: fissali e chiama il 112", "Non soffiare sulla ferita e non toccarla a mani nude"]),
      ],
    });
    s.addNotes("La vaccinazione antitetanica va richiamata ogni 10 anni. Il tetano è causato da un batterio presente nella terra e nella ruggine. Un oggetto conficcato può fare da tappo: toglierlo può scatenare un'emorragia.");
  })();

  // Emorragie (2 slide)
  {
    const s = contentSlide(S3, "Le emorragie: i tipi");
    lead(s, "L'emorragia è la fuoriuscita di sangue dai vasi. Perdere molto sangue in poco tempo è pericoloso e può portare allo shock.", 1.2, 0.5);
    const types = [
      ["Arteriosa", HEX.accent1, T.red, "Sangue rosso vivo che esce a fiotti, al ritmo del battito. È la più grave e va fermata subito."],
      ["Venosa", HEX.dk2, T.red, "Sangue rosso scuro che esce in modo continuo e lento."],
      ["Capillare", HEX.accent3, T.amber, "Il sangue esce a gocce o a velo, come in una sbucciatura. Si ferma facilmente."],
    ];
    const ty = 1.9, th = 2.1, tw = (W - 2 * M - 0.6) / 3;
    for (let i = 0; i < types.length; i++) {
      const [t, col, tint, d] = types[i];
      const x = M + i * (tw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ty, w: tw, h: th, rectRadius: 0.12, fill: { color: tint }, line: { color: tint }, objectName: "Emorragia " + t });
      await iconCircle(s, "drop", x + 0.25, ty + 0.25, 0.6, col, "FFFFFF", "Goccia " + t);
      s.addText(t, { x: x + 1.0, y: ty + 0.25, w: tw - 1.2, h: 0.6, fontSize: 20, bold: true, color: col, valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: x + 0.25, y: ty + 1.0, w: tw - 0.5, h: th - 1.1, fontSize: 16, color: HEX.dk1, valign: "top", margin: 0, isTextBox: true });
    }
    s.addText("In base a dove finisce il sangue", { x: M, y: 4.25, w: 8, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
    const where = [
      ["Esterna", "Il sangue esce da una ferita della pelle ed è visibile."],
      ["Interna", "Il sangue resta dentro il corpo (addome, torace, cranio). Non si vede: si riconosce dai segni dello shock."],
      ["Esteriorizzata", "Il sangue è interno ma esce da un'apertura naturale: naso, orecchio, bocca."],
    ];
    where.forEach(([t, d], i) => {
      const x = M + i * (tw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.8, w: tw, h: 1.95, rectRadius: 0.12, fill: { color: T.gray }, line: { color: T.gray }, objectName: "Sede " + t });
      s.addText([{ text: t, options: { bold: true, color: HEX.dk2, fontSize: 17, breakLine: true } }, { text: d, options: { fontSize: 16, color: HEX.dk1 } }], { x: x + 0.25, y: 4.95, w: tw - 0.5, h: 1.7, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 4 });
    });
    s.addNotes("Un'emorragia interna è insidiosa perché non si vede: dopo un trauma forte all'addome o al torace bisogna sospettarla se compaiono pallore, sudore freddo, polso rapido e debole. In quel caso: 112 subito.");
  }

  {
    const s = contentSlide(S3, "Come fermare un'emorragia esterna");
    const steps = [
      ["1", "Proteggiti", "Indossa guanti monouso, fai sdraiare la vittima"],
      ["2", "Comprimi", "Premi forte sulla ferita con garza o un panno pulito per almeno 10 minuti, senza mai sollevarla per controllare"],
      ["3", "Solleva", "Se è un arto e non c'è frattura, tienilo più in alto del cuore"],
      ["4", "Aggiungi", "Se il sangue passa, aggiungi altre garze sopra, senza togliere le prime, e fai un bendaggio compressivo"],
      ["5", "Punti di compressione", "Se non basta, premi l'arteria a monte contro l'osso: braccio (arteria omerale) o inguine (femorale)"],
    ];
    const sx = M, sw = 7.6, sy = 1.35, sh = 1.0;
    steps.forEach(([n, t, d], i) => {
      const y = sy + i * (sh + 0.08);
      s.addShape(pres.shapes.OVAL, { x: sx, y: y + 0.15, w: 0.62, h: 0.62, fill: { color: HEX.accent1 }, line: { color: HEX.accent1 }, objectName: "Passo " + n });
      s.addText(n, { x: sx, y: y + 0.15, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: t, options: { bold: true, color: HEX.dk2, fontSize: 17, breakLine: true } }, { text: d, options: { fontSize: 14, color: HEX.dk1 } }], { x: sx + 0.85, y, w: sw - 0.85, h: sh, valign: "middle", margin: 0, isTextBox: true });
    });
    const rx = 8.6, rw = W - M - rx;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: 1.35, w: rw, h: 2.95, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Laccio emostatico" });
    await iconCircle(s, "warn", rx + 0.25, 1.55, 0.55, "FFFFFF", HEX.accent1, "Icona attenzione");
    s.addText("Il laccio emostatico", { x: rx + 0.95, y: 1.55, w: rw - 1.15, h: 0.55, fontSize: 18, bold: true, color: "FFFFFF", valign: "middle", margin: 0, isTextBox: true });
    s.addText("Solo come ultima risorsa, per un'emorragia grave di un arto che non si ferma in altro modo. Va stretto a monte della ferita, annotando l'ora. Non va mai allentato: lo toglie il medico.", { x: rx + 0.25, y: 2.25, w: rw - 0.5, h: 1.95, fontSize: 14, color: "FFFFFF", valign: "top", margin: 0, isTextBox: true });
    await card(s, { x: rx, y: 4.5, w: rw, h: 2.25, tint: T.red, color: HEX.accent1, iconName: "phone", label: "Chiama il 112", items: ["Se il sangue esce a fiotti o non si ferma", "Se la vittima è pallida, sudata e confusa"], fontSize: 14, space: 4 });
    s.addNotes("La compressione diretta ferma la grande maggioranza delle emorragie esterne. L'errore più comune è sollevare la garza per guardare: così si stacca il coagulo che si stava formando.");
  }

  // Epistassi
  await (async () => {
    const s = await scheda(S3, "L'emorragia nasale (epistassi)", {
      def: "Uscita di sangue dal naso, frequente nello sport. Può essere causata da un colpo, da aria secca, dal soffiarsi il naso con forza o da pressione alta.",
      cols: [
        FARE(["Fai sedere la persona con la testa leggermente in avanti", "Stringi la parte morbida del naso tra pollice e indice per 10 minuti", "Falla respirare con la bocca", "Metti ghiaccio avvolto in un panno sulla radice del naso", "Se non si ferma in 20 minuti, o dopo un colpo forte, vai al pronto soccorso"]),
        NONFARE(["Non piegare la testa all'indietro: il sangue scende in gola e può causare vomito", "Non far sdraiare la persona", "Non soffiare il naso per alcune ore", "Non infilare cotone nelle narici"]),
        { tint: T.amber, color: HEX.accent3, iconName: "warn", label: "Attenzione", items: ["Dopo un trauma alla testa, un liquido chiaro o sangue da naso o orecchie può indicare una frattura del cranio: chiama il 112", "Se il naso è deformato dopo un colpo, potrebbe essere rotto"] },
      ],
    });
    s.addNotes("È un errore molto diffuso mettere la testa all'indietro: il sangue non si ferma, scende in gola, può andare nelle vie aeree o essere ingerito e far vomitare. La testa va in avanti.");
  })();

  // =========================================================
  // PARTE 4 — TRAUMI APPARATO LOCOMOTORE
  // =========================================================
  pres.addSection({ title: S4 });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: S4 });
    s.addText("PARTE 4", { placeholder: "body" });
    s.addText("Traumi dell'apparato locomotore", { placeholder: "title" });
    await iconCircle(s, "bone", 9.6, 2.0, 2.8, HEX.accent3, "FFFFFF", "Icona osso");
    s.addNotes("Traumi a muscoli, tendini, articolazioni e ossa: sono gli infortuni più frequenti durante l'attività sportiva.");
  }

  // PRICE
  {
    const s = contentSlide(S4, "Il protocollo P.R.I.C.E.");
    lead(s, "Le prime cure per quasi tutti i traumi di muscoli, tendini e articolazioni, da applicare nelle prime 48–72 ore.", 1.2, 0.5);
    const L = [
      ["P", "Protezione", "shield", "Proteggi la parte da altri danni: stop all'attività, eventuale tutore"],
      ["R", "Riposo", "bed", "Non caricare e non muovere la parte infortunata"],
      ["I", "Ice: ghiaccio", "ice", "15–20 minuti ogni 2–3 ore, sempre avvolto in un panno"],
      ["C", "Compressione", "bandage", "Bendaggio elastico, non troppo stretto, per limitare il gonfiore"],
      ["E", "Elevazione", "arrowUp", "Tieni la parte sollevata sopra il livello del cuore"],
    ];
    const ty = 1.95, th = 3.55, tw = (W - 2 * M - 4 * 0.25) / 5;
    for (let i = 0; i < L.length; i++) {
      const [l, t, ico, d] = L[i];
      const x = M + i * (tw + 0.25);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ty, w: tw, h: th, rectRadius: 0.12, fill: { color: T.amber }, line: { color: T.amber }, objectName: "Lettera " + l });
      s.addText(l, { x, y: ty + 0.2, w: tw, h: 0.95, fontSize: 54, bold: true, color: C.accent3, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
      s.addText(t, { x: x + 0.15, y: ty + 1.2, w: tw - 0.3, h: 0.45, fontSize: 17, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(d, { x: x + 0.2, y: ty + 1.75, w: tw - 0.4, h: th - 1.9, fontSize: 16, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.75, w: W - 2 * M, h: 1.0, rectRadius: 0.12, fill: { color: T.red }, line: { color: T.red }, objectName: "Regola nessun calore" });
    await iconCircle(s, "times", M + 0.2, 5.97, 0.56, HEX.accent1, "FFFFFF", "Icona divieto");
    s.addText([{ text: "Nei primi giorni evita calore, massaggi e alcol: ", options: { bold: true } }, { text: "aumentano il gonfiore e il sanguinamento interno." }], { x: M + 1.0, y: 5.75, w: W - 2 * M - 1.2, h: 1.0, fontSize: 16, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    s.addNotes("P.R.I.C.E. è l'acronimo inglese di Protection, Rest, Ice, Compression, Elevation. Va bene per contusioni, stiramenti, strappi e distorsioni. Per fratture e lussazioni la priorità è immobilizzare e chiamare i soccorsi.");
  }

  // Muscoli: contrattura, stiramento, strappo
  {
    const s = contentSlide(S4, "Contrattura, stiramento e strappo muscolare");
    lead(s, "Tre lesioni del muscolo di gravità crescente: dalla contrattura (nessuna fibra rotta) allo strappo (fibre rotte).", 1.2, 0.5);
    const cols = [
      ["Contrattura", HEX.accent3, T.amber, "Lieve", ["**Cos'è:** contrazione involontaria e prolungata del muscolo, che resta duro e teso; le fibre non sono rotte", "**Segni:** dolore che compare gradualmente, muscolo rigido, si continua a muoversi ma con fastidio", "**Cosa fare:** riposo, calore dopo qualche ora, massaggio leggero, stretching dolce"]],
      ["Stiramento", HEX.accent1, T.red, "Media", ["**Cos'è:** il muscolo viene allungato oltre il limite; le fibre si danneggiano ma non si rompono", "**Segni:** dolore acuto durante lo sforzo, che costringe a rallentare", "**Cosa fare:** stop, P.R.I.C.E., riposo per 1–2 settimane"]],
      ["Strappo", HEX.dk2, T.red, "Grave", ["**Cos'è:** rottura di un numero più o meno grande di fibre muscolari", "**Segni:** fitta improvvisa, come una coltellata; impossibile continuare; poi ematoma e a volte un avvallamento", "**Cosa fare:** stop, P.R.I.C.E., visita medica ed ecografia; recupero di settimane"]],
    ];
    const cy = 1.9, ch = 4.85, cw = (W - 2 * M - 0.6) / 3;
    cols.forEach(([t, col, tint, g, it], i) => {
      const x = M + i * (cw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: tint }, line: { color: tint }, objectName: "Lesione " + t });
      s.addText(t, { x: x + 0.25, y: cy + 0.2, w: cw - 1.6, h: 0.55, fontSize: 21, bold: true, color: col, valign: "middle", margin: 0, isTextBox: true });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + cw - 1.3, y: cy + 0.27, w: 1.05, h: 0.42, rectRadius: 0.2, fill: { color: col }, line: { color: col }, objectName: "Gravità " + t });
      s.addText(g, { x: x + cw - 1.3, y: cy + 0.27, w: 1.05, h: 0.42, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText(richBullets(it, { space: 10 }), { x: x + 0.25, y: cy + 0.95, w: cw - 0.5, h: ch - 1.1, fontSize: 16, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    });
    s.addNotes("Le cause principali sono un riscaldamento insufficiente, la stanchezza, il freddo e la disidratazione. Attenzione: la contrattura può essere trattata col calore, lo strappo no. Nel dubbio, nei primi giorni, sempre ghiaccio. Prevenzione: riscaldamento, gradualità, idratazione e stretching.");
  }

  // Tendinopatie
  await (async () => {
    const s = await scheda(S4, "Le tendinopatie", {
      def: "Sofferenza del tendine (la struttura che collega il muscolo all'osso) dovuta al sovraccarico: movimenti ripetuti nel tempo, non un singolo trauma. Un tempo si chiamavano tendiniti.",
      cols: [
        CAUSE(["Allenamenti troppo intensi o aumentati troppo in fretta", "Gesti ripetuti sempre uguali", "Scarpe o attrezzi inadatti, terreni duri", "Postura o tecnica sbagliata", "Esempi: tendine d'Achille, tendine rotuleo, gomito del tennista"]),
        SEGNI(["Dolore lungo il tendine, che aumenta con il movimento", "Rigidità al mattino o all'inizio dell'attività", "A volte gonfiore e scricchiolio", "Dolore alla pressione del tendine"]),
        FARE(["Riduci o sospendi l'attività che provoca dolore", "Ghiaccio nella fase dolorosa", "Visita medica e fisioterapia", "Riprendi gradualmente", "Prevenzione: riscaldamento, progressione dei carichi, scarpe adatte"]),
      ],
    });
    s.addNotes("Le tendinopatie tendono a cronicizzare se si continua ad allenarsi col dolore. Il messaggio chiave è: il dolore è un segnale da ascoltare.");
  })();

  // Distorsione
  await (async () => {
    const s = await scheda(S4, "La distorsione", {
      def: "Trauma di un'articolazione in cui le ossa si spostano per un attimo oltre il normale e poi tornano al loro posto. Si possono lesionare legamenti e capsula. Le più frequenti: caviglia e ginocchio.",
      cols: [
        SEGNI(["Dolore all'articolazione", "Gonfiore rapido", "Ematoma nelle ore successive", "Difficoltà a muovere l'articolazione o a camminare", "Gravità: dalla lesione lieve dei legamenti (1° grado) alla rottura completa (3° grado)"]),
        FARE(["Interrompi subito l'attività", "Applica il P.R.I.C.E.", "Immobilizza l'articolazione con una benda o un tutore", "Non caricare il peso sull'arto", "Vai dal medico: può servire una radiografia per escludere una frattura"]),
        NONFARE(["Non continuare a giocare \"per scaldarla\"", "Non massaggiare e non applicare calore", "Non togliere subito la scarpa se la caviglia è molto gonfia e dolente: può contenere il gonfiore fino all'arrivo dei soccorsi"]),
      ],
    });
    s.addNotes("La distorsione di caviglia avviene quasi sempre con il piede che si gira verso l'interno. Se non viene curata bene si rischia l'instabilità cronica e altre distorsioni. Il rinforzo e la propriocezione aiutano a prevenirle.");
  })();

  // Lussazione
  await (async () => {
    const s = await scheda(S4, "La lussazione", {
      def: "Le ossa di un'articolazione perdono il contatto e restano fuori posto. È più grave della distorsione: si lesionano capsula e legamenti. Le più frequenti: spalla, dita, gomito.",
      cols: [
        SEGNI(["Dolore molto forte", "Deformazione evidente dell'articolazione", "Impossibilità di muoverla", "Gonfiore", "Possibili formicolio o perdita di sensibilità se sono coinvolti nervi o vasi"]),
        FARE(["Chiama il 112 o porta la persona al pronto soccorso", "Immobilizza l'articolazione nella posizione in cui si trova, con bende o un foulard", "Applica ghiaccio avvolto in un panno", "Tieni la persona tranquilla e non darle da bere né da mangiare"]),
        NONFARE(["Non tentare mai di rimettere a posto l'articolazione: puoi rompere l'osso o danneggiare nervi e vasi", "Non far muovere l'arto", "Non sottovalutarla se si riduce da sola: serve comunque una visita"]),
      ],
    });
    s.addNotes("Nella lussazione della spalla la vittima tiene di solito il braccio attaccato al corpo e la spalla appare squadrata. La riduzione deve farla solo un medico, spesso dopo una radiografia. Niente cibo né bevande: potrebbe servire un'anestesia.");
  })();

  // Frattura
  await (async () => {
    const s = await scheda(S4, "La frattura", {
      def: "Rottura di un osso. Chiusa se la pelle è intatta, esposta se l'osso esce dalla pelle. Composta se i pezzi restano allineati, scomposta se si spostano.",
      cols: [
        SEGNI(["Dolore forte che aumenta con il movimento", "Deformità o posizione innaturale dell'arto", "Gonfiore ed ematoma", "Impossibilità di usare l'arto (impotenza funzionale)", "A volte uno scricchiolio (crepitio); nella frattura esposta, osso visibile e sanguinamento"]),
        FARE(["Chiama il 112", "Non muovere la vittima e immobilizza l'arto nella posizione in cui si trova", "Blocca anche le articolazioni sopra e sotto la frattura, con stecche o cuscini", "Frattura esposta: copri con garza sterile e controlla l'emorragia", "Applica ghiaccio intorno, mai sopra l'osso esposto"]),
        NONFARE(["Non tentare di raddrizzare l'arto", "Non far rientrare l'osso nella frattura esposta", "Non spostare chi ha dolore a collo o schiena: sospetta una lesione della colonna", "Non dare da bere né da mangiare"]),
      ],
    });
    s.addNotes("In caso di sospetta lesione della colonna vertebrale (caduta dall'alto, forte dolore a collo o schiena, formicolio o perdita di forza alle gambe) la vittima non va mai mossa: si mantiene la testa ferma in asse con il corpo fino all'arrivo dei soccorsi.");
  })();

  // Riepilogo tabella
  {
    const s = contentSlide(S4, "Riepilogo: come distinguerli");
    const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.dk2 }, fontSize: 15, valign: "middle" } });
    const c0 = (t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: T.gray }, fontSize: 15, valign: "middle" } });
    const c = (t) => ({ text: t, options: { color: HEX.dk1, fontSize: 14, valign: "middle" } });
    const rows = [
      [hdr("Trauma"), hdr("Cosa si lesiona"), hdr("Segno tipico"), hdr("Primo intervento")],
      [c0("Contusione"), c("Pelle e muscoli, senza ferita"), c("Livido e gonfiore"), c("Ghiaccio e riposo")],
      [c0("Contrattura"), c("Muscolo teso, nessuna fibra rotta"), c("Muscolo duro, dolore graduale"), c("Riposo, calore, stretching dolce")],
      [c0("Strappo"), c("Fibre muscolari rotte"), c("Fitta improvvisa, si deve smettere"), c("P.R.I.C.E. e medico")],
      [c0("Tendinopatia"), c("Tendine sovraccaricato"), c("Dolore che cresce con l'uso"), c("Riposo, ghiaccio, medico")],
      [c0("Distorsione"), c("Legamenti; l'articolazione torna in sede"), c("Gonfiore rapido all'articolazione"), c("P.R.I.C.E., immobilizzare, medico")],
      [c0("Lussazione"), c("L'articolazione resta fuori sede"), c("Deformità dell'articolazione"), c("Immobilizzare così com'è, 112")],
      [c0("Frattura"), c("Osso rotto"), c("Deformità dell'arto, dolore forte"), c("Non muovere, immobilizzare, 112")],
    ];
    s.addTable(rows, { x: M, y: 1.35, w: W - 2 * M, colW: [2.2, 3.6, 3.3, 3.03], rowH: 0.66, border: { type: "solid", pt: 0.75, color: "D1D5DB" }, margin: [0.05, 0.12, 0.05, 0.12], fontFace: THEME.bodyFontFace, objectName: "Tabella riepilogo traumi" });
    s.addNotes("Tabella utile per lo studio. Nel dubbio tra distorsione, lussazione e frattura ci si comporta come per la frattura: si immobilizza e si chiede un controllo medico.");
  }

  // =========================================================
  // PARTE 5 — EMERGENZE
  // =========================================================
  pres.addSection({ title: S5 });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: S5 });
    s.addText("PARTE 5", { placeholder: "body" });
    s.addText("Le emergenze", { placeholder: "title" });
    await iconCircle(s, "siren", 9.6, 2.0, 2.8, HEX.accent6, "FFFFFF", "Icona emergenza");
    s.addNotes("Situazioni in cui sono in pericolo le funzioni vitali: coscienza, respiro, circolazione. Qui chiamare subito il 112 è fondamentale.");
  }

  // Lipotimia
  await (async () => {
    const s = await scheda(S5, "La lipotimia (lo svenimento)", {
      def: "Breve perdita di coscienza, o sensazione di stare per svenire, perché arriva poco sangue al cervello. Di solito si risolve da sola in pochi secondi o minuti.",
      cols: [
        CAUSE(["Caldo e ambienti affollati", "Stare a lungo in piedi fermi", "Emozioni forti, paura, dolore, vista del sangue", "Digiuno o calo di zuccheri", "Alzarsi in piedi di scatto"]),
        SEGNI(["Prima: pallore, sudore freddo, vista annebbiata, nausea, ronzio nelle orecchie, debolezza", "Poi la persona si accascia", "Polso debole, respiro presente"]),
        FARE(["Ai primi segni fai sdraiare la persona", "Solleva le gambe di circa 30–40 cm", "Allenta gli abiti stretti e fai circolare aria", "Quando si riprende, falla rialzare lentamente; poi acqua o qualcosa di zuccherato", "Se non si riprende in 1–2 minuti o non respira normalmente: 112"]),
      ],
    });
    s.addNotes("Sollevare le gambe fa tornare più sangue verso il cuore e il cervello. Non dare mai da bere a una persona non completamente cosciente: può soffocare. Se lo svenimento avviene durante uno sforzo fisico va sempre fatto un controllo medico, perché può nascondere un problema cardiaco.");
  })();

  // Arresto cardiaco (2 slide)
  {
    const s = contentSlide(S5, "L'arresto cardiaco");
    lead(s, "Il cuore smette di pompare sangue in modo efficace: il cervello non riceve più ossigeno. La persona è incosciente e non respira, o respira in modo anomalo (boccheggia).", 1.2, 0.8);
    const stats = [
      ["4–6 min", "dopo i quali senza ossigeno il cervello inizia a subire danni"],
      ["−10%", "di probabilità di sopravvivenza per ogni minuto senza intervento"],
      ["×2–3", "aumenta la sopravvivenza se qualcuno inizia subito la RCP"],
    ];
    const sy = 2.2, sh = 1.9, sw = (W - 2 * M - 0.6) / 3;
    stats.forEach(([n, l], i) => {
      const x = M + i * (sw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: sy, w: sw, h: sh, rectRadius: 0.12, fill: { color: T.red }, line: { color: T.red }, objectName: "Dato " + n });
      s.addText(n, { x: x + 0.2, y: sy + 0.2, w: sw - 0.4, h: 0.85, fontSize: 40, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
      s.addText(l, { x: x + 0.3, y: sy + 1.05, w: sw - 0.6, h: 0.75, fontSize: 14, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
    });
    s.addText("La catena della sopravvivenza", { x: M, y: 4.35, w: 8, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
    const chain = [["phone", "Allarme precoce", "chiama il 112"], ["heartplus", "RCP precoce", "compressioni toraciche subito"], ["bolt", "Defibrillazione precoce", "con il DAE"], ["ambulance", "Soccorso avanzato", "medici e ospedale"]];
    const cw = (W - 2 * M - 3 * 0.25) / 4, cy = 4.9, ch = 1.85;
    for (let i = 0; i < chain.length; i++) {
      const [ico, t, d] = chain[i];
      const x = M + i * (cw + 0.25);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Anello " + t });
      await iconCircle(s, ico, x + 0.25, cy + 0.25, 0.6, "FFFFFF", HEX.dk2, "Icona " + t);
      s.addText(String(i + 1), { x: x + cw - 0.75, y: cy + 0.2, w: 0.5, h: 0.6, fontSize: 26, bold: true, color: "F2C9CF", align: "right", valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: t, options: { bold: true, fontSize: 16, breakLine: true } }, { text: d, options: { fontSize: 14 } }], { x: x + 0.25, y: cy + 0.95, w: cw - 0.5, h: 0.8, color: "FFFFFF", valign: "top", margin: 0, isTextBox: true });
    }
    s.addNotes("Nello sport l'arresto cardiaco può colpire anche giovani apparentemente sani, a causa di malattie cardiache non diagnosticate: per questo esistono la visita medico-sportiva e l'obbligo del defibrillatore negli impianti sportivi. I soccorsi impiegano in media diversi minuti per arrivare: chi è sul posto fa la differenza.");
  }

  {
    const s = contentSlide(S5, "La rianimazione cardiopolmonare (RCP)");
    const steps = [
      ["1", "Sicurezza e coscienza", "Verifica che la scena sia sicura. Chiama la vittima e scuotila: non risponde"],
      ["2", "Respiro", "Testa indietro e mento in su; guarda, ascolta e senti per 10 secondi: non respira normalmente"],
      ["3", "112 e DAE", "Chiama il 112 (meglio in vivavoce) e fai portare un defibrillatore"],
      ["4", "30 compressioni", "Al centro del torace, braccia tese, mani sovrapposte: 5–6 cm di profondità, 100–120 al minuto"],
      ["5", "2 ventilazioni", "Se sei addestrato, alterna 30 compressioni e 2 ventilazioni; altrimenti fai solo compressioni"],
      ["6", "Non fermarti", "Accendi il DAE e segui le sue istruzioni vocali; continua fino all'arrivo dei soccorsi"],
    ];
    const lw = 8.0, sh = 0.84, sy = 1.35;
    steps.forEach(([n, t, d], i) => {
      const y = sy + i * (sh + 0.06);
      s.addShape(pres.shapes.OVAL, { x: M, y: y + 0.12, w: 0.6, h: 0.6, fill: { color: i >= 3 ? HEX.accent1 : HEX.dk2 }, line: { color: i >= 3 ? HEX.accent1 : HEX.dk2 }, objectName: "Passo " + n });
      s.addText(n, { x: M, y: y + 0.12, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
      s.addText([{ text: t, options: { bold: true, color: HEX.dk2, fontSize: 17, breakLine: true } }, { text: d, options: { fontSize: 14, color: HEX.dk1 } }], { x: M + 0.8, y, w: lw - 0.8, h: sh, valign: "middle", margin: 0, isTextBox: true });
    });
    const rx = M + lw + 0.4, rw = W - M - rx;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: 1.35, w: rw, h: 3.05, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Ritmo RCP" });
    s.addText("Il ritmo da ricordare", { x: rx + 0.25, y: 1.5, w: rw - 0.5, h: 0.45, fontSize: 16, bold: true, color: "F2C9CF", margin: 0, isTextBox: true });
    s.addText("30 : 2", { x: rx + 0.25, y: 1.95, w: rw - 0.5, h: 1.1, fontSize: 60, bold: true, color: "FFFFFF", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
    s.addText("compressioni : ventilazioni\n100–120 compressioni al minuto\nprofondità 5–6 cm", { x: rx + 0.25, y: 3.05, w: rw - 0.5, h: 1.25, fontSize: 14, color: "FFFFFF", margin: 0, isTextBox: true, valign: "top" });
    await card(s, { x: rx, y: 4.6, w: rw, h: 2.15, tint: T.amber, color: HEX.accent3, iconName: "bolt", label: "Il DAE", items: ["Defibrillatore semiautomatico: analizza il cuore e dà la scarica solo se serve", "Può usarlo chiunque"], fontSize: 14, space: 4 });
    s.addNotes("Le compressioni devono permettere al torace di riespandersi completamente tra una e l'altra. Per tenere il ritmo si può pensare a una canzone a circa 100–120 battiti al minuto. Il DAE non può far male: dà la scarica solo se riconosce un ritmo defibrillabile. Durante la scarica nessuno deve toccare la vittima.");
  }

  // Shock
  await (async () => {
    const s = await scheda(S5, "Lo shock", {
      def: "Grave insufficienza della circolazione: gli organi non ricevono abbastanza sangue e ossigeno. Può essere causato da una forte emorragia (ipovolemico), da una reazione allergica (anafilattico), da un problema al cuore, da gravi ustioni o infezioni.",
      cols: [
        SEGNI(["Pelle pallida, fredda e sudata", "Polso rapido e debole", "Respiro rapido e superficiale", "Sete, agitazione, poi confusione e sonnolenza", "Nello shock anafilattico: gonfiore di viso e gola, difficoltà a respirare, macchie sulla pelle"]),
        FARE(["Chiama subito il 112", "Ferma eventuali emorragie esterne", "Fai sdraiare la persona e solleva le gambe (se non ci sono traumi)", "Coprila per evitare che perda calore", "Rassicurala e controlla coscienza e respiro"]),
        NONFARE(["Non dare da bere né da mangiare", "Non lasciarla sola", "Non sollevare le gambe se ci sono fratture alle gambe, al bacino o un trauma cranico", "Non perdere tempo: lo shock peggiora rapidamente"]),
      ],
    });
    s.addNotes("La posizione antishock (sdraiati con le gambe sollevate) aiuta il sangue a tornare verso gli organi vitali. Chi soffre di allergie gravi può avere con sé una penna di adrenalina autoiniettabile: si può aiutarlo a usarla.");
  })();

  // Trauma cranico
  await (async () => {
    const s = await scheda(S5, "Il trauma cranico", {
      def: "Un colpo alla testa che può danneggiare il cervello. Nello sport la forma più comune è la commozione cerebrale: un disturbo temporaneo del cervello, anche senza perdere i sensi.",
      cols: [
        { tint: T.amber, color: HEX.accent3, iconName: "warn", label: "Segnali di allarme", items: ["Perdita di coscienza, anche breve", "Mal di testa forte o che peggiora, vomito", "Confusione, non ricorda cosa è successo", "Sonnolenza, difficoltà a parlare o a camminare", "Pupille diverse, sangue o liquido chiaro da naso o orecchie"] },
        FARE(["Ferma subito l'attività: se c'è un dubbio, l'atleta esce dal campo", "Chiama il 112 se c'è uno dei segnali di allarme", "Tieni ferma la testa in asse con il corpo se si sospetta un trauma al collo", "Ghiaccio sul bernoccolo", "Non lasciare sola la persona per almeno 24 ore e osservala"]),
        NONFARE(["Non farla tornare a giocare lo stesso giorno", "Non muovere testa e collo se è a terra e ha dolore al collo", "Non togliere il casco a un motociclista", "Non dare farmaci contro il dolore senza parere medico"]),
      ],
    });
    s.addNotes("I sintomi di un trauma cranico possono comparire anche ore dopo il colpo: per questo la persona va osservata. Il ritorno allo sport dopo una commozione cerebrale deve essere graduale e autorizzato da un medico.");
  })();

  // Soffocamento
  {
    const s = contentSlide(S5, "Il soffocamento (ostruzione delle vie aeree)");
    lead(s, "Un corpo estraneo, di solito cibo, blocca la trachea. Il segno tipico: la persona si porta le mani alla gola.", 1.2, 0.5);
    const cw = (W - 2 * M - 0.3) / 2, cy = 1.9, ch = 3.0;
    await card(s, { x: M, y: cy, w: cw, h: ch, tint: T.amber, color: HEX.accent3, iconName: "cough", label: "Ostruzione parziale", items: ["**Segni:** la persona tossisce, parla e respira, anche con fatica", "**Cosa fare:** incoraggiala a tossire con forza e resta con lei", "**Non** darle colpi sulla schiena e non darle da bere"], fontSize: 16, space: 8 });
    await card(s, { x: M + cw + 0.3, y: cy, w: cw, h: ch, tint: T.red, color: HEX.accent1, iconName: "hand", label: "Ostruzione completa", items: ["**Segni:** non riesce a parlare, tossire né respirare; diventa rossa e poi bluastra", "**Cosa fare:** intervieni subito con colpi sulla schiena e manovra di Heimlich", "Fai chiamare il 112"], fontSize: 16, space: 8 });
    const steps = [
      ["5 colpi sulla schiena", "Sostieni il torace, falla piegare in avanti e dai 5 colpi decisi tra le scapole con il palmo della mano"],
      ["5 compressioni addominali", "Manovra di Heimlich: alle sue spalle, pugno sopra l'ombelico, spingi forte verso l'interno e verso l'alto"],
      ["Ripeti o RCP", "Alterna 5 e 5 finché il corpo estraneo esce. Se perde coscienza: a terra, 112 e RCP"],
    ];
    const sy = 5.1, sh = 1.65, sw = (W - 2 * M - 0.6) / 3;
    steps.forEach(([t, d], i) => {
      const x = M + i * (sw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: sy, w: sw, h: sh, rectRadius: 0.12, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: "Manovra " + t });
      s.addText([{ text: (i + 1) + ". " + t, options: { bold: true, color: "F2C9CF", fontSize: 16, breakLine: true } }, { text: d, options: { color: "FFFFFF", fontSize: 14 } }], { x: x + 0.25, y: sy + 0.12, w: sw - 0.5, h: sh - 0.24, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 3 });
    });
    s.addNotes("Le manovre per i lattanti (sotto 1 anno) sono diverse: colpi sulla schiena e compressioni sul torace, mai la manovra di Heimlich. Chi ha subito compressioni addominali va comunque controllato da un medico.");
  }

  // Ipertermia
  await (async () => {
    const s = await scheda(S5, "L'ipertermia (colpo di calore)", {
      def: "La temperatura del corpo sale troppo, fino a oltre 40 °C, perché il corpo non riesce più a disperdere il calore. Il colpo di sole è la forma causata dall'esposizione diretta al sole sulla testa.",
      cols: [
        { tint: T.amber, color: HEX.accent3, iconName: "sun", label: "Cause e segni", items: ["Sforzo fisico con caldo umido, disidratazione, abiti pesanti", "Pelle calda e arrossata, prima sudata poi secca", "Mal di testa, nausea, vertigini, crampi", "Confusione, svenimento nei casi gravi"] },
        FARE(["Porta la persona all'ombra o in un luogo fresco", "Toglile gli indumenti in eccesso", "Raffreddala con panni bagnati su testa, collo, ascelle e inguine, e fai aria", "Se è cosciente, falla bere acqua fresca a piccoli sorsi", "Se è confusa o perde coscienza: 112"]),
        { tint: T.teal, color: HEX.accent2, iconName: "shield", label: "Prevenzione", items: ["Bevi prima, durante e dopo l'attività, anche senza sete", "Evita le ore più calde", "Abiti leggeri, chiari e traspiranti; cappello al sole", "Abituati al caldo gradualmente"] },
      ],
    });
    s.addNotes("Il colpo di calore è un'emergenza vera: se la temperatura non viene abbassata rapidamente può danneggiare cervello e organi. Non dare da bere a chi non è completamente cosciente. Le bevande ghiacciate bevute tutte d'un fiato possono dare crampi allo stomaco: meglio acqua fresca a piccoli sorsi.");
  })();

  // Ipotermia
  await (async () => {
    const s = await scheda(S5, "L'ipotermia", {
      def: "La temperatura del corpo scende sotto i 35 °C perché il corpo perde più calore di quanto ne produca. Può capitare in montagna, in acqua fredda o con abiti bagnati e vento.",
      cols: [
        SEGNI(["Brividi intensi, che nei casi gravi smettono", "Pelle pallida e fredda, labbra bluastre", "Rigidità muscolare e movimenti impacciati", "Sonnolenza, confusione, difficoltà a parlare", "Respiro e polso lenti"]),
        FARE(["Porta la persona in un luogo riparato e caldo", "Togli gli indumenti bagnati e coprila con coperte o con la coperta termica", "Riscaldala gradualmente, partendo dal tronco", "Se è cosciente, bevande calde e zuccherate", "Muovila con delicatezza e chiama il 112 se è grave"]),
        NONFARE(["Non dare alcolici: disperdono calore", "Non frizionare né massaggiare braccia e gambe", "Non usare fonti di calore molto forti direttamente sulla pelle", "Non farla camminare o sforzare"]),
      ],
    });
    s.addNotes("Frizionare gli arti fa tornare verso il cuore sangue molto freddo e può causare problemi al ritmo cardiaco. La coperta termica si usa con il lato dorato verso l'esterno per trattenere il calore. Bambini e anziani perdono calore più facilmente.");
  })();

  // =========================================================
  // Conclusione
  // =========================================================
  pres.addSection({ title: "Conclusione" });
  {
    const s = pres.addSlide({ masterName: "Sezione", sectionTitle: "Conclusione" });
    s.addText("DA RICORDARE", { placeholder: "body" });
    s.addText("Le regole d'oro del soccorritore", { placeholder: "title" });
    const rules = [
      "Prima la tua sicurezza, poi quella della vittima",
      "Chiama il 112 e spiega dove sei e cosa è successo",
      "Fai solo ciò che sai fare: a volte non fare è meglio che fare male",
      "Non muovere chi ha un trauma, salvo pericolo immediato",
      "Non dare da bere a chi non è pienamente cosciente",
      "Resta con la vittima e rassicurala fino all'arrivo dei soccorsi",
    ];
    s.addText(rules.map((r, i) => ({ text: r, options: { bullet: { type: "number" }, breakLine: i < rules.length - 1, paraSpaceAfter: 8 } })), { x: M, y: 4.75, w: 8.6, h: 2.2, fontSize: 16, color: "FFFFFF", valign: "top", margin: 0, isTextBox: true, objectName: "Regole d'oro" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.6, y: 2.0, w: 3.13, h: 3.9, rectRadius: 0.15, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" }, objectName: "Numero emergenza" });
    await iconCircle(s, "phone", 10.56, 2.35, 1.2, HEX.accent1, "FFFFFF", "Icona telefono");
    s.addText("112", { x: 9.6, y: 3.6, w: 3.13, h: 1.2, fontSize: 66, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: THEME.headFontFace });
    s.addText("Numero Unico di Emergenza\ngratuito, anche senza credito", { x: 9.8, y: 4.85, w: 2.73, h: 0.9, fontSize: 15, color: C.text1, align: "center", valign: "top", margin: 0, isTextBox: true });
    s.addNotes("Riassunto finale. Il primo soccorso non richiede di essere medici: richiede calma, conoscenza di poche regole chiare e la prontezza di chiamare aiuto. Le manovre di RCP e di disostruzione vanno imparate con un corso pratico, ad esempio BLSD.");
  }

  await pres.writeFile({ fileName: OUT });
  // pptxgenjs writes an extra <a:pPr> before every run of a multi-run paragraph: drop those after the first run
  {
    const JSZip = require("jszip");
    const fs = require("fs");
    const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
    let fixed = 0;
    for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
      const xml = await zip.file(name).async("string");
      const out = xml.replace(/(<\/a:r>)<a:pPr\b[^>]*>.*?<\/a:pPr>/g, (m, r) => { fixed++; return r; }).replace(/(<\/a:r>)<a:pPr\b[^>]*\/>/g, (m, r) => { fixed++; return r; });
      zip.file(name, out);
    }
    fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("pPr fixed:", fixed);
  }
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
