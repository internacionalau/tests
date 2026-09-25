// Builds the student-facing + teacher lesson deck:
// "O que você fez ontem?" — everyday verbs in the pretérito perfeito.
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const OUT = process.argv[2] || "aula.pptx";

// ---------- design tokens ----------
const W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
const C = {
  dark: "0F3D2E", green: "1B7A55", mint: "E3F1EA", mintBg: "F1F8F4",
  yellow: "F5B700", ySoft: "FFF4D1", coral: "C8453C", cSoft: "FBE7E4",
  text: "1D2B26", muted: "5F6F69", panel: "F3F6F4", white: "FFFFFF", line: "D5E0DA",
};
const T = { bg: "262B35", card: "353B48", card2: "404757", accent: "F28C28", text: "F2F4F7", muted: "B4BCC8" };
const HEAD = "Arial", BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "O que você fez ontem? — Verbos do dia a dia no passado";

// ---------- icons ----------
const ICONS = {};
async function icon(name, color) {
  const key = name + color;
  if (!ICONS[key]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: 256 }));
    const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    ICONS[key] = "image/png;base64," + buf.toString("base64");
  }
  return ICONS[key];
}
const ICON_LIST = [
  ["FaComments", "FFFFFF"], ["FaComments", C.dark], ["FaBullseye", "FFFFFF"], ["FaListCheck", "FFFFFF"],
  ["FaLanguage", "FFFFFF"], ["FaCircleQuestion", "FFFFFF"], ["FaMicrophone", "FFFFFF"], ["FaMugHot", C.dark],
  ["FaRotate", C.dark], ["FaStar", C.dark], ["FaHouse", C.dark], ["FaUtensils", C.dark],
  ["FaTriangleExclamation", C.dark], ["FaLanguage", C.dark], ["FaPuzzlePiece", C.dark], ["FaMagnifyingGlass", C.dark],
  ["FaListCheck", C.dark], ["FaPenToSquare", C.dark], ["FaArrowRightArrowLeft", C.dark], ["FaBug", C.dark],
  ["FaLightbulb", C.dark], ["FaCircleQuestion", C.dark], ["FaMicrophone", C.dark], ["FaSun", C.dark],
  ["FaCloudSun", C.dark], ["FaMoon", C.dark], ["FaCartShopping", "FFFFFF"], ["FaPlaneArrival", "FFFFFF"],
  ["FaFaceDizzy", "FFFFFF"], ["FaBolt", C.dark], ["FaFlagCheckered", C.dark], ["FaHouseLaptop", C.dark],
  ["FaCheck", "FFFFFF"], ["FaArrowRight", C.dark], ["FaChalkboardUser", T.accent], ["FaWhatsapp", "FFFFFF"],
  ["FaBook", "FFFFFF"], ["FaLayerGroup", "FFFFFF"], ["FaEarthAmericas", "FFFFFF"], ["FaPhone", C.dark],
  ["FaClock", C.green], ["FaUserPen", C.dark],
];

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.10, blur: 6, offset: 2, angle: 90 });
function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, color: C.text, margin: 0, valign: "top", fontSize: 16 }, o));
}
function card(s, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: o.r ?? 0.12,
    fill: { color: o.fill || C.white, transparency: o.transparency || 0 },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
    shadow: o.shadow === false ? undefined : shadow(),
  });
}
function chipW(t, size = 12) { return t.length * size * 0.0085 + 0.4; }
function chip(s, x, y, t, o = {}) {
  const size = o.size || 12, w = o.w || chipW(t, size), h = o.h || size * 0.03 + 0.02;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: o.fill || C.mint }, line: o.line ? { color: o.line, width: 1 } : { type: "none" } });
  txt(s, t, { x, y, w, h, fontSize: size, bold: o.bold ?? true, color: o.color || C.green, align: "center", valign: "middle", charSpacing: o.cs ?? 1, fontFace: o.font || BODY });
  return w;
}
function badge(s, x, y, t, o = {}) {
  const d = o.d || 0.42;
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: o.fill || C.yellow }, line: { type: "none" } });
  txt(s, String(t), { x, y, w: d, h: d, fontSize: o.size || 15, bold: true, color: o.color || C.dark, align: "center", valign: "middle", fontFace: HEAD });
}
async function iconCircle(s, x, y, d, name, o = {}) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: o.fill || C.yellow }, line: { type: "none" } });
  const p = d * 0.26;
  s.addImage({ data: await icon(name, o.color || C.dark), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
// segments → text runs. string | {a} blank | {x,a} error fix | {b} bold | {g} green bold
function runs(parts, reveal, base = {}) {
  const out = [];
  for (const p of parts) {
    if (typeof p === "string") out.push({ text: p, options: { ...base } });
    else if (p.x !== undefined) {
      if (reveal) {
        out.push({ text: p.x, options: { ...base, strike: "sngStrike", color: C.coral } });
        out.push({ text: " " + p.a, options: { ...base, bold: true, color: C.green } });
      } else out.push({ text: p.x, options: { ...base } });
    } else if (p.a !== undefined) {
      out.push(reveal
        ? { text: p.a, options: { ...base, bold: true, color: C.green, underline: { style: "sng", color: C.green } } }
        : { text: p.blank || "_______", options: { ...base, color: C.muted } });
    } else if (p.b !== undefined) out.push({ text: p.b, options: { ...base, bold: true } });
    else if (p.g !== undefined) out.push({ text: p.g, options: { ...base, bold: true, color: C.green } });
  }
  return out;
}

let slideNo = 0;
const N = {}; // named slide numbers (for teacher map)
async function studentSlide({ id, stage, title, sub, ic, reveal, notes }) {
  const s = pres.addSlide();
  slideNo++;
  if (id) N[id] = slideNo;
  s.background = { color: reveal ? C.mintBg : C.white };
  chip(s, M, 0.45, stage);
  if (reveal) {
    const t = "✓  RESPOSTAS · ANSWERS";
    chip(s, W - M - 0.8 - chipW(t) - 0.25, 0.53, t, { fill: C.green, color: C.white });
  }
  if (ic) await iconCircle(s, W - M - 0.8, 0.4, 0.8, ic);
  txt(s, title, { x: M, y: 0.92, w: CW - 1.2, h: 0.7, fontSize: 32, bold: true, fontFace: HEAD, color: C.dark, valign: "middle" });
  if (sub) txt(s, sub, { x: M, y: 1.6, w: CW - 1.2, h: 0.4, fontSize: 16, italic: true, color: C.muted });
  txt(s, String(slideNo), { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontSize: 10, color: C.muted, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}
// For exercise + reveal pairs: build twice with the same drawing function.
async function pair(meta, draw, notesQ, notesA) {
  const s1 = await studentSlide({ ...meta, notes: notesQ });
  await draw(s1, false);
  const s2 = await studentSlide({ ...meta, id: meta.id ? meta.id + "_ans" : undefined, reveal: true, notes: notesA || "Revele as respostas. Peça ao aluno para ler as frases corrigidas em voz alta." });
  await draw(s2, true);
}

// =====================================================================
async function build() {
  for (const [n, c] of ICON_LIST) await icon(n, c);

  // ---------- 1. TITLE ----------
  {
    const s = pres.addSlide(); slideNo++;
    s.background = { color: C.dark };
    s.addShape(pres.shapes.OVAL, { x: 8.7, y: -1.6, w: 6.4, h: 6.4, fill: { color: C.green }, line: { type: "none" } });
    s.addShape(pres.shapes.OVAL, { x: 9.9, y: 0.9, w: 2.6, h: 2.6, fill: { color: C.yellow }, line: { type: "none" } });
    s.addImage({ data: await icon("FaComments", C.dark), x: 10.55, y: 1.55, w: 1.3, h: 1.3 });
    chip(s, M, 0.9, "PORTUGUÊS BRASILEIRO · AULA DE 50 MIN", { fill: C.green, color: C.white });
    txt(s, "O que você\nfez ontem?", { x: M, y: 1.6, w: 8.5, h: 2.4, fontSize: 64, bold: true, fontFace: HEAD, color: C.white, valign: "middle" });
    txt(s, "Verbos do dia a dia no passado", { x: M, y: 4.1, w: 8.5, h: 0.55, fontSize: 28, bold: true, color: C.yellow });
    txt(s, "Everyday verbs in the past · pretérito perfeito", { x: M, y: 4.65, w: 8.5, h: 0.45, fontSize: 18, italic: true, color: "C9DDD3" });
    let x = M;
    for (const v of ["fui", "fiz", "tive", "vi", "vim", "disse", "saí", "comi"]) {
      x += chip(s, x, 6.05, v, { fill: C.dark, line: "5E8C79", color: C.white, size: 16, cs: 0, bold: false, h: 0.5, w: chipW(v, 16) + 0.25 }) + 0.15;
    }
    s.addNotes("Abertura (30 s). Apresente o tema: hoje o foco é USAR os verbos no passado, não estudar regras. Diga a meta: 15 verbos diferentes falando.");
  }

  // ---------- 2. OBJECTIVES ----------
  {
    const s = await studentSlide({ stage: "OBJETIVOS · LESSON GOALS", title: "Hoje você vai…", sub: "Today you will…", ic: "FaBullseye", notes: "Leia rapidamente os objetivos (1 min). Mostre o plano da aula embaixo para o aluno saber o ritmo." });
    const rows = [
      ["FaLayerGroup", "Reconhecer 27 verbos do dia a dia", "Recognize 27 everyday verbs"],
      ["FaRotate", "Usar os verbos no passado", "Use them in the past: eu fui, eu fiz, eu comi…"],
      ["FaCircleQuestion", "Fazer e responder perguntas", "Ask and answer questions about the past"],
      ["FaMicrophone", "Contar o que você fez", "Talk about your day and your weekend"],
    ];
    let y = 2.2;
    for (const [ic, pt, en] of rows) {
      await iconCircle(s, M, y, 0.72, ic, { fill: C.green, color: "FFFFFF" });
      txt(s, pt, { x: M + 0.95, y: y + 0.02, w: 6.6, h: 0.4, fontSize: 21, bold: true });
      txt(s, en, { x: M + 0.95, y: y + 0.42, w: 6.6, h: 0.3, fontSize: 15, italic: true, color: C.muted });
      y += 0.95;
    }
    card(s, 8.75, 2.2, W - M - 8.75, 3.55, { fill: C.dark });
    chip(s, 9.1, 2.5, "META DA AULA", { fill: C.green, color: C.white });
    txt(s, "15+", { x: 9.1, y: 2.95, w: 3.3, h: 1.2, fontSize: 80, bold: true, fontFace: HEAD, color: C.yellow });
    txt(s, "verbos diferentes no passado — falando, não só lendo!", { x: 9.1, y: 4.2, w: 3.3, h: 1.2, fontSize: 17, color: C.white });
    const plan = [["Aquecimento", "5'"], ["Verbos", "10'"], ["Prática", "10'"], ["Perguntas", "10'"], ["Desafio oral", "15'"]];
    const pw = (CW - 4 * 0.2) / 5;
    plan.forEach(([n, t], i) => {
      const px = M + i * (pw + 0.2);
      card(s, px, 6.15, pw, 0.62, { fill: i === 4 ? C.ySoft : C.panel, shadow: false });
      badge(s, px + 0.12, 6.25, i + 1, { d: 0.42, size: 14 });
      txt(s, [{ text: n + "  ", options: { bold: true } }, { text: t, options: { color: C.muted } }], { x: px + 0.65, y: 6.15, w: pw - 0.7, h: 0.62, fontSize: 16, valign: "middle" });
    });
  }

  // ---------- 3. WARM-UP ----------
  {
    const s = await studentSlide({ id: "warm", stage: "1 · AQUECIMENTO · 5 MIN", title: "E ontem?", sub: "Warm-up · Answer quickly — short answers are OK!", ic: "FaMugHot",
      notes: "5 min. Comece pelo Nível 1 e só suba quando as respostas saírem fáceis. Responda você também, rapidamente, para dar modelo (\"Eu acordei às 6h, fiz café…\"). Observe: ele responde com o verbo? Usa a forma do EU ou repete a forma da pergunta (trabalhou)?" });
    const cols = [
      ["NÍVEL 1", "Sim / não", ["Você trabalhou ontem?", "Você saiu de casa ontem?", "Você dormiu bem?"]],
      ["NÍVEL 2", "Informação", ["Que horas você acordou hoje?", "O que você comeu no almoço ontem?", "O que você bebeu hoje de manhã?"]],
      ["NÍVEL 3", "Resposta aberta", ["O que você fez ontem à noite?", "Qual foi a melhor parte do seu dia ontem?"]],
    ];
    const cw = (CW - 2 * 0.3) / 3;
    cols.forEach(([lv, lab, qs], i) => {
      const x = M + i * (cw + 0.3);
      card(s, x, 2.2, cw, 3.55, { fill: i === 2 ? C.ySoft : C.panel, shadow: false });
      chip(s, x + 0.3, 2.45, lv, { fill: i === 2 ? C.yellow : C.green, color: i === 2 ? C.dark : C.white });
      txt(s, lab, { x: x + 0.3 + chipW(lv) + 0.15, y: 2.45, w: 2, h: 0.38, fontSize: 14, italic: true, color: C.muted, valign: "middle" });
      txt(s, qs.map((q, k) => ({ text: q, options: { breakLine: k < qs.length - 1, paraSpaceAfter: 14 } })), { x: x + 0.3, y: 3.1, w: cw - 0.6, h: 2.5, fontSize: 20, bold: true, color: C.dark });
    });
    card(s, M, 6.05, CW, 0.8, { fill: C.mint, shadow: false });
    await iconCircle(s, M + 0.2, 6.15, 0.6, "FaLightbulb");
    txt(s, [
      { text: "Responda com o verbo:   ", options: { bold: true, color: C.green } },
      { text: "Você trabalhou ontem?  →  " },
      { text: "Sim, trabalhei.", options: { bold: true } },
      { text: "  /  " },
      { text: "Não, não trabalhei.", options: { bold: true } },
    ], { x: M + 1.0, y: 6.05, w: CW - 1.2, h: 0.8, fontSize: 19, valign: "middle" });
  }

  // ---------- 4. PRESENT -> PAST ----------
  {
    const s = await studentSlide({ stage: "REVISÃO RÁPIDA · QUICK REVIEW", title: "Todo dia × ontem", sub: "Present = habit · Past = finished action", ic: "FaRotate",
      notes: "1 min. Não explique regra longa. Mostre o contraste e as pistas de tempo — o aluno vai usar essas palavras para decidir presente × passado o resto da aula." });
    const pw = 5.55;
    card(s, M, 2.2, pw, 2.75, { fill: C.panel, shadow: false });
    chip(s, M + 0.35, 2.45, "PRESENTE · HABIT", { fill: C.white, color: C.muted });
    txt(s, [{ text: "Todo dia eu " }, { text: "faço", options: { bold: true, color: C.dark } }, { text: " café." }], { x: M + 0.35, y: 3.0, w: pw - 0.7, h: 0.8, fontSize: 32, color: C.text });
    txt(s, "todo dia · sempre · normalmente · às vezes · aos sábados", { x: M + 0.35, y: 4.0, w: pw - 0.7, h: 0.7, fontSize: 16, italic: true, color: C.muted });
    const rx = W - M - pw;
    card(s, rx, 2.2, pw, 2.75, { fill: C.dark });
    chip(s, rx + 0.35, 2.45, "PASSADO · FINISHED", { fill: C.yellow, color: C.dark });
    txt(s, [{ text: "Ontem eu " }, { text: "fiz", options: { bold: true, color: C.yellow } }, { text: " café." }], { x: rx + 0.35, y: 3.0, w: pw - 0.7, h: 0.8, fontSize: 32, color: C.white });
    txt(s, "ontem · anteontem · no sábado · ontem à noite · semana passada · hoje de manhã", { x: rx + 0.35, y: 4.0, w: pw - 0.7, h: 0.7, fontSize: 16, italic: true, color: "C9DDD3" });
    await iconCircle(s, W / 2 - 0.4, 3.18, 0.8, "FaArrowRight");
    const ends = [["-AR", "-ei", "trabalhar → trabalhei"], ["-ER", "-i", "comer → comi"], ["-IR", "-i", "dormir → dormi"]];
    const ew = (CW - 2 * 0.3) / 3;
    ends.forEach(([a, b, ex], i) => {
      const x = M + i * (ew + 0.3);
      card(s, x, 5.3, ew, 1.45, { fill: C.white, line: C.line, shadow: false });
      txt(s, [{ text: a + "  →  eu ", options: { color: C.muted } }, { text: b, options: { bold: true, color: C.green } }], { x: x + 0.3, y: 5.45, w: ew - 0.6, h: 0.5, fontSize: 24, fontFace: HEAD });
      txt(s, ex, { x: x + 0.3, y: 6.0, w: ew - 0.6, h: 0.5, fontSize: 18 });
    });
  }

  // ---------- 5. EU x VOCÊ ----------
  {
    const s = await studentSlide({ id: "euvoce", stage: "REVISÃO RÁPIDA · QUICK REVIEW", title: "Pergunta com você, resposta com eu", sub: "The question uses você… — your answer uses eu!", ic: "FaComments",
      notes: "1 min. Este é o erro nº 1 da aula: responder \"Sim, trabalhou\". Leia os pares com o aluno (você pergunta, ele responde). Combine um gesto (polegar para trás) para 'passado' e outro para 'use eu'." });
    const pairs = [["Você trabalhou ontem?", "Sim, trabalhei."], ["Você comeu em casa?", "Não, comi num restaurante."], ["Você saiu no sábado?", "Saí, sim!"], ["Você fez o jantar?", "Fiz. Fiz um macarrão."]];
    let y = 2.2;
    for (const [q, a] of pairs) {
      card(s, M, y, 3.7, 0.78, { fill: C.panel, shadow: false, r: 0.35 });
      txt(s, q, { x: M + 0.25, y, w: 3.4, h: 0.78, fontSize: 18, valign: "middle" });
      s.addImage({ data: await icon("FaArrowRight", C.dark), x: M + 3.85, y: y + 0.24, w: 0.3, h: 0.3 });
      card(s, M + 4.3, y, 3.55, 0.78, { fill: C.green, shadow: false, r: 0.35 });
      txt(s, a, { x: M + 4.55, y, w: 3.2, h: 0.78, fontSize: 18, bold: true, color: C.white, valign: "middle" });
      y += 0.98;
    }
    const x = 9.0, w = W - M - x;
    card(s, x, 2.2, w, 3.72, { fill: C.ySoft, shadow: false });
    txt(s, "Terminações", { x: x + 0.3, y: 2.4, w: w - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.dark });
    txt(s, [
      { text: "você", options: { bold: true } }, { text: "  trabalhou · comeu · saiu", options: { breakLine: true } },
      { text: "eu", options: { bold: true, color: C.green } }, { text: "  trabalhei · comi · saí", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 8 } },
      { text: "Irregulares:", options: { bold: true, breakLine: true } },
      { text: "você foi · fez · teve · viu · veio", options: { breakLine: true } },
      { text: "eu fui · fiz · tive · vi · vim", options: { color: C.green, bold: true } },
    ], { x: x + 0.3, y: 2.9, w: w - 0.6, h: 2.9, fontSize: 16, paraSpaceAfter: 4 });
    txt(s, "✗  Você trabalhou? — Sim, trabalhou.", { x: M, y: 6.2, w: CW, h: 0.45, fontSize: 18, color: C.coral, bold: true });
  }

  // ---------- 6-7. SINGLE MC ----------
  await pair({ id: "mc1", stage: "QUIZ RÁPIDO", title: "Qual é a forma certa?", sub: "Which one is correct?", ic: "FaCircleQuestion" }, async (s, r) => {
    card(s, M + 1.2, 2.25, CW - 2.4, 1.7, { fill: C.panel, shadow: false });
    txt(s, runs(["Ontem eu ", { a: "fiz", blank: "________" }, " café."], r), { x: M + 1.2, y: 2.35, w: CW - 2.4, h: 1.05, fontSize: 48, bold: true, fontFace: HEAD, align: "center", valign: "middle", color: C.dark });
    txt(s, "Yesterday I ___ coffee.", { x: M + 1.2, y: 3.35, w: CW - 2.4, h: 0.45, fontSize: 18, italic: true, color: C.muted, align: "center" });
    const opts = [["A", "faço"], ["B", "fiz"], ["C", "fazer"]];
    const ow = 3.3, gap = 0.45, x0 = (W - (3 * ow + 2 * gap)) / 2;
    for (let i = 0; i < 3; i++) {
      const [L, t] = opts[i], x = x0 + i * (ow + gap), ok = r && L === "B";
      card(s, x, 4.35, ow, 1.35, { fill: ok ? C.green : C.white, line: ok ? undefined : C.line, transparency: r && !ok ? 0 : 0 });
      badge(s, x + 0.3, 4.8, L, { fill: ok ? C.white : C.yellow, color: C.dark });
      txt(s, t, { x: x + 0.95, y: 4.35, w: ow - 1.1, h: 1.35, fontSize: 34, bold: true, fontFace: HEAD, valign: "middle", color: ok ? C.white : (r ? "A9B5B0" : C.dark) });
      if (ok) s.addImage({ data: await icon("FaCheck", "FFFFFF"), x: x + ow - 0.65, y: 4.8, w: 0.4, h: 0.4 });
    }
    if (r) txt(s, [{ text: "fazer → eu faço (presente) → " }, { text: "eu fiz", options: { bold: true, color: C.green } }, { text: " (passado)   ·   ontem = passado" }], { x: M, y: 6.1, w: CW, h: 0.5, fontSize: 20, align: "center" });
  }, "Deixe o aluno escolher e justificar (\"por causa de 'ontem'\"). Não revele antes de ele decidir.",
     "Resposta: B) fiz. Reforce: 'ontem' = pista de passado. fazer é irregular — memorizar: faço → fiz.");

  // ---------- 8-9. IRREGULAR VERB BANK ----------
  const irregular = [
    [["fazer", "to do / to make", "faço", "fiz", "Fiz café de manhã."], ["ir", "to go", "vou", "fui", "Fui à padaria."], ["ter", "to have", "tenho", "tive", "Tive uma reunião."],
     ["ver", "to see / to watch", "vejo", "vi", "Vi um filme ótimo."], ["vir", "to come (here)", "venho", "vim", "Vim pro Brasil em março."], ["ser", "to be (description)", "sou", "fui", "A festa foi ótima!"]],
    [["estar", "to be (place / temporary)", "estou", "estive", "Estive em Salvador."], ["dar", "to give", "dou", "dei", "Dei um presente pra ela."], ["dizer", "to say / to tell", "digo", "disse", "Disse \"obrigado\"."],
     ["poder", "can / to be able to", "posso", "pude", "Não pude ir à festa."], ["querer", "to want", "quero", "quis", "Não quis sair ontem."], null],
  ];
  for (let k = 0; k < 2; k++) {
    const s = await studentSlide({ id: "bank" + (k + 1), stage: "2 · VERBOS · IRREGULARES ⭐", title: `Os irregulares superstars (${k + 1}/2)`, sub: "The most frequent past forms — learn them by using them", ic: "FaStar",
      notes: k === 0
        ? "Verb bank (irregulares 1/2). Leia só a forma do passado em voz alta, aluno repete 2x. Depois: você diz o infinitivo, ele diz o passado (sem olhar). Não demore: ~1,5 min."
        : "Verb bank (irregulares 2/2). Mesmo procedimento. Destaque o card amarelo: fui = ir e ser. Mencione que no dia a dia 'não consegui' e 'eu queria' também são muito usados, sem entrar no imperfeito." });
    const cw = (CW - 2 * 0.3) / 3, ch = 2.1;
    for (let i = 0; i < 6; i++) {
      const x = M + (i % 3) * (cw + 0.3), y = 2.2 + Math.floor(i / 3) * (ch + 0.3);
      const v = irregular[k][i];
      if (!v) {
        card(s, x, y, cw, ch, { fill: C.ySoft, shadow: false });
        txt(s, [{ text: "fui", options: { bold: true, color: C.green } }, { text: " = ir + ser" }], { x: x + 0.3, y: y + 0.2, w: cw - 0.6, h: 0.6, fontSize: 28, bold: true, fontFace: HEAD, color: C.dark });
        txt(s, [{ text: "Fui ao mercado.", options: { bold: true } }, { text: "  (ir)", options: { color: C.muted, breakLine: true } }, { text: "A festa foi ótima!", options: { bold: true } }, { text: "  (ser)", options: { color: C.muted } }], { x: x + 0.3, y: y + 0.95, w: cw - 0.6, h: 1.0, fontSize: 17, paraSpaceAfter: 6 });
        continue;
      }
      const [inf, en, pr, pa, ex] = v;
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      txt(s, inf, { x: x + 0.3, y: y + 0.18, w: 1.6, h: 0.45, fontSize: 24, bold: true, fontFace: HEAD, color: C.dark });
      txt(s, en, { x: x + 1.6, y: y + 0.22, w: cw - 1.9, h: 0.4, fontSize: 13, italic: true, color: C.muted, align: "right", valign: "middle" });
      txt(s, [{ text: "eu " + pr + "  →  ", options: { color: C.muted, fontSize: 18 } }, { text: "eu " + pa, options: { bold: true, color: C.green, fontSize: 28, fontFace: HEAD } }], { x: x + 0.3, y: y + 0.72, w: cw - 0.6, h: 0.6, valign: "middle" });
      txt(s, ex, { x: x + 0.3, y: y + 1.45, w: cw - 0.6, h: 0.45, fontSize: 16 });
    }
  }

  // ---------- 10-11. REGULAR TABLES ----------
  const tables = [
    ["Rotina", "Daily routine", "FaHouse", [["acordar", "to wake up", "acordo", "acordei", "Acordei às 7h."], ["dormir", "to sleep", "durmo", "dormi", "Dormi mal essa noite."], ["trabalhar", "to work", "trabalho", "trabalhei", "Trabalhei de casa."], ["estudar", "to study", "estudo", "estudei", "Estudei português à noite."], ["sair", "to leave / to go out", "saio", "saí", "Saí de casa às 8h."], ["chegar", "to arrive", "chego", "cheguei", "Cheguei em casa tarde."], ["voltar", "to come back", "volto", "voltei", "Voltei de Uber."], ["ficar", "to stay / to get (tired…)", "fico", "fiquei", "Fiquei em casa."]]],
    ["Vida social e consumo", "Social life & shopping", "FaUtensils", [["comer", "to eat", "como", "comi", "Comi pão de queijo."], ["beber", "to drink", "bebo", "bebi", "Bebi uma cerveja."], ["comprar", "to buy", "compro", "comprei", "Comprei um chip de celular."], ["falar", "to speak / to talk", "falo", "falei", "Falei com minha mãe."], ["encontrar", "to meet up / to find", "encontro", "encontrei", "Encontrei uns amigos."], ["conhecer", "to meet (1st time)", "conheço", "conheci", "Conheci uma pessoa legal."], ["assistir", "to watch (TV, games)", "assisto", "assisti", "Assisti o jogo do Brasil."], ["gostar (de)", "to like", "gosto", "gostei", "Gostei do restaurante."]]],
  ];
  for (const [ti, en, ic, rows] of tables) {
    const s = await studentSlide({ id: "tab_" + ic, stage: "2 · VERBOS · REGULARES", title: ti, sub: en + " · regular verbs — same pattern, easy to use", ic,
      notes: "~1 min. Não leia a tabela inteira. Peça ao aluno para ler só a coluna 'Eu (passado)' em voz alta e escolher 2 exemplos que são verdade para ele ('Eu também dormi mal!')." });
    const hdr = ["Infinitivo", "English", "Eu (presente)", "Eu (passado)", "Exemplo"].map(t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, fontSize: 15 } }));
    const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: {
      fill: { color: i % 2 ? C.panel : C.white }, fontSize: j === 3 ? 19 : 16, bold: j === 0 || j === 3,
      color: j === 3 ? C.green : (j === 1 || j === 2 ? C.muted : C.text), italic: j === 1 } })));
    s.addTable([hdr, ...body], { x: M, y: 2.2, w: CW, colW: [1.9, 2.75, 1.85, 1.95, 3.68], rowH: 0.48, fontFace: BODY, valign: "middle", margin: [0, 0.12, 0, 0.12], border: { type: "none" } });
  }

  // ---------- 12. CUIDADO ----------
  {
    const s = await studentSlide({ id: "cuidado", stage: "2 · VERBOS · ATENÇÃO", title: "Cuidado! Mesma tradução, sentido diferente", sub: "Watch out: these verbs are easy to mix up", ic: "FaTriangleExclamation",
      notes: "1 min. Só os pontos que forem relevantes para o aluno. Pergunte: 'Quando você veio pro Brasil?' / 'Você conheceu alguém aqui?' para testar vim×fui e conheci×encontrei." });
    const cards = [
      ["vim  ×  fui", [["Vim pro Brasil em março.", "I came (here, where I am now)"], ["Fui pra Argentina.", "I went (there)"]]],
      ["conheci  ×  encontrei", [["Conheci a Ana na festa.", "I met Ana (for the 1st time)"], ["Encontrei a Ana no shopping.", "I ran into / met up with Ana"]]],
      ["ficar", [["Fiquei em casa.", "I stayed home"], ["Fiquei cansado.", "I got tired"]]],
      ["sair", [["Saí do trabalho às 6h.", "I left work at 6"], ["Saí com uns amigos.", "I went out with some friends"]]],
    ];
    const cw = (CW - 0.3) / 2, ch = 1.85;
    cards.forEach(([t, ex], i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 2.2 + Math.floor(i / 2) * (ch + 0.25);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      txt(s, t, { x: x + 0.3, y: y + 0.15, w: cw - 0.6, h: 0.45, fontSize: 22, bold: true, fontFace: HEAD, color: C.green });
      txt(s, ex.flatMap(([pt, e], k) => [{ text: pt, options: { bold: true } }, { text: "   " + e, options: { italic: true, color: C.muted, fontSize: 15, breakLine: k === 0 } }]), { x: x + 0.3, y: y + 0.72, w: cw - 0.6, h: 1.0, fontSize: 18, paraSpaceAfter: 8 });
    });
    card(s, M, 6.35, CW, 0.55, { fill: C.ySoft, shadow: false });
    txt(s, [{ text: "Ortografia:  ", options: { bold: true } }, { text: "chegar → che" }, { text: "gu", options: { bold: true, color: C.coral } }, { text: "ei   ·   ficar → fi" }, { text: "qu", options: { bold: true, color: C.coral } }, { text: "ei   ·   sair → sa" }, { text: "í", options: { bold: true, color: C.coral } }, { text: " (with accent!)" }],
      { x: M + 0.3, y: 6.35, w: CW - 0.6, h: 0.55, fontSize: 18, valign: "middle" });
  }

  // ---------- 13-14. EN -> PT WORDS ----------
  const enpt = [["to go", "ir", "fui"], ["to have", "ter", "tive"], ["to see", "ver", "vi"], ["to come", "vir", "vim"], ["to say", "dizer", "disse"], ["to stay", "ficar", "fiquei"], ["to arrive", "chegar", "cheguei"], ["to meet (1st time)", "conhecer", "conheci"]];
  await pair({ id: "enpt", stage: "3 · VOCABULÁRIO · EN → PT", title: "Como se diz…?", sub: "Say the verb in Portuguese — then say it in the past", ic: "FaLanguage" }, async (s, r) => {
    const cw = (CW - 3 * 0.3) / 4, ch = 1.9;
    enpt.forEach(([en, inf, pa], i) => {
      const x = M + (i % 4) * (cw + 0.3), y = 2.25 + Math.floor(i / 4) * (ch + 0.35);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      badge(s, x + 0.2, y + 0.2, i + 1, { d: 0.38, size: 13 });
      txt(s, en, { x: x + 0.7, y: y + 0.18, w: cw - 0.85, h: 0.45, fontSize: 20, bold: true, color: C.dark, valign: "middle" });
      txt(s, runs([{ a: inf, blank: "________" }], r), { x: x + 0.25, y: y + 0.8, w: cw - 0.5, h: 0.45, fontSize: 20 });
      txt(s, runs(["eu ", { a: pa, blank: "________" }], r), { x: x + 0.25, y: y + 1.28, w: cw - 0.5, h: 0.45, fontSize: 20 });
    });
  }, "Ritmo rápido (quick fire). O aluno fala o infinitivo e o passado. Se travar, dê a primeira sílaba.");

  // ---------- 15-16. MATCHING ----------
  const matchInf = ["fazer", "sair", "dar", "querer", "dormir", "estar", "comer", "voltar"];
  const matchPast = [["A", "dormi"], ["B", "quis"], ["C", "fiz"], ["D", "voltei"], ["E", "estive"], ["F", "saí"], ["G", "comi"], ["H", "dei"]];
  const matchAns = ["C", "F", "H", "B", "A", "E", "G", "D"];
  await pair({ id: "match", stage: "3 · VOCABULÁRIO · MATCHING", title: "Qual é o passado?", sub: "Match each verb with its past form (eu)", ic: "FaPuzzlePiece" }, async (s, r) => {
    matchInf.forEach((v, i) => {
      const col = Math.floor(i / 4), x = M + col * 3.55, y = 2.25 + (i % 4) * 1.08;
      card(s, x, y, 3.25, 0.85, { fill: C.panel, shadow: false });
      badge(s, x + 0.2, y + 0.21, i + 1);
      txt(s, v, { x: x + 0.8, y, w: 1.4, h: 0.85, fontSize: 22, bold: true, valign: "middle", color: C.dark });
      card(s, x + 2.2, y + 0.15, 0.85, 0.55, { fill: r ? C.green : C.white, line: r ? undefined : C.line, shadow: false });
      if (r) txt(s, matchAns[i], { x: x + 2.2, y: y + 0.15, w: 0.85, h: 0.55, fontSize: 20, bold: true, color: C.white, align: "center", valign: "middle" });
    });
    const bx = 8.2;
    card(s, bx, 2.25, W - M - bx, 4.1, { fill: C.ySoft, shadow: false });
    txt(s, "Passado", { x: bx + 0.3, y: 2.4, w: 3, h: 0.4, fontSize: 16, bold: true, color: C.muted });
    matchPast.forEach(([L, p], i) => {
      const x = bx + 0.3 + (i % 2) * 2.2, y = 2.95 + Math.floor(i / 2) * 0.82;
      txt(s, [{ text: L + "   ", options: { bold: true, color: C.muted } }, { text: p, options: { bold: true, color: C.dark } }], { x, y, w: 2.1, h: 0.6, fontSize: 24, valign: "middle" });
    });
  }, "O aluno diz o par em voz alta: '1 é C — fazer, fiz'. Peça uma frase rápida com 2 ou 3 deles.");

  // ---------- 17-18. PT -> EN ----------
  const pten = [["Conheci a Ana na festa.", "I met Ana (for the first time) at the party."], ["Encontrei a Ana no shopping.", "I ran into / met up with Ana at the mall."], ["Fiquei cansado depois do trabalho.", "I got tired after work."], ["Saí do trabalho às 6h.", "I left work at 6."], ["Vim pro Brasil em 2024.", "I came to Brazil in 2024."], ["O show foi incrível!", "The show was amazing! (foi = ser)"]];
  await pair({ id: "pten", stage: "3 · VOCABULÁRIO · PT → EN", title: "O que significa?", sub: "What does it mean in English? Pay attention to the verb!", ic: "FaMagnifyingGlass" }, async (s, r) => {
    const cw = (CW - 0.3) / 2, ch = 1.2;
    pten.forEach(([pt, en], i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 2.25 + Math.floor(i / 2) * (ch + 0.25);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      badge(s, x + 0.25, y + 0.22, i + 1);
      txt(s, pt, { x: x + 0.85, y: y + 0.15, w: cw - 1.1, h: 0.5, fontSize: 21, bold: true, color: C.dark, valign: "middle" });
      txt(s, r ? en : "→  ____________________________", { x: x + 0.85, y: y + 0.66, w: cw - 1.1, h: 0.4, fontSize: 17, italic: r, bold: false, color: r ? C.green : C.muted });
    });
  }, "Foco no sentido: conheci × encontrei, fiquei (got), saí (left), vim (came), foi (was).");

  // ---------- 19-20. MC x4 ----------
  const mc = [
    [["Ontem eu ", { a: "fui" }, " ao shopping."], ["vou", "fui", "fiz"], 1],
    [["Você ", { a: "viu" }, " o jogo do Flamengo ontem?"], ["viu", "vi", "vê"], 0],
    [["Normalmente eu ", { a: "acordo" }, " às 7h, mas ontem eu ", { a: "acordei" }, " às 10h."], ["acordei / acordo", "acordo / acordei", "acordo / acordou"], 1],
    [["Eu não ", { a: "pude" }, " ir à festa porque ", { a: "tive" }, " que trabalhar."], ["pude / tive", "posso / tive", "pude / tenho"], 0],
  ];
  await pair({ id: "mc4", stage: "4 · PRÁTICA · MÚLTIPLA ESCOLHA", title: "Escolha a alternativa correta", sub: "Choose the correct option — look for the time clues!", ic: "FaListCheck" }, async (s, r) => {
    const ch = 1.02;
    mc.forEach(([sent, opts, ok], i) => {
      const x = M, y = 2.2 + i * (ch + 0.18);
      card(s, x, y, CW, ch, { fill: C.white, line: C.line });
      badge(s, x + 0.25, y + 0.3, i + 1);
      txt(s, runs(sent, r), { x: x + 0.85, y: y + 0.1, w: CW - 1.1, h: 0.42, fontSize: 21, color: C.dark, valign: "middle" });
      let ox = x + 0.85;
      opts.forEach((o, j) => {
        const L = "ABC"[j] + ")  " + o, good = r && j === ok, w = chipW(L, 16) + 0.1;
        chip(s, ox, y + 0.56, L, { size: 16, h: 0.38, w, fill: good ? C.green : C.panel, color: good ? C.white : (r ? "9AA6A1" : C.text), cs: 0, bold: good });
        ox += w + 0.25;
      });
    });
  }, "Peça para o aluno dizer qual palavra da frase mostra o tempo (ontem, normalmente…).");

  // ---------- 21-22. FILL IN ----------
  const fill = [
    ["ir · comprar", ["Ontem eu ", { a: "fui" }, " ao supermercado e ", { a: "comprei" }, " frutas."]],
    ["ter", ["Ontem eu ", { a: "tive" }, " muito trabalho."]],
    ["sair", ["Eu ", { a: "saí" }, " de casa às 8h."]],
    ["chegar", ["Eu ", { a: "cheguei" }, " atrasado na aula."]],
    ["dizer", ["Eu ", { a: "disse" }, " \"bom dia\" pro porteiro."]],
    ["ficar", ["No domingo eu ", { a: "fiquei" }, " em casa o dia todo."]],
  ];
  await pair({ id: "fill", stage: "4 · PRÁTICA · COMPLETE", title: "Complete no passado", sub: "Fill in the blanks — use the verb in the past (eu)", ic: "FaPenToSquare" }, async (s, r) => {
    fill.forEach(([v, sent], i) => {
      const y = 2.2 + i * 0.75;
      card(s, M, y, CW, 0.62, { fill: i % 2 ? C.white : C.panel, shadow: false });
      badge(s, M + 0.15, y + 0.1, i + 1);
      chip(s, M + 0.75, y + 0.12, v, { size: 15, h: 0.38, w: 1.9, fill: C.ySoft, color: C.dark, cs: 0 });
      txt(s, runs(sent, r), { x: M + 2.95, y, w: CW - 3.1, h: 0.62, fontSize: 21, valign: "middle" });
    });
  }, "Oral primeiro, depois o aluno pode escrever se quiser. Atenção à ortografia de cheguei e ao acento de saí.");

  // ---------- 23-24. TRANSFORM ----------
  await pair({ id: "transf", stage: "4 · PRÁTICA · PRESENTE → PASSADO", title: "Todo dia… mas ontem…", sub: "Retell the text in the past: start with \"Ontem…\"", ic: "FaArrowRightArrowLeft" }, async (s, r) => {
    const cw = (CW - 0.4) / 2;
    card(s, M, 2.2, cw, 4.55, { fill: C.panel, shadow: false });
    chip(s, M + 0.35, 2.45, "PRESENTE · TODO DIA", { fill: C.white, color: C.muted });
    txt(s, runs(["Todo dia eu ", { b: "acordo" }, " às 7h. Eu ", { b: "faço" }, " café e ", { b: "como" }, " um pão. Depois eu ", { b: "vou" }, " pro trabalho de ônibus. À noite, eu ", { b: "vejo" }, " uma série e ", { b: "durmo" }, " às 23h."], false),
      { x: M + 0.35, y: 3.05, w: cw - 0.7, h: 3.5, fontSize: 24, lineSpacingMultiple: 1.3 });
    const x = M + cw + 0.4;
    card(s, x, 2.2, cw, 4.55, { fill: C.dark });
    chip(s, x + 0.35, 2.45, "PASSADO · ONTEM", { fill: C.yellow, color: C.dark });
    const parts = ["Ontem eu ", { a: "acordei" }, " às 7h. Eu ", { a: "fiz" }, " café e ", { a: "comi" }, " um pão. Depois eu ", { a: "fui" }, " pro trabalho de ônibus. À noite, eu ", { a: "vi" }, " uma série e ", { a: "dormi" }, " às 23h."];
    const rr = runs(parts, r, { color: C.white }).map(o => (o.options.color === C.green ? { ...o, options: { ...o.options, color: C.yellow, underline: { style: "sng", color: C.yellow } } } : o.options.color === C.muted ? { ...o, options: { ...o.options, color: "8FB5A4" } } : o));
    txt(s, rr, { x: x + 0.35, y: 3.05, w: cw - 0.7, h: 3.5, fontSize: 24, lineSpacingMultiple: 1.3, color: C.white });
  }, "O aluno lê o texto do presente e fala a versão no passado sem escrever. Depois, peça para ele mudar detalhes para contar o ontem DELE.");

  // ---------- 25-26. EN -> PT SENTENCES ----------
  const tr = [["I went to the bakery and bought bread.", "Eu fui à padaria e comprei pão."], ["I had a meeting in the morning.", "Eu tive uma reunião de manhã."], ["I met a Brazilian guy at the party. (1st time!)", "Eu conheci um brasileiro na festa."], ["I stayed home and watched a series.", "Eu fiquei em casa e assisti uma série."], ["What did you do on the weekend?", "O que você fez no fim de semana?"]];
  await pair({ id: "trad", stage: "4 · PRÁTICA · EN → PT", title: "Traduza para o português", sub: "Say it in Portuguese — it doesn't need to be word for word", ic: "FaLanguage" }, async (s, r) => {
    tr.forEach(([en, pt], i) => {
      const y = 2.2 + i * 0.92;
      card(s, M, y, CW, 0.8, { fill: i % 2 ? C.white : C.panel, shadow: false });
      badge(s, M + 0.2, y + 0.19, i + 1);
      txt(s, en, { x: M + 0.85, y, w: 5.9, h: 0.8, fontSize: 19, valign: "middle", color: C.dark });
      txt(s, r ? pt : "→  ______________________________", { x: M + 6.85, y, w: CW - 7.0, h: 0.8, fontSize: 19, valign: "middle", bold: r, color: r ? C.green : C.muted });
    });
  }, "Aceite variações naturais (fui na padaria, vi uma série). O importante é o verbo no passado.");

  // ---------- 27-28. ERROR CORRECTION ----------
  const errs = [["Ontem eu ", { x: "fazi", a: "fiz" }, " o jantar."], ["— Você trabalhou ontem? — Sim, ", { x: "trabalhou", a: "trabalhei" }, "."], ["Semana passada eu ", { x: "vou", a: "fui" }, " ao cinema."], ["Ontem eu ", { x: "chegei", a: "cheguei" }, " muito tarde."], ["Anteontem eu ", { x: "dormei", a: "dormi" }, " dez horas."], ["Eu fui ao mercado e ", { x: "comprou", a: "comprei" }, " arroz."]];
  await pair({ id: "erro", stage: "4 · PRÁTICA · ENCONTRE O ERRO", title: "Encontre o erro", sub: "Each sentence has ONE mistake. Find it and fix it.", ic: "FaBug" }, async (s, r) => {
    const cw = (CW - 0.3) / 2, ch = 1.2;
    errs.forEach((sent, i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 2.2 + Math.floor(i / 2) * (ch + 0.25);
      card(s, x, y, cw, ch, { fill: r ? C.white : C.cSoft, shadow: false, line: r ? C.line : undefined });
      badge(s, x + 0.25, y + 0.39, i + 1, { fill: r ? C.green : C.coral, color: C.white });
      txt(s, runs(sent, r), { x: x + 0.85, y, w: cw - 1.05, h: ch, fontSize: 21, valign: "middle", color: C.dark });
    });
  }, "Deixe o aluno achar sozinho. Se travar, dê uma pista: 'o erro está no verbo 2'. O nº 2 é o erro mais comum da aula!");

  // ---------- 29. SENTENCE CREATION ----------
  {
    const s = await studentSlide({ id: "criar", stage: "4 · PRÁTICA · CRIE FRASES", title: "Crie uma frase com os dois verbos", sub: "Make ONE true sentence in the past with both verbs", ic: "FaUserPen",
      notes: "Produção livre. Peça frases VERDADEIRAS sobre a vida dele. Modelos nas teacher notes. Se sobrar tempo, peça duas frases por card." });
    const items = [["ir + comprar", "no sábado"], ["acordar + fazer", "hoje de manhã"], ["sair + encontrar", "ontem à noite"], ["chegar + ficar", "ontem"], ["ver + gostar", "semana passada"], ["vir + conhecer", "quando…"]];
    const cw = (CW - 2 * 0.3) / 3, ch = 1.6;
    items.forEach(([v, when], i) => {
      const x = M + (i % 3) * (cw + 0.3), y = 2.2 + Math.floor(i / 3) * (ch + 0.3);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      txt(s, v, { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.6, fontSize: 26, bold: true, fontFace: HEAD, color: C.dark });
      chip(s, x + 0.3, y + 0.98, when, { size: 14, fill: C.ySoft, color: C.dark, cs: 0 });
    });
    card(s, M, 6.1, CW, 0.7, { fill: C.mint, shadow: false });
    txt(s, [{ text: "Exemplo:  ", options: { bold: true, color: C.green } }, { text: "voltar + dormir  →  " }, { text: "Ontem eu voltei pra casa às 10h e dormi logo.", options: { bold: true } }], { x: M + 0.3, y: 6.1, w: CW - 0.6, h: 0.7, fontSize: 19, valign: "middle" });
  }

  // ---------- 30-31. Q&A ROUND 1 ----------
  const qa = [
    ["Que horas você acordou hoje?", ["acordar", "ficar", "sair"]], ["O que você comeu no café da manhã?", ["comer", "beber", "fazer"]],
    ["O que você fez ontem à noite?", ["ficar", "assistir", "sair", "dormir"]], ["Você trabalhou ou estudou ontem?", ["trabalhar", "estudar", "ter"]],
    ["Onde você foi no fim de semana?", ["ir", "ver", "encontrar", "voltar"]], ["Com quem você falou ontem? Sobre o quê?", ["falar", "dizer", "encontrar"]],
    ["Você assistiu alguma coisa essa semana?", ["assistir", "ver", "gostar"]], ["Você comprou alguma coisa recentemente?", ["comprar", "ir", "dar"]],
    ["O que você fez depois do trabalho ontem?", ["sair", "voltar", "ir", "fazer"]], ["Que horas você chegou em casa ontem?", ["chegar", "voltar", "ficar"]],
    ["Você conheceu alguém novo recentemente?", ["conhecer", "falar", "encontrar"]], ["Você comeu comida brasileira essa semana?", ["comer", "gostar", "ser"]],
    ["Você dormiu bem essa noite?", ["dormir", "acordar", "ter"]], ["Quando e por que você veio para o Brasil?", ["vir", "querer", "conhecer"]],
  ];
  for (let k = 0; k < 2; k++) {
    const s = await studentSlide({ id: "qa" + (k + 1), stage: "5 · PERGUNTAS · RODADA 1", title: `Eu pergunto, você responde (${k + 1}/2)`, sub: "Answer with full sentences + one extra detail. The verbs can help you.", ic: "FaComments",
      notes: "~3 min por slide. Depois de cada resposta, faça UM follow-up curto: Com quem? E depois? Gostou? Por quê? Não precisa fazer todas — se a conversa fluir, deixe fluir. Corrija com recast (repita a forma certa naturalmente)." });
    const list = qa.slice(k * 7, k * 7 + 7);
    const cw = (CW - 0.3) / 2, ch = 1.02;
    list.forEach(([q, vs], i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 2.2 + Math.floor(i / 2) * (ch + 0.16);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      badge(s, x + 0.2, y + 0.14, k * 7 + i + 1, { d: 0.38, size: 13 });
      txt(s, q, { x: x + 0.72, y: y + 0.08, w: cw - 0.9, h: 0.5, fontSize: 18, bold: true, color: C.dark, valign: "middle" });
      let cx = x + 0.72;
      for (const v of vs) { cx += chip(s, cx, y + 0.6, v, { size: 13, h: 0.3, fill: C.mint, color: C.green, cs: 0, bold: false }) + 0.1; }
    });
    const x = M + cw + 0.3, y = 2.2 + 3 * (ch + 0.16);
    card(s, x, y, cw, ch, { fill: C.ySoft, shadow: false });
    txt(s, [{ text: "Conecte as ideias:  ", options: { bold: true } }, { text: "primeiro… · depois… · aí… · mais tarde… · porque…" }], { x: x + 0.3, y, w: cw - 0.6, h: ch, fontSize: 16, valign: "middle" });
  }

  // ---------- 32-33. Q&A ROUND 2 ----------
  const cues = [["what / do / last night", "O que você fez ontem à noite?"], ["what time / wake up / today", "Que horas você acordou hoje?"], ["where / go / on the weekend", "Onde você foi no fim de semana?"], ["what / eat / for lunch", "O que você comeu no almoço?"], ["watch / anything good", "Você assistiu alguma coisa boa?"], ["buy / something / recently", "Você comprou alguma coisa recentemente?"],
    ["who / talk to / today", "Com quem você falou hoje?"], ["go out / on Saturday", "Você saiu no sábado?"], ["like / the last movie", "Você gostou do último filme que viu?"], ["meet (1st time) / someone", "Você conheceu alguém interessante?"], ["sleep well / last night", "Você dormiu bem ontem à noite?"], ["How was… / your weekend?", "Como foi o seu fim de semana?"]];
  await pair({ id: "qa3", stage: "5 · PERGUNTAS · RODADA 2", title: "Agora VOCÊ pergunta!", sub: "Ask your teacher in Portuguese. Then react — Sério? Que legal! — and ask one more.", ic: "FaCircleQuestion" }, async (s, r) => {
    const cw = (CW - 3 * 0.25) / 4, ch = 1.18;
    cues.forEach(([en, pt], i) => {
      const x = M + (i % 4) * (cw + 0.25), y = 2.2 + Math.floor(i / 4) * (ch + 0.2);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      txt(s, en, { x: x + 0.2, y: y + 0.12, w: cw - 0.4, h: 0.35, fontSize: r ? 13 : 17, italic: r, bold: !r, color: r ? C.muted : C.dark });
      if (r) txt(s, pt, { x: x + 0.2, y: y + 0.45, w: cw - 0.4, h: 0.68, fontSize: 15, bold: true, color: C.green });
    });
    card(s, M, 6.3, CW, 0.58, { fill: C.mint, shadow: false });
    txt(s, [{ text: "Comece com:  ", options: { bold: true, color: C.green } }, { text: "O que você…?  ·  Onde você…?  ·  Que horas você…?  ·  Com quem você…?  ·  Você…?  ·  Como foi…?" }], { x: M + 0.3, y: 6.3, w: CW - 0.6, h: 0.58, fontSize: 17, valign: "middle" });
  }, "4 min. O aluno pergunta, você responde DE VERDADE e com detalhes, usando de propósito fui, fiz, vi, tive, disse. Ele deve reagir e fazer 1 follow-up.",
     "Modelos de pergunta — mostre só depois que o aluno tentou todas. Variações naturais são corretas (Aonde você foi…?).");

  // ---------- 34. MEU ONTEM ----------
  {
    const s = await studentSlide({ id: "ontem", stage: "6 · DESAFIO ORAL · 6 MIN", title: "Meu ontem", sub: "Tell me about your day yesterday — from morning to night", ic: "FaMicrophone",
      notes: "Deixe o aluno falar 1–2 min SEM interromper. Conte os verbos diferentes (meta 8+). Anote no máximo 3 erros para o final. Depois use as perguntas do próximo slide." });
    chip(s, M, 2.15, "1–2 minutos", { size: 16, h: 0.46, fill: C.yellow, color: C.dark, cs: 0 });
    chip(s, M + chipW("1–2 minutos", 16) + 0.2, 2.15, "8+ verbos diferentes", { size: 16, h: 0.46, fill: C.yellow, color: C.dark, cs: 0 });
    const cols = [["FaSun", "De manhã", ["acordei", "fiz", "bebi", "saí", "tive"]], ["FaCloudSun", "À tarde", ["comi", "trabalhei", "fui", "encontrei", "comprei"]], ["FaMoon", "À noite", ["voltei", "cheguei", "assisti", "fiquei", "dormi"]]];
    const cw = (CW - 2 * 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const [ic, t, vs] = cols[i], x = M + i * (cw + 0.3), y = 2.9;
      card(s, x, y, cw, 2.75, { fill: i === 2 ? C.dark : C.panel, shadow: false });
      await iconCircle(s, x + 0.3, y + 0.25, 0.65, ic);
      txt(s, t, { x: x + 1.1, y: y + 0.25, w: cw - 1.3, h: 0.65, fontSize: 24, bold: true, fontFace: HEAD, valign: "middle", color: i === 2 ? C.white : C.dark });
      vs.forEach((v, j) => {
        const cx = x + 0.3 + (j % 2) * 1.65, cy = y + 1.15 + Math.floor(j / 2) * 0.52;
        chip(s, cx, cy, "☐  " + v, { size: 16, h: 0.4, w: 1.5, fill: i === 2 ? "1E5A45" : C.white, color: i === 2 ? C.white : C.dark, cs: 0, bold: false });
      });
    }
    card(s, M, 5.95, CW, 0.9, { fill: C.mint, shadow: false });
    txt(s, [{ text: "Primeiro… · Depois… · Aí… · Mais tarde… · No final do dia…", options: { bold: true, breakLine: true } }, { text: "Para ganhar tempo: Deixa eu pensar… · Hmm… · Tipo… · Então…", options: { italic: true, color: C.muted, fontSize: 15 } }], { x: M + 0.3, y: 5.95, w: CW - 0.6, h: 0.9, fontSize: 18, valign: "middle" });
  }

  // ---------- 35. FOLLOW-UP ----------
  {
    const s = await studentSlide({ id: "follow", stage: "6 · DESAFIO ORAL · CONVERSA", title: "Me conta mais…", sub: "Follow-up questions — answer in the past and keep the conversation going", ic: "FaComments",
      notes: "Faça 4–6 perguntas conforme o que o aluno contou. Termine com o role-play: você é um amigo brasileiro ligando na segunda-feira." });
    const qs = ["Que horas você saiu de casa?", "O que você comeu no almoço? Onde?", "Você falou com alguém interessante?", "Você comprou alguma coisa?", "O que você fez depois do trabalho?", "Qual foi a melhor parte do seu dia? E a pior?"];
    const cw = (CW - 0.3) / 2;
    qs.forEach((q, i) => {
      const x = M + (i % 2) * (cw + 0.3), y = 2.2 + Math.floor(i / 2) * 0.92;
      card(s, x, y, cw, 0.75, { fill: C.panel, shadow: false, r: 0.3 });
      txt(s, q, { x: x + 0.35, y, w: cw - 0.7, h: 0.75, fontSize: 19, valign: "middle", color: C.dark });
    });
    card(s, M, 5.2, CW, 1.6, { fill: C.dark });
    await iconCircle(s, M + 0.35, 5.55, 0.9, "FaPhone");
    chip(s, M + 1.55, 5.45, "ROLE-PLAY", { fill: C.yellow, color: C.dark });
    txt(s, "\"E aí, beleza? Como foi o fim de semana? Conta tudo!\"", { x: M + 1.55, y: 5.9, w: CW - 1.9, h: 0.55, fontSize: 24, bold: true, color: C.white });
    txt(s, "Your Brazilian friend calls you on Monday. Talk for 1–2 minutes.", { x: M + 1.55, y: 6.4, w: CW - 1.9, h: 0.35, fontSize: 15, italic: true, color: "C9DDD3" });
  }

  // ---------- 36-38. SCENARIOS ----------
  const scen = [
    ["FaCartShopping", "Um sábado muito corrido", "You had a VERY busy Saturday. You didn't stop all day!",
      ["Que horas você acordou? Por que tão cedo?", "Onde você foi primeiro?", "O que você comprou? Onde?", "Quem você encontrou?", "Onde e o que você comeu?", "Que horas você voltou pra casa?", "Como você ficou no final do dia?"],
      ["feira", "padaria", "academia", "shopping", "aniversário", "morto de cansaço", "correria"]],
    ["FaPlaneArrival", "Meu primeiro dia no Brasil", "Remember (or imagine) your first day in Brazil.",
      ["Quando você chegou? Em que cidade?", "Como você foi do aeroporto pro hotel?", "Você falou português? O que você disse?", "O que você comeu pela primeira vez?", "O que você comprou no primeiro dia?", "Você conheceu alguém?", "Do que você gostou? E do que não gostou?"],
      ["pedi um Uber", "chip de celular", "calor", "açaí", "pão de queijo", "fila", "não entendi nada"]],
    ["FaFaceDizzy", "Um dia em que deu tudo errado", "Yesterday was a terrible day. Everything went wrong!",
      ["Que horas você acordou? Perdeu a hora?", "O que aconteceu no caminho?", "Você chegou atrasado? O que seu chefe disse?", "O que você esqueceu ou perdeu?", "O que você comeu? Estava bom?", "O que você fez pra resolver?", "Como o dia terminou?"],
      ["perdi a hora", "perdi o ônibus", "esqueci", "choveu", "trânsito", "a bateria acabou", "que dia!"]],
  ];
  for (let k = 0; k < 3; k++) {
    const [ic, t, en, qs, words] = scen[k];
    const s = await studentSlide({ id: "scen" + (k + 1), stage: `6 · DESAFIO ORAL · O QUE ACONTECEU? ${k + 1}/3`, title: "O que aconteceu?", sub: "What happened? Tell the story in the past (eu…). You can invent everything!", ic: "FaMicrophone",
      notes: "Dê 30 s para o aluno olhar o cenário e depois ele conta a história. As perguntas são apoio — não precisa seguir a ordem. Se o aluno estiver avançado: peça para contar em 3ª pessoa (o Pedro…) ou inverta os papéis." });
    card(s, M, 2.2, 4.3, 4.65, { fill: C.dark });
    await iconCircle(s, M + 0.4, 2.55, 1.0, ic);
    chip(s, M + 0.4, 3.85, "CENÁRIO " + (k + 1), { fill: C.green, color: C.white });
    txt(s, t, { x: M + 0.4, y: 4.35, w: 3.5, h: 1.3, fontSize: 28, bold: true, fontFace: HEAD, color: C.white });
    txt(s, en, { x: M + 0.4, y: 5.75, w: 3.5, h: 0.9, fontSize: 15, italic: true, color: "C9DDD3" });
    const x = M + 4.6, w = W - M - x;
    txt(s, qs.map((q, j) => ({ text: q, options: { bullet: true, breakLine: j < qs.length - 1 } })), { x, y: 2.2, w, h: 3.5, fontSize: 19, paraSpaceAfter: 6, color: C.dark });
    card(s, x, 5.85, w, 1.0, { fill: C.ySoft, shadow: false });
    txt(s, [{ text: "Palavras úteis:  ", options: { bold: true } }, { text: words.join("  ·  ") }], { x: x + 0.3, y: 5.85, w: w - 0.6, h: 1.0, fontSize: 16, valign: "middle" });
  }

  // ---------- 39-40. SPEED ROUND ----------
  const speed = [["ir", "fui"], ["fazer", "fiz"], ["ter", "tive"], ["ver", "vi"], ["vir", "vim"], ["dizer", "disse"], ["sair", "saí"], ["chegar", "cheguei"], ["ficar", "fiquei"], ["dar", "dei"], ["querer", "quis"], ["estar", "estive"]];
  await pair({ id: "speed", stage: "7 · REVISÃO FINAL", title: "Rodada relâmpago", sub: "Speed round: say the past (eu) as fast as you can!", ic: "FaBolt" }, async (s, r) => {
    const cw = (CW - 5 * 0.22) / 6, ch = 1.75;
    speed.forEach(([inf, pa], i) => {
      const x = M + (i % 6) * (cw + 0.22), y = 2.3 + Math.floor(i / 6) * (ch + 0.35);
      card(s, x, y, cw, ch, { fill: i < 6 ? C.white : C.panel, line: i < 6 ? C.line : undefined, shadow: i < 6 });
      txt(s, inf, { x, y: y + 0.2, w: cw, h: 0.5, fontSize: 22, bold: true, color: C.dark, align: "center" });
      txt(s, r ? "eu " + pa : "eu _____", { x, y: y + 0.95, w: cw, h: 0.55, fontSize: 22, bold: r, color: r ? C.green : C.muted, align: "center" });
    });
  }, "1 min. Você diz o infinitivo em ordem aleatória, o aluno responde. Repita os que ele errou.");

  // ---------- 41. 3-2-1 ----------
  {
    const s = await studentSlide({ id: "321", stage: "7 · REVISÃO FINAL", title: "Para fechar: 3 · 2 · 1", sub: "Exit ticket — say it out loud", ic: "FaFlagCheckered",
      notes: "2 min. Depois do 3-2-1, apresente as 3 correções que você anotou durante a aula (no máximo 3!)." });
    const items = [["3", "verbos irregulares que eu usei hoje", "3 irregular verbs I used today"], ["2", "frases sobre algo que eu fiz hoje", "2 sentences about something I did today"], ["1", "pergunta no passado para um amigo brasileiro", "1 question in the past for a Brazilian friend"]];
    const cw = (CW - 2 * 0.3) / 3;
    items.forEach(([n, pt, en], i) => {
      const x = M + i * (cw + 0.3);
      card(s, x, 2.3, cw, 4.2, { fill: i === 0 ? C.dark : (i === 1 ? C.green : C.yellow), shadow: true });
      txt(s, n, { x: x + 0.4, y: 2.5, w: 2, h: 1.6, fontSize: 96, bold: true, fontFace: HEAD, color: i === 2 ? C.dark : (i === 0 ? C.yellow : C.white) });
      txt(s, pt, { x: x + 0.4, y: 4.25, w: cw - 0.8, h: 1.2, fontSize: 22, bold: true, color: i === 2 ? C.dark : C.white });
      txt(s, en, { x: x + 0.4, y: 5.5, w: cw - 0.8, h: 0.8, fontSize: 15, italic: true, color: i === 2 ? C.dark : "D6E7DE" });
    });
  }

  // ---------- 42. HOMEWORK ----------
  {
    const s = await studentSlide({ id: "hw", stage: "LIÇÃO DE CASA · HOMEWORK", title: "Para a próxima aula", sub: "Homework — practice a little every day", ic: "FaHouseLaptop",
      notes: "Combine prazo e formato do áudio de WhatsApp. Peça que ele traga os 2 verbos que ouviu de um brasileiro — ótimo aquecimento para a próxima aula." });
    const hw = [["FaWhatsapp", "Áudio de WhatsApp", "1–2 min: \"Meu fim de semana\". Use 10+ verbos no passado (4 irregulares)."], ["FaBook", "Mini diário · 3 dias", "Toda noite, 3 frases sobre o seu dia. Ex.: Hoje eu acordei cedo e fui à academia."], ["FaLayerGroup", "Flashcards", "Os 11 irregulares: frente fazer · verso eu fiz / você fez. 2 min por dia."], ["FaEarthAmericas", "Desafio real", "Pergunte a um brasileiro: \"O que você fez no fim de semana?\" Anote 2 verbos que ouvir."]];
    const cw = (CW - 0.3) / 2, ch = 1.95;
    for (let i = 0; i < 4; i++) {
      const [ic, t, d] = hw[i], x = M + (i % 2) * (cw + 0.3), y = 2.2 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch, { fill: C.white, line: C.line });
      await iconCircle(s, x + 0.3, y + 0.3, 0.8, ic, { fill: C.green, color: "FFFFFF" });
      txt(s, t, { x: x + 1.35, y: y + 0.3, w: cw - 1.6, h: 0.5, fontSize: 22, bold: true, color: C.dark });
      txt(s, d, { x: x + 1.35, y: y + 0.85, w: cw - 1.6, h: 0.95, fontSize: 17 });
    }
  }
  const lastStudent = slideNo;

  // =================== TEACHER SLIDES ===================
  async function teacherSlide(title, sub) {
    const s = pres.addSlide(); slideNo++;
    s.background = { color: T.bg };
    chip(s, M, 0.45, "TEACHER ONLY · NÃO COMPARTILHAR", { fill: T.accent, color: T.bg });
    txt(s, title, { x: M, y: 0.92, w: CW, h: 0.7, fontSize: 30, bold: true, fontFace: HEAD, color: T.text, valign: "middle" });
    if (sub) txt(s, sub, { x: M, y: 1.58, w: CW, h: 0.4, fontSize: 15, italic: true, color: T.muted });
    txt(s, String(slideNo), { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontSize: 10, color: T.muted, align: "right" });
    return s;
  }
  const tBox = (s, x, y, w, h, title, lines, o = {}) => {
    card(s, x, y, w, h, { fill: T.card, shadow: false });
    txt(s, title, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.35, fontSize: 15, bold: true, color: T.accent });
    txt(s, lines.map((l, i) => {
      const parts = Array.isArray(l) ? l : [l];
      return parts.map((p, j) => ({ text: p, options: { bold: j === 0 && parts.length > 1, color: j === 0 && parts.length > 1 ? T.text : "DDE2EA", breakLine: j === parts.length - 1 && i < lines.length - 1 } }));
    }).flat(), { x: x + 0.25, y: y + 0.55, w: w - 0.5, h: h - 0.65, fontSize: o.size || 13, color: "DDE2EA", paraSpaceAfter: o.ps ?? 3 });
  };
  const tTable = (s, rows, colW, y, o = {}) => {
    const hdr = rows[0].map(t => ({ text: t, options: { bold: true, color: T.bg, fill: { color: T.accent } } }));
    const body = rows.slice(1).map((r, i) => r.map((c, j) => ({ text: c, options: { fill: { color: i % 2 ? T.card : T.card2 }, color: j === 0 ? T.text : "DDE2EA", bold: j === 0 } })));
    s.addTable([hdr, ...body], { x: M, y, w: CW, colW, fontFace: BODY, fontSize: o.size || 12, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1], border: { type: "solid", color: T.bg, pt: 1 }, rowH: o.rowH });
  };

  // T1 divider
  {
    const s = pres.addSlide(); slideNo++;
    s.background = { color: T.bg };
    s.addShape(pres.shapes.OVAL, { x: 9.2, y: 1.2, w: 3.4, h: 3.4, fill: { color: T.card }, line: { type: "none" } });
    s.addImage({ data: await icon("FaChalkboardUser", T.accent), x: 10.15, y: 2.15, w: 1.5, h: 1.5 });
    chip(s, M, 1.2, "TEACHER ONLY · NÃO COMPARTILHAR", { fill: T.accent, color: T.bg });
    txt(s, "Para o professor", { x: M, y: 1.85, w: 8, h: 1.0, fontSize: 48, bold: true, fontFace: HEAD, color: T.text });
    txt(s, `Pare de compartilhar a tela antes desta parte. Os slides do aluno terminam no slide ${lastStudent}.`, { x: M, y: 2.9, w: 8, h: 0.8, fontSize: 17, italic: true, color: T.muted });
    txt(s, ["Plano da aula e o que observar", "Gabarito completo", "Respostas-modelo (produção livre)", "Erros comuns e correções sugeridas", "Como corrigir e como adaptar", "Mapa de reciclagem dos verbos", "Perguntas extras de follow-up"].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
      { x: M, y: 3.95, w: 8, h: 3.0, fontSize: 18, color: T.text, paraSpaceAfter: 6 });
    s.addNotes("Slides só para o professor. Use o modo apresentador ou pare de compartilhar a tela.");
  }

  // T2 plan
  {
    const s = await teacherSlide("Plano da aula e o que observar", "20% exposição · 80% prática. Timing sugerido — ajuste ao ritmo do aluno.");
    tTable(s, [
      ["Etapa", "Slides", "Tempo", "Objetivo", "O que observar"],
      ["1. Aquecimento", `${N.warm}`, "5'", "Ativar o que ele já sabe; diagnosticar automatização.", "Responde com o verbo ou só 'sim'? Usa a forma do EU ou repete a da pergunta?"],
      ["2. Verbos (exposição)", `${N.euvoce - 1}–${N.cuidado}`, "10'", "Input com significado + manipulação imediata. Foco nos 11 irregulares.", "Quais formas ele trava ou regulariza (fazi, dizei). Anote 3–4 para reciclar."],
      ["3. Vocabulário", `${N.enpt}–${N.pten_ans}`, "(dentro dos 10')", "Reconhecer significado e forma; EN↔PT.", "Confusões de sentido: conheci × encontrei, vim × fui, ficar."],
      ["4. Prática controlada", `${N.mc4}–${N.criar}`, "10'", "Fixar forma e pistas de tempo com baixa pressão.", "Velocidade. Se acerta mas demora, faça tudo oral e mais rápido."],
      ["5. Perguntas", `${N.qa1}–${N.qa3_ans}`, "10'", "Responder E perguntar: forma você × eu.", "Pergunta sem ler? Reage? Faz follow-up?"],
      ["6. Desafio oral", `${N.ontem}–${N.scen3}`, "13'", "Produção longa, narrativa, vários verbos seguidos.", "Nº de verbos diferentes, conectores, se o passado se mantém quando ele foca no conteúdo."],
      ["7. Revisão + casa", `${N.speed}–${N.hw}`, "2'", "Consolidar e dar 3 correções finais.", "Máximo 3 erros na correção final."],
    ], [2.1, 1.0, 1.1, 3.9, 4.03], 2.15, { size: 13 });
  }

  // T3 answer key 1
  {
    const s = await teacherSlide("Gabarito completo (1/2)", "Vocabulário e prática controlada");
    const cw = (CW - 2 * 0.25) / 3;
    tBox(s, M, 2.15, cw, 2.2, `Quiz rápido · slide ${N.mc1}`, [["Ontem eu fiz café.", " → B) fiz"]], { size: 14 });
    tBox(s, M, 4.5, cw, 2.35, `Como se diz? · slide ${N.enpt}`, ["1 ir / fui · 2 ter / tive", "3 ver / vi · 4 vir / vim", "5 dizer / disse · 6 ficar / fiquei", "7 chegar / cheguei", "8 conhecer / conheci"], { size: 14 });
    tBox(s, M + cw + 0.25, 2.15, cw, 2.2, `Matching · slide ${N.match}`, ["1 fazer → C fiz · 2 sair → F saí", "3 dar → H dei · 4 querer → B quis", "5 dormir → A dormi · 6 estar → E estive", "7 comer → G comi · 8 voltar → D voltei"], { size: 14 });
    tBox(s, M + cw + 0.25, 4.5, cw, 2.35, `Múltipla escolha · slide ${N.mc4}`, ["1 B) fui", "2 A) viu", "3 B) acordo / acordei", "4 A) pude / tive  (natural: não consegui)"], { size: 14 });
    tBox(s, M + 2 * (cw + 0.25), 2.15, cw, 2.2, `Complete · slide ${N.fill}`, ["1 fui · comprei", "2 tive · 3 saí", "4 cheguei · 5 disse", "6 fiquei"], { size: 14 });
    tBox(s, M + 2 * (cw + 0.25), 4.5, cw, 2.35, `Presente → passado · slide ${N.transf}`, ["acordo → acordei", "faço → fiz · como → comi", "vou → fui", "vejo → vi · durmo → dormi"], { size: 14 });
  }

  // T4 answer key 2
  {
    const s = await teacherSlide("Gabarito completo (2/2)", "Tradução, erros, PT → EN e respostas-modelo para a produção livre");
    const cw = (CW - 0.25) / 2;
    tBox(s, M, 2.15, cw, 2.35, `PT → EN · slide ${N.pten}`, ["1 I met Ana (1st time) at the party.", "2 I ran into / met up with Ana at the mall.", "3 I got tired after work.  4 I left work at 6.", "5 I came to Brazil in 2024.  6 The show was amazing! (foi = ser)"], { size: 13 });
    tBox(s, M, 4.65, cw, 2.2, `EN → PT · slide ${N.trad}`, ["1 Eu fui à padaria (na padaria) e comprei pão.", "2 Eu tive uma reunião de manhã.", "3 Eu conheci um brasileiro na festa.", "4 Eu fiquei em casa e assisti (a) uma série / vi uma série.", "5 O que você fez no fim de semana?"], { size: 13 });
    tBox(s, M + cw + 0.25, 2.15, cw, 2.35, `Encontre o erro · slide ${N.erro}`, ["1 fazi → fiz · 2 trabalhou → trabalhei", "3 vou → fui · 4 chegei → cheguei", "5 dormei → dormi · 6 comprou → comprei"], { size: 13 });
    tBox(s, M + cw + 0.25, 4.65, cw, 2.2, `Crie frases (modelos) · slide ${N.criar}`, ["No sábado eu fui ao shopping e comprei um tênis.", "Hoje de manhã eu acordei cedo e fiz café.", "Ontem à noite eu saí e encontrei uns amigos.", "Ontem eu cheguei tarde e fiquei muito cansado.", "Semana passada eu vi um filme e gostei muito.", "Quando eu vim pro Brasil, conheci muita gente."], { size: 12, ps: 1 });
  }

  // T5 common mistakes
  {
    const s = await teacherSlide("Erros mais prováveis e correções sugeridas", "Priorize os erros 1–4: são frequentes e afetam a comunicação");
    tTable(s, [
      ["O aluno diz…", "Forma correta", "Por que acontece / como corrigir"],
      ["— Você comeu? — Sim, comeu.", "Sim, comi.", "Copia a forma da pergunta. Erro nº 1: corrija SEMPRE, na hora (gesto 'eu').  Slides " + N.euvoce + ", " + N.erro],
      ["eu foi · eu comeu · eu fez", "eu fui · eu comi · eu fiz", "3ª pessoa no lugar do eu. Recast: \"Ah, você FOI…\" e peça para repetir a frase."],
      ["fazi · dizei · teni · vei · podei", "fiz · disse · tive · vim · pude", "Regulariza o irregular. Volte aos slides " + N.bank1 + "–" + N.bank2 + "; recicle na rodada relâmpago."],
      ["Ontem eu vou ao mercado.", "Ontem eu fui…", "Ignora a pista de tempo. Aponte a palavra 'ontem' (slide de revisão todo dia × ontem)."],
      ["chegei · ficei · sai (= saí)", "cheguei · fiquei · saí", "Ortografia gu/qu e acento. Na fala, marque o -EI e o sa-Í."],
      ["Eu fui para o Brasil (estando aqui)", "Eu vim para o Brasil", "Direção: vim = para cá; fui = para lá. Gesto aqui × lá."],
      ["Encontrei minha esposa em 2015.", "Conheci minha esposa…", "conhecer = primeira vez; encontrar = rever / achar."],
      ["Ontem eu estive cansado.", "Fiquei cansado / estava cansado", "'estive' soa estranho aqui. Só reformule; não entre no imperfeito agora."],
    ], [3.3, 2.9, 5.93], 2.15, { size: 13 });
  }

  // T6 correction + adaptation
  {
    const s = await teacherSlide("Como corrigir e como adaptar", "Na prática controlada, corrija na hora. Na conversa, priorize a fluência.");
    const cw = (CW - 2 * 0.25) / 3;
    tBox(s, M, 2.15, cw, 4.7, "Corrigir sem interromper", [
      ["Recast: ", "repita certo, como reação natural. Aluno: \"Eu fazi uma pizza.\" → \"Ah, você FEZ uma pizza! De quê?\""],
      ["Autocorreção: ", "diga só o começo — \"Ontem eu…?\" — ou repita o erro com tom de dúvida: \"Fazi?\""],
      ["Gesto: ", "polegar para trás = 'passado!'. Combine no início."],
      ["Correção adiada: ", "nas etapas 5 e 6, anote e corrija no fim. Máximo 3 erros."],
      ["Exceção: ", "corrija na hora o 'Sim, comeu' e erros que impedem a compreensão."],
    ], { size: 13, ps: 6 });
    tBox(s, M + cw + 0.25, 2.15, cw, 4.7, "Se o aluno tiver dificuldade", [
      "Reduza para 15 verbos: os 11 irregulares + acordar, comer, trabalhar, sair.",
      "Comece com perguntas de sim/não e aceite respostas curtas (Sim, fui.).",
      "Deixe o banco de verbos visível (volte aos slides " + N.bank1 + "–" + N["tab_FaUtensils"] + ").",
      "Meu ontem: 1 min de preparação com palavras-chave; meta de 5 verbos.",
      "Faça só o Cenário 1, em formato de entrevista (pergunta por pergunta).",
      "Faça a transformação frase por frase.",
    ], { size: 13, ps: 6 });
    tBox(s, M + 2 * (cw + 0.25), 2.15, cw, 4.7, "Se o aluno estiver avançando rápido", [
      "Esconda o banco de verbos depois de 1 minuto.",
      "Exija 2 verbos + 1 detalhe em cada resposta.",
      "Acrescente ele/ela, a gente (a gente foi) e eles (eles foram, fizeram).",
      "Cenários em 3ª pessoa (\"O que aconteceu com o Pedro?\") ou com perguntas surpresa.",
      "Desafio: fim de semana em 90 s com 12 verbos diferentes.",
      "Teaser: Eu fui à praia × Eu ia à praia todo sábado — sem explicar ainda.",
    ], { size: 13, ps: 6 });
  }

  // T7 recycling map
  {
    const s = await teacherSlide("Verbos-chave e mapa de reciclagem", "Cada verbo de alta frequência aparece em 3+ atividades — números = slides");
    const map = [
      ["fazer → fiz", "Quiz " + N.mc1, "Matching " + N.match, "Transformação " + N.transf, "Erro " + N.erro, "Criar " + N.criar, "Q&A · Meu ontem"],
      ["ir → fui", "Como se diz " + N.enpt, "Múlt. escolha " + N.mc4, "Complete " + N.fill, "Transformação " + N.transf, "Tradução " + N.trad, "Erro " + N.erro + " · Relâmpago " + N.speed],
      ["ter → tive", "Como se diz " + N.enpt, "Múlt. escolha " + N.mc4, "Complete " + N.fill, "Tradução " + N.trad, "Relâmpago " + N.speed, ""],
      ["ver → vi", "Como se diz " + N.enpt, "Múlt. escolha " + N.mc4, "Transformação " + N.transf, "Criar " + N.criar, "Relâmpago " + N.speed, ""],
      ["vir → vim", "Como se diz " + N.enpt, "PT→EN " + N.pten, "Criar " + N.criar, "Q&A " + N.qa2, "Relâmpago " + N.speed, ""],
      ["sair → saí", "Matching " + N.match, "PT→EN " + N.pten, "Complete " + N.fill, "Criar " + N.criar, "Relâmpago " + N.speed, ""],
      ["chegar → cheguei", "Como se diz " + N.enpt, "Complete " + N.fill, "Erro " + N.erro, "Criar " + N.criar, "Relâmpago " + N.speed, ""],
      ["ficar → fiquei", "Como se diz " + N.enpt, "PT→EN " + N.pten, "Complete " + N.fill, "Tradução " + N.trad, "Criar " + N.criar, "Relâmpago " + N.speed],
      ["conhecer → conheci", "Como se diz " + N.enpt, "PT→EN " + N.pten, "Tradução " + N.trad, "Criar " + N.criar, "Q&A " + N.qa2, ""],
      ["dizer → disse", "Como se diz " + N.enpt, "Complete " + N.fill, "Relâmpago " + N.speed, "Q&A " + N.qa1, "", ""],
      ["comprar → comprei", "Complete " + N.fill, "Tradução " + N.trad, "Erro " + N.erro, "Criar " + N.criar, "", ""],
      ["acordar / dormir / comer", "Múlt. escolha " + N.mc4, "Matching " + N.match, "Transformação " + N.transf, "Erro " + N.erro, "Meu ontem " + N.ontem, ""],
    ];
    tTable(s, [["Verbo", "Atividade 1", "Atividade 2", "Atividade 3", "Atividade 4", "Atividade 5", "Atividade 6"], ...map], [2.3, 1.63, 1.63, 1.63, 1.63, 1.63, 1.68], 2.15, { size: 11.5, rowH: 0.34 });
  }

  // T8 follow-up questions
  {
    const s = await teacherSlide("Perguntas extras de follow-up", "Use quando a conversa esfriar, ou para desafiar um aluno mais rápido");
    const cw = (CW - 2 * 0.25) / 3;
    tBox(s, M, 2.15, cw, 4.7, "Depois de \"Meu ontem\"", ["Que horas você dormiu?", "Você falou português com alguém ontem? Com quem?", "Você pediu comida ou cozinhou?", "Você usou o celular muito tempo?", "O que você fez que não faz todo dia?", "Se pudesse repetir o dia, o que mudaria?"], { size: 14, ps: 8 });
    tBox(s, M + cw + 0.25, 2.15, cw, 4.7, "Fim de semana e vida no Brasil", ["Qual foi o último lugar novo que você conheceu?", "Qual foi a última coisa que você comprou online?", "Você já foi a um churrasco? Como foi?", "Qual foi a comida mais estranha que você já comeu aqui?", "Você já se perdeu numa cidade brasileira?", "O que você fez no último feriado?"], { size: 14, ps: 8 });
    tBox(s, M + 2 * (cw + 0.25), 2.15, cw, 4.7, "Reações para modelar", ["Sério? · Que legal! · Que chato!", "Nossa! · Que delícia! · Imagina!", "E aí? O que aconteceu depois?", "E você gostou?", "Com quem você foi?", "Dica: peça que o aluno use 2 reações por conversa — deixa o diálogo natural."], { size: 14, ps: 8 });
  }

  await pres.writeFile({ fileName: OUT });
  console.log("slides:", slideNo, "student:", lastStudent, OUT);
}
build().catch(e => { console.error(e); process.exit(1); });
