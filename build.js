const pptxgen = require("pptxgenjs");

// ---------- Palette ----------
const DARK_BG = "07373A";
const PRIMARY = "028090";
const SECONDARY = "00A896";
const ACCENT = "02C39A";
const WHITE = "FFFFFF";
const TEXT_DARK = "1C2B2C";
const TEXT_MUTED = "6E7C7D";
const CARD_BG = "F2F7F7";
const PROBLEM = "C1440E";
const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";
const FONT_CODE = "Courier New";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
const PAGE_W = 13.333;
const PAGE_H = 7.5;

function footer(slide, label) {
  slide.addText(label, {
    x: 0.5, y: PAGE_H - 0.42, w: 8, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, color: TEXT_MUTED, isTextBox: true, margin: 0,
  });
}

function pageNum(slide, n) {
  slide.addText(String(n), {
    x: PAGE_W - 0.9, y: PAGE_H - 0.42, w: 0.5, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, color: TEXT_MUTED, align: "right", isTextBox: true, margin: 0,
  });
}

// ---------- 1. Title ----------
function slideTitle() {
  const s = pres.addSlide();
  s.background = { color: DARK_BG };
  s.addShape(pres.ShapeType.ellipse, { x: 10.6, y: -1.8, w: 5, h: 5, fill: { color: SECONDARY, transparency: 88 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: -1.5, y: 4.8, w: 4.2, h: 4.2, fill: { color: ACCENT, transparency: 90 }, line: { type: "none" } });

  s.addText("GUÍA DE ESTUDIO", {
    x: 0.9, y: 2.15, w: 8, h: 0.4, fontFace: FONT_BODY, fontSize: 14, color: ACCENT, charSpacing: 3, bold: true, isTextBox: true, margin: 0,
  });
  s.addText("Dos proyectos de IA,\nde cero a producción", {
    x: 0.85, y: 2.55, w: 10.5, h: 2.0, fontFace: FONT_HEAD, fontSize: 40, bold: true, color: WHITE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.05,
  });
  s.addText("Chatbot RAG  +  Clasificador de Residuos Reciclables", {
    x: 0.9, y: 4.55, w: 10, h: 0.5, fontFace: FONT_BODY, fontSize: 18, color: "CFEFEC", isTextBox: true, margin: 0,
  });
  s.addText("Arquitectura · decisiones de diseño · bugs reales y cómo se resolvieron", {
    x: 0.9, y: 5.05, w: 10, h: 0.4, fontFace: FONT_BODY, fontSize: 13, italic: true, color: "9FC9C6", isTextBox: true, margin: 0,
  });
}

// ---------- Section divider ----------
function slideSection(number, title, subtitle) {
  const s = pres.addSlide();
  s.background = { color: DARK_BG };
  s.addText(number, {
    x: 0.9, y: 2.35, w: 3, h: 2.2, fontFace: FONT_HEAD, fontSize: 110, bold: true, color: SECONDARY, isTextBox: true, margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 3.7, y: 2.55, w: 0.04, h: 2.0, fill: { color: ACCENT }, line: { type: "none" } });
  s.addText(title, {
    x: 4.0, y: 2.55, w: 8.4, h: 1.1, fontFace: FONT_HEAD, fontSize: 38, bold: true, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText(subtitle, {
    x: 4.0, y: 3.65, w: 8.2, h: 0.9, fontFace: FONT_BODY, fontSize: 16, color: "CFEFEC", isTextBox: true, margin: 0,
  });
}

// ---------- Generic content title bar ----------
function contentHeader(s, kicker, title) {
  s.background = { color: WHITE };
  s.addText(kicker, {
    x: 0.6, y: 0.4, w: 10, h: 0.35, fontFace: FONT_BODY, fontSize: 12, bold: true, color: SECONDARY, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: 0.6, y: 0.72, w: 12, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: TEXT_DARK, isTextBox: true, margin: 0,
  });
}

// ---------- 2. Agenda ----------
function slideAgenda() {
  const s = pres.addSlide();
  contentHeader(s, "AGENDA", "Qué vamos a repasar");
  const items = [
    ["01", "Proyecto 1 — Chatbot RAG", "Arquitectura, decisiones, bugs y deploy"],
    ["02", "Proyecto 2 — Clasificador de Residuos", "Transfer learning, ONNX, bugs y deploy"],
    ["03", "Lecciones transversales", "Qué se repite entre ambos proyectos"],
    ["04", "Próximos pasos", "Hacia dónde sigue el portafolio de IA"],
  ];
  let y = 1.85;
  items.forEach(([num, title, desc]) => {
    s.addShape(pres.ShapeType.roundRect, { x: 0.7, y, w: 0.75, h: 0.75, rectRadius: 0.12, fill: { color: PRIMARY }, line: { type: "none" } });
    s.addText(num, { x: 0.7, y, w: 0.75, h: 0.75, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 20, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(title, { x: 1.75, y: y - 0.02, w: 7.5, h: 0.4, fontFace: FONT_BODY, fontSize: 17, bold: true, color: TEXT_DARK, isTextBox: true, margin: 0 });
    s.addText(desc, { x: 1.75, y: y + 0.36, w: 8.5, h: 0.35, fontFace: FONT_BODY, fontSize: 13, color: TEXT_MUTED, isTextBox: true, margin: 0 });
    y += 1.15;
  });
  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, 2);
}

// ---------- Concept (problem/solution two-column) ----------
function slideConcept(kicker, title, leftLabel, leftText, rightLabel, rightText, pageN) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);
  const colY = 2.1, colH = 4.4, colW = 5.7;
  // Left = problem
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: colY, w: colW, h: colH, rectRadius: 0.08, fill: { color: CARD_BG }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: 0.95, y: colY + 0.4, w: 0.6, h: 0.6, fill: { color: PROBLEM }, line: { type: "none" } });
  s.addText("!", { x: 0.95, y: colY + 0.4, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 22, bold: true, color: WHITE, isTextBox: true, margin: 0 });
  s.addText(leftLabel, { x: 1.75, y: colY + 0.45, w: 4.2, h: 0.5, fontFace: FONT_BODY, fontSize: 17, bold: true, color: PROBLEM, isTextBox: true, margin: 0 });
  s.addText(leftText, { x: 0.95, y: colY + 1.25, w: colW - 0.7, h: colH - 1.6, fontFace: FONT_BODY, fontSize: 14.5, color: TEXT_DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });

  const rx = 6.95;
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: colY, w: colW, h: colH, rectRadius: 0.08, fill: { color: PRIMARY }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: rx + 0.35, y: colY + 0.4, w: 0.6, h: 0.6, fill: { color: ACCENT }, line: { type: "none" } });
  s.addText("✓", { x: rx + 0.35, y: colY + 0.4, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 22, bold: true, color: DARK_BG, isTextBox: true, margin: 0 });
  s.addText(rightLabel, { x: rx + 1.15, y: colY + 0.45, w: 4.2, h: 0.5, fontFace: FONT_BODY, fontSize: 17, bold: true, color: WHITE, isTextBox: true, margin: 0 });
  s.addText(rightText, { x: rx + 0.35, y: colY + 1.25, w: colW - 0.7, h: colH - 1.6, fontFace: FONT_BODY, fontSize: 14.5, color: "E9FBF8", isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });

  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Flow diagram (phases) ----------
function flowBox(s, x, y, w, h, title, sub, color, textColor) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color }, line: { type: "none" } });
  s.addText(title, { x: x + 0.12, y: y + 0.08, w: w - 0.24, h: h - (sub ? 0.42 : 0.16), align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 12.5, bold: true, color: textColor, isTextBox: true, margin: 0, lineSpacingMultiple: 1.05 });
  if (sub) {
    s.addText(sub, { x: x + 0.1, y: y + h - 0.34, w: w - 0.2, h: 0.3, align: "center", fontFace: FONT_BODY, fontSize: 9.5, color: textColor, isTextBox: true, margin: 0 });
  }
}
function arrowRight(s, x, y, w) {
  s.addShape(pres.ShapeType.rightArrow, { x, y, w, h: 0.22, fill: { color: TEXT_MUTED }, line: { type: "none" } });
}

function slideDiagram(kicker, title, rowATitle, rowASteps, rowBTitle, rowBSteps, note, pageN) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);

  const boxW = 2.05, boxH = 1.0, gap = 0.35, arrowW = 0.3;
  const startX = 0.6;
  let rowY = 2.15;

  s.addText(rowATitle, { x: startX, y: rowY - 0.4, w: 8, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
  let x = startX;
  rowASteps.forEach((step, i) => {
    flowBox(s, x, rowY, boxW, boxH, step, "", i === rowASteps.length - 1 ? SECONDARY : CARD_BG, i === rowASteps.length - 1 ? WHITE : TEXT_DARK);
    x += boxW;
    if (i < rowASteps.length - 1) { arrowRight(s, x, rowY + boxH / 2 - 0.11, arrowW); x += arrowW; }
  });

  rowY = 4.15;
  s.addText(rowBTitle, { x: startX, y: rowY - 0.4, w: 8, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
  x = startX;
  rowBSteps.forEach((step, i) => {
    flowBox(s, x, rowY, boxW, boxH, step, "", i === rowBSteps.length - 1 ? ACCENT : CARD_BG, i === rowBSteps.length - 1 ? DARK_BG : TEXT_DARK);
    x += boxW;
    if (i < rowBSteps.length - 1) { arrowRight(s, x, rowY + boxH / 2 - 0.11, arrowW); x += arrowW; }
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.55, w: 12.1, h: 1.05, rectRadius: 0.08, fill: { color: CARD_BG }, line: { type: "none" } });
  s.addText(note, { x: 0.9, y: 5.68, w: 11.5, h: 0.85, fontFace: FONT_BODY, fontSize: 13, color: TEXT_DARK, valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Steps deep-dive (icon rows) ----------
function slideSteps(kicker, title, steps, pageN, codeBlock) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);
  const hasCode = !!codeBlock;
  const colW = hasCode ? 6.9 : 12.1;
  let y = 1.95;
  const rowH = (5.15) / steps.length;
  steps.forEach((step, i) => {
    const cy = y + i * rowH;
    s.addShape(pres.ShapeType.ellipse, { x: 0.6, y: cy, w: 0.55, h: 0.55, fill: { color: i % 2 === 0 ? PRIMARY : SECONDARY }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 0.6, y: cy, w: 0.55, h: 0.55, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 16, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(step.title, { x: 1.35, y: cy - 0.06, w: colW - 0.8, h: 0.35, fontFace: FONT_BODY, fontSize: 14.5, bold: true, color: TEXT_DARK, isTextBox: true, margin: 0 });
    s.addText(step.desc, { x: 1.35, y: cy + 0.28, w: colW - 0.8, h: rowH - 0.35, fontFace: FONT_BODY, fontSize: 12, color: TEXT_MUTED, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  });

  if (hasCode) {
    s.addShape(pres.ShapeType.roundRect, { x: 7.85, y: 1.95, w: 4.85, h: 5.15, rectRadius: 0.08, fill: { color: DARK_BG }, line: { type: "none" } });
    s.addText(codeBlock.label, { x: 8.1, y: 2.1, w: 4.4, h: 0.3, fontFace: FONT_BODY, fontSize: 11, bold: true, color: ACCENT, isTextBox: true, margin: 0 });
    s.addText(codeBlock.code, { x: 8.1, y: 2.45, w: 4.4, h: 4.5, fontFace: FONT_CODE, fontSize: 10.5, color: "E9FBF8", isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });
  }

  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Glossary grid ----------
function slideGlossary(kicker, title, items, pageN) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);
  const cols = 2, cardW = 5.85, cardH = 2.15, gapX = 0.4, gapY = 0.3;
  const startX = 0.6, startY = 1.9;
  items.forEach((item, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    s.addShape(pres.ShapeType.roundRect, { x, y, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: CARD_BG }, line: { type: "none" } });
    s.addShape(pres.ShapeType.rect, { x: x + 0.28, y: y + 0.28, w: 0.5, h: 0.06, fill: { color: ACCENT }, line: { type: "none" } });
    s.addText(item.term, { x: x + 0.28, y: y + 0.4, w: cardW - 0.56, h: 0.4, fontFace: FONT_BODY, fontSize: 15.5, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
    s.addText(item.def, { x: x + 0.28, y: y + 0.85, w: cardW - 0.56, h: cardH - 1.05, fontFace: FONT_BODY, fontSize: 12, color: TEXT_DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  });
  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Challenges timeline ----------
function slideChallenges(kicker, title, items, pageN) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);
  const startX = 0.6, startY = 1.95, rowH = 1.24;
  s.addShape(pres.ShapeType.rect, { x: 1.05, y: startY + 0.1, w: 0.035, h: rowH * items.length - 0.2, fill: { color: CARD_BG }, line: { type: "none" } });
  items.forEach((item, i) => {
    const y = startY + i * rowH;
    s.addShape(pres.ShapeType.ellipse, { x: 0.85, y: y + 0.08, w: 0.4, h: 0.4, fill: { color: PROBLEM }, line: { type: "none", color: WHITE, width: 2 } });
    s.addText("✕", { x: 0.85, y: y + 0.08, w: 0.4, h: 0.4, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 13, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(item.problem, { x: 1.55, y: y, w: 5.1, h: rowH - 0.15, fontFace: FONT_BODY, fontSize: 12.5, color: TEXT_DARK, valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    s.addShape(pres.ShapeType.rightArrow, { x: 6.75, y: y + rowH / 2 - 0.35 - 0.11, w: 0.4, h: 0.22, fill: { color: ACCENT }, line: { type: "none" } });
    s.addShape(pres.ShapeType.roundRect, { x: 7.3, y: y, w: 5.4, h: rowH - 0.15, rectRadius: 0.06, fill: { color: CARD_BG }, line: { type: "none" } });
    s.addText(item.fix, { x: 7.55, y: y, w: 4.9, h: rowH - 0.15, fontFace: FONT_BODY, fontSize: 12.5, color: TEXT_DARK, valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  });
  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Stack + deploy summary ----------
function slideStack(kicker, title, stackItems, deployLines, resultLine, pageN) {
  const s = pres.addSlide();
  contentHeader(s, kicker, title);

  s.addText("STACK TÉCNICO", { x: 0.6, y: 1.85, w: 6, h: 0.3, fontFace: FONT_BODY, fontSize: 12, bold: true, color: SECONDARY, charSpacing: 1.5, isTextBox: true, margin: 0 });
  let y = 2.25;
  stackItems.forEach((item) => {
    s.addShape(pres.ShapeType.ellipse, { x: 0.6, y, w: 0.12, h: 0.12, fill: { color: ACCENT }, line: { type: "none" } });
    s.addText(item, { x: 0.85, y: y - 0.1, w: 5.7, h: 0.35, fontFace: FONT_BODY, fontSize: 13, color: TEXT_DARK, isTextBox: true, margin: 0 });
    y += 0.42;
  });

  s.addShape(pres.ShapeType.roundRect, { x: 6.95, y: 1.85, w: 5.75, h: 3.55, rectRadius: 0.08, fill: { color: DARK_BG }, line: { type: "none" } });
  s.addText("DEPLOY", { x: 7.25, y: 2.05, w: 5, h: 0.3, fontFace: FONT_BODY, fontSize: 12, bold: true, color: ACCENT, charSpacing: 1.5, isTextBox: true, margin: 0 });
  s.addText(deployLines, { x: 7.25, y: 2.4, w: 5.2, h: 2.9, fontFace: FONT_BODY, fontSize: 13, color: "E9FBF8", isTextBox: true, margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.75, w: 12.1, h: 0.95, rectRadius: 0.08, fill: { color: PRIMARY }, line: { type: "none" } });
  s.addText(resultLine, { x: 0.9, y: 5.75, w: 11.5, h: 0.95, valign: "middle", fontFace: FONT_BODY, fontSize: 14, bold: true, color: WHITE, isTextBox: true, margin: 0 });

  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, pageN);
}

// ---------- Lessons learned ----------
function slideLessons() {
  const s = pres.addSlide();
  contentHeader(s, "SÍNTESIS", "Lecciones que se repiten en ambos proyectos");
  const lessons = [
    ["Liviano en producción, pesado solo en desarrollo", "PyTorch/sentence-transformers entrenan y prueban local; ONNX Runtime / FastEmbed corren en el servidor. La regla: nada de torch en el deploy final."],
    ["Fija versiones, no confíes en \"latest\"", "Python 3.10 explícito, versiones exactas en requirements.txt. Los free tiers cambian de Python por defecto sin avisar, y eso rompe builds silenciosamente."],
    ["\"Compila\" no es \"funciona\"", "Los dos proyectos parecían listos tras el build — pero fallaban en runtime real (memoria, bug de esquema). Siempre se probó con una predicción/pregunta real contra el servidor desplegado."],
    ["Los free tiers tienen límites reales", "512MB de RAM, cold starts de 30-60s, cambios de precios de un día para otro (Hugging Face Spaces). Hay que diseñar para esas restricciones, no asumirlas."],
  ];
  let y = 1.95;
  const rowH = 1.28;
  lessons.forEach(([title, desc], i) => {
    s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 12.1, h: rowH - 0.15, rectRadius: 0.08, fill: { color: i % 2 === 0 ? CARD_BG : WHITE }, line: i % 2 === 0 ? { type: "none" } : { color: CARD_BG, width: 1.5 } });
    s.addShape(pres.ShapeType.ellipse, { x: 0.85, y: y + 0.19, w: 0.5, h: 0.5, fill: { color: ACCENT }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 0.85, y: y + 0.19, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 16, bold: true, color: DARK_BG, isTextBox: true, margin: 0 });
    s.addText(title, { x: 1.6, y: y + 0.1, w: 10.8, h: 0.35, fontFace: FONT_BODY, fontSize: 14.5, bold: true, color: TEXT_DARK, isTextBox: true, margin: 0 });
    s.addText(desc, { x: 1.6, y: y + 0.44, w: 10.8, h: rowH - 0.6, fontFace: FONT_BODY, fontSize: 11.5, color: TEXT_MUTED, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    y += rowH;
  });
  footer(s, "Guía de estudio · IA en Python");
  pageNum(s, 20);
}

// ---------- Closing ----------
function slideClosing() {
  const s = pres.addSlide();
  s.background = { color: DARK_BG };
  s.addShape(pres.ShapeType.ellipse, { x: 9.8, y: 4.5, w: 5.5, h: 5.5, fill: { color: SECONDARY, transparency: 89 }, line: { type: "none" } });
  s.addText("PRÓXIMOS PASOS", { x: 0.9, y: 1.6, w: 8, h: 0.4, fontFace: FONT_BODY, fontSize: 13, bold: true, color: ACCENT, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText("De 2 proyectos a un\nportafolio de IA completo", { x: 0.85, y: 2.05, w: 10.5, h: 1.7, fontFace: FONT_HEAD, fontSize: 32, bold: true, color: WHITE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.05 });
  const bullets = [
    "Agente con tool-use — el siguiente escalón natural después del RAG",
    "Sistema de recomendación — fundamentos de ML clásico",
    "Whisper (audio) y YOLO (detección) — nuevas modalidades",
    "Fine-tuning con LoRA — el más avanzado, para el final",
  ];
  let y = 4.0;
  bullets.forEach((b) => {
    s.addShape(pres.ShapeType.ellipse, { x: 0.9, y: y + 0.08, w: 0.1, h: 0.1, fill: { color: ACCENT }, line: { type: "none" } });
    s.addText(b, { x: 1.15, y, w: 9.5, h: 0.4, fontFace: FONT_BODY, fontSize: 15, color: "E9FBF8", isTextBox: true, margin: 0 });
    y += 0.55;
  });
  s.addText("Este portafolio de IA vivirá en un repo nuevo, separado, con un enlace al portafolio web actual.", {
    x: 0.9, y: 6.5, w: 10.5, h: 0.6, fontFace: FONT_BODY, fontSize: 12.5, italic: true, color: "9FC9C6", isTextBox: true, margin: 0,
  });
}

// ===================== BUILD DECK =====================
slideTitle();
slideAgenda();

// Project 1
slideSection("01", "Chatbot RAG", "Responde preguntas sobre mi experiencia, basado en mis propios documentos");
slideConcept(
  "PROYECTO 1 · CONCEPTO",
  "¿Por qué RAG y no solo un LLM?",
  "El problema",
  "Un LLM normal solo responde con lo que memorizó en su entrenamiento — no sabe nada sobre mí en particular. Preguntarle directamente \"¿quién es Nicolas?\" no da ninguna información real.",
  "La solución: RAG",
  "Retrieval-Augmented Generation: antes de responder, se le entrega al LLM un puñado de fragmentos relevantes de mis propios documentos, directamente en el prompt. El modelo no se re-entrena — solo lee el contexto en el momento.",
  3
);
slideDiagram(
  "PROYECTO 1 · ARQUITECTURA",
  "Dos fases separadas",
  "FASE A — Indexado (una vez, al arrancar)",
  ["Cargar\ndocumentos", "Trocear en\nchunks", "Generar\nembeddings", "Guardar en\nChroma"],
  "FASE B — Consulta (en cada mensaje)",
  ["Pregunta del\nusuario", "Buscar chunks\nsimilares", "Insertar en\nel prompt", "El LLM\nresponde"],
  "El vector store construido en la Fase A queda en memoria y se reutiliza en cada consulta de la Fase B — por eso indexar es \"una vez\" y consultar es \"cada mensaje\".",
  4
);
slideSteps(
  "PROYECTO 1 · FASE A",
  "Indexado: de documentos a vectores",
  [
    { title: "Cargar documentos (load_documents)", desc: "Recorre data/, usa TextLoader para .md/.txt y PyPDFLoader para .pdf. Hoy: un solo archivo con mi bio, experiencia y proyectos." },
    { title: "Trocear en chunks (RecursiveCharacterTextSplitter)", desc: "chunk_size=800, chunk_overlap=100. Chunks pequeños se embeben con más precisión que un documento gigante." },
    { title: "Generar embeddings (FastEmbedEmbeddings)", desc: "sentence-transformers/all-MiniLM-L6-v2 vía ONNX Runtime — corre local y gratis, sin PyTorch." },
    { title: "Guardar en Chroma (Chroma.from_documents)", desc: "Base de datos vectorial en memoria: guarda pares (texto, vector) y busca los más cercanos por similitud." },
  ],
  5
);
slideSteps(
  "PROYECTO 1 · FASE B",
  "Consulta: de pregunta a respuesta",
  [
    { title: "Embeber la pregunta y buscar (retriever.invoke)", desc: "Mismo modelo de embeddings; k=4 → los 4 chunks más similares semánticamente (no por palabras clave)." },
    { title: "Armar el prompt con el contexto", desc: "Los 4 chunks se insertan en el system prompt, que instruye al modelo a responder solo desde ahí." },
    { title: "El LLM genera la respuesta (llm.invoke)", desc: "ChatGroq, modelo openai/gpt-oss-20b, temperature=0.2 (respuestas factuales, no creativas)." },
  ],
  6,
  { label: "SYSTEM PROMPT (app.py)", code: "\"You are a helpful\nassistant answering\nquestions about\nNicolas... based\nonly on the context\nprovided below.\nIf the answer isn't\nin the context, say\nyou don't have that\ninformation instead\nof guessing.\n\nContext:\n{context}\"" }
);
slideGlossary(
  "PROYECTO 1 · CONCEPTOS CLAVE",
  "Glosario para no perderse",
  [
    { term: "Embedding", def: "Lista de ~384 números que representa el significado de un texto. Textos similares → vectores cercanos en ese espacio." },
    { term: "Chunking", def: "Trocear documentos largos en fragmentos pequeños antes de embeberlos, para que cada vector represente una idea, no un promedio borroso de todo." },
    { term: "Búsqueda semántica", def: "Buscar por significado, no por palabras exactas. \"¿A qué se dedica?\" puede hacer match con \"Ingeniero en Informática\"." },
    { term: "Grounding", def: "Anclar al modelo al contexto real vía el system prompt, para que admita no saber en vez de alucinar (inventar)." },
  ],
  7
);
slideChallenges(
  "PROYECTO 1 · TROUBLESHOOTING",
  "4 problemas reales en el camino al deploy",
  [
    { problem: "El modelo llama-3.1-8b-instant ya no existía en Groq (catálogo cambió)", fix: "Se listaron los modelos vigentes vía API y se cambió a openai/gpt-oss-20b" },
    { problem: "Out of memory en Render: sentence-transformers (PyTorch) pasaba los 512MB", fix: "Se cambió a FastEmbed (ONNX Runtime) — bajó de >512MB a ~330MB" },
    { problem: "fastembed==0.4.2 no existía en PyPI para el entorno de Render", fix: "Se corrigió a una versión real y publicada (0.7.3)" },
    { problem: "Conflicto de numpy: langchain pide <2.0, fastembed pide ≥2.1 en Python 3.12+", fix: "Se fijó Python 3.10.13 explícitamente en render.yaml" },
  ],
  8
);
slideStack(
  "PROYECTO 1 · RESUMEN",
  "Stack, deploy y resultado",
  [
    "LangChain — orquestación del pipeline RAG",
    "Chroma — vector store en memoria",
    "FastEmbed — embeddings vía ONNX Runtime",
    "Groq (openai/gpt-oss-20b) — generación de respuestas",
    "Gradio — interfaz de chat",
  ],
  "Plataforma: Render (free tier)\nBlueprint: render.yaml\nVariable: GROQ_API_KEY (secreta)\nPython fijado: 3.10.13",
  "✓ Verificado en producción: pregunta real sobre mi experiencia → respuesta correcta y bien fundamentada, citando proyectos reales del contexto.",
  9
);

// Project 2
slideSection("02", "Clasificador de Residuos", "Clasifica una foto en 6 categorías de reciclaje usando transfer learning");
slideConcept(
  "PROYECTO 2 · CONCEPTO",
  "¿Por qué transfer learning?",
  "El problema",
  "Entrenar una red de reconocimiento de imágenes desde cero requiere millones de fotos etiquetadas y días de cómputo en GPU — nada realista para un proyecto de práctica.",
  "La solución: transfer learning",
  "Partir de una red ya entrenada en un problema genérico gigante (1000 categorías, 1.2 millones de fotos) y solo reenseñarle la última capa con mis 6 categorías y unos cientos de imágenes.",
  11
);
slideDiagram(
  "PROYECTO 2 · ARQUITECTURA",
  "Entrenamiento vs. inferencia",
  "FASE A — Entrenamiento (train.py, una vez, en mi máquina)",
  ["Descargar\ndataset", "Red\npre-entrenada", "Congelar y\nreemplazar capa", "Exportar\na ONNX"],
  "FASE B — Inferencia (app.py, en cada foto subida)",
  ["Usuario sube\nuna foto", "Preprocesar\nimagen", "Pasar por\nla red", "Softmax →\npredicción"],
  "El modelo entrenado (waste_classifier.onnx) es el único puente entre las dos fases — por eso el servidor de producción nunca necesita instalar PyTorch.",
  12
);
slideSteps(
  "PROYECTO 2 · FASE A (parte 1)",
  "Dataset balanceado y modelo base",
  [
    { title: "Dataset desbalanceado → muestreo parejo", desc: "garythung/trashnet tiene clases muy dispares (ej. \"trash\" mucho menor que \"paper\"). Se tomaron 150 imágenes por clase (900 en total) para evitar que el modelo aprenda a favorecer las clases más comunes." },
    { title: "MobileNetV2 pre-entrenado (ImageNet)", desc: "Arquitectura pequeña y rápida, pensada para celulares. Ya sabe detectar bordes, texturas y formas genéricas." },
    { title: "Congelar model.features", desc: "requires_grad = False en todas las capas convolucionales — ese conocimiento visual genérico no se toca." },
    { title: "Reemplazar model.classifier[1]", desc: "Nueva capa final con 6 salidas (antes 1000, las clases de ImageNet). Solo esta capa se entrena." },
  ],
  13
);
slideSteps(
  "PROYECTO 2 · FASE A (parte 2)",
  "Entrenamiento y exportación",
  [
    { title: "Bucle de entrenamiento (5 épocas)", desc: "Cada época: predecir → medir error (cross-entropy) → ajustar solo los pesos de la capa final. Accuracy: 50% (época 1) → 78% (época 5)." },
    { title: "Data augmentation (RandomHorizontalFlip)", desc: "Voltear imágenes al azar durante el entrenamiento para generalizar mejor con pocos datos." },
    { title: "Exportar a ONNX (torch.onnx.export)", desc: "Formato estándar ejecutable con ONNX Runtime, sin necesitar PyTorch instalado en el servidor de producción." },
  ],
  14,
  { label: "RESULTADO DEL ENTRENAMIENTO", code: "Epoch 1/5\n  loss 1.36 - acc 49.8%\nEpoch 2/5\n  loss 0.88 - acc 68.7%\nEpoch 3/5\n  loss 0.71 - acc 75.3%\nEpoch 4/5\n  loss 0.67 - acc 77.9%\nEpoch 5/5\n  loss 0.62 - acc 77.9%\n\nExportado:\nwaste_classifier.onnx\n(8.9 MB)" }
);
slideSteps(
  "PROYECTO 2 · FASE B",
  "Inferencia: la parte más delicada",
  [
    { title: "Cargar el modelo ONNX (una sola vez)", desc: "InferenceSession al arrancar el servidor — igual que el vectorstore del chatbot." },
    { title: "Preprocesar EXACTO igual que en el entrenamiento", desc: "Mismo tamaño (160×160), mismo rango (÷255), misma normalización (mean/std de ImageNet), mismo orden de ejes (HWC→CHW). Si algo no calza, el modelo no falla — predice mal en silencio." },
    { title: "Softmax convierte logits en probabilidades", desc: "6 números crudos del modelo se convierten en 6 probabilidades que suman 100%, listas para mostrar en gr.Label." },
  ],
  15,
  { label: "SOFTMAX EN LA PRÁCTICA", code: "logits crudos:\n  glass:     2.3\n  metal:     1.8\n  cardboard: -0.5\n\n↓ softmax ↓\n\nprobabilidades:\n  glass:     77%\n  metal:     16%\n  cardboard:  2%\n\n(siempre suman 100%,\nincluso si la foto no\nes ninguna categoría\nconocida)" }
);
slideGlossary(
  "PROYECTO 2 · CONCEPTOS CLAVE",
  "Glosario para no perderse",
  [
    { term: "Transfer learning", def: "Reutilizar una red ya entrenada en un problema grande y genérico, adaptando solo la parte final a un problema nuevo y específico." },
    { term: "Capas congeladas", def: "Pesos que no se actualizan durante el entrenamiento (requires_grad=False) — se preserva el conocimiento visual ya aprendido." },
    { term: "Logits vs. probabilidades", def: "Los logits son la salida cruda del modelo (cualquier magnitud); softmax los convierte en probabilidades que suman 100%." },
    { term: "ONNX Runtime", def: "Motor de inferencia liviano que ejecuta modelos exportados sin necesitar la librería original (PyTorch) — clave para deploys con poca RAM." },
  ],
  16
);
slideChallenges(
  "PROYECTO 2 · TROUBLESHOOTING",
  "2 problemas reales en el camino al deploy",
  [
    { problem: "datasets con streaming=True se caía con MemoryError decodificando ciertas imágenes", fix: "Se cambió a descarga normal + muestreo balanceado manual sobre el dataset completo" },
    { problem: "gradio==5.7.1 crasheaba generando el schema de la API para gr.Label", fix: "Se actualizó a gradio==5.50.0, donde el bug ya estaba corregido" },
  ],
  17
);
slideStack(
  "PROYECTO 2 · RESUMEN",
  "Stack, deploy y resultado",
  [
    "PyTorch + torchvision — entrenamiento (solo local/dev)",
    "MobileNetV2 — arquitectura base (ImageNet)",
    "Hugging Face datasets — garythung/trashnet",
    "ONNX Runtime — inferencia en producción",
    "Gradio — interfaz de subida de foto",
  ],
  "Plataforma: Render (free tier)\nBlueprint: render.yaml\nSin variables secretas —\ntodo corre local, sin API externa\nPython fijado: 3.10.13",
  "✓ Verificado en producción: fotos reales de prueba → \"cardboard\" 99.6%, \"glass\" 76.7%, \"paper\" 94.0% de confianza.",
  18
);

slideLessons();
slideClosing();

pres.writeFile({ fileName: "dos-proyectos-de-ia.pptx" }).then(() => console.log("done"));
