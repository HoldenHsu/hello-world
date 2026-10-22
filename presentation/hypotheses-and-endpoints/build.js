const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Fi = require("react-icons/fi");
const sharp = require("sharp");
const fs = require("fs");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");

// ---------- theme ----------
const THEME = {
  name: "Clinical Trials Lecture",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0A1723", lt1: "FFFFFF", dk2: "0F3344", lt2: "EEF4F6",
    accent1: "156878", accent2: "2E6FCC", accent3: "D0452F",
    accent4: "6E8290", accent5: "E8A33D", accent6: "496474",
    hlink: "156878", folHlink: "496474",
  },
};
const HX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Yao-Chun Hsu";
pres.title = "Research Hypotheses and Study Endpoints";
const C = pres.SchemeColor;
const NAVY = C.text1, DEEP = C.text2, WHITE = C.background1, TINT = C.background2;
const TEAL = C.accent1, BLUE = C.accent2, RED = C.accent3, GRAY = C.accent4, AMBER = C.accent5, SLATE = C.accent6;

const W = 13.333, H = 7.5, L = 0.75, CW = W - 2 * L; // content width 11.83
const FOOT = "CLINICAL TRIALS  /  HYPOTHESES AND ENDPOINTS";

// ---------- icons ----------
const iconCache = {};
async function icon(name, hex) {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  const Comp = Fi[name];
  if (!Comp) throw new Error("no icon " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256, strokeWidth: 2 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}
// icon in a colored circle
async function iconCircle(slide, name, x, y, d, circleColor, iconHex, label) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: circleColor }, line: { color: circleColor, width: 0 }, objectName: label + " circle" });
  const data = await icon(name, iconHex);
  const pad = d * 0.25;
  slide.addImage({ data, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: label + " icon" });
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE", background: { color: HX.lt1 },
  objects: [],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: HX.lt1 },
  objects: [
    { placeholder: { options: { name: "label", type: "body", x: L, y: 0.32, w: CW, h: 0.3, fontSize: 11, bold: true, color: TEAL, margin: 0, valign: "middle" }, text: "SECTION" } },
    { placeholder: { options: { name: "title", type: "title", x: L, y: 0.66, w: CW, h: 0.75, fontSize: 32, bold: true, color: NAVY, margin: 0, valign: "middle", align: "left" }, text: "Title" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: L, y: 1.45, w: CW, h: 0.4, fontSize: 15, color: SLATE, margin: 0, valign: "middle" }, text: "" } },
    { line: { x: L, y: 2.0, w: CW, h: 0, line: { color: "D5DCE0", width: 0.75 } } },
    { text: { text: FOOT, options: { x: L, y: 7.05, w: 8, h: 0.25, fontSize: 9, color: GRAY, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 12.0, y: 7.05, w: 0.58, h: 0.25, fontSize: 9, color: GRAY, align: "right" },
});

// ---------- helpers ----------
let sectionName = "";
function section(name) { sectionName = name; pres.addSection({ title: name }); }
function slideC({ label, title, subtitle, takeaway, source, notes }) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sectionName });
  s.addText(label.toUpperCase(), { placeholder: "label" });
  s.addText(title, { placeholder: "title" });
  if (subtitle) s.addText(subtitle, { placeholder: "subtitle" });
  if (takeaway) {
    s.addShape(pres.ShapeType.rect, { x: L, y: 6.3, w: CW, h: 0.5, fill: { color: TINT }, line: { color: TINT, width: 0 }, objectName: "Takeaway bar" });
    s.addText(takeaway, { x: L + 0.18, y: 6.3, w: CW - 0.36, h: 0.5, fontSize: 14, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: "Takeaway" });
  }
  if (source) s.addText(source, { x: L, y: 6.84, w: CW, h: 0.2, fontSize: 9.5, color: GRAY, margin: 0, valign: "middle", isTextBox: true, objectName: "Source" });
  if (notes) s.addNotes(notes);
  return s;
}
function txt(s, text, o) {
  return s.addText(text, Object.assign({ fontSize: 15, color: NAVY, margin: 0, isTextBox: true, valign: "top" }, o));
}
function bullets(s, items, o) {
  const arr = items.map((t, i) => (typeof t === "string" ? { text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 6 } } : t));
  return s.addText(arr, Object.assign({ fontSize: 15, color: NAVY, margin: 0, isTextBox: true, valign: "top" }, o));
}
function card(s, x, y, w, h, fillColor, name) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fillColor }, line: { color: fillColor, width: 0 }, objectName: name });
}
function box(s, text, x, y, w, h, fillColor, textColor, o) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fillColor }, line: { color: fillColor, width: 0 }, objectName: "Box " + text });
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 15, bold: true, color: textColor, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: "Box text " + text }, o));
}
function arrow(s, x1, y1, x2, y2, color) {
  // right-pointing horizontal arrow
  s.addShape(pres.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: color || HX.accent4, width: 1.5, endArrowType: "triangle" }, objectName: "Arrow" });
}
const tblHead = (t) => ({ text: t, options: { bold: true, color: HX.lt1, fill: { color: HX.dk2 }, fontSize: 13 } });
const tblCell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 13, color: HX.dk1 }, o || {}) });
function table(s, rows, o) {
  const body = rows.map((r, ri) => r.map((c, ci) => {
    if (ri === 0) return tblHead(c);
    const opt = { fill: { color: ri % 2 === 1 ? HX.lt2 : HX.lt1 } };
    if (ci === 0) opt.bold = true;
    return tblCell(c, opt);
  }));
  s.addTable(body, Object.assign({ x: L, w: CW, border: { type: "solid", color: "FFFFFF", pt: 1 }, margin: 0.07, valign: "middle", fontFace: "Calibri" }, o));
}
const chartBase = {
  catAxisLabelColor: HX.accent6, valAxisLabelColor: HX.accent6, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: HX.dk1,
  valGridLine: { color: "D5DCE0", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, catAxisTitleFontFace: "+mn-lt", valAxisTitleFontFace: "+mn-lt",
  catAxisTitleColor: HX.dk1, valAxisTitleColor: HX.dk1, catAxisTitleFontSize: 12, valAxisTitleFontSize: 12, catAxisLineColor: "D5DCE0", valAxisLineColor: "D5DCE0",
};

(async () => {
  // =============== TITLE ===============
  section("Introduction");
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: sectionName });
    s.addImage({ path: "cover.jpg", x: 7.67, y: 0, w: 5.667, h: 7.5, sizing: { type: "cover", w: 5.667, h: 7.5 }, objectName: "Cover photo" });
    txt(s, "GRADUATE INSTITUTE OF MEDICINE  ·  I-SHOU UNIVERSITY", { x: 0.67, y: 0.45, w: 6.6, h: 0.3, fontSize: 12, bold: true, color: TEAL });
    txt(s, "Design of\nClinical Trials", { x: 0.67, y: 1.05, w: 6.6, h: 1.7, fontSize: 44, bold: true, color: NAVY, lineSpacingMultiple: 0.95 });
    txt(s, "Principles III", { x: 0.67, y: 2.8, w: 6.6, h: 0.45, fontSize: 24, bold: true, color: TEAL });
    txt(s, "Research hypotheses\nand study endpoints", { x: 0.67, y: 3.5, w: 6.6, h: 0.9, fontSize: 22, bold: true, color: NAVY, lineSpacingMultiple: 1.0 });
    card(s, 0.6, 4.6, 6.5, 1.95, HX.lt2, "Author card");
    s.addImage({ path: "headshot.jpg", x: 0.82, y: 4.8, w: 1.35, h: 1.55, sizing: { type: "cover", w: 1.35, h: 1.55 }, rounding: false, objectName: "Headshot" });
    txt(s, "Yao-Chun Hsu", { x: 2.4, y: 4.8, w: 4.5, h: 0.4, fontSize: 22, bold: true, color: NAVY });
    txt(s, "MD, MSc, PhD, FAASLD", { x: 2.4, y: 5.2, w: 4.5, h: 0.3, fontSize: 14, bold: true, color: TEAL });
    txt(s, "Graduate Institute of Medicine & School of Medicine, I-Shou University\nDepartment of Medical Research & Clinical Trial Center, E-Da Hospital", { x: 2.4, y: 5.55, w: 4.5, h: 0.85, fontSize: 11, bold: true, color: SLATE });
    txt(s, "FIRST SEMESTER  ·  ACADEMIC YEAR 115  ·  2026–2027", { x: 0.67, y: 6.85, w: 6.6, h: 0.3, fontSize: 11, bold: true, color: GRAY });
    s.addNotes("Third session on principles of clinical trial design. Today: how to turn a clinical question into a testable hypothesis, and how to choose, define, and manage the endpoints that will answer it.");
  }

  // =============== OBJECTIVES ===============
  {
    const s = slideC({ label: "Learning objectives", title: "Four goals for this session", takeaway: "A hypothesis, its endpoint, and its analysis are designed together, not one after another." });
    const items = [
      ["FiTarget", "Formulate", "Turn a clinical problem into a precise, testable hypothesis with explicit H₀ and H₁."],
      ["FiFlag", "Select", "Explain the roles of primary, secondary, and exploratory endpoints and choose an appropriate one."],
      ["FiLayers", "Manage", "Handle multiple endpoints: multiplicity adjustment, hierarchical testing, composite and co-primary endpoints."],
      ["FiCheckCircle", "Apply", "Apply good practice: prespecification, clinical relevance, and transparent documentation."],
    ];
    const cw = 2.78, gap = 0.24;
    for (let i = 0; i < 4; i++) {
      const x = L + i * (cw + gap);
      card(s, x, 2.35, cw, 3.6, HX.lt2, "Objective card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.25, 2.6, 0.7, HX.accent1, "FFFFFF", "Objective " + (i + 1));
      txt(s, items[i][1], { x: x + 0.25, y: 3.5, w: cw - 0.5, h: 0.4, fontSize: 20, bold: true, color: TEAL });
      txt(s, items[i][2], { x: x + 0.25, y: 3.95, w: cw - 0.5, h: 1.8, fontSize: 14, color: NAVY });
    }
  }

  // =============== FROM QUESTION TO HYPOTHESIS ===============
  section("Research hypotheses");
  {
    const s = slideC({ label: "Research hypotheses", title: "From clinical problem to testable hypothesis", subtitle: "The question narrows at each step; the hypothesis is the last and most precise statement.", takeaway: "A hypothesis you cannot measure is a wish, not a design." });
    const steps = [
      ["FiSearch", "Clinical problem", "Broad uncertainty seen in practice"],
      ["FiBookOpen", "Literature review", "Identify the knowledge gap"],
      ["FiTarget", "Objective(s)", "Primary and secondary aims"],
      ["FiEdit3", "Hypothesis", "H₀ versus H₁, measurable"],
      ["FiActivity", "Endpoint", "What is measured, when, how"],
    ];
    const bw = 2.05, gap = 0.395;
    for (let i = 0; i < 5; i++) {
      const x = L + i * (bw + gap);
      const fill = i === 3 ? HX.accent1 : HX.dk2;
      s.addShape(pres.ShapeType.roundRect, { x, y: 2.45, w: bw, h: 1.0, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: "Step " + (i + 1) });
      await iconCircle(s, steps[i][0], x + 0.15, 2.6, 0.7, i === 3 ? HX.dk2 : HX.accent1, "FFFFFF", "Step " + (i + 1));
      txt(s, steps[i][1], { x: x + 0.95, y: 2.45, w: bw - 1.0, h: 1.0, fontSize: 14, bold: true, color: WHITE, valign: "middle" });
      txt(s, steps[i][2], { x, y: 3.6, w: bw, h: 0.6, fontSize: 13, color: SLATE });
      if (i < 4) arrow(s, x + bw + 0.05, 2.95, x + bw + gap - 0.05, 2.95);
    }
    txt(s, "Broad", { x: L, y: 4.35, w: 2, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    txt(s, "Narrow", { x: W - L - 2, y: 4.35, w: 2, h: 0.3, fontSize: 12, bold: true, color: GRAY, align: "right" });
    s.addShape(pres.ShapeType.line, { x: L + 0.7, y: 4.5, w: CW - 1.6, h: 0, line: { color: HX.accent4, width: 1, endArrowType: "triangle" }, objectName: "Broad to narrow" });
    txt(s, [
      { text: "Example  ", options: { bold: true, color: TEAL } },
      { text: "“Does hydroxychloroquine help in COVID-19?”  →  “Among adults within 4 days of a high-risk exposure, does hydroxychloroquine reduce symptomatic, laboratory-confirmed COVID-19 by day 14 compared with placebo?”" },
    ], { x: L, y: 4.95, w: CW, h: 1.2, fontSize: 15, color: NAVY });
  }

  // =============== PICO ===============
  {
    const s = slideC({ label: "Research hypotheses", title: "PICOT: the anatomy of an answerable question", subtitle: "Every element of the question reappears later as an element of the design.", takeaway: "If one PICOT element is vague, the hypothesis, the endpoint, or the sample size will be vague too." });
    const rows = [
      ["P", "Population", "Who is studied; eligibility criteria", "Adults ≤4 days after household or occupational exposure to SARS-CoV-2"],
      ["I", "Intervention", "The test treatment, dose, duration", "Hydroxychloroquine, 5-day course"],
      ["C", "Comparator", "Placebo, active control, standard care", "Matching placebo (folate)"],
      ["O", "Outcome", "The endpoint and its measurement", "Symptomatic illness compatible with COVID-19 or PCR-confirmed infection"],
      ["T", "Time", "When the outcome is assessed", "Within 14 days of enrollment"],
    ];
    const y0 = 2.5, rh = 0.72;
    txt(s, "Element", { x: L + 0.85, y: y0 - 0.35, w: 2, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    txt(s, "Design content", { x: L + 3.2, y: y0 - 0.35, w: 3, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    txt(s, "NCT04308668 example", { x: L + 7.0, y: y0 - 0.35, w: 4, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    rows.forEach((r, i) => {
      const y = y0 + i * rh;
      s.addShape(pres.ShapeType.ellipse, { x: L, y: y + 0.07, w: 0.56, h: 0.56, fill: { color: HX.accent1 }, line: { color: HX.accent1, width: 0 }, objectName: "PICOT " + r[0] });
      txt(s, r[0], { x: L, y: y + 0.07, w: 0.56, h: 0.56, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, r[1], { x: L + 0.85, y, w: 2.2, h: 0.7, fontSize: 17, bold: true, color: NAVY, valign: "middle" });
      txt(s, r[2], { x: L + 3.2, y, w: 3.6, h: 0.7, fontSize: 14, color: SLATE, valign: "middle" });
      txt(s, r[3], { x: L + 7.0, y, w: 4.8, h: 0.7, fontSize: 14, color: NAVY, valign: "middle" });
      if (i < rows.length - 1) s.addShape(pres.ShapeType.line, { x: L + 0.85, y: y + rh, w: CW - 0.85, h: 0, line: { color: "E3E8EB", width: 0.75 }, objectName: "Row line " + i });
    });
    s.addNotes("Boulware DR et al. NEJM 2020;383:517-525. Use the example to show students how each PICOT element maps to a protocol section.");
  }

  // =============== H0 / H1 ===============
  {
    const s = slideC({ label: "Research hypotheses", title: "Null and alternative hypotheses", subtitle: "The trial is designed to reject H₀; it never proves H₁ and never proves H₀.", takeaway: "Write the hypothesis on the scale of the endpoint, with the direction of benefit stated explicitly." });
    card(s, L, 2.3, 5.75, 1.95, HX.lt2, "H0 card");
    txt(s, "H₀  Null hypothesis", { x: L + 0.3, y: 2.45, w: 5.2, h: 0.4, fontSize: 20, bold: true, color: RED });
    txt(s, "The treatment has no effect (or no effect beyond the margin). It is the position the data must argue against.", { x: L + 0.3, y: 2.95, w: 5.2, h: 0.8, fontSize: 15, color: NAVY });
    txt(s, "Example  H₀: incidence of COVID-19 with HCQ = incidence with placebo", { x: L + 0.3, y: 3.75, w: 5.2, h: 0.55, fontSize: 14, italic: true, color: SLATE });
    card(s, L + 6.08, 2.3, 5.75, 1.95, HX.lt2, "H1 card");
    txt(s, "H₁  Alternative hypothesis", { x: L + 6.38, y: 2.45, w: 5.2, h: 0.4, fontSize: 20, bold: true, color: TEAL });
    txt(s, "The treatment has an effect. It is what the investigator hopes to demonstrate, and it determines the sample size.", { x: L + 6.38, y: 2.95, w: 5.2, h: 0.8, fontSize: 15, color: NAVY });
    txt(s, "Example  H₁: incidence with HCQ ≠ (or <) incidence with placebo", { x: L + 6.38, y: 3.75, w: 5.2, h: 0.55, fontSize: 14, italic: true, color: SLATE });
    // link to Principles II
    txt(s, "Link to Principles II", { x: L, y: 4.4, w: 4, h: 0.35, fontSize: 13, bold: true, color: TEAL });
    table(s, [
      ["Claim", "H₀", "H₁", "Evidence required"],
      ["Superiority", "d ≤ 0", "d > 0", "Lower CI bound > 0"],
      ["Non-inferiority", "d ≤ −Δ", "d > −Δ", "Lower CI bound > −Δ"],
      ["Equivalence", "d ≤ −Δ or d ≥ +Δ", "−Δ < d < +Δ", "Entire CI within (−Δ, +Δ)"],
    ], { y: 4.72, colW: [2.6, 3.0, 2.8, 3.43], rowH: 0.28, fontSize: 12 });
  }

  // =============== TESTABLE HYPOTHESIS ===============
  {
    const s = slideC({ label: "Research hypotheses", title: "What makes a hypothesis testable", subtitle: "Practical checks before the protocol is written", takeaway: "Statistical input belongs at the hypothesis stage, not after the data arrive." });
    const items = [
      ["FiCrosshair", "Specific", "One population, one comparison, one primary outcome, one time point."],
      ["FiBarChart2", "Measurable", "The outcome can be observed with a defined instrument and a defined rule for success."],
      ["FiTrendingUp", "Directional", "States whether the aim is superiority, non-inferiority, or equivalence, with any margin."],
      ["FiClock", "Feasible", "Event rates, recruitment, follow-up, and budget allow the required sample size."],
      ["FiShield", "Ethical and lawful", "Clinical equipoise exists; the design meets regulatory and IRB expectations."],
      ["FiFileText", "Documented", "Written in the protocol and statistical analysis plan, and registered before enrollment."],
    ];
    const cw = 3.78, ch = 1.75, gx = 0.245, gy = 0.25;
    for (let i = 0; i < 6; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = L + col * (cw + gx), y = 2.3 + row * (ch + gy);
      card(s, x, y, cw, ch, HX.lt2, "Check card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.22, y + 0.22, 0.6, HX.accent1, "FFFFFF", "Check " + (i + 1));
      txt(s, items[i][1], { x: x + 0.98, y: y + 0.22, w: cw - 1.15, h: 0.6, fontSize: 17, bold: true, color: NAVY, valign: "middle" });
      txt(s, items[i][2], { x: x + 0.22, y: y + 0.95, w: cw - 0.44, h: 0.75, fontSize: 13.5, color: SLATE });
    }
  }

  // =============== PITFALLS ===============
  {
    const s = slideC({ label: "Research hypotheses", title: "Common pitfalls and their fixes", takeaway: "Each weak hypothesis can be repaired by adding the missing PICOT element." });
    const rows = [
      ["Overly broad", "“Drug X improves outcomes in liver disease.”", "Specific", "“Drug X reduces 48-week fibrosis progression in compensated MASH cirrhosis.”"],
      ["Ambiguous language", "“Patients feel better on treatment.”", "Clear", "“Mean change in the FACIT-F fatigue score at week 24 is greater with treatment.”"],
      ["Unfeasible objective", "“Drug X lowers 20-year mortality.”", "Reasonable", "“Drug X lowers 3-year liver-related events, a validated surrogate pathway.”"],
      ["Ethically questionable", "“Placebo for patients with an effective standard of care.”", "Lawful", "“Add-on to standard care versus standard care alone.”"],
    ];
    const y0 = 2.5, rh = 0.9;
    txt(s, "Pitfall", { x: L, y: y0 - 0.38, w: 5, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    txt(s, "Fix", { x: L + 6.3, y: y0 - 0.38, w: 5, h: 0.3, fontSize: 12, bold: true, color: GRAY });
    rows.forEach((r, i) => {
      const y = y0 + i * rh;
      card(s, L, y, 5.45, rh - 0.15, "FBEDEA", "Pitfall " + (i + 1));
      txt(s, r[0], { x: L + 0.2, y: y + 0.08, w: 5.1, h: 0.3, fontSize: 14, bold: true, color: RED });
      txt(s, r[1], { x: L + 0.2, y: y + 0.38, w: 5.1, h: 0.4, fontSize: 13, italic: true, color: NAVY });
      s.addShape(pres.ShapeType.rightArrow, { x: L + 5.6, y: y + 0.2, w: 0.55, h: 0.4, fill: { color: HX.accent4 }, line: { color: HX.accent4, width: 0 }, objectName: "Fix arrow " + (i + 1) });
      card(s, L + 6.3, y, 5.53, rh - 0.15, "E6F2F3", "Fix " + (i + 1));
      txt(s, r[2], { x: L + 6.5, y: y + 0.08, w: 5.2, h: 0.3, fontSize: 14, bold: true, color: TEAL });
      txt(s, r[3], { x: L + 6.5, y: y + 0.38, w: 5.2, h: 0.4, fontSize: 13, italic: true, color: NAVY });
    });
  }

  // =============== WORKED EXAMPLE HCQ ===============
  {
    const s = slideC({ label: "Clinical example", title: "Hydroxychloroquine post-exposure prophylaxis", subtitle: "NCT04308668: a randomized, double-blind, placebo-controlled trial", takeaway: "Two objectives, two hypotheses, two endpoints: each hypothesis has its own measurable outcome.", source: "Boulware DR et al. N Engl J Med 2020;383:517–525" });
    card(s, L, 2.3, CW, 1.25, HX.lt2, "Hypothesis quote");
    txt(s, "“We hypothesized that hydroxychloroquine could potentially be used as postexposure prophylaxis, to prevent symptomatic infection after exposure to Covid-19.”", { x: L + 0.25, y: 2.3, w: CW - 0.5, h: 1.25, fontSize: 16, italic: true, color: NAVY, valign: "middle" });
    const cols = [
      ["Objective 1", "Post-exposure prophylaxis", "Prevent symptomatic COVID-19 after a known high-risk exposure", "Primary endpoint: incidence of laboratory-confirmed COVID-19 or compatible illness within 14 days"],
      ["Objective 2", "Early preemptive therapy", "Prevent disease progression in persons already symptomatic", "Endpoints: hospitalization, symptom severity (0–10 visual analog scale) over 14 days"],
    ];
    cols.forEach((c, i) => {
      const x = L + i * (5.9 + 0.03);
      txt(s, c[0], { x, y: 3.8, w: 5.6, h: 0.3, fontSize: 12, bold: true, color: GRAY });
      txt(s, c[1], { x, y: 4.1, w: 5.6, h: 0.4, fontSize: 19, bold: true, color: TEAL });
      txt(s, c[2], { x, y: 4.55, w: 5.6, h: 0.6, fontSize: 14, color: NAVY });
      txt(s, c[3], { x, y: 5.2, w: 5.6, h: 0.9, fontSize: 14, bold: true, color: NAVY });
    });
    s.addShape(pres.ShapeType.line, { x: L + 5.75, y: 3.85, w: 0, h: 2.25, line: { color: "D5DCE0", width: 0.75 }, objectName: "Divider" });
  }

  // =============== ENDPOINT DEFINITION ===============
  section("Study endpoints");
  {
    const s = slideC({ label: "Study endpoints", title: "An endpoint is the hypothesis made measurable", subtitle: "A specific outcome, measured in a defined way at a defined time, used to judge the effect of the intervention", takeaway: "The primary endpoint fixes the analysis and the sample size; change it and the whole design changes." });
    const chain = [["Question", HX.dk2], ["Hypothesis", HX.dk2], ["Endpoint", HX.accent1], ["Analysis", HX.dk2], ["Sample size", HX.dk2]];
    const bw = 2.0, gap = 0.4575;
    chain.forEach((c, i) => {
      const x = L + i * (bw + gap);
      box(s, c[0], x, 2.4, bw, 0.8, c[1], WHITE, { fontSize: 16 });
      if (i < 4) arrow(s, x + bw + 0.06, 2.8, x + bw + gap - 0.06, 2.8);
    });
    // three considerations
    const cons = [
      ["FiHeart", "Clinical relevance", "Reflects how a patient feels, functions, or survives, or a validated surrogate of it."],
      ["FiBarChart2", "Measurability", "Objective, reproducible, with a validated instrument and a prespecified rule for success."],
      ["FiClock", "Feasibility and acceptability", "Can be observed in the available time and population, and is acceptable to patients and regulators."],
    ];
    const cw = 3.78, gx = 0.245;
    for (let i = 0; i < 3; i++) {
      const x = L + i * (cw + gx);
      card(s, x, 3.6, cw, 2.5, HX.lt2, "Consideration " + (i + 1));
      await iconCircle(s, cons[i][0], x + 0.25, 3.85, 0.6, HX.accent1, "FFFFFF", "Consideration " + (i + 1));
      txt(s, cons[i][1], { x: x + 0.25, y: 4.6, w: cw - 0.5, h: 0.4, fontSize: 17, bold: true, color: NAVY });
      txt(s, cons[i][2], { x: x + 0.25, y: 5.05, w: cw - 0.5, h: 1.0, fontSize: 13.5, color: SLATE });
    }
  }

  // =============== HIERARCHY ===============
  {
    const s = slideC({ label: "Study endpoints", title: "Three tiers of endpoints", subtitle: "The tier determines the statistical rigor, not the clinical interest", takeaway: "Only the primary endpoint carries the confirmatory claim; the others support, explain, or generate hypotheses." });
    // pyramid on left
    const px = L + 0.2, py = 2.35, pw = 4.6, ph = 3.7;
    const tiers = [["Primary", HX.dk2, 1], ["Secondary", HX.accent1, 2], ["Exploratory", HX.accent6, 3]];
    tiers.forEach((t, i) => {
      const h = ph / 3, y = py + i * h, wfrac = (i + 1) / 3;
      const w = pw * (0.35 + 0.65 * wfrac * 0.95);
      const x = px + (pw - w) / 2;
      s.addShape(pres.ShapeType.trapezoid, { x, y, w, h: h - 0.06, fill: { color: t[1] }, line: { color: t[1], width: 0 }, flipV: false, objectName: "Tier " + t[0] });
      txt(s, t[0], { x, y, w, h: h - 0.06, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
    });
    // descriptions on right
    const desc = [
      ["Primary", TEAL, "The main outcome of interest. Basis for the hypothesis test, the success criterion, and the sample size. Usually one."],
      ["Secondary", TEAL, "Prespecified additional outcomes: support the primary result, describe other effects, or characterize safety. Usually several; formal claims need multiplicity control."],
      ["Exploratory", TEAL, "Hypothesis-generating. Not powered, flexible, may be post hoc. Findings are preliminary and not a basis for regulatory claims."],
    ];
    desc.forEach((d, i) => {
      const y = 2.35 + i * 1.25;
      txt(s, d[0], { x: L + 5.4, y, w: 6.4, h: 0.35, fontSize: 17, bold: true, color: d[1] });
      txt(s, d[2], { x: L + 5.4, y: y + 0.38, w: 6.4, h: 0.8, fontSize: 13.5, color: NAVY });
    });
    txt(s, "▲ Statistical rigor", { x: L + 0.2, y: 6.05, w: 2.3, h: 0.25, fontSize: 11, bold: true, color: GRAY });
    txt(s, "Number of endpoints ▼", { x: L + 2.5, y: 6.05, w: 2.3, h: 0.25, fontSize: 11, bold: true, color: GRAY, align: "right" });
  }

  // =============== COMPARISON TABLE ===============
  {
    const s = slideC({ label: "Study endpoints", title: "Primary, secondary, and exploratory endpoints compared", takeaway: "Secondary endpoints can support a label claim only if they are prespecified and multiplicity is controlled." });
    table(s, [
      ["Aspect", "Primary endpoint", "Secondary endpoints", "Exploratory endpoints"],
      ["Purpose", "Test the main hypothesis; answer the primary objective", "Evaluate additional prespecified effects", "Generate new hypotheses; explore further questions"],
      ["Statistical power", "Trial is powered to detect the targeted difference", "May be powered, but usually to a lesser extent", "Not powered; analyses are often descriptive"],
      ["Prespecification", "Clearly defined before the trial begins", "Defined before the trial; analyses planned", "May be prespecified or identified during analysis"],
      ["Regulatory impact", "Central to approval and labeling claims", "May support claims; provide additional information", "Generally not used for regulatory decisions"],
      ["Multiplicity", "Adjustment required if more than one", "Adjustment needed for formal claims", "Typically no adjustment; findings are preliminary"],
    ], { y: 2.3, colW: [2.2, 3.21, 3.21, 3.21], rowH: 0.55 });
  }

  // =============== EXAMPLE OUTCOME MEASURES ===============
  {
    const s = slideC({ label: "Clinical example", title: "Registered outcome measures: NCT04308668", subtitle: "As listed on ClinicalTrials.gov, with time frame", takeaway: "Registration before enrollment makes the primary endpoint publicly verifiable.", source: "ClinicalTrials.gov NCT04308668; Boulware DR et al. N Engl J Med 2020;383:517–525" });
    txt(s, "Primary outcome measures", { x: L, y: 2.2, w: 6, h: 0.3, fontSize: 13, bold: true, color: TEAL });
    table(s, [
      ["Outcome measure", "Description", "Time frame"],
      ["COVID-19 disease at day 14 (asymptomatic at baseline)", "Participants with active, confirmed or compatible disease", "14 days"],
      ["Change in disease severity (symptomatic at baseline)", "Visual analog scale 0–10 for overall symptom severity", "Baseline and 14 days"],
    ], { y: 2.5, colW: [4.6, 5.23, 2.0], rowH: 0.34, fontSize: 12 });
    txt(s, "Secondary outcome measures", { x: L, y: 3.95, w: 6, h: 0.3, fontSize: 13, bold: true, color: TEAL });
    table(s, [
      ["Outcome measure", "Description", "Time frame"],
      ["Hospitalization", "Participants requiring hospitalization for COVID-19", "14 days"],
      ["Death", "Participants who die of COVID-19-related disease", "About 30 days"],
      ["Confirmed SARS-CoV-2 detection", "Participants with confirmed infection", "14 days"],
      ["Medication discontinuation", "Participants who stop study medication for any reason", "14 days"],
    ], { y: 4.28, colW: [4.6, 5.23, 2.0], rowH: 0.32, fontSize: 12 });
  }

  // =============== DATA TYPES ===============
  {
    const s = slideC({ label: "Study endpoints", title: "The endpoint's data type decides the effect measure", subtitle: "Link to Principles II: the hypothesis is written on the scale of the effect measure", takeaway: "Choose the data type first; the test, the margin, and the sample size formula follow from it." });
    table(s, [
      ["Data type", "Typical endpoints", "Summary per arm", "Effect measure", "Sample size driver"],
      ["Binary", "Response, cure, infection, death by a fixed time", "Proportion", "Risk difference, risk ratio, odds ratio", "Control event rate"],
      ["Continuous", "Blood pressure, HbA1c, score change from baseline", "Mean (SD)", "Mean difference", "Standard deviation"],
      ["Time-to-event", "Overall survival, progression, first hospitalization", "Median, survival curve", "Hazard ratio", "Number of events"],
      ["Count or ordinal", "Exacerbations per year, modified Rankin scale", "Rate, distribution", "Rate ratio, common odds ratio", "Dispersion, distribution"],
      ["Patient-reported", "Symptoms, fatigue, quality of life", "Mean score, responder %", "Mean difference, responder ratio", "SD, minimal important difference"],
    ], { y: 2.3, colW: [1.9, 3.3, 2.1, 2.63, 1.9], rowH: 0.58 });
  }

  // =============== HARD VS SURROGATE ===============
  {
    const s = slideC({ label: "Study endpoints", title: "Clinical outcomes versus surrogate endpoints", subtitle: "A surrogate substitutes for a clinical outcome; it is only as good as its validation", takeaway: "A surrogate must predict the clinical benefit, not merely correlate with disease severity.", source: "Fleming TR, DeMets DL. Ann Intern Med 1996;125:605–613; CAST Investigators. N Engl J Med 1989;321:406–412" });
    card(s, L, 2.3, 5.75, 1.9, HX.lt2, "Clinical card");
    txt(s, "Clinical (“hard”) outcome", { x: L + 0.3, y: 2.42, w: 5.2, h: 0.4, fontSize: 19, bold: true, color: TEAL });
    bullets(s, ["How a patient feels, functions, or survives", "Directly meaningful; rarely disputed", "Often infrequent and late: large trials, long follow-up"], { x: L + 0.3, y: 2.9, w: 5.2, h: 1.6, fontSize: 14 });
    card(s, L + 6.08, 2.3, 5.75, 1.9, HX.lt2, "Surrogate card");
    txt(s, "Surrogate endpoint", { x: L + 6.38, y: 2.42, w: 5.2, h: 0.4, fontSize: 19, bold: true, color: BLUE });
    bullets(s, ["Biomarker or intermediate outcome on the causal path", "Earlier, more frequent, cheaper: smaller and shorter trials", "Risk: treatment effect on surrogate may not translate into benefit"], { x: L + 6.38, y: 2.9, w: 5.2, h: 1.6, fontSize: 14 });
    table(s, [
      ["Surrogate", "Clinical outcome", "Validation status"],
      ["LDL cholesterol", "Myocardial infarction, cardiovascular death", "Accepted for statins and PCSK9 inhibitors"],
      ["HBV DNA suppression; HCV sustained virologic response", "Cirrhosis, hepatocellular carcinoma, liver death", "Accepted; supported by long-term cohorts"],
      ["Suppression of ventricular ectopy (CAST)", "Sudden cardiac death", "Failed: encainide and flecainide increased mortality"],
    ], { y: 4.4, colW: [4.4, 3.5, 3.93], rowH: 0.32, fontSize: 12 });
  }

  // =============== MEASUREMENT QUALITY ===============
  {
    const s = slideC({ label: "Study endpoints", title: "Measuring the endpoint well", subtitle: "Practical considerations that protect the validity of the comparison", takeaway: "An endpoint is only as credible as the procedure that measures it." });
    const items = [
      ["FiCheckSquare", "Validated instruments", "Use tools with established reliability, validity, and responsiveness, such as PASI, MELD, or FACIT-F."],
      ["FiList", "Standard procedures", "Same timing, device, laboratory, and reader across arms and sites; central reading where feasible."],
      ["FiUsers", "Blinded adjudication", "Independent committee classifies events (for example, cardiovascular death) without knowledge of allocation."],
      ["FiMessageSquare", "Patient-reported outcomes", "Capture symptoms and function from the patient; need validated questionnaires and a defined responder threshold."],
      ["FiRefreshCw", "Planned adaptation", "Allow for intercurrent events: dropout, rescue therapy, death before assessment. Define handling in advance."],
      ["FiDatabase", "Data integrity", "Audit trail, source verification, and missing-data strategy specified in the protocol."],
    ];
    const cw = 3.78, ch = 1.75, gx = 0.245, gy = 0.25;
    for (let i = 0; i < 6; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = L + col * (cw + gx), y = 2.3 + row * (ch + gy);
      card(s, x, y, cw, ch, HX.lt2, "Measure card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.22, y + 0.22, 0.6, HX.dk2, "FFFFFF", "Measure " + (i + 1));
      txt(s, items[i][1], { x: x + 0.98, y: y + 0.22, w: cw - 1.15, h: 0.6, fontSize: 16, bold: true, color: NAVY, valign: "middle" });
      txt(s, items[i][2], { x: x + 0.22, y: y + 0.95, w: cw - 0.44, h: 0.78, fontSize: 13, color: SLATE });
    }
  }

  // =============== ESTIMAND ===============
  {
    const s = slideC({ label: "Study endpoints", title: "The estimand: what exactly is being estimated", subtitle: "ICH E9(R1) asks the protocol to define five attributes, not just the endpoint", takeaway: "Two trials with the same endpoint can estimate different things if they handle intercurrent events differently.", source: "ICH E9(R1) Addendum on estimands and sensitivity analysis in clinical trials (2019)" });
    const attrs = [
      ["Treatment", "The regimen being compared, including background therapy"],
      ["Population", "Patients targeted by the clinical question"],
      ["Variable", "The endpoint obtained for each patient"],
      ["Intercurrent events", "Rescue medication, discontinuation, death: how each is handled"],
      ["Summary measure", "The population-level comparison, e.g., hazard ratio or risk difference"],
    ];
    const bw = 2.17, gap = 0.245;
    attrs.forEach((a, i) => {
      const x = L + i * (bw + gap);
      const fill = i === 3 ? HX.accent1 : HX.dk2;
      s.addShape(pres.ShapeType.roundRect, { x, y: 2.4, w: bw, h: 0.85, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: "Attr " + (i + 1) });
      txt(s, String(i + 1), { x: x + 0.15, y: 2.4, w: 0.45, h: 0.85, fontSize: 26, bold: true, color: WHITE, valign: "middle" });
      txt(s, a[0], { x: x + 0.6, y: 2.4, w: bw - 0.7, h: 0.85, fontSize: 15, bold: true, color: WHITE, valign: "middle" });
      txt(s, a[1], { x, y: 3.4, w: bw, h: 1.0, fontSize: 13, color: SLATE });
    });
    txt(s, "Strategies for intercurrent events", { x: L, y: 4.42, w: 6, h: 0.3, fontSize: 13, bold: true, color: TEAL });
    table(s, [
      ["Strategy", "Question answered", "Example"],
      ["Treatment policy", "Effect regardless of what happened after randomization", "Intention-to-treat analysis of mortality"],
      ["Hypothetical", "Effect had the event not occurred", "HbA1c change as if no rescue medication had been given"],
      ["Composite", "Event itself counts as a (bad) outcome", "Rescue therapy counted as treatment failure"],
    ], { y: 4.75, colW: [2.4, 5.0, 4.43], rowH: 0.28, fontSize: 12 });
  }

  // =============== MULTIPLICITY: WHY ===============
  section("Multiple endpoints");
  {
    const s = slideC({ label: "Multiple endpoints", title: "Why multiple endpoints are a problem", subtitle: "Family-wise error rate with k independent tests, each at α = 0.05 and all null hypotheses true", takeaway: "With five unadjusted endpoints there is roughly a one-in-four chance of at least one false-positive claim.", source: "Family-wise error rate = 1 − (1 − α)ᵏ, assuming independent tests" });
    const ks = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    const fwer = ks.map(k => Math.round((1 - Math.pow(0.95, k)) * 1000) / 10);
    s.addChart(pres.ChartType.bar, [{ name: "FWER", labels: ks.map(String), values: fwer }], Object.assign({}, chartBase, {
      x: L, y: 2.25, w: 7.6, h: 3.9, barDir: "col", chartColors: [HX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"',
      valAxisMinVal: 0, valAxisMaxVal: 50, valAxisLabelFormatCode: '0"%"', showCatAxisTitle: true, catAxisTitle: "Number of endpoints tested (k)", showValAxisTitle: true, valAxisTitle: "Probability of ≥1 false positive",
      valAxisMajorUnit: 10,
    }));
    txt(s, "k = 1", { x: L + 8.0, y: 2.4, w: 3.8, h: 0.3, fontSize: 13, bold: true, color: GRAY });
    txt(s, "5%", { x: L + 8.0, y: 2.65, w: 3.8, h: 0.6, fontSize: 36, bold: true, color: TEAL });
    txt(s, "k = 5", { x: L + 8.0, y: 3.45, w: 3.8, h: 0.3, fontSize: 13, bold: true, color: GRAY });
    txt(s, "23%", { x: L + 8.0, y: 3.7, w: 3.8, h: 0.6, fontSize: 36, bold: true, color: AMBER });
    txt(s, "k = 10", { x: L + 8.0, y: 4.5, w: 3.8, h: 0.3, fontSize: 13, bold: true, color: GRAY });
    txt(s, "40%", { x: L + 8.0, y: 4.75, w: 3.8, h: 0.6, fontSize: 36, bold: true, color: RED });
    txt(s, "Each extra endpoint inflates the chance of a false claim unless the testing plan controls it.", { x: L + 8.0, y: 5.5, w: 3.8, h: 0.7, fontSize: 13, color: SLATE });
  }

  // =============== WHEN PLURAL ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "When a trial needs more than one primary endpoint", takeaway: "Multiplicity is a cost to be planned for, not a reason to avoid clinically necessary endpoints." });
    card(s, L, 2.3, 5.75, 3.8, HX.lt2, "When card");
    txt(s, "When to consider", { x: L + 0.3, y: 2.45, w: 5.2, h: 0.4, fontSize: 19, bold: true, color: TEAL });
    const when = [["FiActivity", "Complex diseases", "Benefit has several dimensions, e.g., histology and symptoms in MASH"], ["FiShield", "Regulatory requirements", "Agencies may require evidence on two domains, e.g., cognition and function in Alzheimer disease"], ["FiLayers", "Comprehensive evaluation", "Efficacy and durability, or efficacy and safety, both matter for the decision"]];
    for (let i = 0; i < 3; i++) {
      const y = 3.0 + i * 1.0;
      await iconCircle(s, when[i][0], L + 0.3, y, 0.5, HX.accent1, "FFFFFF", "When " + (i + 1));
      txt(s, when[i][1], { x: L + 1.0, y, w: 4.5, h: 0.3, fontSize: 15, bold: true, color: NAVY });
      txt(s, when[i][2], { x: L + 1.0, y: y + 0.3, w: 4.5, h: 0.6, fontSize: 12.5, color: SLATE });
    }
    card(s, L + 6.08, 2.3, 5.75, 3.8, "FBEDEA", "Challenge card");
    txt(s, "Challenges", { x: L + 6.38, y: 2.45, w: 5.2, h: 0.4, fontSize: 19, bold: true, color: RED });
    const ch = [["FiAlertTriangle", "Multiplicity", "Type I error inflates with every additional test"], ["FiBarChart2", "Statistical complexity", "Correlated endpoints, alpha allocation, power for the combination"], ["FiHelpCircle", "Interpretation", "What does the trial conclude if one endpoint succeeds and the other fails?"]];
    for (let i = 0; i < 3; i++) {
      const y = 3.0 + i * 1.0;
      await iconCircle(s, ch[i][0], L + 6.38, y, 0.5, HX.accent3, "FFFFFF", "Challenge " + (i + 1));
      txt(s, ch[i][1], { x: L + 7.08, y, w: 4.5, h: 0.3, fontSize: 15, bold: true, color: NAVY });
      txt(s, ch[i][2], { x: L + 7.08, y: y + 0.3, w: 4.5, h: 0.6, fontSize: 12.5, color: SLATE });
    }
  }

  // =============== STRATEGIES OVERVIEW ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Four strategies for multiple endpoints", subtitle: "All must be written into the protocol and the statistical analysis plan before unblinding", takeaway: "The strategy is chosen by the clinical question: at least one success, all successes, or a single combined outcome." });
    const items = [
      ["FiSliders", "Multiplicity adjustment", "Split or adjust α across endpoints (Bonferroni, Holm, Hochberg). Trial succeeds if any endpoint passes its adjusted threshold.", "Claim: at least one"],
      ["FiGitMerge", "Hierarchical testing", "Test endpoints in a prespecified order at full α; stop at the first failure. No alpha is spent on later tests.", "Claim: ordered sequence"],
      ["FiLayers", "Composite endpoint", "Combine several events into one outcome (first occurrence of any). One test, more events, higher power.", "Claim: single combined outcome"],
      ["FiCheckCircle", "Co-primary endpoints", "All endpoints must be significant, each at full α. No adjustment needed, but power falls and sample size rises.", "Claim: all"],
    ];
    const cw = 5.78, ch = 1.8, gx = 0.27, gy = 0.25;
    for (let i = 0; i < 4; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = L + col * (cw + gx), y = 2.3 + row * (ch + gy);
      card(s, x, y, cw, ch, HX.lt2, "Strategy card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.25, y + 0.25, 0.65, HX.accent1, "FFFFFF", "Strategy " + (i + 1));
      txt(s, items[i][1], { x: x + 1.1, y: y + 0.25, w: cw - 1.3, h: 0.35, fontSize: 18, bold: true, color: NAVY });
      txt(s, items[i][3], { x: x + 1.1, y: y + 0.6, w: cw - 1.3, h: 0.3, fontSize: 12, bold: true, color: TEAL });
      txt(s, items[i][2], { x: x + 0.25, y: y + 1.0, w: cw - 0.5, h: 0.75, fontSize: 13, color: SLATE });
    }
  }

  // =============== BONFERRONI / HOLM ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Adjusting for multiple comparisons", subtitle: "Worked example: four hypotheses, overall α = 0.05", takeaway: "Adjustment protects the Type I error but lowers power; the smallest p-values must clear a harder bar.", notes: "Walk through: order p-values, compute Holm thresholds 0.05/4, 0.05/3, 0.05/2, 0.05/1. Reject sequentially until the first failure (0.03 > 0.025), then stop: H3 and H2 are not rejected." });
    card(s, L, 2.3, 5.6, 1.55, HX.lt2, "Bonferroni card");
    txt(s, "Bonferroni", { x: L + 0.25, y: 2.4, w: 5.1, h: 0.35, fontSize: 18, bold: true, color: TEAL });
    txt(s, [{ text: "α", options: { italic: true } }, { text: "adj", options: { subscript: true } }, { text: " = α / m        for every test" }], { x: L + 0.25, y: 2.8, w: 5.1, h: 0.4, fontSize: 18, bold: true, color: NAVY });
    txt(s, "Simple and conservative; power falls quickly as m grows.", { x: L + 0.25, y: 3.25, w: 5.1, h: 0.5, fontSize: 13, color: SLATE });
    card(s, L, 4.0, 5.6, 2.1, HX.lt2, "Holm card");
    txt(s, "Holm (step-down)", { x: L + 0.25, y: 4.1, w: 5.1, h: 0.35, fontSize: 18, bold: true, color: TEAL });
    txt(s, [{ text: "α" }, { text: "(k)", options: { subscript: true } }, { text: " = α / (m − k + 1)      k = rank of the p-value" }], { x: L + 0.25, y: 4.5, w: 5.1, h: 0.4, fontSize: 17, bold: true, color: NAVY });
    txt(s, "Order p-values from smallest to largest; test each against its threshold; stop at the first p-value that fails. Uniformly more powerful than Bonferroni.", { x: L + 0.25, y: 4.95, w: 5.1, h: 1.1, fontSize: 13, color: SLATE });
    table(s, [
      ["Rank", "Hypothesis", "p-value", "Bonferroni", "Holm", "Decision"],
      ["1", "H₄", "0.005", "0.0125", "0.0125", "Reject"],
      ["2", "H₁", "0.010", "0.0125", "0.0167", "Reject"],
      ["3", "H₃", "0.030", "0.0125", "0.0250", "Stop"],
      ["4", "H₂", "0.040", "0.0125", "0.0500", "Not tested"],
    ], { x: L + 5.95, y: 2.3, w: 5.88, colW: [0.7, 1.1, 0.95, 1.08, 0.95, 1.1], rowH: 0.36, fontSize: 12 });
    txt(s, "Thresholds: Bonferroni α/4 for every test; Holm α/(4 − k + 1) by rank. Unadjusted testing would have declared all four significant at 0.05. Both methods reject H₄ and H₁ only; Holm’s later thresholds are more generous, but H₃ still fails and testing stops.", { x: L + 5.95, y: 4.5, w: 5.88, h: 1.6, fontSize: 13, color: NAVY });
  }

  // =============== HIERARCHICAL ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Hierarchical (gatekeeping) testing", subtitle: "Each endpoint is tested at the full α, but only if every endpoint before it succeeded", takeaway: "Order the hierarchy by clinical importance and by the probability of success; a failure early closes every later gate." });
    const steps = ["Primary endpoint", "Key secondary 1", "Key secondary 2", "Key secondary 3"];
    const bw = 2.3, gap = 0.87, y = 2.55;
    steps.forEach((t, i) => {
      const x = L + i * (bw + gap);
      const fill = i < 2 ? HX.accent1 : i === 2 ? HX.accent3 : "C9D3D9";
      box(s, t, x, y, bw, 0.8, fill, WHITE, { fontSize: 15 });
      txt(s, i < 2 ? "p < 0.05 ✓" : i === 2 ? "p = 0.08 ✗" : "Not tested", { x, y: y + 0.9, w: bw, h: 0.35, fontSize: 14, bold: true, color: i < 2 ? TEAL : i === 2 ? RED : GRAY, align: "center" });
      txt(s, i < 2 ? "Claim allowed" : i === 2 ? "Gate closed" : "No claim, even if p < 0.05", { x, y: y + 1.25, w: bw, h: 0.4, fontSize: 12, color: SLATE, align: "center" });
      if (i < 3) {
        arrow(s, x + bw + 0.08, y + 0.4, x + bw + gap - 0.08, y + 0.4, i < 2 ? HX.accent1 : HX.accent4);
        txt(s, "if significant", { x: x + bw, y: y - 0.35, w: gap, h: 0.3, fontSize: 10, color: GRAY, align: "center" });
      }
    });
    card(s, L, 4.55, 5.75, 1.55, HX.lt2, "Pros");
    txt(s, "Strengths", { x: L + 0.25, y: 4.62, w: 5.2, h: 0.3, fontSize: 15, bold: true, color: TEAL });
    bullets(s, ["Full α at every step; no power loss for the primary endpoint", "Natural fit for a primary followed by key secondary claims", "Common in regulatory submissions"], { x: L + 0.25, y: 4.95, w: 5.2, h: 1.15, fontSize: 13 });
    card(s, L + 6.08, 4.55, 5.75, 1.55, "FBEDEA", "Cons");
    txt(s, "Limitations", { x: L + 6.33, y: 4.62, w: 5.2, h: 0.3, fontSize: 15, bold: true, color: RED });
    bullets(s, ["Order must be fixed before unblinding and cannot be changed", "A late endpoint with a strong effect yields no claim if an earlier one fails", "Variants (fallback, graphical approaches) can recycle α"], { x: L + 6.33, y: 4.95, w: 5.2, h: 1.15, fontSize: 13 });
  }

  // =============== LEAP-002 ===============
  {
    const s = slideC({ label: "Clinical example", title: "LEAP-002: dual primary endpoints with split alpha", subtitle: "Lenvatinib plus pembrolizumab versus lenvatinib plus placebo in advanced hepatocellular carcinoma", takeaway: "Both hazard ratios favored the combination, yet neither crossed its prespecified boundary: a negative trial by design.", source: "Llovet JM et al. Lancet Oncol 2023;24:1399–1410", notes: "The trial allocated one-sided alpha: OS 0.019 and PFS 0.002 at the final analysis (after earlier interim looks). A 'nominal' p of 0.023 for OS looks close but does not meet the threshold. Emphasize that the thresholds were set before unblinding." });
    const cols = [
      ["Overall survival", "HR 0.84", "95% CI 0.71–1.00", "One-sided p = 0.023", "Threshold p ≤ 0.019", false],
      ["Progression-free survival", "HR 0.87", "95% CI 0.73–1.02", "One-sided p = 0.047", "Threshold p ≤ 0.002", false],
    ];
    cols.forEach((c, i) => {
      const x = L + i * 4.1;
      card(s, x, 2.3, 3.85, 3.8, HX.lt2, "LEAP card " + (i + 1));
      txt(s, c[0], { x: x + 0.25, y: 2.42, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: TEAL });
      txt(s, c[1], { x: x + 0.25, y: 2.85, w: 3.4, h: 0.7, fontSize: 40, bold: true, color: NAVY });
      txt(s, c[2], { x: x + 0.25, y: 3.6, w: 3.4, h: 0.35, fontSize: 15, color: NAVY });
      txt(s, c[3], { x: x + 0.25, y: 4.15, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: NAVY });
      txt(s, c[4], { x: x + 0.25, y: 4.5, w: 3.4, h: 0.35, fontSize: 15, color: SLATE });
      txt(s, "Not met ✗", { x: x + 0.25, y: 5.3, w: 3.4, h: 0.5, fontSize: 20, bold: true, color: RED });
    });
    txt(s, "Design logic", { x: L + 8.45, y: 2.35, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: TEAL });
    bullets(s, ["Two primary endpoints, each with its own α allocation and event-driven analysis", "Success on either endpoint would have supported the claim", "The PFS boundary was very strict after interim analyses, so nearly all α rested on OS", "Interpretation: “did not meet prespecified significance”; findings do not support a change in practice"], { x: L + 8.45, y: 2.75, w: 3.4, h: 3.4, fontSize: 13 });
  }

  // =============== COMPOSITE ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Composite endpoints", subtitle: "Several clinically related events combined into one outcome, usually time to the first of any", takeaway: "Always report each component: the composite tells you whether something happened, not what happened." });
    card(s, L, 2.3, 5.75, 1.75, HX.lt2, "Why card");
    txt(s, "Why use one", { x: L + 0.3, y: 2.4, w: 5.2, h: 0.35, fontSize: 17, bold: true, color: TEAL });
    bullets(s, ["More events: smaller sample size and shorter follow-up", "One primary test: no multiplicity problem", "Captures the net clinical burden of a disease"], { x: L + 0.3, y: 2.8, w: 5.2, h: 1.2, fontSize: 13.5 });
    card(s, L, 4.2, 5.75, 1.9, "FBEDEA", "Caution card");
    txt(s, "Cautions", { x: L + 0.3, y: 4.3, w: 5.2, h: 0.35, fontSize: 17, bold: true, color: RED });
    bullets(s, ["Components should have similar clinical importance and respond similarly to treatment", "Often driven by the most frequent, least serious component", "Only the first event counts; later, worse events are ignored", "Can be difficult to interpret and to communicate to patients"], { x: L + 0.3, y: 4.7, w: 5.2, h: 1.4, fontSize: 13 });
    // diagram on the right: three events feeding into composite
    const ex = L + 6.3;
    const comps = ["Cardiovascular death", "Non-fatal myocardial infarction", "Non-fatal stroke"];
    comps.forEach((c, i) => {
      const y = 2.45 + i * 1.0;
      box(s, c, ex, y, 2.6, 0.75, HX.dk2, WHITE, { fontSize: 13 });
      s.addShape(pres.ShapeType.line, { x: ex + 2.6, y: y + 0.375, w: 0.6, h: 0, line: { color: HX.accent4, width: 1.5 }, objectName: "Comp line " + i });
    });
    s.addShape(pres.ShapeType.line, { x: ex + 3.2, y: 2.825, w: 0, h: 2.0, line: { color: HX.accent4, width: 1.5 }, objectName: "Comp join" });
    arrow(s, ex + 3.2, 3.825, ex + 3.75, 3.825, HX.accent1);
    box(s, "MACE\n(time to first event)", ex + 3.8, 3.3, 1.75, 1.05, HX.accent1, WHITE, { fontSize: 13 });
    txt(s, "SELECT: 3-point major adverse cardiovascular events (MACE)", { x: ex, y: 5.5, w: 5.5, h: 0.5, fontSize: 12, italic: true, color: SLATE });
  }

  // =============== SELECT FOREST ===============
  {
    const s = slideC({ label: "Clinical example", title: "SELECT: reading a composite with its components", subtitle: "Semaglutide 2.4 mg versus placebo in overweight or obesity with cardiovascular disease; 17,604 patients, 40 months", takeaway: "The composite was clearly positive; the components show which events drove it and which are uncertain.", source: "Lincoff AM et al. N Engl J Med 2023;389:2221–2232" });
    const rows = [
      ["Primary composite (MACE)", 0.80, 0.72, 0.90, "p < 0.001", true],
      ["Cardiovascular death", 0.85, 0.71, 1.01, "p = 0.07", false],
      ["Non-fatal myocardial infarction", 0.72, 0.61, 0.85, "", false],
      ["Non-fatal stroke", 0.93, 0.74, 1.15, "", false],
      ["Heart failure composite", 0.82, 0.71, 0.96, "", false],
      ["Death from any cause", 0.81, 0.71, 0.93, "", false],
    ];
    const ax = L + 4.6, aw = 4.6, y0 = 2.4, rh = 0.5;
    const lo = 0.5, hi = 1.4;
    const xOf = v => ax + aw * (Math.log(v) - Math.log(lo)) / (Math.log(hi) - Math.log(lo));
    // axis
    const yAxis = y0 + rows.length * rh;
    s.addShape(pres.ShapeType.line, { x: ax, y: yAxis, w: aw, h: 0, line: { color: HX.accent4, width: 1 }, objectName: "Forest axis" });
    [0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.2, 1.4].forEach(v => {
      s.addShape(pres.ShapeType.line, { x: xOf(v), y: yAxis, w: 0, h: 0.08, line: { color: HX.accent4, width: 1 }, objectName: "Tick " + v });
      txt(s, v.toFixed(1), { x: xOf(v) - 0.3, y: yAxis + 0.1, w: 0.6, h: 0.25, fontSize: 11, color: SLATE, align: "center" });
    });
    s.addShape(pres.ShapeType.line, { x: xOf(1), y: y0 - 0.15, w: 0, h: rows.length * rh + 0.15, line: { color: HX.accent3, width: 1, dashType: "dash" }, objectName: "Null line" });
    txt(s, "Favors semaglutide", { x: ax, y: yAxis + 0.36, w: 2.3, h: 0.25, fontSize: 11, bold: true, color: GRAY });
    txt(s, "Favors placebo", { x: ax + aw - 2.3, y: yAxis + 0.36, w: 2.3, h: 0.25, fontSize: 11, bold: true, color: GRAY, align: "right" });
    txt(s, "Hazard ratio (95% CI), log scale", { x: ax + aw + 0.25, y: y0 - 0.3, w: 2.4, h: 0.25, fontSize: 11, bold: true, color: GRAY });
    rows.forEach((r, i) => {
      const yc = y0 + i * rh + rh / 2;
      txt(s, r[0], { x: L, y: yc - 0.2, w: 4.4, h: 0.4, fontSize: 13, bold: r[5], color: NAVY, valign: "middle" });
      s.addShape(pres.ShapeType.line, { x: xOf(r[2]), y: yc, w: xOf(r[3]) - xOf(r[2]), h: 0, line: { color: r[5] ? HX.accent1 : HX.accent6, width: 2 }, objectName: "CI " + i });
      const d = r[5] ? 0.2 : 0.15;
      s.addShape(r[5] ? pres.ShapeType.diamond : pres.ShapeType.ellipse, { x: xOf(r[1]) - d / 2, y: yc - d / 2, w: d, h: d, fill: { color: r[5] ? HX.accent1 : HX.accent6 }, line: { color: r[5] ? HX.accent1 : HX.accent6, width: 0 }, objectName: "Point " + i });
      txt(s, `${r[1].toFixed(2)} (${r[2].toFixed(2)}–${r[3].toFixed(2)})  ${r[4]}`, { x: ax + aw + 0.25, y: yc - 0.2, w: 2.4, h: 0.4, fontSize: 12, color: NAVY, valign: "middle", bold: r[5] });
    });
  }

  // =============== CO-PRIMARY ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Co-primary endpoints: every endpoint must succeed", subtitle: "Used when a treatment is clinically useful only if it works on two distinct domains", takeaway: "No alpha adjustment is needed, but the price is paid in power and sample size." });
    // logic diagram
    const y = 2.45;
    box(s, "Endpoint A\nsignificant at α", L, y, 2.6, 0.95, HX.accent1, WHITE, { fontSize: 14 });
    txt(s, "AND", { x: L + 2.7, y, w: 0.8, h: 0.95, fontSize: 16, bold: true, color: NAVY, align: "center", valign: "middle" });
    box(s, "Endpoint B\nsignificant at α", L + 3.6, y, 2.6, 0.95, HX.accent1, WHITE, { fontSize: 14 });
    arrow(s, L + 6.3, y + 0.475, L + 6.95, y + 0.475, HX.accent1);
    box(s, "Trial succeeds", L + 7.0, y, 2.3, 0.95, HX.dk2, WHITE, { fontSize: 14 });
    txt(s, "Either fails → trial fails", { x: L + 9.5, y, w: 2.4, h: 0.95, fontSize: 14, bold: true, color: RED, valign: "middle" });
    // power chart
    s.addChart(pres.ChartType.bar, [
      { name: "Power of each endpoint", labels: ["80%", "85%", "90%", "95%"], values: [80, 85, 90, 95] },
      { name: "Power that both succeed", labels: ["80%", "85%", "90%", "95%"], values: [64, 72.3, 81, 90.3] },
    ], Object.assign({}, chartBase, {
      x: L, y: 3.55, w: 6.3, h: 2.4, barDir: "col", barGrouping: "clustered", chartColors: [HX.accent6, HX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
      valAxisMinVal: 0, valAxisMaxVal: 100, valAxisLabelFormatCode: '0"%"', showLegend: true, legendPos: "t", legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HX.dk1,
      showCatAxisTitle: true, catAxisTitle: "Power for each endpoint alone", valAxisMajorUnit: 25,
    }));
    txt(s, "Joint power ≈ product of the individual powers when endpoints are independent; higher when correlated.", { x: L, y: 6.0, w: 6.3, h: 0.25, fontSize: 10.5, color: GRAY });
    card(s, L + 6.6, 3.6, 5.23, 2.5, HX.lt2, "Coprimary notes");
    txt(s, "Design implications", { x: L + 6.85, y: 3.7, w: 4.8, h: 0.35, fontSize: 16, bold: true, color: TEAL });
    bullets(s, ["Each endpoint tested at full α (no split): the “all must win” rule is already conservative", "Sample size must give adequate power for the pair, so it rises above the larger single requirement", "Typical settings: Alzheimer disease (cognition and function), IBS (pain and stool), MASH (histology)", "Different from multiple primary endpoints, where one success suffices and α is split"], { x: L + 6.85, y: 4.1, w: 4.8, h: 2.0, fontSize: 12.5 });
  }

  // =============== SMART TRIAL ===============
  {
    const s = slideC({ label: "Clinical example", title: "SMART: co-primary endpoints with different claims", subtitle: "Self-expanding (SEV) versus balloon-expandable (BEV) transcatheter aortic valves in small annuli; 716 patients, 12 months", takeaway: "One endpoint tested for non-inferiority, the other for superiority: both had to succeed, and both did.", source: "Herrmann HC et al. N Engl J Med 2024;390:1959–1971" });
    const cols = [
      ["Co-primary 1: clinical", "Death, disabling stroke, or rehospitalization for heart failure", "9.4% vs 10.6%", "Difference −1.2 points (90% CI −4.9 to 2.5)", "Non-inferiority: p < 0.001", "Met ✓"],
      ["Co-primary 2: valve function", "Bioprosthetic-valve dysfunction", "9.4% vs 41.6%", "Difference −32.2 points (95% CI −38.7 to −25.6)", "Superiority: p < 0.001", "Met ✓"],
    ];
    cols.forEach((c, i) => {
      const x = L + i * 4.1;
      card(s, x, 2.3, 3.85, 3.8, HX.lt2, "SMART card " + (i + 1));
      txt(s, c[0], { x: x + 0.25, y: 2.42, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: TEAL });
      txt(s, c[1], { x: x + 0.25, y: 2.8, w: 3.4, h: 0.6, fontSize: 13, color: SLATE });
      txt(s, c[2], { x: x + 0.25, y: 3.45, w: 3.4, h: 0.65, fontSize: 32, bold: true, color: NAVY });
      txt(s, c[3], { x: x + 0.25, y: 4.15, w: 3.4, h: 0.55, fontSize: 13, color: NAVY });
      txt(s, c[4], { x: x + 0.25, y: 4.75, w: 3.4, h: 0.35, fontSize: 14, bold: true, color: NAVY });
      txt(s, c[5], { x: x + 0.25, y: 5.35, w: 3.4, h: 0.5, fontSize: 20, bold: true, color: TEAL });
    });
    txt(s, "Why co-primary here", { x: L + 8.45, y: 2.35, w: 3.4, h: 0.35, fontSize: 15, bold: true, color: TEAL });
    bullets(s, ["A better-functioning valve is useful only if clinical outcomes are not worse", "The two domains answer different questions and use different hypotheses (NI versus superiority)", "Both tests run at their own α; success requires both", "Reporting states both results and both claims explicitly"], { x: L + 8.45, y: 2.75, w: 3.4, h: 3.4, fontSize: 13 });
  }

  // =============== DECISION GUIDE ===============
  {
    const s = slideC({ label: "Multiple endpoints", title: "Choosing a strategy: match the claim to the method", takeaway: "Decide what the trial must show to change practice, then pick the testing rule that delivers exactly that claim." });
    table(s, [
      ["What must the trial show?", "Strategy", "α handling", "Power impact", "Example"],
      ["Benefit on at least one of several outcomes", "Multiple primary endpoints with adjustment", "Split or adjusted (Bonferroni, Holm, Hochberg)", "Reduced per endpoint", "LEAP-002: OS or PFS"],
      ["Primary benefit, then ordered secondary claims", "Hierarchical testing", "Full α at each step, sequential", "None for primary; later steps depend on earlier", "Primary then key secondary endpoints"],
      ["Net effect on related events", "Composite endpoint", "Single test", "Increased (more events)", "SELECT: 3-point MACE"],
      ["Benefit on every domain", "Co-primary endpoints", "Full α for each; all must succeed", "Reduced jointly; larger sample", "SMART: clinical NI and valve superiority"],
    ], { y: 2.35, colW: [2.75, 2.35, 2.45, 2.15, 2.13], rowH: 0.7, fontSize: 12.5 });
  }

  // =============== CLINICAL RELEVANCE ===============
  section("Good practice");
  {
    const s = slideC({ label: "Good practice", title: "Aligning hypotheses and endpoints with clinical relevance", takeaway: "A statistically significant endpoint that patients and clinicians do not value will not change practice." });
    const items = [
      ["FiHeart", "Patient-centered outcomes", "Focus on what matters to patients: survival, symptoms, function, and quality of life. Involve patients in endpoint selection."],
      ["FiActivity", "Disease-specific measures", "Tailor the endpoint to the condition and its stage: histology in early MASH, decompensation events in cirrhosis."],
      ["FiUsers", "Stakeholder input", "Engage clinicians, patients, statisticians, and regulators early; use core outcome sets where they exist."],
    ];
    const cw = 3.78, gx = 0.245;
    for (let i = 0; i < 3; i++) {
      const x = L + i * (cw + gx);
      card(s, x, 2.3, cw, 3.8, HX.lt2, "Relevance card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.3, 2.6, 0.8, HX.accent1, "FFFFFF", "Relevance " + (i + 1));
      txt(s, items[i][1], { x: x + 0.3, y: 3.6, w: cw - 0.6, h: 0.75, fontSize: 18, bold: true, color: NAVY });
      txt(s, items[i][2], { x: x + 0.3, y: 4.4, w: cw - 0.6, h: 1.6, fontSize: 14, color: SLATE });
    }
  }

  // =============== GOOD PRACTICE CHECKLIST ===============
  {
    const s = slideC({ label: "Good practice", title: "Good practice checklist", subtitle: "What reviewers, regulators, and journals will look for", takeaway: "Prespecify, register, document, and report: the credibility of the result rests on the paper trail." });
    const items = [
      ["FiBookOpen", "Guideline compliance", "ICH E9 and E9(R1), CONSORT, SPIRIT, and disease-specific regulatory guidance"],
      ["FiFileText", "Documentation", "Hypotheses, endpoints, testing hierarchy, and sample size in the protocol and statistical analysis plan"],
      ["FiEye", "Transparency", "Registration before enrollment; any change to the primary endpoint dated, justified, and reported"],
      ["FiAlertCircle", "Risk-benefit analysis", "Endpoints capture harms as well as benefits; safety endpoints prespecified"],
      ["FiDatabase", "Data integrity", "Blinded assessment, independent adjudication, audit trail, missing-data strategy"],
      ["FiUsers", "Collaboration", "Clinicians, statisticians, patients, and regulators involved from the design stage"],
    ];
    const cw = 5.78, ch = 1.15, gx = 0.27, gy = 0.17;
    for (let i = 0; i < 6; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = L + col * (cw + gx), y = 2.3 + row * (ch + gy);
      card(s, x, y, cw, ch, HX.lt2, "Practice card " + (i + 1));
      await iconCircle(s, items[i][0], x + 0.22, y + 0.27, 0.6, HX.dk2, "FFFFFF", "Practice " + (i + 1));
      txt(s, items[i][1], { x: x + 1.0, y: y + 0.15, w: cw - 1.2, h: 0.35, fontSize: 16, bold: true, color: NAVY });
      txt(s, items[i][2], { x: x + 1.0, y: y + 0.5, w: cw - 1.2, h: 0.6, fontSize: 12.5, color: SLATE });
    }
  }

  // =============== CLASS EXERCISE ===============
  {
    const s = slideC({ label: "Apply the concept", title: "Class exercise: design the hypothesis and endpoints", subtitle: "Phase 3 trial of a new oral agent for MASH with stage F2–F3 fibrosis, 52 weeks, versus placebo", takeaway: "Prepare a two-minute justification of your primary endpoint and your multiplicity plan." });
    card(s, L, 2.3, 5.75, 3.8, HX.lt2, "Scenario card");
    txt(s, "Scenario", { x: L + 0.3, y: 2.42, w: 5.2, h: 0.35, fontSize: 17, bold: true, color: TEAL });
    bullets(s, ["Regulators accept two histologic endpoints: MASH resolution without worsening of fibrosis, and fibrosis improvement ≥1 stage without worsening of MASH", "Sponsor also wants claims on non-invasive tests (liver stiffness, ALT) and a patient-reported fatigue score", "Paired biopsies are available for about 85% of patients"], { x: L + 0.3, y: 2.85, w: 5.2, h: 3.2, fontSize: 13.5 });
    const qs = [["1", "Write the primary hypothesis as H₀ and H₁ on the scale of the endpoint."], ["2", "Should the two histologic endpoints be co-primary or multiple primary with split α? Justify."], ["3", "Place the non-invasive and PRO endpoints in a testing hierarchy."], ["4", "Which intercurrent event worries you most, and which estimand strategy would you choose?"]];
    qs.forEach((q, i) => {
      const y = 2.35 + i * 0.95;
      s.addShape(pres.ShapeType.ellipse, { x: L + 6.1, y: y + 0.08, w: 0.5, h: 0.5, fill: { color: HX.accent1 }, line: { color: HX.accent1, width: 0 }, objectName: "Q " + q[0] });
      txt(s, q[0], { x: L + 6.1, y: y + 0.08, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(s, q[1], { x: L + 6.8, y, w: 5.0, h: 0.85, fontSize: 14, color: NAVY, valign: "middle" });
    });
  }

  // =============== KEY TAKEAWAYS ===============
  {
    const s = slideC({ label: "Take-home points", title: "Key principles", takeaway: "A well-designed trial can be described in one sentence: the hypothesis, the primary endpoint, and the rule for success." });
    const rows = [
      ["The hypothesis", "Starts from a knowledge gap, is written as H₀ and H₁ on a measurable scale, and states its direction and margin."],
      ["The endpoint", "Makes the hypothesis measurable; its data type determines the effect measure, the test, and the sample size."],
      ["The tiers", "Primary carries the confirmatory claim; secondary endpoints support it; exploratory endpoints generate hypotheses."],
      ["Multiplicity", "Every additional claim costs Type I error: adjust, order, combine, or require all, and decide before unblinding."],
      ["The practice", "Clinical relevance, validated measurement, prespecification, registration, and transparent reporting."],
    ];
    rows.forEach((r, i) => {
      const y = 2.3 + i * 0.76;
      txt(s, r[0], { x: L, y, w: 2.6, h: 0.65, fontSize: 17, bold: true, color: TEAL, valign: "middle" });
      txt(s, r[1], { x: L + 2.8, y, w: CW - 2.8, h: 0.65, fontSize: 14.5, color: NAVY, valign: "middle" });
      if (i < rows.length - 1) s.addShape(pres.ShapeType.line, { x: L, y: y + 0.72, w: CW, h: 0, line: { color: "E3E8EB", width: 0.75 }, objectName: "KP line " + i });
    });
  }

  // =============== Q&A ===============
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: sectionName });
    s.background = { color: HX.dk2 };
    txt(s, "QUESTIONS AND DISCUSSION", { x: 0.75, y: 2.4, w: 11.8, h: 0.35, fontSize: 13, bold: true, color: "8CC5CF" });
    txt(s, "Which endpoint would change your practice?", { x: 0.75, y: 2.85, w: 11.8, h: 1.2, fontSize: 40, bold: true, color: WHITE });
    txt(s, "Yao-Chun Hsu  ·  Graduate Institute of Medicine, I-Shou University  ·  E-Da Hospital", { x: 0.75, y: 4.3, w: 11.8, h: 0.4, fontSize: 14, color: "C9DEE3" });
    txt(s, FOOT, { x: 0.75, y: 7.05, w: 8, h: 0.25, fontSize: 9, color: "8CC5CF" });
  }

  await pres.writeFile({ fileName: "Hypotheses_and_Endpoints_2026.pptx" });
  await applyTheme("Hypotheses_and_Endpoints_2026.pptx", THEME);
  console.log("written");
})().catch(e => { console.error(e); process.exit(1); });
