// Slide deck: "Screening and Prevention of Liver Cancer:
// Future Strategies for the European Union from Taiwan's Experience"
const pptxgen = require("pptxgenjs");
const path = require("path");
const ICON = (n) => path.join(__dirname, "icons", `${n}.png`);

const C = {
  burgDark: "4E1B23",
  burg: "6D1F2C",
  burgMid: "8C2B35",
  burgSeries: "B3323F",
  blue: "2A78D6",
  blueDark: "1C5CAB",
  gold: "EDA100",
  goldDark: "C98500",
  aqua: "1BAF7A",
  green: "0CA30C",
  red: "D03B3B",
  ink: "232323",
  sub: "52514E",
  mut: "898781",
  grid: "E1E0D9",
  axis: "C3C2B7",
  white: "FFFFFF",
  rose: "E9CDD2",
  tintBurg: "F8F0F1",
  tintBlue: "EDF4FC",
  tintGold: "FBF5E4",
  tintGray: "F4F3F1",
  tintAqua: "EAF7F1",
};
const TITLE = "Cambria";
const BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5

// ---------- helpers (always return / build fresh option objects) ----------
function lightSlide() {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  return s;
}
function darkSlide() {
  const s = pres.addSlide();
  s.background = { color: C.burgDark };
  return s;
}
function kickerTitle(s, kicker, title, opts = {}) {
  s.addText(kicker, {
    x: 0.6, y: 0.34, w: 12.1, h: 0.3, fontFace: BODY, fontSize: 12, bold: true,
    color: opts.kcolor || C.burgMid, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: 0.6, y: 0.66, w: 12.1, h: 1.15, fontFace: TITLE, fontSize: 30, bold: true,
    color: opts.tcolor || C.ink, isTextBox: true, margin: 0, valign: "top",
  });
}
function source(s, txt, color) {
  s.addText(txt, {
    x: 0.6, y: 7.08, w: 12.1, h: 0.3, fontFace: BODY, fontSize: 8.5,
    color: color || C.mut, isTextBox: true, margin: 0,
  });
}
function iconCircle(s, icon, x, y, d, fill) {
  s.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill } });
  const pad = d * 0.27;
  s.addImage({ path: ICON(icon), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
function statTile(s, o) {
  s.addShape("roundRect", {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.07,
    fill: { color: o.tint || C.tintGray },
  });
  s.addText(o.value, {
    x: o.x + 0.22, y: o.y + 0.1, w: o.w - 0.44, h: o.h * 0.48, fontFace: TITLE,
    fontSize: o.vsize || 28, bold: true, color: o.vcolor || C.burg,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText(o.label, {
    x: o.x + 0.22, y: o.y + o.h * 0.52, w: o.w - 0.44, h: o.h * 0.45,
    fontFace: BODY, fontSize: o.lsize || 11, color: C.sub, isTextBox: true, margin: 0, valign: "top",
  });
}
function chartHeader(s, txt, x, y, w) {
  s.addText(txt, {
    x, y, w, h: 0.32, fontFace: BODY, fontSize: 13, bold: true, color: C.ink,
    isTextBox: true, margin: 0,
  });
}
function baseChart(extra) {
  return Object.assign(
    {
      chartColors: [C.blue],
      showLegend: false,
      showTitle: false,
      catAxisLabelColor: C.sub,
      catAxisLabelFontSize: 11,
      catAxisLabelFontFace: BODY,
      catAxisLineColor: C.axis,
      catGridLine: { style: "none" },
      valAxisLabelColor: C.mut,
      valAxisLabelFontSize: 10,
      valAxisLabelFontFace: BODY,
      valAxisLineShow: false,
      valGridLine: { color: C.grid, size: 0.5 },
      showValue: true,
      dataLabelColor: C.ink,
      dataLabelFontSize: 11,
      dataLabelFontFace: BODY,
      dataLabelPosition: "outEnd",
      barGapWidthPct: 80,
    },
    extra
  );
}
function iconRow(s, o) {
  // small icon circle + bold lead + description, one row
  iconCircle(s, o.icon, o.x, o.y + 0.02, 0.52, o.fill);
  s.addText(
    [
      { text: o.lead, options: { bold: true, color: C.ink } },
      { text: o.rest ? "  " + o.rest : "", options: { color: C.sub } },
    ],
    {
      x: o.x + 0.72, y: o.y, w: o.w - 0.72, h: o.h || 0.62, fontFace: BODY,
      fontSize: o.size || 12.5, isTextBox: true, margin: 0, valign: "middle",
    }
  );
}

// =====================================================================
// SLIDE 1 — Title (dark)
// =====================================================================
{
  const s = darkSlide();
  s.addImage({ path: ICON("liver-white"), x: 10.35, y: 4.55, w: 2.4, h: 2.4, transparency: 82 });
  s.addText("EUROPEAN–TAIWAN CANCER SCREENING CONFERENCE 2026", {
    x: 0.9, y: 0.75, w: 11.5, h: 0.35, fontFace: BODY, fontSize: 13, bold: true,
    color: C.rose, charSpacing: 3, isTextBox: true, margin: 0,
  });
  s.addText("Screening and Prevention of Liver Cancer", {
    x: 0.9, y: 1.75, w: 11.6, h: 1.0, fontFace: TITLE, fontSize: 40, bold: true,
    color: C.white, isTextBox: true, margin: 0,
  });
  s.addText("Future Strategies for the European Union from Taiwan's Experience", {
    x: 0.9, y: 2.8, w: 11.0, h: 0.6, fontFace: TITLE, fontSize: 21,
    color: C.rose, isTextBox: true, margin: 0,
  });
  const pillars = [
    ["syringe-white", "Vaccinate"],
    ["pills-white", "Cure"],
    ["ultrasound-white", "Screen"],
  ];
  pillars.forEach(([ic, lb], i) => {
    const x = 0.95 + i * 1.75;
    iconCircle(s, ic, x, 4.35, 0.85, C.burgMid);
    s.addText(lb, {
      x: x - 0.35, y: 5.28, w: 1.55, h: 0.3, align: "center", fontFace: BODY,
      fontSize: 12.5, color: C.white, isTextBox: true, margin: 0,
    });
  });
  s.addText("Holden Hsu  ·  [Affiliation]", {
    x: 0.9, y: 6.45, w: 8.0, h: 0.35, fontFace: BODY, fontSize: 14, color: C.white,
    isTextBox: true, margin: 0,
  });
  s.addNotes(
    "Good morning. Over the next twenty minutes I will make one argument: liver cancer is a policy-sensitive disease. " +
    "Taiwan has spent forty years proving that a country can vaccinate, cure and screen its way out of a liver cancer epidemic. " +
    "Europe now faces its own — with different causes — and the question is whether the EU will build an organised response before the burden doubles. " +
    "I will show Taiwan's three pillars and their measured impact, then translate them into three concrete future strategies for the EU."
  );
}

// =====================================================================
// SLIDE 2 — The problem: Europe's rising burden
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "THE PROBLEM", "The cancer Europe is losing ground on");
  // 2x2 stat tiles left
  const tx = 0.6, ty = 1.75, tw = 2.85, th = 1.62, gx = 0.25, gy = 0.25;
  statTile(s, { x: tx, y: ty, w: tw, h: th, value: "54,600", label: "liver-cancer deaths, EU-27, 2022", tint: C.tintBlue, vcolor: C.blueDark });
  statTile(s, { x: tx + tw + gx, y: ty, w: tw, h: th, value: "6th", label: "cause of cancer death in the EU", tint: C.tintBlue, vcolor: C.blueDark });
  statTile(s, { x: tx, y: ty + th + gy, w: tw, h: th, value: "~12%", label: "5-year survival in Europe", tint: C.tintBurg, vcolor: C.burg });
  statTile(s, { x: tx + tw + gx, y: ty + th + gy, w: tw, h: th, value: "+22%", label: "further burden growth projected by 2040", tint: C.tintBurg, vcolor: C.burg });
  s.addText(
    [
      { text: "While overall EU cancer mortality keeps falling, ", options: { color: C.sub } },
      { text: "liver-cancer deaths are rising.", options: { bold: true, color: C.ink } },
    ],
    { x: 0.6, y: 5.65, w: 5.95, h: 0.8, fontFace: BODY, fontSize: 14, isTextBox: true, margin: 0 }
  );
  // chart right
  chartHeader(s, "Liver-cancer deaths, WHO European region", 7.0, 1.7, 5.7);
  s.addChart(pres.ChartType.bar, [
    { name: "Deaths", labels: ["2000", "2023"], values: [43000, 69000] },
  ], baseChart({
    x: 7.0, y: 2.1, w: 5.7, h: 4.1, barDir: "col",
    chartColors: [C.blue], dataLabelFormatCode: "#,##0",
    valAxisMinVal: 0, valAxisMaxVal: 80000, barGapWidthPct: 110,
    valAxisLabelFormatCode: "#,##0",
  }));
  s.addText("+52% in one generation", {
    x: 7.0, y: 6.28, w: 5.7, h: 0.3, fontFace: BODY, fontSize: 12, bold: true,
    color: C.burg, isTextBox: true, margin: 0, align: "center",
  });
  source(s, "Sources: ECIS / European Commission (EU-27, 2022) · 2nd EASL–Lancet Liver Commission, The Lancet 2026 · EUROCARE-5 · UEG Week 2025.");
  s.addNotes(
    "Scale first. Liver cancer killed about 54,600 people in the EU-27 in 2022 — the 6th cause of cancer death — and five-year survival is still around 12%. " +
    "The alarming part is the direction of travel: total EU cancer mortality has been falling for three decades, yet liver-cancer deaths in the WHO European region rose from about 43,000 in 2000 to about 69,000 in 2023 — up more than fifty percent — and UEG projects a further ~22% rise by 2040 if nothing changes. " +
    "This is the one major cancer where Europe is moving backwards."
  );
}

// =====================================================================
// SLIDE 3 — Preventable cancer without a program
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "THE PARADOX", "The most preventable major cancer — without a programme");
  chartHeader(s, "Causes of liver-cancer deaths, global (GBD 2021)", 0.6, 1.75, 6.0);
  s.addChart(pres.ChartType.bar, [
    {
      name: "Share of deaths",
      labels: ["Hepatitis B", "Hepatitis C", "Alcohol", "MASH / NASH", "Other"],
      values: [37.4, 30.3, 19.1, 8.5, 4.3],
    },
  ], baseChart({
    x: 0.6, y: 2.15, w: 6.0, h: 3.9, barDir: "bar",
    chartColors: [C.blue], dataLabelFormatCode: '0.0"%"',
    valAxisMaxVal: 45, catAxisLabelFontSize: 11.5,
  }));
  s.addText("Every major cause is identifiable — and actionable — decades before the tumour.", {
    x: 0.6, y: 6.2, w: 6.0, h: 0.6, fontFace: BODY, fontSize: 12.5, italic: true,
    color: C.sub, isTextBox: true, margin: 0,
  });
  // right: hero + checklist
  s.addShape("roundRect", { x: 7.1, y: 1.75, w: 5.6, h: 1.7, rectRadius: 0.07, fill: { color: C.tintBurg } });
  s.addText("≥ 60%", {
    x: 7.35, y: 1.9, w: 2.4, h: 1.4, fontFace: TITLE, fontSize: 48, bold: true,
    color: C.burg, isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("of liver cancers are preventable\n(Lancet Commission on Liver Cancer, 2025)", {
    x: 9.75, y: 1.9, w: 2.8, h: 1.4, fontFace: BODY, fontSize: 12.5, color: C.sub,
    isTextBox: true, margin: 0, valign: "middle",
  });
  const rows = [
    ["check-green", "Causes known:", "HBV · HCV · alcohol · MASLD"],
    ["check-green", "Risk groups identifiable", "decades ahead of cancer"],
    ["check-green", "Proven test:", "6-monthly ultrasound ± AFP"],
    ["cross-red", "No organised EU programme:", "liver excluded from the 2022 Council Recommendation on cancer screening"],
  ];
  rows.forEach(([ic, lead, rest], i) => {
    const y = 3.75 + i * 0.72;
    s.addImage({ path: ICON(ic), x: 7.1, y: y + 0.06, w: 0.34, h: 0.34 });
    s.addText(
      [
        { text: lead + " ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.sub } },
      ],
      { x: 7.6, y, w: 5.1, h: 0.66, fontFace: BODY, fontSize: 12.5, isTextBox: true, margin: 0, valign: "top" }
    );
  });
  source(s, "Sources: GBD 2021 · Lancet Commission on Liver Cancer 2025 · EU Council Recommendation on cancer screening, 9 Dec 2022 · EU Group of Chief Scientific Advisors 2022.");
  s.addNotes(
    "Here is the paradox. Liver cancer has the clearest causal chain in oncology: hepatitis B, hepatitis C, alcohol and metabolic liver disease together explain over ninety percent of deaths, and the 2025 Lancet Commission judged at least sixty percent of cases preventable. " +
    "We can identify the at-risk population decades in advance, and a mortality-reducing surveillance test exists. " +
    "Yet when the EU updated its cancer-screening recommendation in December 2022 — adding lung, prostate and gastric — liver was left out, on grounds of 'insufficient evidence'. " +
    "Taiwan's forty years of data speak directly to that evidence gap. Let me show you."
  );
}

// =====================================================================
// SLIDE 4 — Taiwan's starting point
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  WHERE IT BEGAN", "1980s Taiwan: a hyperendemic baseline");
  const ty = 1.8, th = 1.5, tw = 5.5;
  statTile(s, { x: 0.6, y: ty, w: tw, h: th, value: "15–20%", label: "of adults were chronic HBV carriers", tint: C.tintBurg });
  statTile(s, { x: 0.6, y: ty + 1.7, w: tw, h: th, value: "#1", label: "cancer killer in Taiwan for decades", tint: C.tintBurg });
  statTile(s, { x: 0.6, y: ty + 3.4, w: tw, h: th, value: "1.8 M", label: "people still live with chronic HBV today — prevention is a marathon", tint: C.tintGray, vcolor: C.sub });
  chartHeader(s, "Annual HCC incidence per 100,000 men (Beasley cohort, n = 22,707)", 6.6, 1.75, 6.1);
  s.addChart(pres.ChartType.bar, [
    { name: "HCC incidence", labels: ["HBsAg carriers", "Non-carriers"], values: [1158, 5] },
  ], baseChart({
    x: 6.6, y: 2.15, w: 6.1, h: 4.0, barDir: "col",
    chartColors: [C.burgSeries], dataLabelFormatCode: "#,##0",
    valAxisMinVal: 0, valAxisMaxVal: 1200, barGapWidthPct: 110,
    valAxisLabelFormatCode: "#,##0",
  }));
  s.addText("≈ 100× relative risk — the causal proof that made a national policy possible", {
    x: 6.6, y: 6.3, w: 6.1, h: 0.32, fontFace: BODY, fontSize: 12, bold: true,
    color: C.burg, isTextBox: true, margin: 0, align: "center",
  });
  source(s, "Sources: Beasley RP et al., Lancet 1981 · Taiwan Business TOPICS / MOHW 2026 · Chen DS, historical seroprevalence surveys.");
  s.addNotes(
    "Taiwan in the early 1980s looked worse than any EU country does today: fifteen to twenty percent of adults were chronic hepatitis B carriers, and liver cancer was the island's leading cancer killer. " +
    "The turning point was epidemiology done at home: Beasley's cohort of 22,707 Taiwanese men, published in the Lancet in 1981, showed HCC incidence of 1,158 per 100,000 person-years in HBsAg carriers versus 5 in non-carriers — roughly a hundred-fold risk. " +
    "That single, unambiguous causal estimate is what convinced the government to act at population scale. Lesson number zero: national cohort evidence converts into national policy."
  );
}

// =====================================================================
// SLIDE 5 — Pillar 1: Vaccinate
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  PILLAR 1 — VACCINATE", "1984: the world's first universal newborn HBV vaccination");
  chartHeader(s, "HBsAg carrier prevalence after 35 years of vaccination", 0.6, 1.75, 5.9);
  s.addChart(pres.ChartType.bar, [
    { name: "HBsAg prevalence", labels: ["Unvaccinated cohorts", "Vaccinated cohorts"], values: [9.78, 0.64] },
  ], baseChart({
    x: 0.6, y: 2.15, w: 5.9, h: 3.7, barDir: "col",
    chartColors: [C.burgSeries], dataLabelFormatCode: '0.0#"%"',
    valAxisMinVal: 0, valAxisMaxVal: 10, barGapWidthPct: 110,
  }));
  s.addText("national survey, n = 4,717, ages 1–60 (Liver International 2024)", {
    x: 0.6, y: 5.95, w: 5.9, h: 0.3, fontFace: BODY, fontSize: 10.5, color: C.mut,
    isTextBox: true, margin: 0, align: "center",
  });
  const rx = 6.9, rw = 5.8, rh = 1.32;
  statTile(s, { x: rx, y: 1.78, w: rw, h: rh, value: "−69%", label: "HCC in vaccinated cohorts (RR 0.31; 1,958 cases, ages 6–29)", tint: C.tintBurg });
  statTile(s, { x: rx, y: 1.78 + 1.5, w: rw, h: rh, value: "−63% / −82%", label: "HCC mortality / chronic liver-disease mortality in immunised birth cohorts", tint: C.tintBurg, vsize: 24 });
  statTile(s, { x: rx, y: 1.78 + 3.0, w: rw, h: rh, value: "≥ 98%", label: "third-dose coverage sustained since 1989", tint: C.tintGray, vcolor: C.sub });
  s.addShape("roundRect", { x: 0.6, y: 6.35, w: 12.1, h: 0.58, rectRadius: 0.07, fill: { color: C.burg } });
  s.addImage({ path: ICON("syringe-white"), x: 0.95, y: 6.47, w: 0.34, h: 0.34 });
  s.addText("The first vaccine ever shown to prevent a human cancer.", {
    x: 1.5, y: 6.35, w: 11.0, h: 0.58, fontFace: BODY, fontSize: 14, italic: true, bold: true,
    color: C.white, isTextBox: true, margin: 0, valign: "middle",
  });
  source(s, "");
  s.addText("Sources: Chang MH et al., NEJM 1997; JNCI 2009 · Chang KC et al., Liver Int 2024 · J Formos Med Assoc 2025 · Coalition for Global Hepatitis Elimination 2023.", {
    x: 0.6, y: 7.08, w: 12.1, h: 0.3, fontFace: BODY, fontSize: 8.5, color: C.mut, isTextBox: true, margin: 0,
  });
  s.addNotes(
    "Pillar one: primary prevention. In July 1984 Taiwan launched the world's first universal newborn hepatitis B vaccination programme, extended to all newborns by 1986, with third-dose coverage above 98% since 1989. " +
    "Thirty-five years on, the national serosurvey shows carrier prevalence of 0.64% in vaccinated cohorts versus 9.78% in unvaccinated ones. " +
    "The cancer endpoint followed: the landmark NEJM 1997 paper showed childhood HCC falling by half, and the twenty-year JNCI follow-up quantified a 69% reduction — relative risk 0.31. Immunised cohorts now show 63% lower HCC mortality. " +
    "This was the first proof anywhere that a vaccine prevents a human cancer — proof generated by exactly the kind of registry infrastructure I'll return to."
  );
}

// =====================================================================
// SLIDE 6 — Pillar 2: Cure (HCV elimination)
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  PILLAR 2 — CURE", "Hepatitis C eliminated in 2025 — five years ahead of WHO");
  const ty = 1.72, tw = 2.87, th = 1.35, gx = 0.21;
  statTile(s, { x: 0.6, y: ty, w: tw, h: th, value: "182,000+", label: "treated with state-funded DAAs since 2017", tint: C.tintBurg, vsize: 24 });
  statTile(s, { x: 0.6 + (tw + gx), y: ty, w: tw, h: th, value: "98.3%", label: "real-world cure rate (SVR12, national registry)", tint: C.tintBurg, vsize: 24 });
  statTile(s, { x: 0.6 + 2 * (tw + gx), y: ty, w: tw, h: th, value: "90 / 90", label: "diagnosed 90.2% · treated 90.6% — WHO gold tier", tint: C.tintGold, vcolor: C.goldDark, vsize: 24 });
  statTile(s, { x: 0.6 + 3 * (tw + gx), y: ty, w: tw, h: th, value: "€830 M", label: "invested (NT$28.7 bn) — an 'investment in health'", tint: C.tintGray, vcolor: C.sub, vsize: 24 });
  chartHeader(s, "10-year cumulative HCC risk in cirrhosis, by treatment outcome", 0.6, 3.45, 6.3);
  s.addChart(pres.ChartType.bar, [
    { name: "HCC risk", labels: ["Cured (SVR)", "Not cured"], values: [21.7, 38.7] },
  ], baseChart({
    x: 0.6, y: 3.85, w: 6.3, h: 2.85, barDir: "col",
    chartColors: [C.burgSeries], dataLabelFormatCode: '0.0"%"',
    valAxisMaxVal: 45, barGapWidthPct: 110,
  }));
  s.addText("nationwide cohort, n = 27,577 (Dig Dis Sci 2024)", {
    x: 0.6, y: 6.72, w: 6.3, h: 0.28, fontFace: BODY, fontSize: 10.5, color: C.mut,
    isTextBox: true, margin: 0, align: "center",
  });
  const rows = [
    ["people-burgundy", "Micro-elimination first:", "prisons (100% cure per protocol), dialysis units, PWID, remote villages", C.tintBurg],
    ["gov-burgundy", "Single-payer muscle:", "one national insurer lifted all fibrosis restrictions by 2019 — treatment free at point of care", C.tintBurg],
    ["trophy-gold", "Dec 2025:", "Taiwan announces WHO 2030 goal achieved in 2025; elimination report submitted to WHO WPRO", C.tintGold],
  ];
  rows.forEach(([ic, lead, rest, fill], i) => {
    const y = 3.6 + i * 1.06;
    iconCircle(s, ic, 7.2, y, 0.56, fill);
    s.addText(
      [
        { text: lead + " ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.sub } },
      ],
      { x: 7.95, y: y - 0.1, w: 4.75, h: 0.95, fontFace: BODY, fontSize: 12.5, isTextBox: true, margin: 0, valign: "middle" }
    );
  });
  source(s, "Sources: Taiwan MOHW, 17 Dec 2025 · TACR registry, Liver Int 2021 · Dig Dis Sci 2024 · National Hepatitis C Program 2016–2025.");
  s.addNotes(
    "Pillar two: cure the already-infected. Taiwan created a National Hepatitis C Program office in 2016 and began fully reimbursed DAA therapy in January 2017, removing every fibrosis restriction by 2019. " +
    "Over 182,000 patients have been treated at a real-world cure rate of 98.3%, for roughly 830 million euros — framed explicitly as an investment, not a cost. " +
    "Micro-elimination came first: prisons, dialysis units, people who inject drugs, remote villages. " +
    "Cure translates into cancer prevented: ten-year HCC risk in cirrhosis drops from 38.7% to 21.7% after SVR. " +
    "And in December 2025 Taiwan announced it had reached the WHO 2030 elimination goal five years early — diagnosis 90.2%, treatment 90.6%, both above the gold-tier bar. Four EU countries are on track; twenty-three are not."
  );
}

// =====================================================================
// SLIDE 7 — Pillar 3: Screen & surveil
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  PILLAR 3 — SCREEN", "Find the at-risk, then watch them every six months");
  // left: 4-step flow
  const steps = [
    ["search-burgundy", "Test once, free, for life", "national HBV/HCV screening, ages 39–79 — 7.55 M screened"],
    ["link-burgundy", "Reflex + registry", "automatic HCV-RNA confirmation; every positive enrolled for recall"],
    ["ultrasound-burgundy", "Ultrasound + AFP, every 6 months", "for carriers and cirrhosis; 3-monthly if very high risk"],
    ["stetho-burgundy", "Treat early", "curative resection / ablation / transplant while it still works"],
  ];
  steps.forEach(([ic, lead, rest], i) => {
    const y = 1.85 + i * 1.22;
    iconCircle(s, ic, 0.6, y, 0.6, C.tintBurg);
    s.addText(
      [
        { text: lead + "\n", options: { bold: true, color: C.ink, fontSize: 13 } },
        { text: rest, options: { color: C.sub, fontSize: 11.5 } },
      ],
      { x: 1.4, y: y - 0.14, w: 4.9, h: 1.0, fontFace: BODY, isTextBox: true, margin: 0, valign: "middle" }
    );
    if (i < steps.length - 1) {
      s.addImage({ path: ICON("arrow-down-green"), x: 0.78, y: y + 0.72, w: 0.24, h: 0.24 });
    }
  });
  // right: mortality stat + survival chart
  statTile(s, {
    x: 6.7, y: 1.78, w: 6.0, h: 1.05, value: "−31% HCC mortality",
    label: "with invited community ultrasound screening (Changhua, 45–69 y)", tint: C.tintBurg, vsize: 22,
  });
  chartHeader(s, "5-year survival by stage at diagnosis (BCLC) — why early matters", 6.7, 3.05, 6.0);
  s.addChart(pres.ChartType.bar, [
    { name: "5-yr survival", labels: ["0", "A", "B", "C", "D"], values: [70, 58, 34, 11, 4] },
  ], baseChart({
    x: 6.7, y: 3.45, w: 6.0, h: 3.0, barDir: "col",
    chartColors: [C.burgSeries], dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 80,
  }));
  s.addText("Taiwan Cancer Registry, 73,817 patients (2011–2019) · earliest-stage (BCLC 0) detection nearly doubled: 6.2% → 11.3%", {
    x: 6.7, y: 6.5, w: 6.0, h: 0.45, fontFace: BODY, fontSize: 10.5, color: C.mut, isTextBox: true, margin: 0,
  });
  source(s, "Sources: Yeh YP et al., Hepatology 2014 · MOHW / HPA 2025 · Viruses 2023 (registry trends) · JAMA Netw Open 2021 (6-monthly interval).");
  s.addNotes(
    "Pillar three: organised screening and surveillance. The cascade is simple. Every adult aged 39 to 79 gets one free lifetime hepatitis B and C test — over 7.5 million people screened so far. " +
    "Positives get reflex RNA confirmation and registry enrolment, then biannual ultrasound with AFP; very-high-risk patients go three-monthly. " +
    "Does it save lives? The Changhua community programme, with 80% attendance, cut HCC mortality by about 31% in invited residents — randomised-quality evidence at population scale. " +
    "And the registry shows the mechanism: earliest-stage detection has nearly doubled, and five-year survival runs 70% at stage 0 versus 4% at stage D. Six-monthly beats annual — Taiwan tested that too."
  );
}

// =====================================================================
// SLIDE 8 — The payoff (timeline + outcomes)
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  OUTCOMES", "Four decades of policy bent the curve");
  // timeline
  const nodes = [
    ["syringe-burgundy", "1984", "universal newborn vaccination"],
    ["book-burgundy", "1991", "REVEAL cohorts begin"],
    ["pills-burgundy", "2017", "DAAs free for all"],
    ["search-burgundy", "2020", "adult HBV/HCV screening"],
    ["trophy-gold", "2025", "HCV eliminated (WHO goal)"],
  ];
  s.addShape("line", { x: 1.15, y: 2.32, w: 10.9, h: 0, line: { color: C.axis, width: 1.5 } });
  nodes.forEach(([ic, yr, lb], i) => {
    const x = 0.85 + i * 2.72;
    iconCircle(s, ic, x, 2.02, 0.6, i === 4 ? C.tintGold : C.tintBurg);
    s.addText(yr, {
      x: x - 0.66, y: 2.72, w: 1.9, h: 0.3, align: "center", fontFace: TITLE, fontSize: 15,
      bold: true, color: C.burg, isTextBox: true, margin: 0,
    });
    s.addText(lb, {
      x: x - 0.76, y: 3.02, w: 2.1, h: 0.55, align: "center", fontFace: BODY, fontSize: 10.5,
      color: C.sub, isTextBox: true, margin: 0,
    });
  });
  // bottom: chart + tiles
  chartHeader(s, "HCC incidence per 100,000", 0.6, 3.95, 4.3);
  s.addChart(pres.ChartType.bar, [
    { name: "Incidence", labels: ["2013", "2021"], values: [57.8, 45.0] },
  ], baseChart({
    x: 0.6, y: 4.35, w: 4.3, h: 2.5, barDir: "col",
    chartColors: [C.burgSeries], dataLabelFormatCode: "0.0",
    valAxisMinVal: 0, valAxisMaxVal: 60, barGapWidthPct: 110,
  }));
  const tx2 = 5.3, tw2 = 3.6, th2 = 1.18;
  statTile(s, { x: tx2, y: 4.05, w: tw2, h: th2, value: "14 years", label: "of consecutively falling incidence — Taiwan's only major cancer in decline", tint: C.tintBurg, vsize: 22 });
  statTile(s, { x: tx2, y: 4.05 + 1.38, w: tw2, h: th2, value: "Out of top 10", label: "chronic liver disease left Taiwan's top-10 causes of death (2023)", tint: C.tintBurg, vsize: 20 });
  statTile(s, { x: tx2 + tw2 + 0.25, y: 4.05, w: tw2, h: th2, value: "7,513", label: "liver-cancer deaths in 2024 — now 2nd, still falling", tint: C.tintGray, vcolor: C.sub, vsize: 22 });
  statTile(s, { x: tx2 + tw2 + 0.25, y: 4.05 + 1.38, w: tw2, h: th2, value: "−38%", label: "further fall in male incidence projected 2015 → 2035", tint: C.tintGray, vcolor: C.sub, vsize: 22 });
  source(s, "Sources: HPA Cancer Registry 2023 (rel. Dec 2025) · Clinics & Practice 2024 · MOHW 2024 · J Formos Med Assoc 2019 (projection).");
  s.addNotes(
    "Put the three pillars on one timeline — 1984 vaccination, the REVEAL evidence cohorts from 1991, free DAAs from 2017, universal adult screening from 2020, HCV elimination in 2025 — and read the outcomes. " +
    "HCC incidence fell from 57.8 to 45 per 100,000 between 2013 and 2021; the latest registry shows fourteen consecutive years of decline, making liver cancer the only major cancer in Taiwan that is falling. " +
    "Chronic liver disease has dropped out of the national top-ten causes of death for the first time in decades. " +
    "Deaths — 7,513 in 2024 — are still substantial; this is a marathon, and the projection is a further 38% fall in male incidence by 2035. But the curve is unmistakably bent. No other country has this depth of proof."
  );
}

// =====================================================================
// SLIDE 9 — Lessons
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TAIWAN'S EXPERIENCE  ·  LESSONS", "What actually made it work");
  const cards = [
    ["gov-burgundy", "Political commitment, single-payer financing", "Screening was never launched without free treatment behind it. The NHI pays; nobody is found positive and then abandoned."],
    ["link-burgundy", "Registries with call & recall", "Vaccination, hepatitis, screening and cancer registries are linked. Surveillance is organised, not left to whoever remembers."],
    ["microscope-burgundy", "Home-grown evidence engines", "Beasley, the vaccine trials, REVEAL — national cohorts generated the proof, then policy consumed it within years."],
    ["target-burgundy", "Sequence the attack", "Vaccinate the newborn, cure the infected, surveil the at-risk — three pillars, one cascade, each feeding the next."],
  ];
  cards.forEach(([ic, head, body], i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.8 + Math.floor(i / 2) * 2.55, w = 5.9, h = 2.3;
    s.addShape("roundRect", { x, y, w, h, rectRadius: 0.09, fill: { color: C.tintBurg } });
    iconCircle(s, ic, x + 0.3, y + 0.3, 0.62, C.white);
    s.addText(head, {
      x: x + 1.15, y: y + 0.24, w: w - 1.45, h: 0.75, fontFace: BODY, fontSize: 14.5, bold: true,
      color: C.ink, isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(body, {
      x: x + 0.3, y: y + 1.12, w: w - 0.6, h: 1.05, fontFace: BODY, fontSize: 12, color: C.sub,
      isTextBox: true, margin: 0,
    });
  });
  source(s, "Synthesis of: J Formos Med Assoc 2025 (national strategy review) · Viruses 2020 (lessons from Taiwan) · NHCP policy papers 2023–2025.");
  s.addNotes(
    "Before translating to Europe, distil why it worked — because none of it was a technology miracle. " +
    "First, political commitment backed by single-payer financing: screening was always coupled to free treatment. " +
    "Second, linked registries with call and recall — surveillance is a system, not a reminder in a busy clinician's head. " +
    "Third, Taiwan built its own evidence engines — Beasley's cohort, the vaccine follow-up studies, REVEAL — and policy consumed that evidence within a few years of publication. " +
    "Fourth, sequencing: vaccinate newborns, cure the infected, surveil the at-risk. Each pillar feeds the next through the same infrastructure. Keep these four in mind — they are the transferable part."
  );
}

// =====================================================================
// SLIDE 10 — Europe is different
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TRANSLATING TO THE EU", "Europe's liver cancer is a different disease");
  // Taiwan card
  s.addShape("roundRect", { x: 0.6, y: 1.8, w: 5.95, h: 4.15, rectRadius: 0.09, fill: { color: C.tintBurg } });
  iconCircle(s, "liver-burgundy", 0.95, 2.1, 0.6, C.white);
  s.addText("Taiwan: one dominant driver", {
    x: 1.75, y: 2.08, w: 4.6, h: 0.62, fontFace: BODY, fontSize: 16, bold: true, color: C.burg,
    isTextBox: true, margin: 0, valign: "middle",
  });
  const twRows = [
    ["blood-burgundy", "HBV", "~70% of HCC carries hepatitis B — one dominant enemy"],
    ["syringe-burgundy", "Vaccinated", "childhood carrier rate 9.78% → 0.64% since 1984"],
    ["pills-burgundy", "Cured", "HCV eliminated 2025 — the virus-first playbook, completed"],
  ];
  twRows.forEach(([ic, lead, rest], i) => {
    const y = 2.98 + i * 0.98;
    s.addImage({ path: ICON(ic), x: 0.95, y: y + 0.04, w: 0.36, h: 0.36 });
    s.addText(
      [
        { text: lead + ": ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.sub } },
      ],
      { x: 1.47, y, w: 4.9, h: 0.94, fontFace: BODY, fontSize: 12.5, isTextBox: true, margin: 0, valign: "top" }
    );
  });
  // EU card
  s.addShape("roundRect", { x: 6.75, y: 1.8, w: 5.95, h: 4.15, rectRadius: 0.09, fill: { color: C.tintBlue } });
  iconCircle(s, "globe-blue", 7.1, 2.1, 0.6, C.white);
  s.addText("EU: three drivers at once", {
    x: 7.9, y: 2.08, w: 4.6, h: 0.62, fontFace: BODY, fontSize: 16, bold: true, color: C.blueDark,
    isTextBox: true, margin: 0, valign: "middle",
  });
  const euRows = [
    ["blood-blue", "HCV legacy", "dominant in Southern / Western Europe"],
    ["wine-blue", "Alcohol", "up to 50–60% of HCC in Central & Eastern Europe"],
    ["scale-blue", "MASLD", "fastest-growing cause — already #1 in Sweden; 35–50% arises without cirrhosis, outside today's surveillance net"],
  ];
  euRows.forEach(([ic, lead, rest], i) => {
    const y = 2.98 + i * 0.98;
    s.addImage({ path: ICON(ic), x: 7.1, y: y + 0.04, w: 0.36, h: 0.36 });
    s.addText(
      [
        { text: lead + ": ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.sub } },
      ],
      { x: 7.62, y, w: 4.9, h: 0.94, fontFace: BODY, fontSize: 12.5, isTextBox: true, margin: 0, valign: "top" }
    );
  });
  // bottom banner
  s.addShape("roundRect", { x: 0.6, y: 6.2, w: 12.1, h: 0.62, rectRadius: 0.07, fill: { color: C.tintGold } });
  s.addImage({ path: ICON("lightbulb-gold"), x: 0.95, y: 6.33, w: 0.36, h: 0.36 });
  s.addText("Import the system — organised find–treat–surveil — not the virus focus.", {
    x: 1.5, y: 6.2, w: 11.0, h: 0.62, fontFace: BODY, fontSize: 14.5, bold: true, italic: true,
    color: C.ink, isTextBox: true, margin: 0, valign: "middle",
  });
  source(s, "Sources: GBD 2021 · J Hepatol 2019 (ARLD-HCC) · Vaz J et al., Int J Cancer 2025 (Sweden) · Hepatology Communications 2025 · AASLD 2024.");
  s.addNotes(
    "Now the honest caveat: you cannot copy-paste Taiwan, because Europe's liver cancer is aetiologically different. " +
    "Taiwan fought one dominant enemy — seventy percent of its HCC carries hepatitis B — which made a virus-first playbook natural. " +
    "Europe faces three drivers at once: a hepatitis C legacy in the South and West; alcohol, which causes up to half or more of HCC in Central and Eastern Europe; and MASLD, the fastest-growing cause — already the single most common cause of HCC in Sweden. " +
    "Worse, a third to half of MASLD-related HCC arises without cirrhosis, outside every current surveillance criterion. " +
    "So the translation rule is: import the system — organised find, treat, surveil — not the virus focus."
  );
}

// =====================================================================
// SLIDE 11 — The EU gap today
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "TRANSLATING TO THE EU", "The gap: opportunistic care, organised nowhere");
  const rows = [
    ["virus-blue", "5.4 million", "in the EU/EEA live with chronic HBV or HCV — ~60,000 deaths every year"],
    ["search-blue", "Up to 85% / 89%", "of HBV / HCV infections undiagnosed in some Member States"],
    ["cross-red", "Off track", "ECDC 2026: the EU/EEA will miss WHO 2030 hepatitis elimination; only ES · FR · IT · NL on track for HCV"],
    ["warn-amber", "No liver screening programme", "excluded from the 2022 Council Recommendation — surveillance left to chance"],
  ];
  rows.forEach(([ic, lead, rest], i) => {
    const y = 1.85 + i * 1.18;
    s.addImage({ path: ICON(ic), x: 0.6, y: y + 0.05, w: 0.42, h: 0.42 });
    s.addText(
      [
        { text: lead + "  ", options: { bold: true, color: C.ink, fontSize: 15 } },
        { text: rest, options: { color: C.sub, fontSize: 12.5 } },
      ],
      { x: 1.25, y, w: 5.4, h: 1.08, fontFace: BODY, isTextBox: true, margin: 0, valign: "top" }
    );
  });
  chartHeader(s, "HCC surveillance uptake among eligible patients (%)", 7.1, 1.75, 5.6);
  s.addChart(pres.ChartType.bar, [
    { name: "Uptake", labels: ["Study settings", "Real-world average", "Consistent 6-monthly"], values: [70, 39, 24] },
  ], baseChart({
    x: 7.1, y: 2.15, w: 5.6, h: 3.4, barDir: "col",
    chartColors: [C.blue], dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 80, catAxisLabelFontSize: 10.5,
  }));
  statTile(s, {
    x: 7.1, y: 5.75, w: 5.6, h: 1.0, value: "60% incurable",
    label: "of UK cases already beyond cure at first presentation (Lancet Reg Health Eur 2024)",
    tint: C.tintBurg, vsize: 20,
  });
  source(s, "Sources: ECDC 2024/2026 · Polaris Observatory 2025 · J Hepatocell Carcinoma 2024 (adherence meta-analysis) · Lancet Reg Health Eur 2024.");
  s.addNotes(
    "How big is the gap? Five point four million Europeans live with chronic hepatitis B or C, causing about sixty thousand deaths a year — and in some Member States up to 85 and 89 percent of infections are undiagnosed. " +
    "ECDC warned this July that the EU/EEA as a whole will miss the WHO 2030 elimination targets; only Spain, France, Italy and the Netherlands are on track for HCV. " +
    "On the surveillance side, uptake runs about seventy percent in study settings but under forty percent in real life, and barely a quarter of patients get consistent six-monthly scans — because it is opportunistic, with no call-recall system anywhere in the EU. " +
    "The result: in the UK, sixty percent of patients are already incurable at first presentation. Compare that with Taiwan's doubling of stage-0 detection. This is the distance between a programme and a habit."
  );
}

// =====================================================================
// SLIDE 12 — Strategy 1: find & treat
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "FUTURE STRATEGY 1  ·  ELIMINATE THE VIRUSES", "Find the infected — and treat within weeks, free");
  const cards = [
    ["search-blue", "1 · Test once in a lifetime", "Offer every adult one HBV/HCV test (birth-cohort rollout). Taiwan reached 7.55 M with a free lifetime test."],
    ["link-blue", "2 · Reflex + enrol", "Reflex RNA confirmation; every positive auto-enrolled in a national registry with recall — no lost letters."],
    ["people-blue", "3 · Micro-eliminate", "Prisons, PWID, dialysis, migrant communities (EU4Health VH-COMSAVAC shows the model)."],
    ["pills-blue", "4 · Treat free, immediately", "DAA / antiviral linkage in weeks. Screening only works when treatment is free — Taiwan's core rule."],
  ];
  cards.forEach(([ic, head, body], i) => {
    const x = 0.6 + i * 3.13, y = 1.85, w = 2.93, h = 3.3;
    s.addShape("roundRect", { x, y, w, h, rectRadius: 0.09, fill: { color: C.tintBlue } });
    iconCircle(s, ic, x + 0.28, y + 0.28, 0.6, C.white);
    s.addText(head, {
      x: x + 0.25, y: y + 1.0, w: w - 0.5, h: 0.75, fontFace: BODY, fontSize: 13.5, bold: true,
      color: C.blueDark, isTextBox: true, margin: 0,
    });
    s.addText(body, {
      x: x + 0.25, y: y + 1.78, w: w - 0.5, h: 1.4, fontFace: BODY, fontSize: 11, color: C.sub,
      isTextBox: true, margin: 0,
    });
  });
  s.addShape("roundRect", { x: 0.6, y: 5.5, w: 12.1, h: 1.25, rectRadius: 0.07, fill: { color: C.tintGray } });
  s.addText("EU target, borrowed from Taipei:", {
    x: 0.95, y: 5.66, w: 3.6, h: 0.9, fontFace: BODY, fontSize: 13, bold: true, color: C.ink,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText([
    { text: "90% diagnosed · 90% treated by 2030 ", options: { bold: true, fontSize: 20, color: C.blueDark } },
    { text: " — Taiwan hit both in 2025. The 2024 Council Recommendation on vaccine-preventable cancers (HBV) is the policy hook.", options: { fontSize: 12, color: C.sub } },
  ], {
    x: 4.6, y: 5.66, w: 7.8, h: 0.95, fontFace: BODY, isTextBox: true, margin: 0, valign: "middle",
  });
  source(s, "Sources: WHO GHSS 2030 targets · MOHW 2025 · EU4Health VH-COMSAVAC · Council Recommendation on vaccine-preventable cancers, June 2024.");
  s.addNotes(
    "Strategy one is the direct transplant: eliminate the viruses with a four-step cascade. " +
    "One — offer every adult a once-in-a-lifetime hepatitis B and C test, rolled out by birth cohort, exactly as Taiwan did for ages 39 to 79. " +
    "Two — reflex RNA confirmation and automatic registry enrolment, so no positive result dies in an unopened letter. " +
    "Three — micro-elimination in prisons, drug services, dialysis units and migrant communities; the EU4Health VH-COMSAVAC project already pilots this. " +
    "Four — free treatment within weeks; Taiwan's core rule is that screening is only ethical and only effective when treatment is guaranteed. " +
    "The measurable target: ninety percent diagnosed, ninety percent treated by 2030. Taiwan reached both in 2025 — it is achievable within one EU budget cycle."
  );
}

// =====================================================================
// SLIDE 13 — Strategy 2: smarter surveillance
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "FUTURE STRATEGY 2  ·  ORGANISED, RISK-BASED SURVEILLANCE", "Surveillance smart enough for Europe's aetiology");
  chartHeader(s, "Early-stage HCC detection sensitivity (published estimates, %)", 0.6, 1.75, 5.9);
  s.addChart(pres.ChartType.bar, [
    { name: "Sensitivity", labels: ["Ultrasound", "Blood panel (GALAD)", "Abbreviated MRI"], values: [53, 70, 86] },
  ], baseChart({
    x: 0.6, y: 2.15, w: 5.9, h: 3.6, barDir: "col",
    chartColors: [C.blue], dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 100, catAxisLabelFontSize: 10.5,
  }));
  s.addText("different cohorts — indicative, not head-to-head (Acad Radiol 2024; Gastroenterology 2024; meta-analyses 2024–25)", {
    x: 0.6, y: 5.85, w: 5.9, h: 0.45, fontFace: BODY, fontSize: 10, color: C.mut, isTextBox: true, margin: 0,
  });
  const rows = [
    ["target-blue", "Triage by risk score:", "aMAP / PAGE-B (C-index ≈ 0.82) decide who needs watching — and who safely doesn't", C.tintBlue],
    ["ai-blue", "AI-assisted ultrasound:", "Taiwan is deploying 50 AI ultrasound systems to remote areas in 2026 — quality at scale", C.tintBlue],
    ["dna-blue", "LIVERAIM:", "the EU's own engine — 100,000-participant, 6-country RCT of AI-based early detection", C.tintBlue],
    ["check-green", "EASL 2025 already endorses risk-based surveillance", "— what's missing is the organised programme around it: registries, call–recall, QA", C.tintAqua],
  ];
  rows.forEach(([ic, lead, rest, fill], i) => {
    const y = 1.95 + i * 1.13;
    iconCircle(s, ic, 6.9, y, 0.56, fill);
    s.addText(
      [
        { text: lead + " ", options: { bold: true, color: C.ink } },
        { text: rest, options: { color: C.sub } },
      ],
      { x: 7.65, y: y - 0.16, w: 5.1, h: 1.05, fontFace: BODY, fontSize: 12.5, isTextBox: true, margin: 0, valign: "middle" }
    );
  });
  source(s, "Sources: EASL CPG, J Hepatol 2025 · Fan R et al., J Hepatol 2020 (aMAP) · Acad Radiol 2024 (AMRI) · LIVERAIM / IHI 2024 · Taipei Times, Jun 2026.");
  s.addNotes(
    "Strategy two: make surveillance organised and risk-based, so it fits a MASLD-era Europe where the at-risk pool is huge and often non-cirrhotic. " +
    "The toolkit has matured. Risk scores like aMAP and PAGE-B, with C-statistics around 0.82, can triage who needs watching and — just as important for cost — who does not. " +
    "Where ultrasound underperforms — sensitivity around fifty percent for early disease — blood panels like GALAD and abbreviated MRI, at 86% sensitivity, can take over for selected patients. " +
    "AI-assisted ultrasound closes the operator-quality gap; Taiwan is deploying fifty AI systems to remote areas this year. " +
    "And Europe is already building its own evidence engine: LIVERAIM, a hundred-thousand-participant trial across six countries. " +
    "EASL's 2025 guideline endorses risk-based surveillance. The missing piece is not science — it is the programme: registries, call-recall, quality assurance. Exactly the Taiwan lesson."
  );
}

// =====================================================================
// SLIDE 14 — Strategy 3: non-viral frontier
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "FUTURE STRATEGY 3  ·  PREVENT THE NEXT WAVE", "The non-viral frontier: alcohol and MASLD");
  s.addShape("roundRect", { x: 0.6, y: 1.85, w: 4.4, h: 4.6, rectRadius: 0.09, fill: { color: C.tintBurg } });
  s.addText("By 2050", {
    x: 0.95, y: 2.15, w: 3.7, h: 0.4, fontFace: BODY, fontSize: 15, bold: true, color: C.sub,
    isTextBox: true, margin: 0,
  });
  s.addText("1 in 3", {
    x: 0.95, y: 2.55, w: 3.7, h: 1.2, fontFace: TITLE, fontSize: 60, bold: true, color: C.burg,
    isTextBox: true, margin: 0,
  });
  s.addText("liver cancers will stem from alcohol (21%) or MASH (11%)\n\nLancet Commission on Liver Cancer, 2025", {
    x: 0.95, y: 3.85, w: 3.7, h: 2.3, fontFace: BODY, fontSize: 13.5, color: C.sub,
    isTextBox: true, margin: 0,
  });
  const rows = [
    ["wine-burgundy", "Alcohol policy is cancer policy.", "Europe drinks more than any other region; 7.1% of EU cancers are alcohol-attributable. Pricing, labelling and availability rules are liver-cancer screening's silent partner."],
    ["scale-blue", "Find MASLD before cirrhosis.", "MASLD-HCC is projected to rise 47–130% by 2030, and much of it arises without cirrhosis. LiverScreen (30,000 people, 9 countries) shows FibroScan case-finding works at population scale."],
    ["liver-burgundy", "Taiwan is converging on the same problem.", "Its 2026 national plan adds a non-viral early-detection initiative and a tailored liver-cancer risk model — the playbooks are merging."],
  ];
  rows.forEach(([ic, lead, rest], i) => {
    const y = 1.95 + i * 1.55;
    iconCircle(s, ic, 5.35, y, 0.6, i === 1 ? C.tintBlue : C.tintBurg);
    s.addText(
      [
        { text: lead + " ", options: { bold: true, color: C.ink, fontSize: 13.5 } },
        { text: rest, options: { color: C.sub, fontSize: 12 } },
      ],
      { x: 6.15, y: y - 0.12, w: 6.55, h: 1.5, fontFace: BODY, isTextBox: true, margin: 0, valign: "middle" }
    );
  });
  source(s, "Sources: Lancet Commission 2025 · Rumgay H et al., Lancet Oncol 2021 · LiverScreen / Nat Med 2023 · Taipei Times & AmCham TOPICS 2026.");
  s.addNotes(
    "Strategy three looks past the viruses to the wave behind them. By 2050 the Lancet Commission projects one in three liver cancers will stem from alcohol or MASH. " +
    "Two consequences. First, alcohol policy is cancer policy: Europe drinks more than any other region, and seven percent of all EU cancers are alcohol-attributable — pricing, labelling and availability measures belong in the same strategy document as any screening programme. " +
    "Second, find MASLD before cirrhosis: MASLD-driven HCC is projected to rise up to 130 percent by 2030, and much of it develops without cirrhosis, invisible to today's criteria. The LiverScreen project — thirty thousand people FibroScanned across nine countries — proves population case-finding is feasible. " +
    "Interestingly, Taiwan's newest national plan adds exactly this: non-viral early detection and a tailored risk model. The two playbooks are converging — which is precisely why this conference exists."
  );
}

// =====================================================================
// SLIDE 15 — Roadmap
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "PUTTING IT TOGETHER", "A liver-cancer roadmap for the EU, 2026–2040");
  const cols = [
    ["calendar-blue", "2026–2028", "Prove & prepare", [
      "EU feasibility action on liver screening (Beating Cancer Plan mid-term)",
      "Birth-cohort HBV/HCV test-and-treat scale-up",
      "National registries + call–recall build-out",
    ], C.tintBlue],
    ["target-blue", "2029–2033", "Organise", [
      "Risk-based HCC surveillance programme in every Member State",
      "Biomarkers / abbreviated MRI as second line (LIVERAIM read-out)",
      "AI-QA'd ultrasound networks; MASLD case-finding via FibroScan",
    ], C.tintBlue],
    ["flag-blue", "2034–2040", "Deliver", [
      "Hepatitis eliminated as a public-health threat",
      ">50% of HCC detected at a curable stage",
      "Liver-cancer mortality curve bent — Europe's version of Taiwan's 14-year decline",
    ], C.tintAqua],
  ];
  cols.forEach(([ic, yr, head, items, fill], i) => {
    const x = 0.6 + i * 4.13, y = 1.8, w = 3.93, h = 4.0;
    s.addShape("roundRect", { x, y, w, h, rectRadius: 0.09, fill: { color: fill } });
    iconCircle(s, ic, x + 0.28, y + 0.26, 0.56, C.white);
    s.addText(yr, {
      x: x + 1.0, y: y + 0.22, w: w - 1.3, h: 0.34, fontFace: BODY, fontSize: 12.5, bold: true,
      color: C.blueDark, isTextBox: true, margin: 0,
    });
    s.addText(head, {
      x: x + 1.0, y: y + 0.55, w: w - 1.3, h: 0.4, fontFace: TITLE, fontSize: 18, bold: true,
      color: C.ink, isTextBox: true, margin: 0,
    });
    s.addText(
      items.map((t, j) => ({
        text: t,
        options: { bullet: { code: "2022", indent: 10 }, breakLine: j < items.length - 1 },
      })),
      {
        x: x + 0.3, y: y + 1.15, w: w - 0.6, h: 2.7, fontFace: BODY, fontSize: 11.5,
        color: C.sub, isTextBox: true, margin: 0, paraSpaceAfter: 8,
      }
    );
  });
  s.addShape("roundRect", { x: 0.6, y: 6.05, w: 12.1, h: 0.75, rectRadius: 0.07, fill: { color: C.blueDark } });
  s.addText(
    [
      { text: "Measure like Taiwan:  ", options: { bold: true, color: C.white } },
      { text: "90/90 hepatitis diagnosed & treated · >50% early-stage detection · mortality bent by 2035", options: { color: C.white } },
    ],
    { x: 0.95, y: 6.05, w: 11.5, h: 0.75, fontFace: BODY, fontSize: 14, isTextBox: true, margin: 0, valign: "middle" }
  );
  source(s, "Framing: Europe's Beating Cancer Plan 2021 · 2nd EASL–Lancet Commission 2026 · WHO GHSS 2022–2030 · Taiwan MOHW milestones as benchmarks.");
  s.addNotes(
    "Pulling the three strategies into one roadmap. " +
    "2026 to 2028: prove and prepare — an EU feasibility action on liver screening under the Beating Cancer Plan mid-term review, birth-cohort test-and-treat scale-up, and the unglamorous registry and call-recall plumbing. " +
    "2029 to 2033: organise — a risk-based surveillance programme in every Member State, biomarkers and abbreviated MRI as second-line tools once LIVERAIM reads out, AI-quality-assured ultrasound networks, and FibroScan-based MASLD case-finding. " +
    "2034 to 2040: deliver — hepatitis eliminated, more than half of HCC caught at a curable stage, and the mortality curve bent. " +
    "And measure like Taiwan measured: ninety-ninety on hepatitis, early-stage share above fifty percent, mortality bent by 2035. Targets you can put on a dashboard — and be held to."
  );
}

// =====================================================================
// SLIDE 16 — Conclusions (dark)
// =====================================================================
{
  const s = darkSlide();
  s.addText("CONCLUSIONS", {
    x: 0.9, y: 0.55, w: 11.5, h: 0.35, fontFace: BODY, fontSize: 13, bold: true,
    color: C.rose, charSpacing: 3, isTextBox: true, margin: 0,
  });
  s.addText("Liver cancer is a policy-sensitive disease", {
    x: 0.9, y: 0.95, w: 11.6, h: 0.8, fontFace: TITLE, fontSize: 34, bold: true, color: C.white,
    isTextBox: true, margin: 0,
  });
  const rows = [
    ["trophy-white", "Taiwan proved it.", "Vaccinate, cure, screen — HCV eliminated 5 years early, and 14 consecutive years of falling liver-cancer incidence."],
    ["microscope-white", "Europe has the tools — not the programme.", "Risk scores, biomarkers, abbreviated MRI, AI ultrasound, LIVERAIM. What's missing is organisation: registries, call–recall, guaranteed treatment."],
    ["globe-white", "The window is open.", "EASL–Lancet Commission 2026, the Lancet HCC Commission, ECDC's warning and the Beating Cancer Plan mid-term all point one way: put liver on Europe's screening agenda now."],
  ];
  rows.forEach(([ic, lead, rest], i) => {
    const y = 2.15 + i * 1.28;
    iconCircle(s, ic, 0.95, y, 0.62, C.burgMid);
    s.addText(
      [
        { text: lead + "  ", options: { bold: true, color: C.white, fontSize: 15.5 } },
        { text: rest, options: { color: C.rose, fontSize: 13 } },
      ],
      { x: 1.85, y: y - 0.18, w: 10.6, h: 1.2, fontFace: BODY, isTextBox: true, margin: 0, valign: "middle" }
    );
  });
  s.addText("“Taiwan wrote the proof of concept. The European Union can write the sequel.”", {
    x: 0.9, y: 6.1, w: 11.5, h: 0.6, fontFace: TITLE, fontSize: 19, italic: true, color: C.white,
    isTextBox: true, margin: 0,
  });
  s.addText("Thank you  ·  謝謝  ·  Merci  ·  Danke", {
    x: 0.9, y: 6.85, w: 11.5, h: 0.4, fontFace: BODY, fontSize: 14, color: C.rose,
    isTextBox: true, margin: 0,
  });
  s.addNotes(
    "Three sentences to close. Taiwan proved that liver cancer is a policy-sensitive disease: vaccinate, cure and screen, and the curve bends — hepatitis C eliminated five years early, incidence falling for fourteen straight years. " +
    "Europe already has the science and the tools; what it lacks is the organised programme — the registries, the call-recall, the guarantee that a positive test leads to free treatment. " +
    "And the policy window is open right now: two Lancet Commissions, ECDC's warning, and the Beating Cancer Plan mid-term review all converge. " +
    "Taiwan wrote the proof of concept. The European Union can write the sequel. Thank you."
  );
}

// =====================================================================
// SLIDE 17 — Key references
// =====================================================================
{
  const s = lightSlide();
  kickerTitle(s, "FOR THE PROCEEDINGS", "Key references");
  const left = [
    "Beasley RP et al. Hepatocellular carcinoma and HBV: a prospective study of 22,707 men in Taiwan. Lancet 1981.",
    "Chang MH et al. Universal hepatitis B vaccination in Taiwan and the incidence of HCC in children. N Engl J Med 1997.",
    "Chang MH et al. Decreased incidence of HCC in HBV-vaccinated cohorts: 20-year follow-up. J Natl Cancer Inst 2009.",
    "Chen CJ et al. Risk of HCC across a biological gradient of serum HBV DNA level (REVEAL-HBV). JAMA 2006.",
    "Chang KC et al. HBV infection status after 35 years of universal vaccination in Taiwan. Liver Int 2024.",
    "Yeh YP et al. Evaluation of abdominal ultrasonography mass screening for HCC in Taiwan. Hepatology 2014.",
    "National strategy for elimination of liver diseases in Taiwan. J Formos Med Assoc 2025.",
    "Taiwan MOHW. Milestone: hepatitis C elimination achieved (WHO 2030 goal reached in 2025). Press release, 17 Dec 2025.",
  ];
  const right = [
    "Sangro B et al. EASL Clinical Practice Guidelines on the management of hepatocellular carcinoma. J Hepatol 2025.",
    "2nd EASL–Lancet Liver Commission: implementing sustainable liver health in Europe. Lancet 2026.",
    "Lancet Commission on liver cancer: global burden to 2050 and the case for prevention. Lancet 2025.",
    "ECDC. Prevention of hepatitis B and C in the EU/EEA — elimination-target monitoring. 2024 & 2026 reports.",
    "EU Council Recommendation on strengthening cancer screening (2022) · Council Recommendation on vaccine-preventable cancers (2024).",
    "Singal AG et al. HCC surveillance, early detection and survival: systematic review and meta-analysis. J Hepatol 2022.",
    "Fan R et al. aMAP risk score predicts HCC development in patients with chronic hepatitis. J Hepatol 2020.",
    "LIVERAIM Consortium (IHI, 2024) · LiverScreen (Horizon 2020) · Polaris Observatory, Lancet Gastroenterol Hepatol 2025.",
  ];
  const mk = (arr) =>
    arr.map((t, j) => ({
      text: t,
      options: { bullet: { code: "2013", indent: 8 }, breakLine: j < arr.length - 1 },
    }));
  s.addText(mk(left), {
    x: 0.6, y: 1.75, w: 5.95, h: 5.1, fontFace: BODY, fontSize: 10, color: C.sub,
    isTextBox: true, margin: 0, paraSpaceAfter: 7, valign: "top",
  });
  s.addText(mk(right), {
    x: 6.85, y: 1.75, w: 5.9, h: 5.1, fontFace: BODY, fontSize: 10, color: C.sub,
    isTextBox: true, margin: 0, paraSpaceAfter: 7, valign: "top",
  });
  s.addNotes("Reference slide for the proceedings — not presented live.");
}

pres.writeFile({ fileName: path.join(__dirname, "liver-cancer-screening-eu-taiwan.pptx") }).then(() => {
  console.log("deck written");
});
