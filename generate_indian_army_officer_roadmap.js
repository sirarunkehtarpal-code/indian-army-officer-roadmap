import fs from "fs";
import path from "path";
import { fileURLToPath, pathToFileURL } from "url";
import pptxgen from "pptxgenjs";

// ---------- Palette ----------
const DARK = "1E2B1E";
const PRIMARY = "35502B";
const PRIMARY2 = "4C6A3D";
const GOLD = "B8925A";
const CARD_TINT = "EEF1E9";
const CARD_TINT2 = "F4EFE4";
const MUTED = "5B6B57";
const WHITE = "FFFFFF";
const LINE = "D8DED2";
const GRAY = "9AA79A";

const FONT = "Georgia";
const FONT_BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.333;
const PAGE_H = 7.5;

function footer(slide, label) {
  slide.addText(label || "", {
    x: 0.5,
    y: 7.12,
    w: 9,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 9,
    color: MUTED,
    align: "left",
    isTextBox: true,
    margin: 0,
  });
}

function header(slide, kicker, title) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: 0.6,
      y: 0.42,
      w: 11,
      h: 0.32,
      fontFace: FONT_BODY,
      fontSize: 12,
      bold: true,
      color: GOLD,
      charSpacing: 2,
      isTextBox: true,
      margin: 0,
    });
  }

  slide.addText(title, {
    x: 0.6,
    y: kicker ? 0.72 : 0.5,
    w: 11.6,
    h: 0.7,
    fontFace: FONT,
    fontSize: 28,
    bold: true,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });
}

function badge(slide, letter, x, y, size, fill) {
  slide.addShape("ellipse", {
    x,
    y,
    w: size,
    h: size,
    fill: { color: fill || PRIMARY },
    line: { type: "none" },
  });

  slide.addText(letter, {
    x,
    y,
    w: size,
    h: size,
    fontFace: FONT,
    fontSize: size * 32,
    bold: true,
    color: WHITE,
    align: "center",
    valign: "middle",
    isTextBox: true,
    margin: 0,
  });
}

function dataTable(slide, x, y, w, h, colW, headerRow, bodyRows, rowH, fontSize) {
  const rows = [
    headerRow.map((t) => ({
      text: t,
      options: { bold: true, color: WHITE, fill: { color: PRIMARY } },
    })),
  ].concat(
    bodyRows.map((r, i) => {
      const fill = i % 2 === 0 ? WHITE : CARD_TINT;
      return r.map((c, ci) => ({
        text: c,
        options: { color: DARK, fill: { color: fill }, bold: ci === 0 },
      }));
    })
  );

  slide.addTable(rows, {
    x,
    y,
    w,
    h,
    colW,
    fontFace: FONT_BODY,
    fontSize: fontSize || 11.5,
    valign: "middle",
    border: { type: "solid", color: LINE, pt: 0.75 },
    autoPage: false,
    rowH,
  });
}

// ============================================================
// SLIDE 1 — Cover
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: PRIMARY };
  s.addShape("rect", { x: 0, y: 0, w: PAGE_W, h: PAGE_H, fill: { color: PRIMARY }, line: { type: "none" } });
  s.addText("PERMANENT COMMISSION PLANNING", {
    x: 0.9,
    y: 2.0,
    w: 10,
    h: 0.4,
    fontFace: FONT_BODY,
    fontSize: 14,
    bold: true,
    color: GOLD,
    charSpacing: 3,
    isTextBox: true,
    margin: 0,
  });
  s.addText("Indian Army Officer Roadmap", {
    x: 0.85,
    y: 2.4,
    w: 11.5,
    h: 1.3,
    fontFace: FONT,
    fontSize: 44,
    bold: true,
    color: WHITE,
    isTextBox: true,
    margin: 0,
  });
  s.addText("NDA  •  TES  •  CDS (IMA)  •  TGC", {
    x: 0.9,
    y: 3.6,
    w: 10,
    h: 0.5,
    fontFace: FONT_BODY,
    fontSize: 20,
    color: "D9E2D1",
    isTextBox: true,
    margin: 0,
  });
  s.addShape("line", { x: 0.9, y: 4.35, w: 3.2, h: 0, line: { color: GOLD, width: 1.5 } });
  s.addText("DOB 20 September 2011  •  Currently Class 10 (2026–27)", {
    x: 0.9,
    y: 4.55,
    w: 9.5,
    h: 0.45,
    fontFace: FONT_BODY,
    fontSize: 15,
    bold: true,
    color: WHITE,
    isTextBox: true,
    margin: 0,
  });
  s.addText("A decade-long, four-entry plan for Permanent Commission — with age windows calculated exactly against each exam's real cut-off mechanics.", {
    x: 0.9,
    y: 5.05,
    w: 10,
    h: 0.6,
    fontFace: FONT_BODY,
    fontSize: 13,
    color: "C9D4C1",
    isTextBox: true,
    margin: 0,
  });
  s.addText("Prepared as a family planning guide  •  Confirm exact cut-off dates each cycle on joinindianarmy.nic.in and upsc.gov.in", {
    x: 0.9,
    y: 6.75,
    w: 11,
    h: 0.4,
    fontFace: FONT_BODY,
    fontSize: 10.5,
    italic: true,
    color: "9FB093",
    isTextBox: true,
    margin: 0,
  });
}

// ============================================================
// SLIDE 2 — Agenda
// ============================================================
{
  const s = pres.addSlide();
  header(s, "Overview", "What This Plan Covers");

  const items = [
    ["1", "The Four PC Entries", "NDA, TES, CDS (IMA) and TGC — how each one works"],
    ["2", "Why September Matters", "How the real age cut-off mechanics shift his exact attempts"],
    ["3", "Age & Attempt Calculator", "Exact eligibility windows, recalculated precisely"],
    ["4", "Year-by-Year Timeline", "Every stage from Class 10 (now) to age 27 (2038)"],
    ["5", "Preparation Roadmap", "What to focus on, stage by stage, starting now"],
    ["6", "Strategy by Phase", "What to prioritise, and when, across three phases"],
    ["7", "The Complete Journey", "A visual walk-through from school to commissioning"],
    ["8", "Final Takeaways", "The single best path for maximum PC chances"],
  ];

  const colW = 5.65;
  const gapX = 0.35;
  const gapY = 0.2;
  const cardH = 1.12;

  items.forEach((it, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * (colW + gapX);
    const y = 1.55 + row * (cardH + gapY);

    s.addShape("roundRect", { x, y, w: colW, h: cardH, rectRadius: 0.08, fill: { color: CARD_TINT }, line: { type: "none" } });
    badge(s, it[0], x + 0.24, y + 0.24, 0.5, PRIMARY2);
    s.addText(it[1], {
      x: x + 0.95,
      y: y + 0.14,
      w: colW - 1.2,
      h: 0.36,
      fontFace: FONT,
      fontSize: 14.5,
      bold: true,
      color: DARK,
      isTextBox: true,
      margin: 0,
    });
    s.addText(it[2], {
      x: x + 0.95,
      y: y + 0.5,
      w: colW - 1.2,
      h: 0.55,
      fontFace: FONT_BODY,
      fontSize: 11,
      color: MUTED,
      isTextBox: true,
      margin: 0,
      valign: "top",
    });
  });

  footer(s, "");
}

// ============================================================
// SLIDES 3-6 — Entry detail pages
// ============================================================
function entrySlide(letter, name, subtitle, facts, whyTitle, whyBody, badgeColor) {
  const s = pres.addSlide();
  badge(s, letter, 0.6, 0.5, 0.62, badgeColor);
  s.addText(name, {
    x: 1.4,
    y: 0.44,
    w: 9.5,
    h: 0.5,
    fontFace: FONT,
    fontSize: 26,
    bold: true,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });
  s.addText(subtitle, {
    x: 1.4,
    y: 0.92,
    w: 10.5,
    h: 0.35,
    fontFace: FONT_BODY,
    fontSize: 13,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  const leftX = 0.6;
  const leftY = 1.65;
  const leftW = 7.1;
  s.addShape("roundRect", {
    x: leftX,
    y: leftY,
    w: leftW,
    h: 5.15,
    rectRadius: 0.08,
    fill: { color: WHITE },
    line: { color: LINE, width: 1 },
  });

  let fy = leftY + 0.32;
  facts.forEach((f) => {
    s.addText(f[0].toUpperCase(), {
      x: leftX + 0.35,
      y: fy,
      w: leftW - 0.7,
      h: 0.28,
      fontFace: FONT_BODY,
      fontSize: 11,
      bold: true,
      color: GOLD,
      charSpacing: 1,
      isTextBox: true,
      margin: 0,
    });
    s.addText(f[1], {
      x: leftX + 0.35,
      y: fy + 0.28,
      w: leftW - 0.7,
      h: f[2] || 0.55,
      fontFace: FONT_BODY,
      fontSize: 13.5,
      color: DARK,
      isTextBox: true,
      margin: 0,
      valign: "top",
    });
    fy += (f[2] || 0.55) + 0.32;
  });

  const rightX = 7.95;
  const rightY = 1.65;
  const rightW = 4.78;
  s.addShape("roundRect", {
    x: rightX,
    y: rightY,
    w: rightW,
    h: 5.15,
    rectRadius: 0.08,
    fill: { color: CARD_TINT2 },
    line: { type: "none" },
  });
  s.addText(whyTitle, {
    x: rightX + 0.35,
    y: rightY + 0.32,
    w: rightW - 0.7,
    h: 0.6,
    fontFace: FONT,
    fontSize: 17,
    bold: true,
    color: PRIMARY,
    isTextBox: true,
    margin: 0,
  });
  s.addText(whyBody, {
    x: rightX + 0.35,
    y: rightY + 0.95,
    w: rightW - 0.7,
    h: 3.9,
    fontFace: FONT_BODY,
    fontSize: 13,
    color: DARK,
    isTextBox: true,
    margin: 0,
    valign: "top",
    lineSpacingMultiple: 1.25,
  });

  footer(s, "Entry " + letter + " — " + name);
}

entrySlide(
  "A",
  "NDA",
  "National Defence Academy — the earliest, broadest door",
  [
    ["Qualification", "Class 12 pass or appearing (any stream, for Army wing)"],
    ["Age Window", "16½ – 19½ years — checked on the course's start date (1 Jan or 1 Jul), not the exam date", 0.75],
    ["Selection", "UPSC written exam (Maths + General Ability) → SSB Interview → Medical", 0.55],
    ["Training", "3 years at NDA, Khadakwasla → 1 year at IMA Dehradun (4 years total)", 0.55],
    ["Commission", "Permanent Commission, confirmed"],
  ],
  "Why It's the Best Entry",
  "Longest grooming time of any route, and the widest choice of Arms & Services — including the Corps of Signals — decided at the end by merit and preference. He isn't locked into a technical stream early.",
  PRIMARY
);

entrySlide(
  "B",
  "TES",
  "10+2 Technical Entry Scheme — direct to engineering",
  [
    ["Qualification", "Class 12 with PCM, minimum 60% aggregate, plus a valid JEE (Mains) score", 0.55],
    ["Age Window", "Same 16½ – 19½ year band as NDA, on the same course-start reference dates", 0.55],
    ["Selection", "Shortlist on 12th PCM% + JEE Mains rank → SSB Interview → Medical (no written exam)", 0.55],
    ["Training", "3 years B.Tech at CME Pune / MCEME Secunderabad / MCTE Mhow → 1 year IMA (4 years total)", 0.75],
    ["Commission", "Permanent Commission, rank of Lieutenant"],
  ],
  "Signals Advantage",
  "TES is one of the direct feeder entries into the Corps of Signals, Corps of Engineers and EME. Strong PCM / JEE performance genuinely improves his shot at a technical Arm, if that's where his interest lies.",
  PRIMARY2
);

entrySlide(
  "C",
  "CDS (IMA)",
  "Combined Defence Services — the graduate-entry route into IMA",
  [
    ["Qualification", "Graduation in any discipline from a recognised university; final-year students may also apply", 0.55],
    ["Age Window", "19 – 24 years, checked on course commencement — treat the calendar window as approximate within ±6 months at each edge"],
    ["Selection", "UPSC written exam → SSB Interview → Medical"],
    ["Training", "1.5 years at IMA, Dehradun — same academy and same PC status as NDA cadets", 0.55],
    ["Commission", "Permanent Commission (IMA is PC-only; OTA under CDS is SSC and is excluded here)", 0.55],
  ],
  "Why It Matters",
  "The natural backup / second chance if NDA or TES don't work out. It's a twice-a-year exam with a wide, multi-year eligibility window — a generous safety net running well into his twenties.",
  PRIMARY
);

entrySlide(
  "D",
  "TGC",
  "Technical Graduate Course — direct entry for engineers",
  [
    ["Qualification", "BE / B.Tech (or final year) in a notified engineering discipline"],
    ["Age Window", "20 – 27 years, checked on course commencement — also treat as approximate within ±6 months"],
    ["Selection", "Shortlist on engineering marks → SSB Interview → Medical (no written exam)", 0.55],
    ["Training", "1 year at IMA, Dehradun — shorter than CDS, plus 1 year ante-date seniority as a reward for his engineering degree", 0.75],
    ["Commission", "Permanent Commission, rank of Lieutenant"],
  ],
  "Signals Advantage",
  "Like TES, TGC is a preferred feeder for Signals, EME and Engineers. An Electronics / Comm / CS / IT engineering degree is a genuine edge here — and this window stays open the longest of all four entries.",
  PRIMARY2
);

// ============================================================
// SLIDE 7 — Why September Matters (explainer + timeline visual)
// ============================================================
{
  const s = pres.addSlide();
  header(s, "The Fix", "Why September Matters");
  s.addText("UPSC doesn't check age on the exam date — it checks age on the day the course actually starts (1 Jan or 1 Jul for NDA). That fixed reference point, not his birthday's position in the calendar, decides exactly which sittings count.", {
    x: 0.6,
    y: 1.3,
    w: 12.1,
    h: 0.55,
    fontFace: FONT_BODY,
    fontSize: 13,
    color: DARK,
    isTextBox: true,
    margin: 0,
    lineSpacingMultiple: 1.2,
  });

  const lineY = 3.55;
  const x0 = 1.0;
  const x1 = 12.3;
  s.addShape("line", { x: x0, y: lineY, w: x1 - x0, h: 0, line: { color: GRAY, width: 2 } });

  const points = [
    { frac: 0.11, color: GRAY, label: "NDA II 2027\nexam — Sept 2027", above: false, ok: false },
    { frac: 0.28, color: GRAY, label: "NDA I 2028\nexam — Apr 2028", above: true, ok: false },
    { frac: 0.44, color: PRIMARY, label: "NDA II 2028\nexam — Sept 2028", above: false, ok: true },
    { frac: 0.61, color: PRIMARY, label: "NDA I 2029\nexam — Apr 2029", above: true, ok: true },
    { frac: 0.78, color: PRIMARY, label: "NDA II 2029\nexam — Sept 2029", above: false, ok: true },
    { frac: 0.94, color: PRIMARY, label: "NDA I 2030\nexam — Apr 2030", above: true, ok: true },
  ];

  points.forEach((p) => {
    const x = x0 + p.frac * (x1 - x0);
    s.addShape("ellipse", {
      x: x - 0.09,
      y: lineY - 0.09,
      w: 0.18,
      h: 0.18,
      fill: { color: p.color },
      line: { color: WHITE, width: 1.5 },
    });

    const ty = p.above ? lineY - 0.85 : lineY + 0.22;
    s.addText(p.label, {
      x: x - 0.85,
      y: ty,
      w: 1.7,
      h: 0.6,
      fontFace: FONT_BODY,
      fontSize: 9.5,
      bold: p.ok,
      color: p.ok ? PRIMARY : GRAY,
      align: "center",
      isTextBox: true,
      margin: 0,
      valign: p.above ? "bottom" : "top",
    });
  });

  s.addText("16½ yrs\n20 Mar 2028", {
    x: x0 - 0.7,
    y: lineY + 0.28,
    w: 1.6,
    h: 0.5,
    fontFace: FONT_BODY,
    fontSize: 10,
    bold: true,
    color: DARK,
    align: "left",
    isTextBox: true,
    margin: 0,
  });
  s.addText("19½ yrs\n20 Mar 2031", {
    x: x1 - 0.9,
    y: lineY + 0.28,
    w: 1.6,
    h: 0.5,
    fontFace: FONT_BODY,
    fontSize: 10,
    bold: true,
    color: DARK,
    align: "right",
    isTextBox: true,
    margin: 0,
  });

  s.addShape("ellipse", { x: 0.6, y: 5.35, w: 0.14, h: 0.14, fill: { color: PRIMARY }, line: { type: "none" } });
  s.addText("Age-eligible AND he's 12th-appearing/passed — a real, usable attempt", {
    x: 0.85,
    y: 5.27,
    w: 5.6,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 11,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });
  s.addShape("ellipse", { x: 6.7, y: 5.35, w: 0.14, h: 0.14, fill: { color: GRAY }, line: { type: "none" } });
  s.addText("Age-eligible, but he's still in Class 11 — doesn't qualify academically yet", {
    x: 6.95,
    y: 5.27,
    w: 5.6,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 11,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  s.addShape("roundRect", {
    x: 0.6,
    y: 5.85,
    w: 12.13,
    h: 1.05,
    rectRadius: 0.08,
    fill: { color: CARD_TINT2 },
    line: { type: "none" },
  });
  s.addText("Net result: 4 real attempts, not 3 or 6 — NDA II 2028, NDA I 2029, NDA II 2029, NDA I 2030. TES follows the same twice-yearly rhythm alongside these.", {
    x: 0.9,
    y: 5.85,
    w: 11.5,
    h: 1.05,
    fontFace: FONT,
    fontSize: 14,
    bold: true,
    color: PRIMARY,
    isTextBox: true,
    margin: 0,
    valign: "middle",
  });

  footer(s, "Always re-confirm exact cut-off dates against that cycle's official UPSC notification.");
}

// ============================================================
// SLIDE 8 — Age & Attempt Calculator
// ============================================================
{
  const s = pres.addSlide();
  header(s, "Planning Data", "Age & Attempt Calculator");
  s.addText("Based on DOB 20 September 2011, currently Class 10 (2026–27)", {
    x: 0.6,
    y: 1.28,
    w: 9,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  dataTable(
    s,
    0.6,
    1.75,
    12.13,
    3.05,
    [1.5, 1.9, 2.85, 2.65, 3.23],
    ["Entry", "Eligibility (age)", "Exact Age Window", "Realistic Attempts", "Which Sittings"],
    [
      ["NDA", "16½–19½ yrs", "20 Mar 2028 – 20 Mar 2031", "4 solid attempts", "NDA II'28, NDA I'29, NDA II'29, NDA I'30"],
      ["TES", "16½–19½ yrs", "20 Mar 2028 – 20 Mar 2031", "3–4 attempts, parallel to NDA", "Same cycles, from Class 12 (2028–29) onward"],
      ["CDS (IMA)", "19–24 yrs", "~20 Sep 2030 – 20 Sep 2035", "6–8 attempts (twice/yr)", "From final year of degree (~2032) through 2035"],
      ["TGC", "20–27 yrs", "~20 Sep 2031 – 20 Sep 2038", "Many cycles available", "From final year of degree (~2032) through 2038"],
    ],
    0.5,
    11
  );

  s.addShape("roundRect", { x: 0.6, y: 5.05, w: 12.13, h: 1.6, rectRadius: 0.08, fill: { color: CARD_TINT2 }, line: { type: "none" } });
  s.addText("Key Takeaway", {
    x: 0.95,
    y: 5.27,
    w: 5,
    h: 0.35,
    fontFace: FONT,
    fontSize: 15,
    bold: true,
    color: PRIMARY,
    isTextBox: true,
    margin: 0,
  });
  s.addText("Because he's already in Class 10, his first real NDA/TES shot lands in September 2028 — not 2029. CDS and TGC windows are pure age arithmetic from his birthday, but their exact calendar edges can shift by up to 6 months depending on the notification, so treat those two as directional and reconfirm closer to the date.", {
    x: 0.95,
    y: 5.62,
    w: 11.4,
    h: 0.95,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    color: DARK,
    isTextBox: true,
    margin: 0,
    lineSpacingMultiple: 1.2,
  });

  footer(s, "Age & Attempt Calculator");
}

// ============================================================
// SLIDES 9-10 — Year-by-year timeline (calendar years, corrected)
// ============================================================
function timelineSlide(subtitle, dataRows) {
  const s = pres.addSlide();
  header(s, "Planning Data", "Year-by-Year Timeline");
  s.addText(subtitle, {
    x: 0.6,
    y: 1.28,
    w: 9,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  dataTable(
    s,
    0.6,
    1.72,
    12.13,
    4.9,
    [1.0, 0.85, 2.5, 4.13, 3.65],
    ["Year", "Age", "Stage", "Exams He Can Attempt", "Notes"],
    dataRows,
    0.7,
    11
  );

  footer(s, "Year-by-Year Timeline — ages shown are the age he turns via his Sept 20 birthday that year");
}

timelineSlide("Part 1 of 2 — 2026–2032 (currently Class 10)", [
  ["2026", "14–15", "Class 10 (now)", "—", "Foundation phase — start fitness routine and Maths/GK groundwork"],
  ["2027", "15–16", "Class 11", "—", "Lock PCM if leaning TES; begin structured NDA-pattern practice"],
  ["2028", "16–17", "Class 12 (from April)", "NDA II 2028 (Sept)", "His first real attempt — 12th-appearing satisfies the qualification"],
  ["2029", "17–18", "Class 12 boards (Jan–Mar) → degree begins", "NDA I 2029 (Apr), NDA II 2029 (Sept), JEE Mains", "Biggest year — boards, JEE Mains, and two NDA sittings"],
  ["2030", "18–19", "Degree Year 1–2", "NDA I 2030 (Apr) — last shot", "CDS-IMA minimum age (19) crossed this year"],
  ["2031", "19–20", "Degree Year 2–3", "—", "TGC minimum age (20) reached; NDA/TES window has closed"],
  ["2032", "20–21", "Degree Year 3–4 (final year)", "CDS & TGC (if final year)", "First graduate-entry attempts begin, provisionally as a final-year student"],
]);

timelineSlide("Part 2 of 2 — 2033–2038 (graduate-entry years)", [
  ["2033", "21–22", "Graduated (engineering)", "CDS, TGC", "Both entries running fully in parallel"],
  ["2034", "22–23", "Early career / postgrad", "CDS, TGC", "—"],
  ["2035", "23–24", "Early career / postgrad", "CDS (last eligible year), TGC", "Last realistic chance for CDS–IMA"],
  ["2036", "24–25", "Career", "TGC only", "CDS window has closed"],
  ["2037", "25–26", "Career", "TGC only", "—"],
  ["2038", "26–27", "Career", "TGC (last eligible year)", "Final PC door before age 27 closes"],
]);

// ============================================================
// SLIDE 11 — Preparation Roadmap
// ============================================================
{
  const s = pres.addSlide();
  header(s, "Getting Ready", "Preparation Roadmap by Stage");
  s.addText("What to actually focus on, starting now — not just which exams to sit", {
    x: 0.6,
    y: 1.28,
    w: 10,
    h: 0.3,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  dataTable(
    s,
    0.6,
    1.7,
    12.13,
    4.95,
    [2.35, 9.78],
    ["Stage", "Focus Areas"],
    [
      ["Now — Class 10\n(2026–27)", "Build Maths + English fluency for NDA's General Ability paper. Start a daily fitness routine (push-ups, sit-ups, running) aimed at SSB physical standards. Add light daily current-affairs reading."],
      ["Class 11\n(2027–28)", "Firm up PCM if leaning toward TES. Begin structured NDA written-exam mock tests. Track his 1.6 km run timing against SSB benchmarks."],
      ["Class 12\n(2028–29)", "Sit NDA II 2028 in September — his first real attempt. Prep JEE Mains (Jan/Apr 2029) in parallel for TES. Keep PCM aggregate above 60% for TES eligibility. Attempt NDA I 2029 right after boards."],
      ["Post-12th\n(2029–30)", "Attempt NDA II 2029 and NDA I 2030 while starting his engineering degree. If selected at any point, he joins NDA or CME immediately — training begins right away."],
      ["Engineering degree\n(2030–33)", "Keep marks strong — TGC shortlists on engineering marks alone. Maintain fitness continuously. From 2nd–3rd year, start SSB-specific prep: group tasks, psychology tests, mock interviews."],
      ["Final year onward\n(2032+)", "Apply for CDS and TGC every single cycle. Stay on top of medical fitness — eyesight, BMI, dental — the most common reason candidates get dropped after clearing the SSB."],
    ],
    0.78,
    11.5
  );

  footer(s, "Preparation Roadmap");
}

// ============================================================
// SLIDE 12 — Strategy Phase 1
// ============================================================
{
  const s = pres.addSlide();
  badge(s, "1", 0.6, 0.5, 0.62, PRIMARY);
  s.addText("Phase 1 — School-Leaving Attempts (2028–2030)", {
    x: 1.4,
    y: 0.5,
    w: 10.8,
    h: 0.6,
    fontFace: FONT,
    fontSize: 24,
    bold: true,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });
  s.addText("Go for both NDA and TES, starting with NDA II in September 2028", {
    x: 1.4,
    y: 1.05,
    w: 11,
    h: 0.35,
    fontFace: FONT_BODY,
    fontSize: 14,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  const points = [
    ["Focus on NDA first", "It's the broadest entry (any stream), gives the longest training and the best-known SSB conversion track record, and keeps every Arm — including Signals — open based on final merit."],
    ["Run TES in parallel, not instead", "If he studies PCM and keeps JEE Mains prep going anyway (which also helps NDA's Maths paper), TES essentially comes free alongside NDA prep — no extra written exam needed."],
    ["4 real attempts, 2028–2030", "NDA II 2028, NDA I & II 2029, and NDA I 2030 — the widest opening he will have, with the first one arriving as soon as he starts Class 12."],
  ];

  let y = 1.75;
  points.forEach((p) => {
    s.addShape("roundRect", { x: 0.6, y, w: 12.13, h: 1.55, rectRadius: 0.08, fill: { color: CARD_TINT }, line: { type: "none" } });
    s.addShape("roundRect", { x: 0.9, y: y + 0.28, w: 0.08, h: 1.0, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(p[0], { x: 1.2, y: y + 0.2, w: 11, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
    s.addText(p[1], { x: 1.2, y: y + 0.62, w: 11.3, h: 0.85, fontFace: FONT_BODY, fontSize: 13, color: DARK, isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.2 });
    y += 1.75;
  });

  footer(s, "Strategy — Phase 1");
}

// ============================================================
// SLIDE 13 — Strategy Phase 2 (degree choice)
// ============================================================
{
  const s = pres.addSlide();
  badge(s, "2", 0.6, 0.5, 0.62, PRIMARY2);
  s.addText("Phase 2 — Choosing His Degree Wisely (2029 onward)", {
    x: 1.4,
    y: 0.5,
    w: 10.8,
    h: 0.6,
    fontFace: FONT,
    fontSize: 24,
    bold: true,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });
  s.addText("He enrolls in 2029 right after boards — even while still attempting NDA I 2030 as a parallel shot", {
    x: 1.4,
    y: 1.05,
    w: 11.2,
    h: 0.4,
    fontFace: FONT_BODY,
    fontSize: 13.5,
    italic: true,
    color: MUTED,
    isTextBox: true,
    margin: 0,
  });

  const cardY = 1.75;
  const cardW = 5.85;
  const cardH = 3.1;

  s.addShape("roundRect", { x: 0.6, y: cardY, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: CARD_TINT2 }, line: { color: GOLD, width: 1.25 } });
  s.addText("Engineering Degree (B.Tech / BE)", { x: 0.9, y: cardY + 0.25, w: cardW - 0.6, h: 0.6, fontFace: FONT, fontSize: 17, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
  s.addText("Recommended", { x: 0.9, y: cardY + 0.75, w: 2.2, h: 0.35, fontFace: FONT_BODY, fontSize: 11, bold: true, color: WHITE, fill: { color: GOLD }, align: "center", isTextBox: true, margin: 2 });
  s.addText("Keeps BOTH CDS and TGC open. The strictly safer choice — it doesn't shut any Permanent Commission door. Ideally Electronics / Comm / CS / IT if he's drawn to Signals.", {
    x: 0.9,
    y: cardY + 1.35,
    w: cardW - 0.6,
    h: 1.6,
    fontFace: FONT_BODY,
    fontSize: 13,
    color: DARK,
    isTextBox: true,
    margin: 0,
    valign: "top",
    lineSpacingMultiple: 1.25,
  });

  const x2 = 6.88;
  s.addShape("roundRect", { x: x2, y: cardY, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: CARD_TINT }, line: { type: "none" } });
  s.addText("Non-Engineering Degree (BA / BSc / BCom)", { x: x2 + 0.3, y: cardY + 0.25, w: cardW - 0.6, h: 0.6, fontFace: FONT, fontSize: 17, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s.addText("Keeps only CDS open — TGC needs an engineering degree, so this path closes one of his four doors early.", {
    x: x2 + 0.3,
    y: cardY + 1.0,
    w: cardW - 0.6,
    h: 1.9,
    fontFace: FONT_BODY,
    fontSize: 13,
    color: DARK,
    isTextBox: true,
    margin: 0,
    valign: "top",
    lineSpacingMultiple: 1.25,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.15, w: 12.13, h: 1.35, rectRadius: 0.08, fill: { color: PRIMARY }, line: { type: "none" } });
  s.addText("Recommendation: choose engineering — it's strictly the safer choice since it doesn't shut any PC door.", {
    x: 0.95,
    y: 5.15,
    w: 11.4,
    h: 1.35,
    fontFace: FONT,
    fontSize: 16,
    bold: true,
    color: WHITE,
    isTextBox: true,
    margin: 0,
    valign: "middle",
  });

  footer(s, "Strategy — Phase 2");
}

// ============================================================
// SLIDE 14 — Strategy Phase 3 & Backup
// ============================================================
{
  const s = pres.addSlide();
  badge(s, "3", 0.6, 0.5, 0.62, PRIMARY);
  s.addText("Phase 3 — Graduate-Entry Attempts (2032–2038) & Backup Plan", {
    x: 1.4,
    y: 0.5,
    w: 11,
    h: 0.6,
    fontFace: FONT,
    fontSize: 22,
    bold: true,
    color: DARK,
    isTextBox: true,
    margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 1.55, w: 12.13, h: 2.1, rectRadius: 0.08, fill: { color: CARD_TINT }, line: { type: "none" } });
  s.addText("Go for CDS and TGC together", { x: 0.95, y: 1.8, w: 11.4, h: 0.4, fontFace: FONT, fontSize: 17, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
  s.addText("Attempt CDS and TGC simultaneously every cycle from his final year onward (provisional applications are allowed). TGC has the shorter, faster training (1 year) with bonus seniority; CDS-IMA is the broader graduate entry. This phase alone gives him well over 10 combined attempts across several years.", {
    x: 0.95,
    y: 2.25,
    w: 11.4,
    h: 1.25,
    fontFace: FONT_BODY,
    fontSize: 13.5,
    color: DARK,
    isTextBox: true,
    margin: 0,
    valign: "top",
    lineSpacingMultiple: 1.25,
  });

  s.addShape("roundRect", { x: 0.6, y: 3.9, w: 12.13, h: 2.1, rectRadius: 0.08, fill: { color: CARD_TINT2 }, line: { color: GOLD, width: 1 } });
  s.addText("Backup Plan", { x: 0.95, y: 4.15, w: 11.4, h: 0.4, fontFace: FONT, fontSize: 17, bold: true, color: PRIMARY, isTextBox: true, margin: 0 });
  s.addText("Even if NDA and TES don't come through by 2030, the CDS / TGC window (2032–2038) is wide and forgiving — a total of well over 15 realistic attempts across all four entries over his whole eligible life if he stays consistent.", {
    x: 0.95,
    y: 4.6,
    w: 11.4,
    h: 1.25,
    fontFace: FONT_BODY,
    fontSize: 13.5,
    color: DARK,
    isTextBox: true,
    margin: 0,
    valign: "top",
    lineSpacingMultiple: 1.25,
  });

  footer(s, "Strategy — Phase 3 & Backup");
}

// ============================================================
// SLIDE 15 — The Complete Journey (flow diagram)
// ============================================================
{
  const s = pres.addSlide();
  header(s, "Roadmap", "The Complete Journey");

  const boxFill = WHITE;
  const boxLine = LINE;
  const textColor = DARK;

  function box(x, y, w, h, title, sub, fill, tColor) {
    s.addShape("roundRect", { x, y, w, h, rectRadius: 0.07, fill: { color: fill || boxFill }, line: { color: boxLine, width: 1 } });
    s.addText(title, {
      x: x + 0.1,
      y: y + 0.07,
      w: w - 0.2,
      h: h * 0.55,
      fontFace: FONT,
      fontSize: 11,
      bold: true,
      color: tColor || textColor,
      isTextBox: true,
      margin: 0,
      align: "center",
      valign: "middle",
    });
    if (sub) {
      s.addText(sub, {
        x: x + 0.1,
        y: y + h * 0.55,
        w: w - 0.2,
        h: h * 0.42,
        fontFace: FONT_BODY,
        fontSize: 8.5,
        color: tColor || MUTED,
        isTextBox: true,
        margin: 0,
        align: "center",
        valign: "top",
      });
    }
  }

  function arrowRight(x, y, len) {
    s.addShape("line", { x, y, w: len, h: 0, line: { color: MUTED, width: 1.5, endArrowType: "triangle" } });
  }

  function arrowDown(x, y, len) {
    s.addShape("line", { x, y, w: 0, h: len, line: { color: MUTED, width: 1.5, endArrowType: "triangle" } });
  }

  const rowY = 1.55;
  const bh = 0.85;
  const bw = 1.9;
  const gap = 0.35;
  let bx = 0.6;

  box(bx, rowY, bw, bh, "Class 10", "Now, 2026–27");
  bx += bw;
  arrowRight(bx, rowY + bh / 2, gap);
  bx += gap;
  box(bx, rowY, bw, bh, "Class 11", "2027–28");
  bx += bw;
  arrowRight(bx, rowY + bh / 2, gap);
  bx += gap;
  box(bx, rowY, bw + 0.2, bh, "Class 12", "2028–29");
  bx += bw + 0.2;
  arrowRight(bx, rowY + bh / 2, gap);
  bx += gap;
  box(bx, rowY, bw + 0.5, bh, "NDA II'28, NDA I&II'29", "UPSC or PCM% + JEE", CARD_TINT);
  bx += bw + 0.5;
  arrowRight(bx, rowY + bh / 2, gap);
  bx += gap;
  box(bx, rowY, bw - 0.15, bh, "SSB + Medical", "Interview & board", CARD_TINT2);

  const branchX = bx + (bw - 0.15) / 2;
  arrowDown(branchX, rowY + bh, 0.55);

  const rowY2 = rowY + bh + 0.55 + 0.2;
  box(0.6, rowY2, 3.2, bh, "SELECTED", "3–4 yrs training @ NDA/CME + IMA", PRIMARY, WHITE);
  s.addText("Commissioned as Lt (PC) — Arm of choice", { x: 0.6, y: rowY2 + bh + 0.05, w: 3.2, h: 0.35, fontFace: FONT_BODY, fontSize: 10, italic: true, color: PRIMARY, align: "center", isTextBox: true, margin: 0 });

  s.addText("or, if not selected →", { x: 4.0, y: rowY2 + bh / 2 - 0.15, w: 1.5, h: 0.3, fontFace: FONT_BODY, fontSize: 10, italic: true, color: MUTED, isTextBox: true, margin: 0 });
  box(5.6, rowY2, 3.1, bh, "Engineering Degree", "2029/30 — 2033, keeps CDS + TGC open", CARD_TINT);

  arrowRight(8.7, rowY2 + bh / 2, 0.45);
  box(9.15, rowY2, 3.58, bh, "CDS (IMA) & TGC", "Final year onward, 2032–2038", CARD_TINT2);

  const rowY3 = rowY2 + bh + 0.55;
  const branchX2 = 9.15 + 3.58 / 2;
  arrowDown(branchX2, rowY2 + bh, 0.5);
  box(7.6, rowY3, 5.13, 0.7, "SSB Interview + Medical Board", "", CARD_TINT2);
  arrowDown(branchX2, rowY3 + 0.7, 0.35);
  box(7.6, rowY3 + 1.05, 5.13, 0.75, "Commissioned as Lt — Permanent Commission", "Career officer", PRIMARY, WHITE);

  footer(s, "The Complete Journey");
}

// ============================================================
// SLIDE 16 — Final Summary
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: PRIMARY };
  s.addText("FINAL TAKEAWAY", {
    x: 0.8,
    y: 0.55,
    w: 6,
    h: 0.35,
    fontFace: FONT_BODY,
    fontSize: 12.5,
    bold: true,
    color: GOLD,
    charSpacing: 2,
    isTextBox: true,
    margin: 0,
  });
  s.addText("The Best Possible Path for Maximum PC Chances", {
    x: 0.8,
    y: 0.9,
    w: 11.7,
    h: 0.75,
    fontFace: FONT,
    fontSize: 26,
    bold: true,
    color: WHITE,
    isTextBox: true,
    margin: 0,
  });

  const stats = [
    ["4", "Independent PC entries — NDA, TES, CDS, TGC"],
    ["15+", "Realistic attempts spread across roughly a decade"],
    ["2028–2038", "The full eligibility window, start to finish"],
  ];

  const sw = 3.85;
  const sy = 1.95;
  const sh = 1.5;
  const sgap = 0.3;

  stats.forEach((st, i) => {
    const sx = 0.8 + i * (sw + sgap);
    s.addShape("roundRect", { x: sx, y: sy, w: sw, h: sh, rectRadius: 0.08, fill: { color: "2C4224" }, line: { color: GOLD, width: 0.75 } });
    s.addText(st[0], { x: sx, y: sy + 0.15, w: sw, h: 0.7, fontFace: FONT, fontSize: 32, bold: true, color: GOLD, align: "center", isTextBox: true, margin: 0 });
    s.addText(st[1], { x: sx + 0.25, y: sy + 0.85, w: sw - 0.5, h: 0.6, fontFace: FONT_BODY, fontSize: 11.5, color: "D9E2D1", align: "center", isTextBox: true, margin: 0 });
  });

  const lines = [
    "Give NDA everything from September 2028 (his first attempt) through 2030 — it's the widest door and the longest, most complete training.",
    "Run TES alongside it for free, since PCM / JEE prep overlaps with NDA prep.",
    "Choose an engineering degree as his fallback path — the only choice that keeps both CDS and TGC open later.",
    "From 2032 onward, attack CDS and TGC together every cycle until 2035 (CDS closes at 24) and 2038 (TGC closes at 27).",
  ];

  let ly = 3.85;
  lines.forEach((l) => {
    s.addShape("ellipse", { x: 0.85, y: ly + 0.09, w: 0.1, h: 0.1, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(l, { x: 1.15, y: ly - 0.08, w: 11.3, h: 0.55, fontFace: FONT_BODY, fontSize: 13.5, color: WHITE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    ly += 0.62;
  });

  s.addText("Reminder: NDA's age mechanics above are confirmed against the standard UPSC pattern; CDS/TGC calendar edges can shift ±6 months by notification. Always confirm exact dates against the official UPSC / Indian Army notification each year.", {
    x: 0.8,
    y: 6.65,
    w: 11.7,
    h: 0.6,
    fontFace: FONT_BODY,
    fontSize: 10,
    italic: true,
    color: "9FB093",
    isTextBox: true,
    margin: 0,
  });
}

function buildPresentation() {
  const outputDir = path.resolve(__dirname, "output");
  fs.mkdirSync(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, "indian_army_officer_roadmap.pptx");

  return pres.writeFile({ fileName: outputPath });
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  buildPresentation()
    .then(() => {
      console.log("PowerPoint generated at:", path.resolve(__dirname, "output", "indian_army_officer_roadmap.pptx"));
    })
    .catch((err) => {
      console.error("Failed to generate PowerPoint:", err);
      process.exit(1);
    });
}

export { buildPresentation };
