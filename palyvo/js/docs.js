// Формування Word: реєстр «Додаток 6» (A4 альбомно) і акт списання (типова форма № З-2, A4 книжково).
// Перенесено з export_docx.py / export_act.py. D — простір імен бібліотеки docx (window.docx або require("docx")).
import * as C from "./core.js?v=20261026a";
import { litresWords } from "./words.js?v=20261026a";

const FONT = "Times New Roman";
const cm = (x) => Math.round(x * 567);      // см → twips
const pt = (x) => Math.round(x * 20);       // пт → twips

function run(D, text, { size = 9, bold = false, italics = false } = {}) {
  return new D.TextRun({ text, size: size * 2, bold, italics, font: FONT });
}
function runTab(D, a, b, opts = {}) {
  const { size = 12, bold = false } = opts;
  return new D.TextRun({ children: [a, new D.Tab(), b], size: size * 2, bold, font: FONT });
}
function para(D, children, { align = "center", before = 0, after = 0, indentLeft, firstLine, hanging, tabs, keepNext } = {}) {
  const o = { children, alignment: { center: D.AlignmentType.CENTER, left: D.AlignmentType.LEFT, right: D.AlignmentType.RIGHT,
    justify: D.AlignmentType.JUSTIFIED }[align], spacing: { before: pt(before), after: pt(after), line: 240 } };
  if (indentLeft !== undefined || firstLine !== undefined || hanging !== undefined) {
    o.indent = {}; if (indentLeft !== undefined) o.indent.left = indentLeft;
    if (firstLine !== undefined) o.indent.firstLine = firstLine; if (hanging !== undefined) o.indent.hanging = hanging;
  }
  if (tabs) o.tabStops = tabs;
  if (keepNext) o.keepNext = true;
  return new D.Paragraph(o);
}
function cell(D, text, { size = 9, bold = false, vertical = false, colSpan, rowSpan, align = "center", width } = {}) {
  const lines = String(text ?? "").split("\n");
  const o = { children: lines.map((l) => para(D, [run(D, l, { size, bold })], { align })),
    verticalAlign: D.VerticalAlign.CENTER };
  if (vertical) o.textDirection = D.TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT;
  if (colSpan) o.columnSpan = colSpan;
  if (rowSpan) o.rowSpan = rowSpan;
  if (width) o.width = { size: width, type: D.WidthType.DXA };
  return new D.TableCell(o);
}
function splitNames(text) {
  const parts = String(text || "").trim().match(/[^\s,;]+(?:\s+[А-ЯЇІЄҐA-Z]\.\s*[А-ЯЇІЄҐA-Z]\.)?/g);
  return parts && parts.length ? parts.map((s) => s.trim()) : [text];
}
function baseDoc(D, sectionProps, children, defaultSize) {
  return new D.Document({
    styles: { default: { document: { run: { font: FONT, size: defaultSize * 2 } } } },
    sections: [{ properties: sectionProps, children }],
  });
}

// ================================================================= реєстр «Додаток 6»
const W = [480, 560, 1080, 640, 760, 760, 620, 640, 560, 760, 700, 760, 640, 560, 760, 620, 760, 760, 1900];

export function buildRegister(D, s, ym, rows) {
  const usable = cm(29.7 - 3.0);
  const k = usable / W.reduce((a, b) => a + b, 0);
  const widths = W.map((w) => Math.floor(w * k));
  const children = [];
  children.push(para(D, [runTab(D, s.org || "", s.appendix || "", { size: 12 })], { align: "left",
    tabs: [{ type: D.TabStopType.RIGHT, position: usable }] }));
  children.push(para(D, [run(D, s.dept || "", { size: 12 })], { align: "left" }));
  children.push(para(D, [run(D, `${s.title || ""} за ${C.monthTitle(ym)}`, { size: 13, bold: true, italics: true })],
    { before: 4, after: 6 }));

  const hdr = (cells, height, extra = {}) => new D.TableRow({ children: cells, tableHeader: true, cantSplit: true,
    ...(height ? { height: { value: cm(height), rule: D.HeightRule.ATLEAST } } : {}), ...extra });
  const r0 = [
    cell(D, "№\nп/п", { rowSpan: 2 }), cell(D, "Підрозділ", { rowSpan: 2, vertical: true }),
    cell(D, "Марка, днз", { rowSpan: 2 }), cell(D, "Місячна норма пробігу (км)", { rowSpan: 2, vertical: true }),
    cell(D, "Пробіг за звітний місяць (км)", { rowSpan: 2, vertical: true }),
    cell(D, "Залишок палива у баку на початок місяця, (л)", { colSpan: 2 }),
    cell(D, "Видано палива (л)", { colSpan: 5 }),
    cell(D, "Витрачено палива згідно пробігу, (л)", { colSpan: 5 }),
    cell(D, "Залишок палива у баку на кінець місяця, (л)", { colSpan: 2 }),
  ];
  const sub = [["л"], ["вид"], ["АЗС ГУНП талони", 1], ["Талони", 1], ["см/к"], ["АЗС налив"], ["Всього", 1, 1],
    ["АЗС ГУНП талони", 1], ["Талони", 1], ["см/к"], ["благо"], ["Всього", 1, 1], ["л"], ["Вид"]];
  const r1 = sub.map(([t, v, b]) => cell(D, t, { vertical: !!v, bold: !!b }));
  const r2 = Array.from({ length: 19 }, (_, i) => cell(D, String(i + 1)));
  const tableRows = [hdr(r0, 1.3), hdr(r1, 2.1), hdr(r2)];

  rows.forEach((r, i) => {
    const v = r.vehicle;
    const kind = v.fuel || "";
    let vid19 = kind + (v.supplier ? ` (${v.supplier})` : "");
    if (v.mvo_text) vid19 += "\nМВО\n" + splitNames(v.mvo_text).join("\n");
    const vals = [String(i + 1), s.unit || "", v.name, String(v.month_norm_km || ""), String(r.km),
      C.fmtC(r.startC), kind];
    for (const [key] of C.ISSUE_SOURCES) vals.push(C.fmtC(r.issued[key], true));
    vals.push(C.fmtC(r.issuedC));
    for (const [key] of C.CONSUME_SOURCES) vals.push(C.fmtC(r.consBy[key], true));
    vals.push("", C.fmtC(r.consC), C.fmtC(r.endC), vid19);
    tableRows.push(new D.TableRow({ cantSplit: true, children: vals.map((t, ci) => cell(D, t, { bold: ci === 11 || ci === 16 })) }));
  });

  const t = C.totals(rows);
  const tv = { 5: C.fmtC(t.startC), 11: C.fmtC(t.issuedC), 16: C.fmtC(t.consC), 17: C.fmtC(t.endC) };
  C.ISSUE_SOURCES.forEach(([key], j) => { tv[7 + j] = C.fmtC(t.issued[key], true); });
  C.CONSUME_SOURCES.forEach(([key], j) => { tv[12 + j] = C.fmtC(t.consBy[key], true); });
  const totCells = [cell(D, "Всього", { colSpan: 5, bold: true })];
  for (let ci = 5; ci < 19; ci++) totCells.push(cell(D, tv[ci] || "", { bold: true }));
  tableRows.push(new D.TableRow({ cantSplit: true, children: totCells }));

  children.push(new D.Table({ rows: tableRows, columnWidths: widths, width: { size: usable, type: D.WidthType.DXA },
    layout: D.TableLayoutType.FIXED, alignment: D.AlignmentType.CENTER,
    margins: { top: 30, bottom: 30, left: 40, right: 40 } }));

  for (const txt of ["дані графи № 12 повинні відповідати даним граф № 11, 12 звіту про одержане і видане пальне.",
    "дані графи № 17 повинні відповідати кількості пального в акті списання."]) {
    children.push(para(D, [runTab(D, "*", txt, { size: 9 })], { align: "left", before: 4, indentLeft: cm(0.5), hanging: cm(0.5),
      keepNext: true }));
  }
  children.push(para(D, [run(D, "Підпис особи, яка склала реєстр:", { size: 12, bold: true })],
    { align: "left", before: 14, after: 10, keepNext: true }));
  children.push(para(D, [run(D, `______________________________  ${s.compiler || ""}`, { size: 12 })],
    { align: "left", indentLeft: cm(12.5) }));

  return baseDoc(D, { page: { size: { orientation: D.PageOrientation.LANDSCAPE, width: cm(21.0), height: cm(29.7) },
    margin: { top: cm(1), bottom: cm(1), left: cm(1.5), right: cm(1.5) } } }, children, 11);
}

// ================================================================= акт списання (З-2)
const BLANK_DATE = "___.___.";
function joinMembers(members = []) {
  const full = members.map((m) => String(m.full || "").trim()).filter(Boolean);
  if (full.length <= 1) return full.join("");
  return full.slice(0, -1).join(", ") + " та " + full[full.length - 1];
}

export function buildAct(D, s, ym, rows, group = null, now = new Date()) {
  const [y, m] = C.parseYm(ym);
  const yearNow = now.getFullYear();
  const monthTxt = C.MONTHS_UA[m - 1];
  const groups = C.actGroups(rows, group);
  const total = groups.reduce((a, g) => a + g.qtyC, 0);
  const usable = cm(21.0 - 3.5);
  const lw = Math.floor(usable * 0.44), rw = usable - lw;
  const P = (text, o = {}) => para(D, text ? [run(D, text, { size: o.size || 12, bold: o.bold })] : [], { align: "left", ...o });

  // шапка: дві колонки без рамок
  const accW = Math.floor((rw - 200) / 3);
  const acc = new D.Table({ columnWidths: [accW, accW, accW], width: { size: accW * 3, type: D.WidthType.DXA },
    layout: D.TableLayoutType.FIXED, margins: { top: 10, bottom: 10, left: 30, right: 30 },
    rows: [new D.TableRow({ children: ["Номенклатурний номер", "Основний рахунок", "Кореспондентський рахунок"].map((t) => cell(D, t, { size: 7 })) }),
      new D.TableRow({ children: [0, 1, 2].map(() => cell(D, "", { size: 9 })) })] });
  const left = new D.TableCell({ width: { size: lw, type: D.WidthType.DXA }, children: [
    P(s.org || "", { size: 11, before: 6 }), P(`Ідентифікаційний код за ЄДРПОУ ${s.edrpou || ""}`, { size: 11, before: 14 })] });
  const right = new D.TableCell({ width: { size: rw, type: D.WidthType.DXA }, children: [
    P("Типова форма № З-2", { size: 8 }),
    P("Затверджено Наказ Державного казначейства України 18.12.2000 № 130", { size: 8 }),
    para(D, [run(D, "Код за ДКУД  ____________________", { size: 11 })], { align: "left", before: 2 }),
    P("Затверджую:", { before: 8 }), P(s.act_approver_title || ""),
    para(D, [runTab(D, `${BLANK_DATE} ${yearNow}`, s.act_approver || "", { size: 12 })], { align: "left", before: 14,
      tabs: [{ type: D.TabStopType.RIGHT, position: rw - 200 }] }),
    acc, P("", { size: 4 })] });
  const none = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const hdr = new D.Table({ columnWidths: [lw, rw], width: { size: usable, type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED,
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    rows: [new D.TableRow({ children: [left, right] })] });

  const intro = `${BLANK_DATE}${yearNow} комісія призначена ${String(s.act_order || "").trim()} у складі голови комісії: ` +
    `${String(s.act_chair_full || "").trim()} та членів комісії: ${joinMembers(s.act_members)}, здійснили перевірку паливно – ` +
    `мастильних матеріалів, що використані на службові транспортні засоби, які закріплені за ${s.act_owner || ""} ` +
    `за ${monthTxt} ${y} року та встановили, що описані нижче цінності підлягають списанню та вилученню з обліку:`;

  const heads = ["№ з/п", "Номенклатурний номер", "Найменування та опис (марка, гатунок і т.д.)", "Одиниця виміру",
    "Кількість", "Ціна за одиницю", "Сума", "Вид. накладна ГУНП №, дата", "Дата заміни"];
  let w = [560, 1250, 2150, 950, 1100, 950, 900, 1300, 1550];
  const kk = usable / w.reduce((a, b) => a + b, 0); w = w.map((x) => Math.floor(x * kk));
  const period = `Списано згідно шляхових листів за ${monthTxt} ${y}`;
  const tr = [new D.TableRow({ children: heads.map((h) => cell(D, h, { size: 10 })) }),
    new D.TableRow({ children: heads.map((_, i) => cell(D, `${i + 1}.`, { size: 10 })) })];
  const lines = groups.length ? groups : [null];
  lines.forEach((g, i) => {
    const vals = g ? [`${i + 1}.`, "", [g.fuel, g.supplier].filter(Boolean).join(" "), "л", C.fmtC(g.qtyC), "", "",
      g.mvo ? `МВО ${g.mvo}` : "", period] : ["1.", "", "", "л", "0", "", "", "", period];
    tr.push(new D.TableRow({ cantSplit: true, children: vals.map((v) => cell(D, v, { size: 11 })) }));
  });
  tr.push(new D.TableRow({ children: [cell(D, "Разом :", { size: 11, colSpan: 4, align: "left" }),
    cell(D, C.fmtC(total), { size: 11, bold: true }), ...[0, 1, 2, 3].map(() => cell(D, "", { size: 11 }))] }));
  const table = new D.Table({ rows: tr, columnWidths: w, width: { size: usable, type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED,
    alignment: D.AlignmentType.CENTER, margins: { top: 30, bottom: 30, left: 50, right: 50 } });

  const sign = (name) => para(D, [runTab(D, name, "____________________", { size: 12 })], { align: "left", indentLeft: cm(0.6),
    before: 6, keepNext: true, tabs: [{ type: D.TabStopType.RIGHT, position: usable - 300 }] });
  const children = [hdr,
    para(D, [run(D, "Акт списання № ______", { size: 14, bold: true })], { before: 14, after: 8 }),
    para(D, [run(D, intro, { size: 12 })], { align: "justify", firstLine: cm(1.0), after: 6 }),
    table,
    para(D, [run(D, `Усього за цим актом списано ${litresWords(total)} на загальну суму ___________ гривень __________ копійок.`,
      { size: 12 })], { align: "justify", firstLine: cm(1.0), before: 12 }),
    P("Окремі зауваження комісії відсутні.", { indentLeft: cm(0.6), before: 10 }),
    P("Голова комісії:", { indentLeft: cm(0.6), before: 4, keepNext: true }),
    sign(s.act_chair || ""),
    P("Члени комісії:", { indentLeft: cm(0.6), before: 6, keepNext: true }),
    ...(s.act_members || []).filter((mb) => String(mb.short || "").trim()).map((mb) => sign(String(mb.short).trim())),
    P(`${BLANK_DATE}${yearNow}      МВО:  ___________________      ${s.compiler || ""}`, { indentLeft: cm(0.6), before: 30 }),
  ];
  return baseDoc(D, { page: { size: { width: cm(21.0), height: cm(29.7) },
    margin: { top: cm(1), bottom: cm(1), left: cm(2), right: cm(1.5) } } }, children, 12);
}
