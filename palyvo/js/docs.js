// Формування Word: реєстр «Додаток 6» (A4 альбомно) і акт списання (типова форма № З-2, A4 книжково).
// Перенесено з export_docx.py / export_act.py. D — простір імен бібліотеки docx (window.docx або require("docx")).
import * as C from "./core.js?v=20261031b";
import { litresWords } from "./words.js?v=20261031b";

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
// Колонки «Видано» і «Витрачено» будуються з джерел пального підрозділу (усі або лише з рухом — налаштування),
// «благо» — за налаштуванням. Для стандартних джерел вигляд той самий, що й раніше (19 граф).
export function buildRegister(D, s, ym, rows, sources = rows.sources || C.unitSources(s)) {
  const lay = C.registerLayout(rows, sources, s);
  const legacy = !!sources.legacy;
  const nI = lay.iss.length, nC = lay.cons.length + (lay.blago ? 1 : 0);
  // ширини: фіксовані колонки + колонки джерел (у стандартному режимі — як у затвердженому зразку)
  const base = { no: 480, unit: 560, name: 1080, norm: 640, km: 760, sl: 760, sv: legacy ? 620 : 1500, src: legacy ? 640 : 600,
    tot: 760, blago: 560, el: 760, ev: legacy ? 1900 : 1700 };
  const W = [base.no, base.unit, base.name, base.norm, base.km, base.sl, base.sv, ...Array(nI).fill(base.src), base.tot,
    ...Array(lay.cons.length).fill(base.src), ...(lay.blago ? [base.blago] : []), base.tot, base.el, base.ev];
  const usable = cm(29.7 - 3.0);
  const k = usable / W.reduce((a, b) => a + b, 0);
  const widths = W.map((w) => Math.floor(w * k));
  const fs = lay.total > 24 ? 7 : lay.total > 19 ? 8 : 9;      // багато колонок — дрібніший шрифт
  const children = [];
  children.push(para(D, [runTab(D, s.org || "", s.appendix || "", { size: 12 })], { align: "left",
    tabs: [{ type: D.TabStopType.RIGHT, position: usable }] }));
  children.push(para(D, [run(D, s.dept || "", { size: 12 })], { align: "left" }));
  children.push(para(D, [run(D, `${s.title || ""} за ${C.monthTitle(ym)}`, { size: 13, bold: true, italics: true })],
    { before: 4, after: 6 }));

  const c = (t, o = {}) => cell(D, t, { size: fs, ...o });
  const hdr = (cells, height, extra = {}) => new D.TableRow({ children: cells, tableHeader: true, cantSplit: true,
    ...(height ? { height: { value: cm(height), rule: D.HeightRule.ATLEAST } } : {}), ...extra });
  const r0 = [
    c("№\nп/п", { rowSpan: 2 }), c("Підрозділ", { rowSpan: 2, vertical: true }),
    c("Марка, днз", { rowSpan: 2 }), c("Місячна норма пробігу (км)", { rowSpan: 2, vertical: true }),
    c("Пробіг за звітний місяць (км)", { rowSpan: 2, vertical: true }),
    c("Залишок палива у баку на початок місяця, (л)", { colSpan: 2 }),
    c("Видано палива (л)", { colSpan: nI + 1 }),
    c("Витрачено палива згідно пробігу, (л)", { colSpan: nC + 1 }),
    c("Залишок палива у баку на кінець місяця, (л)", { colSpan: 2 }),
  ];
  const sub = [["л"], [legacy ? "вид" : "вид пального на залишку", !legacy]];
  lay.iss.forEach((x) => sub.push([C.sourceTitle(x), 1]));
  sub.push(["Всього", 1, 1]);
  lay.cons.forEach((x) => sub.push([C.sourceTitle(x), 1]));
  if (lay.blago) sub.push(["благо"]);
  sub.push(["Всього", 1, 1], ["л"], [legacy ? "Вид" : "Вид пального на залишку"]);
  const r1 = sub.map(([t, v, b]) => c(t, { vertical: !!v, bold: !!b }));
  const r2 = Array.from({ length: lay.total }, (_, i) => c(String(i + 1)));
  const tableRows = [hdr(r0, 1.3), hdr(r1, legacy ? 2.1 : 3.2), hdr(r2)];
  const boldIdx = new Set([lay.issTotalNo - 1, lay.consTotalNo - 1]);

  rows.forEach((r, i) => {
    const v = r.vehicle;
    let startVid, endVid;
    if (legacy) {
      startVid = C.legacyStartVid(r);
      endVid = (v.fuel || "") + (v.supplier ? ` (${v.supplier})` : "");
      if (v.mvo_text) endVid += "\nМВО\n" + splitNames(v.mvo_text).join("\n");
    } else {
      startVid = C.balanceText(r, "start", s).split(" + ").join(" +\n");
      endVid = C.balanceText(r, "end", s).split(" + ").join(" +\n");
    }
    const vals = [String(i + 1), s.unit || "", v.name, String(v.month_norm_km || ""), String(r.km), C.fmtC(r.startC), startVid,
      ...lay.iss.map((x) => C.fmtC(r.issued[x.id] || 0, true)), C.fmtC(r.issuedC),
      ...lay.cons.map((x) => C.fmtC(r.consBy[x.id] || 0, true)), ...(lay.blago ? [""] : []), C.fmtC(r.consC),
      C.fmtC(r.endC), endVid];
    tableRows.push(new D.TableRow({ cantSplit: true, children: vals.map((t, ci) => c(t, { bold: boldIdx.has(ci) })) }));
  });

  const t = C.totals(rows);
  const tv = [C.fmtC(t.startC), "", ...lay.iss.map((x) => C.fmtC(t.issued[x.id] || 0, true)), C.fmtC(t.issuedC),
    ...lay.cons.map((x) => C.fmtC(t.consBy[x.id] || 0, true)), ...(lay.blago ? [""] : []), C.fmtC(t.consC), C.fmtC(t.endC), ""];
  const totCells = [c("Всього", { colSpan: 5, bold: true }), ...tv.map((x) => c(x, { bold: true }))];
  tableRows.push(new D.TableRow({ cantSplit: true, children: totCells }));

  children.push(new D.Table({ rows: tableRows, columnWidths: widths, width: { size: usable, type: D.WidthType.DXA },
    layout: D.TableLayoutType.FIXED, alignment: D.AlignmentType.CENTER,
    margins: { top: 30, bottom: 30, left: 40, right: 40 } }));

  for (const txt of [`дані графи № ${lay.issTotalNo} повинні відповідати даним граф № 11, 12 звіту про одержане і видане пальне.`,
    `дані графи № ${lay.consTotalNo} повинні відповідати кількості пального в акті списання.`]) {
    children.push(para(D, [runTab(D, "*", txt, { size: 9 })], { align: "left", before: 4, indentLeft: cm(0.5), hanging: cm(0.5),
      keepNext: true }));
  }
  children.push(para(D, [run(D, "Підпис особи, яка склала реєстр:", { size: 12, bold: true })],
    { align: "left", before: 14, after: 10, keepNext: true }));
  for (const nm of [s.compiler, s.compiler2, s.compiler3].map((x) => String(x || "").trim()).filter((x, i) => i === 0 || x))
    children.push(para(D, [run(D, `______________________________  ${nm}`, { size: 12 })], { align: "left", indentLeft: cm(12.5), before: 8, keepNext: true }));

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

/** act — один елемент C.actList(rows, s): {group, mvo, title, lines}. */
export function buildAct(D, s, ym, act, now = new Date()) {
  const [y, m] = C.parseYm(ym);
  const yearNow = now.getFullYear();
  const monthTxt = C.MONTHS_UA[m - 1];
  const groups = act.lines || [];
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
    const vals = g ? [`${i + 1}.`, "", [g.fuel, g.form === "с/к" ? "с/к" : g.form, g.supplier].filter(Boolean).join(" "), "л", C.fmtC(g.qtyC), "", "",
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
    ...(act.mvos && act.mvos.length ? act.mvos : [act.mvo || s.compiler || ""]).map((nm, i) =>
      P(`${i === 0 ? BLANK_DATE + yearNow : " ".repeat(BLANK_DATE.length + 4)}      МВО:  ___________________      ${nm}`,
        { indentLeft: cm(0.6), before: i === 0 ? 30 : 10 })),
  ];
  return baseDoc(D, { page: { size: { width: cm(21.0), height: cm(29.7) },
    margin: { top: cm(1), bottom: cm(1), left: cm(2), right: cm(1.5) } } }, children, 12);
}
