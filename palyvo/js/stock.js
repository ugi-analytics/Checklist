// «Додаток № 2» — звіт про одержане і видане пальне (рух пального в МВО за накладними).
// МВО вносить накладні (№/дата або від кого, джерело, одержано, «в наливі / резерв»); видане — з реєстру,
// розписується на накладні того ж джерела: спершу робочі (від старішої), потім ті, що «в наливі / резерв».
import * as C from "./core.js?v=20261113k";

/** Налив (АЗС налив, налив ОВА) і газ — у графах «літри/(кг)»; талони і с/к — у графах «талони (с/к)». */
export const LIT_FORMS = new Set(["АЗС налив", "Налив ОВА"]);
const FORM_WORD = { "с/к": "с/к", "талони": "талони", "АЗС ГУНП талони": "талони АЗС ГУНП", "АЗС налив": "налив", "Налив ОВА": "налив ОВА" };
export const formWord = (f) => FORM_WORD[f] || f || "";
export const stockKey = (srcId, fuel) => `${srcId}|${C.normFuel(fuel)}`;
/** Вид за накладною (обов'язково при внесенні): графа 5 / 6 і графи «талони (с/к)» чи «літри». */
export const LINE_FORMS = [["талони", "талони"], ["с/к", "смарт-карта (с/к)"], ["АЗС налив", "налив"], ["АЗС ГУНП талони", "талони АЗС ГУНП"], ["Налив ОВА", "налив ОВА"]];
/** Без номера накладної (напр. «Петрунь»): у «Додатку 2» такі рядки одного джерела підсумовуються в один. */
export const noNumber = (doc) => !/\d/.test(String(doc || ""));

const num = (v) => { const x = Number(String(v ?? "").replace(",", ".").replace(/\s/g, "")); return isFinite(x) && x > 0 ? Math.round(x * 100) / 100 : 0; };
/** Рядок накладної: {id, src, fuel, form, supplier, mvo, doc, start, recv, reserve, note, nominal, carried}. */
export function normalizeLine(x = {}) {
  const o = { id: x.id || "n" + Math.random().toString(16).slice(2, 9), src: String(x.src || ""), fuel: String(x.fuel || "").trim(),
    form: String(x.form || "").trim(), supplier: String(x.supplier || "").trim(), mvo: String(x.mvo || "").trim(),
    doc: String(x.doc || "").trim().slice(0, 120), start: num(x.start), recv: num(x.recv), reserve: !!x.reserve,
    note: String(x.note || "").trim().slice(0, 60), nominal: String(x.nominal || "").trim().slice(0, 10), carried: !!x.carried };
  return o;
}
/** Реквізити джерела в рядок накладної (щоб рядок не залежав від пізніших змін у джерелах). */
export function lineFromSource(src, fuel, extra = {}) {
  return normalizeLine({ src: src.id, fuel: src.fuel || fuel || "", form: src.form, supplier: src.supplier, mvo: src.mvo, ...extra });
}
/** Дата з «… від 19.10.2026р.» для порядку списання (старіша — раніше). */
export function docDate(doc) {
  const m = String(doc || "").match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
  if (!m) return 0;
  const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  return y * 10000 + Number(m[2]) * 100 + Number(m[1]);
}
/** Видане за місяць із реєстру: ключ (джерело + пальне) → сотні літра. */
export function issuedFromRows(rows, settings = {}) {
  const out = new Map();
  for (const r of rows) for (const s of r.sources || []) {
    const c = (r.issued || {})[s.id] || 0; if (!c) continue;
    const f = C.sourceFor(s, r.vehicle, settings);
    const k = stockKey(s.id, f.fuel);
    if (!out.has(k)) out.set(k, { src: s, fuel: f.fuel, c: 0 });
    out.get(k).c += c;
  }
  return out;
}
const FUEL_ORDER = (f) => { const n = C.normFuel(f); return n.startsWith("А") ? "1" + n : n === "ДП" ? "2" : n === "ГАЗ" ? "3" : "4" + n; };
export const fuelName = (f) => (C.normFuel(f).startsWith("А") ? `Бензин ${f}` : f || "");

/**
 * Розрахунок звіту: stock = {lines: [...]}, rows — рядки реєстру місяця.
 * Повертає {lines, groups:[{fuel, title, lines, tot}], tot, errors, warnings, mvoCol}.
 */
export function calcStock(stock, rows, sources, settings = {}, ym = "") {
  const lines = (stock && stock.lines || []).map(normalizeLine).map((l, i) => ({ ...l, i, startC: Math.round(l.start * 100),
    recvC: Math.round(l.recv * 100), issC: 0, lit: LIT_FORMS.has(l.form) || C.normFuel(l.fuel) === "ГАЗ" }));
  const errors = [], warnings = [];
  const iss = issuedFromRows(rows, settings);
  for (const [k, it] of iss) {
    const pool = lines.filter((l) => stockKey(l.src, l.fuel) === k)
      // спершу робочі, потім «в наливі / резерв»; у межах — від старішої дати (без номера/дати — першими),
      // тож залишок лишається на накладних, одержаних пізніше
      .sort((a, b) => (a.reserve - b.reserve) || (docDate(a.doc) - docDate(b.doc)) || (a.i - b.i));
    let left = it.c;
    for (const l of pool) { const take = Math.min(left, l.startC + l.recvC); l.issC += take; left -= take; if (!left) break; }
    if (left > 0) {
      const f = C.sourceFor(it.src, { fuel: it.fuel }, settings);
      const l = { ...lineFromSource(it.src, it.fuel, { fuel: f.fuel, doc: "накладну не внесено" }), i: lines.length, startC: 0, recvC: 0,
        issC: left, lit: LIT_FORMS.has(it.src.form) || C.normFuel(f.fuel) === "ГАЗ", missing: true };
      lines.push(l);
      errors.push(`${fuelName(f.fuel)} ${formWord(f.form)}${f.supplier ? " " + f.supplier : ""}${f.mvo ? " (МВО " + f.mvo + ")" : ""}: видано за реєстром ${C.fmtC(left)} л більше, ніж є за накладними — внесіть накладну або залишок.`);
    }
  }
  const [yy, mm] = ym ? C.parseYm(ym) : [0, 0];
  for (const l of lines) {
    l.endC = l.startC + l.recvC - l.issC;
    if (!l.missing && !l.form) errors.push(`Накладна «${l.doc || "без номера"}»: не вказано вид — талони, смарт-карта (с/к) чи налив.`);
    if (yy && l.recvC && docDate(l.doc) > yy * 10000 + mm * 100 + 31) warnings.push(`Накладна «${l.doc}»: дата пізніша за звітний місяць.`);
    if (yy && l.recvC && docDate(l.doc) && docDate(l.doc) < (yy - 1) * 10000 + mm * 100) warnings.push(`Накладна «${l.doc}»: одержано цього місяця, а дата — понад рік тому. Перевірте рік.`);
  }
  const mc = String(settings.stock_mvo_col ?? "auto");
  const mvoCol = mc === "1" ? true : mc === "0" ? false : C.vehicleMvos(sources.filter((s) => s.active !== false)).length > 1;
  const sum = (ls) => {
    const t = { sT: 0, sL: 0, rT: 0, rL: 0, iT: 0, iL: 0, eT: 0, eL: 0 };
    for (const l of ls) for (const [p, x] of l.merged ? [["T", l.T], ["L", l.L]] : [[l.lit ? "L" : "T", l]]) {
      if (!x) continue; t["s" + p] += x.startC; t["r" + p] += x.recvC; t["i" + p] += x.issC; t["e" + p] += x.endC;
    }
    return t;
  };
  // для документа: накладні без номера одного джерела (і виду, і позначки «резерв») — одним рядком
  const docLines = [], mm2 = new Map();
  for (const l of lines) {
    if (l.missing || !noNumber(l.doc)) { docLines.push(l); continue; }
    // усі надходження без номера (з усіх джерел, усіх форм) — одним рядком у межах виду пального;
    // талони/с/к і налив у цьому рядку — кожне у своїх графах
    const k = [C.normFuel(l.fuel), l.reserve ? 1 : 0].join("|");
    if (!mm2.has(k)) {
      const m = { ...l, id: "m:" + k, docs: [], mvos: [], sups: [], forms: [], merged: true, T: null, L: null, startC: 0, recvC: 0, issC: 0, endC: 0 };
      mm2.set(k, m); docLines.push(m);
    }
    const m = mm2.get(k), p = l.lit ? "L" : "T";
    if (!m[p]) m[p] = { startC: 0, recvC: 0, issC: 0, endC: 0 };
    for (const f of ["startC", "recvC", "issC", "endC"]) { m[p][f] += l[f]; m[f] += l[f]; }
    for (const [arr, x] of [[m.docs, l.doc], [m.mvos, l.mvo], [m.sups, l.supplier], [m.forms, l.form]]) if (x && !arr.includes(x)) arr.push(x);
    if (!m.note && l.note) m.note = l.note;
  }
  for (const m of mm2.values()) {
    m.doc = m.docs.join(", ").slice(0, 120) || "без накладної";
    m.mvo = m.mvos.join(", "); m.supplier = m.sups.length === 1 ? m.sups[0] : ""; m.form = m.forms[0] || "";
    m.lit = !m.T; // лише налив — графи «літри»
  }
  const gm = new Map();
  for (const l of docLines) { const k = FUEL_ORDER(l.fuel); if (!gm.has(k)) gm.set(k, { fuel: l.fuel, lines: [] }); gm.get(k).lines.push(l); }
  const groups = [...gm.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([, g]) => ({ ...g, title: fuelName(g.fuel), tot: sum(g.lines) }));
  // у групі: спершу за джерелом (вид, форма, постачальник, МВО), далі — за датою накладної
  const srcOrd = new Map(sources.map((s, i) => [s.id, i]));
  for (const g of groups) g.lines.sort((a, b) => ((srcOrd.get(a.src) ?? 99) - (srcOrd.get(b.src) ?? 99)) || (docDate(a.doc) - docDate(b.doc)) || (a.i - b.i));
  return { lines, groups, tot: sum(lines), errors, warnings, mvoCol };
}
/** Перенесення на наступний місяць: накладні з залишком на кінець > 0 стають залишком на початок. */
export function carryStock(res) {
  return { lines: res.lines.filter((l) => l.endC > 0 && !l.missing).map((l) => normalizeLine({ ...l, start: l.endC / 100, recv: 0, carried: true })) };
}
/** Значення 14 граф для рядка накладної і рядка «відомість». */
export function rowCells(l, settings, mvoCol, unitShort) {
  if (l.merged) {
    const fw = l.forms.map(formWord).join(", ");
    const name = [fuelName(l.fuel), fw, l.supplier].filter(Boolean).join(" ");
    const z = (x, f) => (x ? C.fmtC(x[f]) : "");
    const endT = (x) => (x ? [C.fmtC(x.endC), l.reserve ? l.note : ""].filter(Boolean).join(" ") : "");
    const c5 = mvoCol ? l.mvos.map((m) => `МВО-${m}`).join(", ") : fw, c6 = mvoCol ? fw : "";
    return { inv: [unitShort, name, l.doc, c5, c6, z(l.T, "startC"), z(l.L, "startC"), z(l.T, "recvC"), z(l.L, "recvC"), "", "", endT(l.T), endT(l.L)],
      ved: [unitShort, name, "відомість №", "", mvoCol ? fw : "", "", "", "", "", z(l.T, "issC"), z(l.L, "issC"), "", ""] };
  }
  const name = [fuelName(l.fuel), formWord(l.form), l.supplier].filter(Boolean).join(" ");
  const fw = formWord(l.form);
  const c5 = mvoCol ? (l.mvo ? `МВО-${l.mvo}` : "") : fw;
  const c6 = mvoCol ? fw : l.nominal;
  const z = (c) => C.fmtC(c);
  const T = !l.lit;
  // примітка («резерв», «в наливі») — лише в графі виду накладної: талони/с/к → 13, налив → 14
  const endTxt = [z(l.endC), l.reserve ? l.note : ""].filter(Boolean).join(" ");
  const end = [T ? endTxt : "", T ? "" : endTxt];
  const inv = [unitShort, name, l.doc, c5, c6, T ? z(l.startC) : "", T ? "" : z(l.startC), T ? z(l.recvC) : "", T ? "" : z(l.recvC), "", "", ...end];
  const ved = [unitShort, name, "відомість №", "", mvoCol ? fw : "", "", "", "", "", T ? z(l.issC) : "", T ? "" : z(l.issC), "", ""];
  return { inv, ved };
}
