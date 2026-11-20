// Режим ЦЗ: реєстр (Додаток 6) у Excel — як у реєстрі ЦЗ (служба, марка, номер, джерела за назвами колонок,
// «передано з ТЗ на ТЗ», формули «всього» і залишку, порожня «Примітка»).
import * as C from "./core.js?v=20261120a";
import { buildSheet, ST, ref, col } from "./xlsxcz.js?v=20261120a";

const LOC = ["січні", "лютому", "березні", "квітні", "травні", "червні", "липні", "серпні", "вересні", "жовтні", "листопаді", "грудні"];
const SHORT = { "талони": "т", "т/н": "т/н", "налив": "н", "налив/з": "н/з", "см/к": "см/к", "ова": "ова" };
/** Колонка реєстру ЦЗ — назва джерела («талони», «т/н», «налив/з», «Ярошенко/Кузнівенко»…). */
export const czCol = (s) => (s.label && s.label.trim()) || C.sourceTitle(s);
const short = (s) => { const k = czCol(s).toLowerCase(); return SHORT[k] || (s.form === "с/к" && s.mvo ? `см/к${s.mvo}` : czCol(s)); };
const n2 = (c) => Math.round(c) / 100;
const vin = (w) => (/^[вф]/i.test(w) ? "у " : "в ") + w;
/** Опис форми залишку: «т», «н/з», «50/н+20/т». */
export function czForm(r, which) {
  const own = String((r.entry || {})[which === "start" ? "start_text" : "end_text"] || "").trim();
  if (own) return own;
  const by = which === "start" ? r.startBy : r.endBy;
  const parts = [];
  for (const s of r.sources) {
    let c = by[s.id] || 0;
    if (which === "start" && r.xferTo === s.id) c -= r.xferC;     // передане — не частина залишку на початок
    if (c > 0) parts.push([c, short(s)]);
  }
  for (const k of r.keep || []) { const c = which === "start" ? k.c : k.left; if (c > 0 && k.label !== "передано з ТЗ на ТЗ") parts.push([c, k.label]); }
  const tot = which === "start" ? r.startC : r.endC;
  if (!tot) return "";
  if (parts.length <= 1) return parts.length ? parts[0][1] : "";
  return parts.map(([c, t]) => `${String(n2(c)).replace(".", ",")}/${t}`).join("+");
}
const plateOf = (v) => {
  if (v.plate) return String(v.plate).trim();
  const m = String(v.name || "").match(/\s([A-ZА-ЯІЇЄ]{0,2}\s?\d{3,5}(\s?[A-ZА-ЯІЇЄ]{2})?)\s*$/i);
  return m ? m[1].trim() : "";
};
const nameOf = (v) => (v.plate ? v.name : String(v.name || "").replace(/\s([A-ZА-ЯІЇЄ]{0,2}\s?\d{3,5}(\s?[A-ZА-ЯІЇЄ]{2})?)\s*$/i, "").trim());

export function buildCzRegister(s, ym, rows, sources) {
  const [y, m] = C.parseYm(ym);
  const mon = C.MONTHS_UA[m - 1].toLowerCase();
  // колонки джерел — за назвою (кілька джерел з однаковою назвою, напр. «Ярошенко/Кузнівенко», — в одній колонці)
  const used = sources.filter((x) => x.active !== false || rows.some((r) => (r.issued[x.id] || 0) || (r.consBy[x.id] || 0)));
  const cols = [];
  for (const x of used) { const k = czCol(x); let c = cols.find((y2) => y2.k === k); if (!c) cols.push(c = { k, ids: [] }); c.ids.push(x.id); }
  const nI = cols.length + 2, nC = cols.length + 1;                   // + «передано», «всього» / + «всього»
  const cI = 7, cIx = cI + cols.length, cIt = cIx + 1, cC = cIt + 1, cCt = cC + cols.length, cE = cCt + 1, last = cE + 1;
  const cells = new Map(), merges = [];
  const put = (r, c, v, st = ST.txt, f) => cells.set(`${r},${c}`, { v, s: st, f });
  const mg = (r1, c1, r2, c2) => merges.push(`${ref(c1, r1)}:${ref(c2, r2)}`);
  const box = (r1, c1, r2, c2, st) => { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!cells.has(`${r},${c}`)) put(r, c, "", st); };

  put(0, 0, ` ${s.org || "ГУНП в Полтавській області"}`, ST.title); mg(0, 0, 0, 5);
  put(1, 0, "(найменування підприємства)", ST.cap); mg(1, 0, 1, 5);
  put(2, 0, s.dept || s.unit || "", ST.title); mg(2, 0, 2, 5);
  put(3, 0, "(найменування структурного підрозділу)", ST.cap); mg(3, 0, 3, 5);
  put(4, 0, `Реєстр даних згідно дорожніх листів автотранспорту ГУНП в Полтавській області за ${mon}  ${y} року`, ST.title); mg(4, 0, 4, last);

  const H = 5;
  const hd = (c, t, span = 1, rows2 = true) => { put(H, c, t, ST.head); if (rows2) { box(H, c, H + 1, c + span - 1, ST.head); mg(H, c, H + 1, c + span - 1); } };
  hd(0, "№ п/п"); hd(1, "Підрозділ"); hd(2, "Марка ТЗ"); hd(3, "Номерний знак"); hd(4, `Пробіг за ${mon} м-ць ${y} р`);
  hd(5, "Залишок палива у баку на початок місяця, (л)", 2);
  put(H, cI, `Видано палива ${vin(LOC[m - 1])} м-ці ${y} р.(л)`, ST.head); box(H, cI, H, cIt, ST.head); mg(H, cI, H, cIt);
  cols.forEach((c, i) => put(H + 1, cI + i, c.k, ST.head)); put(H + 1, cIx, "передано з ТЗ на ТЗ", ST.head); put(H + 1, cIt, "всього", ST.head);
  put(H, cC, `Витрачено палива ${vin(LOC[m - 1])} м-ці ${y} р(згідно дор. листів)`, ST.head); box(H, cC, H, cCt, ST.head); mg(H, cC, H, cCt);
  cols.forEach((c, i) => put(H + 1, cC + i, c.k, ST.head)); put(H + 1, cCt, "всього", ST.head);
  hd(cE, "Залишок палива у баку на кінець місяця, (л)", 2);
  for (let c = 0; c <= last; c++) put(H + 2, c, c + 1, ST.ctr);

  let r0 = H + 3, r = r0;
  rows.forEach((row, i) => {
    const v = row.vehicle, R = r + 1;
    put(r, 0, i + 1, ST.ctr); put(r, 1, v.service || "ЦЗ", ST.txt); put(r, 2, nameOf(v), ST.txt); put(r, 3, plateOf(v), ST.ctr);
    put(r, 4, row.km || 0, ST.num); put(r, 5, n2(row.startC), ST.num); put(r, 6, czForm(row, "start"), ST.ctr);
    cols.forEach((c, j) => { const q = c.ids.reduce((a, id) => a + (row.issued[id] || 0), 0); put(r, cI + j, q ? n2(q) : "", ST.num); });
    put(r, cIx, row.xferC ? n2(row.xferC) : "", ST.num);
    put(r, cIt, n2(row.issuedC), ST.num, `SUM(${col(cI)}${R}:${col(cIx)}${R})`);
    cols.forEach((c, j) => { const q = c.ids.reduce((a, id) => a + (row.consBy[id] || 0), 0); put(r, cC + j, q ? n2(q) : "", ST.num); });
    put(r, cCt, n2(row.consC), ST.num, `SUM(${col(cC)}${R}:${col(cCt - 1)}${R})`);
    put(r, cE, n2(row.endC), ST.num, `${col(5)}${R}+${col(cIt)}${R}-${col(cCt)}${R}`);
    put(r, cE + 1, czForm(row, "end"), ST.ctr);
    r++;
  });
  // ВСЬОГО
  put(r, 0, "", ST.totT); put(r, 1, "ВСЬОГО", ST.totT); for (const c of [2, 3, 4, 6, cE + 1]) put(r, c, "", ST.totT);
  const sumC = (c, val) => put(r, c, val, ST.totN, `SUM(${col(c)}${r0 + 1}:${col(c)}${r})`);
  sumC(5, n2(rows.reduce((a, x) => a + x.startC, 0)));
  for (let c = cI; c <= cE; c++) sumC(c, undefined);
  r += 1;
  // Примітка — завжди, порожня (заповнюється вручну)
  put(r, 1, "Примітка", ST.totT); mg(r, 1, r, 2); put(r, 2, "", ST.totT); r++;
  for (let k = 0; k < 10; k++) { box(r, 1, r, cCt, ST.txt); mg(r, cI, r, cCt); r++; }
  r += 1;
  put(r, 1, "Підпис особи, яка склала реєстр:", ST.sign); mg(r, 1, r, 5);
  put(r, 7, String(s.compiler || "").trim(), ST.sign); r++;
  put(r, 5, "(посада та підпис)", ST.cap); mg(r, 5, r, 7);

  const widths = [5, 14, 22, 13, 10, 10, 9, ...cols.map(() => 11), 11, 10, ...cols.map(() => 11), 10, 10, 14];
  return buildSheet({ name: "Реєстр", cells, merges, widths, landscape: true, repeatRows: [H, H + 2], heights: { [H]: 30, [H + 1]: 58, 4: 22 } });
}
