// Реєстр «Додаток 6» у Excel для звичайного підрозділу — той самий вигляд, що й у Word (колонки джерел, «Газ/бенз» двома рядками,
// підсумки за групами), числа — числами, «Всього» в рядку — формулою.
import * as C from "./core.js?v=20261125d";
import { buildSheet, ST, ref, col } from "./xlsxcz.js?v=20261125d";

const n2 = (c) => Math.round(c || 0) / 100;

export function buildRegisterXlsx(s, ym, rows, sources = rows.sources || C.unitSources(s)) {
  // для друку — лише джерела, з яких цього місяця щось видано чи витрачено (як у реєстрі ЦЗ)
  const lay0 = C.registerLayout(rows, sources, s);
  const movI = (x) => rows.some((r) => r.issued[x.id] || 0), movC = (x) => rows.some((r) => r.consBy[x.id] || 0);
  // джерела з однаковою назвою (без різниці великих/малих літер) — в одній колонці, як у реєстрі ЦЗ
  const grpCols = (list, mov) => { const out = []; for (const x of list.filter(mov)) { const t = C.sourceTitle(x).replace(/\s+/g, " ").trim(), k = t.toLowerCase();
    let g = out.find((y) => y.k === k); if (!g) out.push(g = { k, t, ids: [] }); g.ids.push(x.id); } return out; };
  const lay = { ...lay0, iss: grpCols(lay0.iss, movI), cons: grpCols(lay0.cons, movC) };
  const qq = (o, x) => x.ids.reduce((a, id) => a + ((o || {})[id] || 0), 0);
  lay.issTotalNo = 7 + lay.iss.length + 1; lay.consTotalNo = lay.issTotalNo + lay.cons.length + (lay.blago ? 1 : 0) + 1;
  const legacy = !!sources.legacy;
  const nI = lay.iss.length;
  // колонки (0-based): 0 №, 1 підрозділ, 2 марка, 3 норма, 4 пробіг, 5 залишок л, 6 вид, 7.. видано, всього, витрачено.., [благо], всього, залишок л, вид
  const cI = 7, cIt = cI + nI, cC = cIt + 1, cB = cC + lay.cons.length, cCt = cB + (lay.blago ? 1 : 0), cE = cCt + 1, cEv = cE + 1, last = cEv;
  const cells = new Map(), merges = [];
  const put = (r, c, v, st = ST.ctr, f) => cells.set(`${r},${c}`, { v, s: st, f });
  const mg = (r1, c1, r2, c2) => merges.push(`${ref(c1, r1)}:${ref(c2, r2)}`);
  const box = (r1, c1, r2, c2, st) => { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!cells.has(`${r},${c}`)) put(r, c, "", st); };

  put(0, last - 3, s.appendix || "Додаток № 6", ST.titleR); mg(0, last - 3, 0, last);
  put(1, 0, ` ${s.org || "ГУНП в Полтавській області"}`, ST.title); mg(1, 0, 1, 6);
  put(2, 0, "(найменування підприємства)", ST.cap); mg(2, 0, 2, 6);
  put(3, 0, s.dept || s.unit || "", ST.title); mg(3, 0, 3, 6);
  put(4, 0, "(найменування структурного підрозділу)", ST.cap); mg(4, 0, 4, 6);
  put(5, 0, `${s.title || "Реєстр даних згідно дорожніх листів автотранспорту ГУНП в Полтавській області"} за ${C.monthTitle(ym)}`, ST.title); mg(5, 0, 5, last);
  const H = 6;
  const hd2 = (c, t, st = ST.head) => { put(H, c, t, st); box(H, c, H + 1, c, st); mg(H, c, H + 1, c); };
  hd2(0, "№ п/п"); hd2(1, "Підрозділ", ST.head); hd2(2, "Марка, днз"); hd2(3, "Місячна норма пробігу (км)", ST.head); hd2(4, "Пробіг за звітний місяць (км)", ST.head);
  const grp = (c1, c2, t) => { put(H, c1, t, ST.head); box(H, c1, H, c2, ST.head); if (c2 > c1) mg(H, c1, H, c2); };
  grp(5, 6, "Залишок палива у баку на початок місяця, (л)");
  grp(cI, cIt, "Видано палива (л)");
  grp(cC, cCt, "Витрачено палива згідно пробігу, (л)");
  grp(cE, cEv, "Залишок палива у баку на кінець місяця, (л)");
  put(H + 1, 5, "л", ST.head); put(H + 1, 6, legacy ? "вид" : "вид пального на залишку", ST.head);
  lay.iss.forEach((x, i) => put(H + 1, cI + i, x.t, ST.headV10)); put(H + 1, cIt, "Всього", ST.head);
  lay.cons.forEach((x, i) => put(H + 1, cC + i, x.t, ST.headV10)); if (lay.blago) put(H + 1, cB, "благо", ST.head); put(H + 1, cCt, "Всього", ST.head);
  put(H + 1, cE, "л", ST.head); put(H + 1, cEv, legacy ? "Вид" : "Вид пального на залишку", ST.head);
  for (let c = 0; c <= last; c++) put(H + 2, c, c + 1, ST.ctr);

  let r = H + 3;
  // колонки джерел вузькі (назви — вертикально); що більше джерел, то вужчі (числа до 3–4 знаків)
  const SW = nI + lay.cons.length > 16 ? 5.4 : 6.5;
  const W = [4.5, 10, 14, 7, 7, 7, legacy ? 10 : 20, ...Array(nI).fill(SW), 7.5, ...Array(lay.cons.length).fill(SW), ...(lay.blago ? [5.5] : []), 7.5, 7.5, legacy ? 16 : 20];
  // висота: скільки рядків займе перенос (шрифт 11 жирний ≈ 1,15 одиниці ширини на символ; заголовки — шрифт 10 ≈ 1,05)
  const lines = (t, w, k = 1.3) => String(t || "").split(/\n/).reduce((a, part) => a + Math.max(1, Math.ceil(part.length / Math.max(1, Math.floor(w / k)))), 0);
  // заголовки джерел — вертикально: висота за найдовшою назвою (шрифт 10 жирний; у колонці вміщується 2 рядки)
  const srcHead = Math.max(60, ...[...lay.iss, ...lay.cons].map((x) => Math.min(200, Math.ceil(x.t.length / 2) * 6.3 + 16)));
  const heights = { 0: 16, 1: 18, 2: 12, 3: 18, 4: 12, 5: 22, [H]: 46, [H + 1]: srcHead, [H + 2]: 14 };
  // рядок даних (r — рядок excel, 0-based): числа + «Всього» формулою
  const line = (p, vals) => {
    const R = r + 1;
    put(r, 5, n2(p.startC), ST.num); put(r, 6, vals.sv, ST.ctr);
    lay.iss.forEach((x, i) => { const q = qq(p.issued, x); put(r, cI + i, q ? n2(q) : "", ST.num); });
    put(r, cIt, n2(p.issuedC), ST.totN, nI ? `ROUND(SUM(${col(cI)}${R}:${col(cIt - 1)}${R}),2)` : undefined);
    lay.cons.forEach((x, i) => { const q = qq(p.consBy, x); put(r, cC + i, q ? n2(q) : "", ST.num); });
    if (lay.blago) put(r, cB, "", ST.num);
    // витрата з «давніх» залишків (без колонки джерела) — не в колонках, тому «Всього» — значенням
    const keepCons = (p.keep || []).reduce((a, k) => a + (k.c - Math.max(k.left, 0)), 0);
    put(r, cCt, n2(p.consC), ST.totN, cCt > cC && !keepCons ? `ROUND(SUM(${col(cC)}${R}:${col(cCt - 1)}${R}),2)` : undefined);
    put(r, cE, n2(p.endC), ST.num); put(r, cEv, vals.ev, ST.ctr);
    heights[r] = Math.max(heights[r] || 0, 16, lines(vals.sv, W[6]) * 14.5, lines(vals.ev, W[cEv]) * 14.5);
  };
  const dataRows = [];
  const pushRow = (row, i) => {
    const v = row.vehicle;
    const dp = !legacy && C.dualParts(row);
    const head = (span) => {
      const hv = [i + 1, s.unit || "", C.vehName(v), Number(v.month_norm_km) || "", C.isWater(v) ? C.fmtRun(v, row.km) + " м/г" : row.km || 0];
      hv.forEach((t, c) => { put(r, c, t, ST.ctr); if (span > 1) { box(r, c, r + span - 1, c, ST.ctr); mg(r, c, r + span - 1, c); } });
      heights[r] = Math.max(heights[r] || 0, Math.max(lines(C.vehName(v), W[2]), lines(s.unit || "", W[1])) * 14.5 / span);
    };
    if (dp) {
      head(2);
      dp.forEach((p, k) => {
        const sv = (p.startC || k === 0 ? C.balanceText(p, "start", s) : "газ").split(" + ").join(" +\n");
        const ev = (p.endC || k === 0 ? C.balanceText(p, "end", s) : "газ").split(" + ").join(" +\n");
        dataRows.push(r); line(p, { sv, ev }); r++;
      });
      return;
    }
    head(1);
    let sv, ev;
    if (legacy) { sv = C.legacyStartVid(row); ev = v.fuel || ""; }
    else { sv = C.balanceText(row, "start", s).split(" + ").join(" +\n"); ev = C.balanceText(row, "end", s).split(" + ").join(" +\n"); }
    dataRows.push(r); line(row, { sv, ev }); r++;
  };
  const totRow = (label, rr, from) => {
    const t = C.totals(rr);
    put(r, 0, label, ST.totN); box(r, 0, r, 4, ST.totN); mg(r, 0, r, 4);
    // у межах групи — формулою (рядки from..r-1), інакше значенням
    const f = (c) => (from !== undefined ? `ROUND(SUM(${col(c)}${from + 1}:${col(c)}${r}),2)` : undefined);
    put(r, 5, n2(t.startC), ST.totN, f(5)); put(r, 6, "", ST.totN);
    lay.iss.forEach((x, i) => put(r, cI + i, n2(qq(t.issued, x)), ST.totN, f(cI + i))); put(r, cIt, n2(t.issuedC), ST.totN, f(cIt));
    lay.cons.forEach((x, i) => put(r, cC + i, n2(qq(t.consBy, x)), ST.totN, f(cC + i))); if (lay.blago) put(r, cB, "", ST.totN);
    put(r, cCt, n2(t.consC), ST.totN, f(cCt)); put(r, cE, n2(t.endC), ST.totN, f(cE)); put(r, cEv, "", ST.totN);
    heights[r] = 18; r++;
  };
  const groups = C.registerGroups(rows, s);
  groups.forEach((g) => {
    const from = r;
    g.rows.forEach((row, i) => pushRow(row, i));
    if (groups.length > 1 || g.title) totRow(`Всього${g.title ? " (" + g.title + ")" : ""}`, g.rows, from);
  });
  if (groups.length > 1) totRow("Разом", rows);
  else if (!groups[0].title) totRow("Всього", rows, H + 3);

  r++;
  for (const txt of [`* дані графи № ${lay.issTotalNo} повинні відповідати даним граф № 11, 12 звіту про одержане і видане пальне.`,
    `* дані графи № ${lay.consTotalNo} повинні відповідати кількості пального в акті списання.`]) { put(r, 0, txt, ST.plain); mg(r, 0, r, last); r++; }
  r++;
  put(r, 1, "Підпис особи, яка склала реєстр:", ST.sign); mg(r, 1, r, 8); r += 2;
  for (const nm of [s.compiler, s.compiler2, s.compiler3].map((x) => String(x || "").trim()).filter((x, i) => i === 0 || x)) {
    put(r, Math.max(9, cI + 2), `______________________________  ${nm}`, ST.sign); r += 2;
  }
  return buildSheet({ name: "Реєстр", cells, merges, widths: W, landscape: true, repeatRows: [H, H + 2], heights, freeze: [3, H + 3] });
}
