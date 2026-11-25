// Реєстр «Додаток 6» у Excel для звичайного підрозділу — той самий вигляд, що й у Word (колонки джерел, «Газ/бенз» двома рядками,
// підсумки за групами), числа — числами, «Всього» в рядку — формулою.
import * as C from "./core.js?v=20261125a";
import { buildSheet, ST, ref, col } from "./xlsxcz.js?v=20261125a";

const n2 = (c) => Math.round(c || 0) / 100;

export function buildRegisterXlsx(s, ym, rows, sources = rows.sources || C.unitSources(s)) {
  const lay = C.registerLayout(rows, sources, s);
  const legacy = !!sources.legacy;
  const nI = lay.iss.length;
  // колонки (0-based): 0 №, 1 підрозділ, 2 марка, 3 норма, 4 пробіг, 5 залишок л, 6 вид, 7.. видано, всього, витрачено.., [благо], всього, залишок л, вид
  const cI = 7, cIt = cI + nI, cC = cIt + 1, cB = cC + lay.cons.length, cCt = cB + (lay.blago ? 1 : 0), cE = cCt + 1, cEv = cE + 1, last = cEv;
  const cells = new Map(), merges = [];
  const put = (r, c, v, st = ST.ctrS, f) => cells.set(`${r},${c}`, { v, s: st, f });
  const mg = (r1, c1, r2, c2) => merges.push(`${ref(c1, r1)}:${ref(c2, r2)}`);
  const box = (r1, c1, r2, c2, st) => { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!cells.has(`${r},${c}`)) put(r, c, "", st); };

  put(0, 0, s.org || "ГУНП в Полтавській області", ST.titleL); mg(0, 0, 0, 8);
  put(0, last - 3, s.appendix || "Додаток № 6", ST.titleR); mg(0, last - 3, 0, last);
  put(1, 0, s.dept || s.unit || "", ST.titleL); mg(1, 0, 1, 8);
  put(2, 0, `${s.title || "Реєстр даних згідно дорожніх листів автотранспорту ГУНП в Полтавській області"} за ${C.monthTitle(ym)}`, ST.title); mg(2, 0, 2, last);

  const H = 4;
  const hd2 = (c, t, st = ST.headS) => { put(H, c, t, st); box(H, c, H + 1, c, st); mg(H, c, H + 1, c); };
  hd2(0, "№ п/п"); hd2(1, "Підрозділ", ST.headV); hd2(2, "Марка, днз"); hd2(3, "Місячна норма пробігу (км)", ST.headV); hd2(4, "Пробіг за звітний місяць (км)", ST.headV);
  const grp = (c1, c2, t) => { put(H, c1, t, ST.headS); box(H, c1, H, c2, ST.headS); if (c2 > c1) mg(H, c1, H, c2); };
  grp(5, 6, "Залишок палива у баку на початок місяця, (л)");
  grp(cI, cIt, "Видано палива (л)");
  grp(cC, cCt, "Витрачено палива згідно пробігу, (л)");
  grp(cE, cEv, "Залишок палива у баку на кінець місяця, (л)");
  put(H + 1, 5, "л", ST.headS); put(H + 1, 6, legacy ? "вид" : "вид пального на залишку", legacy ? ST.headS : ST.headV);
  lay.iss.forEach((x, i) => put(H + 1, cI + i, C.sourceTitle(x), ST.headV)); put(H + 1, cIt, "Всього", ST.totS);
  lay.cons.forEach((x, i) => put(H + 1, cC + i, C.sourceTitle(x), ST.headV)); if (lay.blago) put(H + 1, cB, "благо", ST.headS); put(H + 1, cCt, "Всього", ST.totS);
  put(H + 1, cE, "л", ST.headS); put(H + 1, cEv, legacy ? "Вид" : "Вид пального на залишку", ST.headS);
  for (let c = 0; c <= last; c++) put(H + 2, c, c + 1, ST.ctrS);

  let r = H + 3;
  const heights = { 0: 17, 1: 17, 2: 20, [H]: 30, [H + 1]: legacy ? 60 : 95, [H + 2]: 13 };
  const lines = (t, w) => String(t || "").split(/\n/).reduce((a, part) => a + Math.max(1, Math.ceil(part.length / Math.max(1, Math.floor(w / 1.05)))), 0);
  const W = [4, 5.5, 14, 6, 6.5, 6.5, legacy ? 8 : 16, ...Array(nI).fill(5.6), 7, ...Array(lay.cons.length).fill(5.6), ...(lay.blago ? [5] : []), 7, 6.5, legacy ? 18 : 16];

  // рядок даних (r — рядок excel, 0-based): числа + «Всього» формулою
  const line = (p, vals) => {
    const R = r + 1;
    put(r, 5, n2(p.startC), ST.numS); put(r, 6, vals.sv, ST.ctrS);
    lay.iss.forEach((x, i) => { const q = p.issued[x.id] || 0; put(r, cI + i, q ? n2(q) : "", ST.numS); });
    put(r, cIt, n2(p.issuedC), ST.totS, nI ? `ROUND(SUM(${col(cI)}${R}:${col(cIt - 1)}${R}),2)` : undefined);
    lay.cons.forEach((x, i) => { const q = p.consBy[x.id] || 0; put(r, cC + i, q ? n2(q) : "", ST.numS); });
    if (lay.blago) put(r, cB, "", ST.numS);
    // витрата з «давніх» залишків (без колонки джерела) — не в колонках, тому «Всього» — значенням
    const keepCons = (p.keep || []).reduce((a, k) => a + (k.c - Math.max(k.left, 0)), 0);
    put(r, cCt, n2(p.consC), ST.totS, cCt > cC && !keepCons ? `ROUND(SUM(${col(cC)}${R}:${col(cCt - 1)}${R}),2)` : undefined);
    put(r, cE, n2(p.endC), ST.numS); put(r, cEv, vals.ev, ST.ctrS);
    heights[r] = Math.max(heights[r] || 0, 15, lines(vals.sv, W[6]) * 12.5, lines(vals.ev, W[cEv]) * 12.5);
  };
  const dataRows = [];
  const pushRow = (row, i) => {
    const v = row.vehicle;
    const dp = !legacy && C.dualParts(row);
    const head = (span) => {
      const hv = [i + 1, s.unit || "", C.vehName(v), Number(v.month_norm_km) || "", C.isWater(v) ? C.fmtRun(v, row.km) + " м/г" : row.km || 0];
      hv.forEach((t, c) => { put(r, c, t, ST.ctrS); if (span > 1) { box(r, c, r + span - 1, c, ST.ctrS); mg(r, c, r + span - 1, c); } });
      heights[r] = Math.max(heights[r] || 0, lines(C.vehName(v), W[2]) * 12.5 / span);
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
    put(r, 0, label, ST.totS); box(r, 0, r, 4, ST.totS); mg(r, 0, r, 4);
    // у межах групи — формулою (рядки from..r-1), інакше значенням
    const f = (c) => (from !== undefined ? `ROUND(SUM(${col(c)}${from + 1}:${col(c)}${r}),2)` : undefined);
    put(r, 5, n2(t.startC), ST.totS, f(5)); put(r, 6, "", ST.totS);
    lay.iss.forEach((x, i) => put(r, cI + i, n2(t.issued[x.id] || 0), ST.totS, f(cI + i))); put(r, cIt, n2(t.issuedC), ST.totS, f(cIt));
    lay.cons.forEach((x, i) => put(r, cC + i, n2(t.consBy[x.id] || 0), ST.totS, f(cC + i))); if (lay.blago) put(r, cB, "", ST.totS);
    put(r, cCt, n2(t.consC), ST.totS, f(cCt)); put(r, cE, n2(t.endC), ST.totS, f(cE)); put(r, cEv, "", ST.totS);
    heights[r] = 16; r++;
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
    `* дані графи № ${lay.consTotalNo} повинні відповідати кількості пального в акті списання.`]) { put(r, 0, txt, ST.note); mg(r, 0, r, last); r++; }
  r++;
  put(r, 1, "Підпис особи, яка склала реєстр:", ST.sign); mg(r, 1, r, 8); r += 2;
  for (const nm of [s.compiler, s.compiler2, s.compiler3].map((x) => String(x || "").trim()).filter((x, i) => i === 0 || x)) {
    put(r, Math.max(9, cI + 2), `______________________________  ${nm}`, ST.sign); r += 2;
  }
  return buildSheet({ name: "Реєстр", cells, merges, widths: W, landscape: true, repeatRows: [H, H + 2], heights, freeze: [3, H + 3] });
}
