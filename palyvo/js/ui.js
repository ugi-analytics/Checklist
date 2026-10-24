// HTML-фрагменти карток, KPI і попереднього перегляду (перенесено з ui.py).
import * as C from "./core.js?v=20261024a";

export const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

export const STATUS_TXT = { ok: "● Заповнено", part: "◐ Заповнено частково", warn: "⚠ Пального не вистачило — списано наявне",
  err: "⛔ Помилка в показниках", empty: "○ Не заповнено" };

export function splitName(name) {
  const p = String(name || "").trim().split(/\s+/);
  if (p.length > 1 && /\d/.test(p[p.length - 1])) return [p.slice(0, -1).join(" "), p[p.length - 1]];
  return [name || "", ""];
}
const fmtNum = (x) => C.fmtC(C.toC(x));

export function idHtml(v, status) {
  const [model, plate] = splitName(v.name);
  const head = plate
    ? `<div class='model'>${esc(model)}</div><div class='fr-plate'><div class='ua'><i></i>UA</div><div class='n'>${esc(plate)}</div></div>`
    : `<div class='fr-name'>${esc(model) || "—"}</div>`;
  const rate = `${fmtNum(v.rate)} л/100 км` + (v.winter_pct ? ` +${fmtNum(v.winter_pct)}%` : "");
  let chips = `<span class='fr-chip y'>${rate}</span>`;
  if (v.fuel) chips += `<span class='fr-chip'>${esc(v.fuel)}</span>`;
  chips += `<span class='fr-chip'>${C.CONSUME_LABEL[v.source] || ""}</span>`;
  const sub = v.month_norm_km ? `норма пробігу ${C.fmtKm(v.month_norm_km)} км/міс` : "";
  return `<div class='fr-id'>${head}<div class='fr-chips'>${chips}</div><div class='fr-sub'>${sub}</div>` +
    `<div class='fr-st ${status}'>${STATUS_TXT[status]}</div></div>`;
}

export function kmHtml(v, r) {
  const norm = v.month_norm_km || 0;
  if (!r.km && !norm) return "";
  const pct = norm ? (r.km / norm) * 100 : 0;
  const right = norm ? `${Math.round(pct)}% від норми ${C.fmtKm(norm)} км` : "";
  const over = pct > 100;
  const bar = over ? "linear-gradient(90deg,#0ea5e9,#f59e0b)" : "linear-gradient(90deg,#0ea5e9,#22d3ee)";
  return `<div class='fr-km'><div class='t'><span>Пробіг за місяць <b>${C.fmtKm(r.km)} км</b></span>` +
    `<span style='color:${over ? "#fbbf24" : "#cbd5e1"}'>${right}</span></div>` +
    `<div class='b'><i style='width:${Math.min(pct, 100).toFixed(0)}%;background:${bar}'></i></div></div>`;
}

export function tankHtml(r, status) {
  if (status === "empty") {
    return "<div class='fr-tank-wrap'><div class='fr-tank'><div class='fr-ring' style='--p:0%'><div>⛽</div></div>" +
      "<div><div class='lb'>Залишок у баку<br>на кінець</div><div class='v' style='color:#94a3b8'>— <small>л</small></div></div></div></div>";
  }
  const avail = r.startC + r.issuedC;
  const p = avail > 0 ? (r.endC / avail) * 100 : 0;
  const sp = avail > 0 ? (r.startC / avail) * 100 : 0;
  return `<div class='fr-tank-wrap'><div class='fr-tank'><div class='fr-ring' style='--p:${p.toFixed(1)}%'><div>⛽</div></div>` +
    `<div><div class='lb'>Залишок у баку<br>на кінець</div><div class='v'>${C.fmtC(r.endC)} <small>л</small></div></div></div>` +
    `<div class='fr-flow'><i style='width:${sp.toFixed(1)}%;background:#38bdf8'></i><i style='width:${(100 - sp).toFixed(1)}%;background:#22c55e'></i></div>` +
    `<div class='fr-leg'><span><span class='fr-dot' style='background:#38bdf8'></span>було <b>${C.fmtC(r.startC)}</b></span>` +
    `<span><span class='fr-dot' style='background:#22c55e'></span>видано <b>${C.fmtC(r.issuedC)}</b></span>` +
    `<span><span class='fr-dot' style='background:#f97316'></span>витрачено <b>${C.fmtC(r.consC)}</b></span>` +
    `<span><span class='fr-dot' style='background:#facc15'></span>лишилось <b>${C.fmtC(r.endC)}</b></span></div></div>`;
}

export function noteHtml(v, r, e) {
  const src = C.CONSUME_LABEL[v.source] || "";
  if (r.empty) return `Витрата рахується за нормою · списання на ${src}`;
  if (e.manual) return `Витрачено вручну <b>${C.fmtC(r.consC)} л</b> · за нормою ${C.fmtC(r.normC)} л · на ${src}`;
  if (r.capped) return `За нормою <b>${C.fmtC(r.normC)} л</b>, в наявності <b>${C.fmtC(r.consC)} л</b> · списано наявне на ${src}`;
  return `Витрачено за нормою <b>${C.fmtC(r.consC)} л</b> · списано на ${src}`;
}

export function msgsHtml(r) {
  return r.errors.map((m) => `<div class='fr-msg e'>⛔ ${esc(m)}</div>`).join("") +
    r.warnings.filter((m) => !m.startsWith("За нормою потрібно")).map((m) => `<div class='fr-msg w'>⚠️ ${esc(m)}</div>`).join("");
}

export function kpiHtml(rows, statuses, monthLabel) {
  const t = C.totals(rows);
  const n = rows.length;
  const done = statuses.filter((s) => s === "ok" || s === "warn").length;
  const left = rows.filter((r, i) => !["ok", "warn"].includes(statuses[i])).map((r) => r.vehicle.name);
  const pct = n ? (done / n) * 100 : 0;
  const leftTxt = left.length ? "Не заповнено: " + left.map(esc).join(", ") : "Усі авто заповнено ✓";
  const tile = (icon, bg, glow, label, value, unit, color = "#f8fafc") =>
    `<div class='fr-k'><div class='glow' style='background:${glow}'></div><div class='l'><span class='ic' style='background:${bg}'>${icon}</span>${label}</div>` +
    `<div class='v' style='color:${color}'>${value} <small>${unit}</small></div></div>`;
  return "<div class='fr-kpis'>" +
    `<div class='fr-k'><div class='l'><span class='ic' style='background:rgba(245,158,11,.15)'>✓</span>Готовність реєстру · ${esc(monthLabel)}</div>` +
    `<div class='v'>${done} <small>з ${n} авто</small></div><div class='bar'><i style='width:${pct.toFixed(0)}%'></i></div><div class='s'>${leftTxt}</div></div>` +
    tile("🛣", "rgba(14,165,233,.15)", "#0ea5e9", "Пробіг", C.fmtKm(t.km), "км") +
    tile("⛽", "rgba(34,197,94,.15)", "#22c55e", "Видано", C.fmtC(t.issuedC), "л") +
    tile("🔥", "rgba(249,115,22,.15)", "#f97316", "Витрачено", C.fmtC(t.consC), "л") +
    tile("🛢", "rgba(250,204,21,.15)", "#facc15", "Залишок у баках", C.fmtC(t.endC), "л", "#facc15") + "</div>" +
    `<div class='fr-eq'>Баланс місяця: <b>${C.fmtC(t.startC)}</b> на початок <span class='op'>+</span><b>${C.fmtC(t.issuedC)}</b> видано ` +
    `<span class='op'>−</span><b>${C.fmtC(t.consC)}</b> витрачено<span class='op'>=</span><b class='rs'>${C.fmtC(t.endC)} л</b> у баках` +
    `<span style='color:#64748b'>· це графа 18, рядок «Всього»</span></div>`;
}

export function previewHtml(rows, unit) {
  const u = esc(unit || "");
  const h = ["<table class='fr-tbl'><tr><th rowspan=2>№ п/п</th><th rowspan=2>Підрозділ</th><th rowspan=2>Марка, днз</th>" +
    "<th rowspan=2>Місячна норма пробігу (км)</th><th rowspan=2>Пробіг за звітний місяць (км)</th>" +
    "<th colspan=2>Залишок на початок місяця, (л)</th><th colspan=5>Видано палива (л)</th>" +
    "<th colspan=5>Витрачено палива згідно пробігу, (л)</th><th colspan=2>Залишок на кінець місяця, (л)</th></tr><tr>" +
    "<th>л</th><th>вид</th><th>АЗС ГУНП талони</th><th>Талони</th><th>см/к</th><th>АЗС налив</th><th>Всього</th>" +
    "<th>АЗС ГУНП талони</th><th>Талони</th><th>см/к</th><th>благо</th><th>Всього</th><th>л</th><th>Вид</th></tr>",
  "<tr class='num'>" + Array.from({ length: 19 }, (_, i) => `<td>${i + 1}</td>`).join("") + "</tr>"];
  rows.forEach((r, i) => {
    const v = r.vehicle;
    let vid = esc((v.fuel || "") + (v.supplier ? ` (${v.supplier})` : ""));
    if (v.mvo_text) vid += "<br>МВО " + esc(v.mvo_text);
    const cells = [i + 1, u, esc(v.name), v.month_norm_km || "", C.fmtKm(r.km), C.fmtC(r.startC), esc(v.fuel || "")];
    for (const [k] of C.ISSUE_SOURCES) cells.push(C.fmtC(r.issued[k], true));
    cells.push(C.fmtC(r.issuedC));
    for (const [k] of C.CONSUME_SOURCES) cells.push(C.fmtC(r.consBy[k], true));
    cells.push("", C.fmtC(r.consC), C.fmtC(r.endC), vid);
    h.push("<tr>" + cells.map((c, j) => `<td class='${j === 11 || j === 16 ? "b" : ""}'>${c}</td>`).join("") + "</tr>");
  });
  const t = C.totals(rows);
  const tc = [C.fmtC(t.startC), "", ...C.ISSUE_SOURCES.map(([k]) => C.fmtC(t.issued[k], true)), C.fmtC(t.issuedC),
    ...C.CONSUME_SOURCES.map(([k]) => C.fmtC(t.consBy[k], true)), "", C.fmtC(t.consC), C.fmtC(t.endC), ""];
  h.push("<tr class='tot'><td colspan=5>Всього</td>" + tc.map((c) => `<td>${c}</td>`).join("") + "</tr></table>");
  return h.join("");
}
