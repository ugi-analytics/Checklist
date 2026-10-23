// Ядро розрахунків «Списання пального для МВО» — перенесено з core.py без змін у правилах.
// Літри зберігаються як цілі «копійки» (сотні долі літра), щоб округлення збігалося з Python Decimal ROUND_HALF_UP.

export const MONTHS_UA = ["січень", "лютий", "березень", "квітень", "травень", "червень",
  "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"];

export const ISSUE_SOURCES = [["azs_gunp", "АЗС ГУНП талони"], ["talony", "Талони"], ["smk", "см/к"], ["azs_smk", "АЗС налив"]];
export const CONSUME_SOURCES = [["azs_gunp", "АЗС ГУНП талони"], ["talony", "Талони"], ["smk", "см/к"]];
export const CONSUME_LABEL = Object.fromEntries(CONSUME_SOURCES);
export const SUPPLIERS = ["UKRNAFTA", "OKKO", "WOG"];
export const ENTRY_FIELDS = ["odo_start", "odo_end", "fuel_start", "azs_gunp", "talony", "smk", "azs_smk", "manual", "consumed_manual"];

export const ACT_GROUPS = {
  tsmk: { title: "Талони + см/к", sources: ["talony", "smk"], file: "Талони_см-к" },
  azs: { title: "АЗС", sources: ["azs_gunp", "azs_smk"], file: "АЗС" },
};
const SOURCE_GROUP = { talony: "tsmk", smk: "tsmk", azs_gunp: "azs", azs_smk: "azs" };

export const DEFAULT_SETTINGS = {
  org: "ГУНП в Полтавській області",
  dept: "",
  unit: "",
  appendix: "Додаток 6",
  title: "Реєстр даних згідно дорожніх листів автотранспорту ГУНП в Полтавській області",
  compiler: "",
  edrpou: "40108630",
  act_approver_title: "",
  act_approver: "",
  act_owner: "",
  act_order: "",
  act_chair_full: "",
  act_chair: "",
  act_members: [],
};
export const SETTING_KEYS = Object.keys(DEFAULT_SETTINGS);

// ---------------------------------------------------------------- числа
const isBlank = (x) => x === null || x === undefined || x === "";
/** Літри → цілі копійки (ROUND_HALF_UP). */
export function toC(x) {
  if (isBlank(x)) return 0;
  const n = Number(String(x).replace(",", "."));
  if (!isFinite(n)) return 0;
  return Math.sign(n) * Math.floor(Math.abs(n) * 100 + 0.5 + 1e-9);
}
/** Ділення невід'ємних цілих з округленням half-up. */
function divHalfUp(num, den) { return den === 0 ? 0 : Math.floor((2 * num + den) / (2 * den)); }

/** Українське написання: 82,44 · 56 · 0 (копійки → рядок). */
export function fmtC(c, blankZero = false) {
  if (c === null || c === undefined) return "";
  if (blankZero && c === 0) return "";
  const neg = c < 0; const a = Math.abs(c);
  let s = `${Math.floor(a / 100)}.${String(a % 100).padStart(2, "0")}`;
  s = s.replace(/0+$/, "").replace(/\.$/, "");
  if (s === "") s = "0";
  return (neg && s !== "0" ? "-" : "") + s.replace(".", ",");
}
export function fmtKm(x) {
  if (isBlank(x)) return "";
  return String(Math.trunc(Number(x))).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}
export const cToNum = (c) => c / 100;

// ---------------------------------------------------------------- місяці
export const ymKey = (y, m) => `${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}`;
export const parseYm = (k) => k.split("-").map(Number);
export function prevYm(k) { const [y, m] = parseYm(k); return m === 1 ? ymKey(y - 1, 12) : ymKey(y, m - 1); }
export function monthTitle(k) { const [y, m] = parseYm(k); return `${MONTHS_UA[m - 1]} ${y} року`; }
export function defaultMonth(d = new Date()) {
  const y = d.getFullYear(), m = d.getMonth() + 1;
  if (d.getDate() <= 10) return m === 1 ? ymKey(y - 1, 12) : ymKey(y, m - 1);
  return ymKey(y, m);
}

// ---------------------------------------------------------------- авто
export function newId() { return "v" + Math.random().toString(16).slice(2, 10); }
export function normalizeVehicle(v = {}) {
  const out = { name: "", rate: 0, month_norm_km: 2640, fuel: "А-95", supplier: "UKRNAFTA", source: "smk", winter_pct: 0,
    mvo_text: "", active: true };
  for (const [k, val] of Object.entries(v)) if (val !== null && val !== undefined) out[k] = val;
  if (!out.id) out.id = newId();
  if (!(out.source in CONSUME_LABEL)) {
    const byLabel = Object.fromEntries(CONSUME_SOURCES.map(([k, l]) => [l, k]));
    out.source = byLabel[out.source] || "smk";
  }
  for (const k of ["rate", "winter_pct"]) { const n = Number(String(out[k] ?? 0).replace(",", ".")); out[k] = isFinite(n) ? n : 0; }
  { const n = Math.trunc(Number(out.month_norm_km || 0)); out.month_norm_km = isFinite(n) ? n : 0; }
  out.active = out.active !== false;
  for (const k of ["name", "fuel", "supplier", "mvo_text"]) out[k] = String(out[k] ?? "").trim();
  return out;
}

// ---------------------------------------------------------------- розрахунок рядка
export function calcRow(vehicle, entry) {
  const e = entry || {};
  const r = { vehicle, entry: e, km: 0, normC: 0, startC: 0, issued: {}, issuedC: 0, consC: 0, consBy: {}, endC: 0,
    capped: false, errors: [], warnings: [], empty: true };
  r.empty = ENTRY_FIELDS.filter((k) => k !== "manual").every((k) => isBlank(e[k]) || Number(e[k]) === 0);
  const s = e.odo_start, en = e.odo_end;
  if (!isBlank(s) && !isBlank(en)) {
    if (Number(en) < Number(s)) r.errors.push("Пробіг на кінець менший за пробіг на початок — перевірте показники одометра.");
    else r.km = Math.trunc(Number(en)) - Math.trunc(Number(s));
  } else if (isBlank(s) !== isBlank(en)) {
    r.warnings.push("Не заповнено пробіг на " + (isBlank(en) ? "кінець" : "початок") + " періоду.");
  }
  const rateC = toC(vehicle.rate);                 // норма ×100
  const wC = toC(vehicle.winter_pct);              // надбавка % ×100
  // норма, коп. = km × rate × (1 + w/100) = km·rateC·(10000+wC) / (100·10000) … ×100 /100
  r.normC = divHalfUp(r.km * rateC * (10000 + wC), 1000000);
  r.startC = toC(e.fuel_start);
  for (const [k] of ISSUE_SOURCES) { r.issued[k] = toC(e[k]); r.issuedC += r.issued[k]; }
  const avail = r.startC + r.issuedC;
  const target = e.manual ? toC(e.consumed_manual) : r.normC;
  if (target > avail) {
    r.capped = true; r.consC = avail;
    r.warnings.push(`За ${e.manual ? "введеною витратою" : "нормою"} потрібно ${fmtC(target)} л, а в наявності ${fmtC(avail)} л — списано ${fmtC(avail)} л, залишок 0.`);
  } else r.consC = target;
  for (const [k] of CONSUME_SOURCES) r.consBy[k] = 0;
  r.consBy[vehicle.source || "smk"] = r.consC;
  r.endC = avail - r.consC;
  if (e.manual && target !== r.normC && !r.capped)
    r.warnings.push(`Витрату введено вручну (${fmtC(target)} л), за нормою — ${fmtC(r.normC)} л.`);
  return r;
}

export function totals(rows) {
  const t = { km: 0, startC: 0, issued: {}, issuedC: 0, consBy: {}, consC: 0, normC: 0, endC: 0 };
  for (const [k] of ISSUE_SOURCES) t.issued[k] = 0;
  for (const [k] of CONSUME_SOURCES) t.consBy[k] = 0;
  for (const r of rows) {
    t.km += r.km; t.startC += r.startC; t.issuedC += r.issuedC; t.consC += r.consC; t.normC += r.normC; t.endC += r.endC;
    for (const k in t.issued) t.issued[k] += r.issued[k];
    for (const k in t.consBy) t.consBy[k] += r.consBy[k];
  }
  return t;
}

export function statusOf(r, e = {}) {
  if (r.errors.length) return "err";
  if (r.empty) return "empty";
  if (isBlank(e.odo_start) || isBlank(e.odo_end)) return "part";
  if (r.capped) return "warn";
  return "ok";
}

// ---------------------------------------------------------------- склад місяця
/** vehicles — довідник підрозділу; month — {entries, snapshot}. */
export function monthVehicles(vehicles, month) {
  const entries = (month && month.entries) || {};
  const snap = (month && month.snapshot) || {};
  const byId = Object.fromEntries(vehicles.map((v) => [v.id, v]));
  const out = [], seen = new Set();
  for (const v of vehicles) {
    const hasData = v.id in entries && !calcRow(v, entries[v.id]).empty;
    if (v.active || hasData) { out.push(v); seen.add(v.id); }
  }
  for (const [vid, e] of Object.entries(entries)) {
    if (!seen.has(vid) && !(vid in byId) && vid in snap) {
      const v = normalizeVehicle(snap[vid]);
      if (!calcRow(v, e).empty) out.push(v);
    }
  }
  return out;
}

/** Початкові дані нового місяця з кінця попереднього. */
export function carryOver(vehicles, prevMonth) {
  if (!prevMonth) return {};
  const out = {};
  for (const v of monthVehicles(vehicles, prevMonth)) {
    const pe = prevMonth.entries[v.id];
    if (!pe) continue;
    const r = calcRow(v, pe);
    out[v.id] = { odo_start: isBlank(pe.odo_end) ? null : pe.odo_end, fuel_start: r.endC / 100 };
  }
  return out;
}

/** Очищення введення: числа ≥ 0 або null. */
export function cleanEntry(raw = {}) {
  const out = {};
  for (const f of ENTRY_FIELDS) {
    const v = raw[f];
    if (f === "manual") { out[f] = !!v; continue; }
    if (isBlank(v)) { out[f] = null; continue; }
    let x = Number(String(v).replace(",", ".").replace(/\s/g, ""));
    if (!isFinite(x)) { out[f] = null; continue; }
    x = Math.max(x, 0);
    out[f] = (f === "odo_start" || f === "odo_end") ? Math.round(x) : Math.round(x * 100) / 100;
  }
  return out;
}

// ---------------------------------------------------------------- акти списання
export function actSplit(r) {
  const out = { tsmk: 0, azs: 0 };
  if (r.consC <= 0) return out;
  const by = {};
  for (const g in ACT_GROUPS) by[g] = ACT_GROUPS[g].sources.reduce((a, k) => a + r.issued[k], 0);
  const tot = by.tsmk + by.azs;
  if (tot <= 0) { out[SOURCE_GROUP[r.vehicle.source || "smk"] || "tsmk"] = r.consC; return out; }
  const a = divHalfUp(r.consC * by.tsmk, tot);
  out.tsmk = a; out.azs = r.consC - a;
  return out;
}

export function actGroups(rows, group = null) {
  const m = new Map();
  for (const r of rows) {
    const q = group === null ? r.consC : actSplit(r)[group];
    if (q <= 0) continue;
    const v = r.vehicle;
    const k = JSON.stringify([v.fuel || "", v.supplier || "", v.mvo_text || ""]);
    m.set(k, (m.get(k) || 0) + q);
  }
  return [...m.entries()].map(([k, q]) => { const [fuel, supplier, mvo] = JSON.parse(k); return { fuel, supplier, mvo, qtyC: q }; });
}

export function shortName(full) {
  const p = String(full || "").trim().split(/\s+/);
  if (p.length >= 3) return `${p[0]} ${p[1][0]}.${p[2][0]}.`;
  if (p.length === 2) return `${p[0]} ${p[1][0]}.`;
  return full || "";
}
