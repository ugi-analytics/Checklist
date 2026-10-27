// Ядро розрахунків «Списання пального для МВО» — перенесено з core.py без змін у правилах.
// Літри зберігаються як цілі «копійки» (сотні долі літра), щоб округлення збігалося з Python Decimal ROUND_HALF_UP.

export const MONTHS_UA = ["січень", "лютий", "березень", "квітень", "травень", "червень",
  "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"];

// ---------------------------------------------------------------- джерела пального («рахунки» МВО)
// Джерело = вид пального + форма видачі + постачальник + МВО. Підрозділ налаштовує їх сам
// («Автопарк і реквізити»); якщо не налаштовано — чотири стандартні джерела, як було раніше.
// У стандартних джерелах вид пального, постачальник і МВО беруться з картки авто.
export const FORMS = ["с/к", "талони", "АЗС ГУНП талони", "АЗС налив"];
export const FORM_GROUP = { "с/к": "tsmk", "талони": "tsmk", "АЗС ГУНП талони": "azs", "АЗС налив": "azs" };
export const LEGACY_SOURCES = [
  { id: "azs_gunp", form: "АЗС ГУНП талони", fuel: "", supplier: "", mvo: "", label: "АЗС ГУНП талони" },
  { id: "talony", form: "талони", fuel: "", supplier: "", mvo: "", label: "Талони" },
  { id: "smk", form: "с/к", fuel: "", supplier: "", mvo: "", label: "см/к" },
  { id: "azs_smk", form: "АЗС налив", fuel: "", supplier: "", mvo: "", label: "АЗС налив" },
];
export const SUPPLIERS = ["UKRNAFTA", "OKKO", "WOG", "BVS"];
/** Авто без основного джерела (заправляється з різних): витрата ділиться між джерелами пропорційно до наявного. */
export const NO_MAIN = "-";

export const ACT_GROUPS = {
  tsmk: { title: "Талони + см/к", file: "Талони_см-к" },
  azs: { title: "АЗС", file: "АЗС" },
};

/** Назва джерела для шапки таблиці (без даних авто). */
export function sourceTitle(src) {
  if (src.label && String(src.label).trim()) return String(src.label).trim();
  return [src.fuel, src.form, src.supplier, src.mvo].map((x) => String(x || "").trim()).filter(Boolean).join(" ") || "Джерело";
}
/** Реквізити джерела для конкретного авто (порожні поля стандартних джерел — з картки авто). */
export function sourceFor(src, v = {}, settings = {}) {
  return { fuel: src.fuel || v.fuel || "", supplier: src.supplier || v.supplier || "",
    mvo: src.mvo || v.mvo_text || settings.compiler || "", form: src.form || "", group: FORM_GROUP[src.form] || "tsmk" };
}
/** Підпис джерела для «Вид пального на залишку»: «А-95 с/к UKRNAFTA Мосійченко». */
export function sourceLabelFor(src, v = {}, settings = {}) {
  const f = sourceFor(src, v, settings);
  const lab = src.label && !LEGACY_SOURCES.some((l) => l.id === src.id) ? String(src.label).trim() : "";
  if (lab) return f.fuel && !normFuel(lab).includes(normFuel(f.fuel)) && !src.fuel ? `${f.fuel} ${lab}` : lab;
  return [f.fuel, f.form, f.supplier, f.mvo].filter(Boolean).join(" ");
}
export function normalizeSource(s = {}) {
  const out = { id: s.id || "s" + Math.random().toString(16).slice(2, 9), fuel: "", form: "с/к", supplier: "", mvo: "", label: "", active: true };
  for (const [k, val] of Object.entries(s)) if (val !== null && val !== undefined) out[k] = val;
  for (const k of ["fuel", "form", "supplier", "mvo", "label"]) out[k] = String(out[k] ?? "").trim();
  out.active = out.active !== false;
  return out;
}
/** Джерела підрозділу (налаштовані або стандартні). */
export function unitSources(settings = {}) {
  const custom = Array.isArray(settings.sources) && settings.sources.length;
  const out = (custom ? settings.sources : LEGACY_SOURCES).map(normalizeSource);
  // стандартний режим (джерела не налаштовано): уся витрата — на основне джерело, залишок — одним числом, як раніше
  out.legacy = !custom;
  // порядок списання: спочатку видане за місяць, потім залишок на початок (налаштування підрозділу)
  out.issuedFirst = custom && !!settings.cons_issued_first;
  return out;
}
/** «А-95», «А95», «A-95» → «А95»; «ДП», «дизель» → «ДП»; «газ», «LPG» → «ГАЗ». */
export function normFuel(f) {
  const s = String(f || "").toUpperCase().replace(/A/g, "А").replace(/[\s\-]/g, "");
  if (!s) return "";
  if (/ДП|ДИЗ/.test(s)) return "ДП";
  if (/ГАЗ|LPG|CNG|ПРОПАН/.test(s)) return "ГАЗ";
  const m = s.match(/А(\d{2})/); if (m) return "А" + m[1];
  return s;
}
/** Чи підходить джерело авто за видом пального (джерело без виду — для всіх). */
export function sourceFits(src, v) {
  const a = normFuel(src.fuel), b = normFuel(v.fuel);
  return !a || !b || a === b || s_in(b, a);
}
const s_in = (x, y) => x.includes(y) || y.includes(x);

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
  if (!out.source) out.source = "smk";
  { const byLabel = Object.fromEntries(LEGACY_SOURCES.map((l) => [l.label, l.id])); if (byLabel[out.source]) out.source = byLabel[out.source]; }
  for (const k of ["rate", "winter_pct"]) { const n = Number(String(out[k] ?? 0).replace(",", ".")); out[k] = isFinite(n) ? n : 0; }
  { const n = Math.trunc(Number(out.month_norm_km || 0)); out.month_norm_km = isFinite(n) ? n : 0; }
  out.active = out.active !== false;
  for (const k of ["name", "fuel", "supplier", "mvo_text"]) out[k] = String(out[k] ?? "").trim();
  return out;
}

// ---------------------------------------------------------------- розрахунок рядка
// Запис місяця по авто: {odo_start, odo_end, manual, consumed_manual,
//   start: {джерело: л} — залишок на початок, iss: {джерело: л} — видано,
//   split: bool, cons: {джерело: л} — ручний розподіл витрати}.
// Старі записи (fuel_start + поля azs_gunp/talony/smk/azs_smk) читаються як є.
const LEGACY_IDS = LEGACY_SOURCES.map((l) => l.id);

/** Джерела, що показуються для авто: підходять за видом пального або вже мають дані. */
export function vehicleSources(v, e, sources) {
  const ee = e || {};
  const has = (id) => [ee.start, ee.iss, ee.cons].some((m) => m && !isBlank(m[id]) && Number(m[id]) !== 0) ||
    (!ee.iss && LEGACY_IDS.includes(id) && !isBlank(ee[id]) && Number(ee[id]) !== 0);
  return sources.filter((s) => s.id === v.source || has(s.id) || (s.active && sourceFits(s, v)));
}
/** МВО, між якими можна розбити залишок авто (з джерел авто). */
export function vehicleMvos(list) { return [...new Set(list.map((s) => String(s.mvo || "").trim()).filter(Boolean))]; }
/** Залишок за МВО (сума за джерелами), напр. для перенесення в наступний місяць. */
export function byMvo(r, which = "end") {
  const by = which === "start" ? r.startBy : r.endBy, out = {};
  for (const s of r.sources) { const c = by[s.id] || 0; if (c && s.mvo) out[s.mvo] = (out[s.mvo] || 0) + c; }
  return out;
}
/** Основне джерело авто (куди йде витрата в першу чергу). */
export function mainSource(v, list) {
  return list.find((s) => s.id === v.source) || list.find((s) => sourceFits(s, v)) || list[0];
}

/** Форма джерела зі слів опису: «т», «тал.», «талони» → талони; «АЗС» → АЗС; «с/к», «см/к» → с/к. */
export function textForms(low) {
  low = String(low || "").toLowerCase();
  if (/талон|(?:^|[\s\d])т(?:ал)?\.?(?=[\s,;]|$)/.test(low)) return ["талони", "АЗС ГУНП талони"];
  if (/азс/.test(low)) return ["АЗС ГУНП талони", "АЗС налив"];
  if (/(?:^|[\s\d])см?\/к/.test(low)) return ["с/к"];
  return null;
}

export function calcRow(vehicle, entry, sources = LEGACY_SOURCES) {
  const e = entry || {};
  const list = vehicleSources(vehicle, e, sources);
  // розбивка залишку на початок за МВО: числа (start_mvo) або довільний опис (start_text, «202 А-95 Мосійченко + 167 … Левошко»)
  const useStartMap = e.start && Object.values(e.start).some((x) => !isBlank(x)) && isBlank(e.fuel_start);
  const splitIn = [];          // [{mvo, c, hint}]
  const textWarn = [];
  if (!useStartMap) {
    const nums = Object.entries(e.start_mvo || {}).filter(([, x]) => !isBlank(x));
    if (nums.length) nums.forEach(([mv, x]) => splitIn.push({ mvo: mv, c: toC(x), hint: "" }));
    else if (!sources.legacy && String(e.start_text || "").trim()) {
      const mvos = [...new Set(sources.map((x) => String(x.mvo || "").trim()).filter(Boolean))];
      for (const part of String(e.start_text).split("+")) {
        const m = part.match(/(?:^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])(\d+(?:[.,]\d+)?)/); if (!m) continue;   // число, але не «А-95»/«А95»
        const low = part.toLowerCase();
        const mv = mvos.find((x) => low.includes(x.toLowerCase().split(/\s+/)[0]));
        if (!mv) {
          // без прізвища: впізнаємо форму — «т», «тал.», «талони» → талони; «АЗС» → АЗС; «с/к», «см/к» → с/к
          const form = textForms(low);
          // давні талони без МВО — окрема частина залишку: не прив'язується до джерела, переноситься як написано
          if (form && form[0] === "талони") { splitIn.push({ keep: true, c: toC(m[1].replace(",", ".")), text: true, label: part.trim() }); continue; }
          const byForm = form && sources.filter((x) => form.includes(x.form))
            .sort((a, b) => (sourceFits(b, vehicle) - sourceFits(a, vehicle)) || (form.indexOf(a.form) - form.indexOf(b.form)) || (list.includes(b) - list.includes(a)))[0];
          if (byForm) { splitIn.push({ mvo: byForm.mvo, c: toC(m[1].replace(",", ".")), hint: low, text: true, src: byForm }); continue; }
        }
        if (!mv) { textWarn.push(`У описі залишку «${part.trim()}» не знайдено МВО чи форму (талони, с/к, АЗС) з джерел — ця частина віднесена на основне джерело.`); continue; }
        splitIn.push({ mvo: mv, c: toC(m[1].replace(",", ".")), hint: low, text: true });
      }
    }
  }
  // джерело для частини залишку: та ж МВО; далі — вид пального, постачальник і форма зі слів опису
  const pickSrc = (sp) => {
    const cand = [...list, ...sources.filter((x) => !list.includes(x))].filter((x) => x.mvo === sp.mvo);
    const score = (x) => (sourceFits(x, vehicle) ? 4 : 0) + (list.includes(x) ? 1 : 0) + (x.id === vehicle.source ? 1 : 0) +
      (sp.hint && x.supplier && sp.hint.includes(x.supplier.toLowerCase()) ? 3 : 0) + (sp.hint && (textForms(sp.hint) || []).includes(x.form) ? 3 : 0) +
      (sp.hint && x.fuel && normFuel(sp.hint).includes(normFuel(x.fuel)) ? 2 : 0);
    return cand.sort((a, b) => score(b) - score(a))[0];
  };
  for (const sp of splitIn) { sp.src = sp.src || (sp.c ? pickSrc(sp) : null); if (sp.src && !list.includes(sp.src)) list.push(sp.src); }
  const free = !sources.legacy && vehicle.source === NO_MAIN;
  const anchor = mainSource(free ? { ...vehicle, source: "" } : vehicle, list);   // куди йде нерозбитий залишок
  const main = free ? null : anchor;
  const r = { vehicle, entry: e, sources: list, mainId: main ? main.id : null, km: 0, normC: 0,
    startBy: {}, startC: 0, issued: {}, issuedC: 0, consBy: {}, consC: 0, endBy: {}, endC: 0,
    capped: false, errors: [], warnings: [], empty: true, keep: [], legacy: !!sources.legacy && !e.split };
  // --- залишок на початок: загальний (fuel_start) + за потреби розбивка за МВО (start_mvo);
  //     старі записи з розбивкою за джерелами (start: {джерело: л}) читаються як є
  const hasStartMap = useStartMap;
  const issMap = e.iss || Object.fromEntries(LEGACY_IDS.map((k) => [k, e[k]]));
  for (const s of list) {
    r.startBy[s.id] = hasStartMap ? toC(e.start[s.id]) : 0;
    r.issued[s.id] = toC(issMap[s.id]); r.issuedC += r.issued[s.id];
    r.consBy[s.id] = 0;
  }
  r.startSplit = hasStartMap && Object.values(e.start).filter((x) => toC(x) > 0).length > 1;
  if (!hasStartMap) {
    const T = toC(e.fuel_start);
    let sum = 0;
    for (const sp of splitIn) {
      sum += sp.c; if (!sp.c) continue;
      if (sp.keep) { r.keep.push({ label: sp.label, c: sp.c, left: sp.c }); continue; }
      if (!sp.src) { r.errors.push(`Залишок за МВО ${sp.mvo}: немає джерела цієї МВО.`); continue; }
      r.startBy[sp.src.id] += sp.c;
    }
    const isText = splitIn.some((x) => x.text);
    if (splitIn.length && sum > T) {
      // більше, ніж загальний залишок, — розбивку не застосовуємо
      for (const sp of splitIn) if (sp.src && !sp.keep) r.startBy[sp.src.id] -= sp.c;
      sum = 0; r.keep = [];
      r.errors.push(`${isText ? "Сума в описі залишку на початок" : "Розбивка залишку на початок за МВО"} (${fmtC(splitIn.reduce((a, x) => a + x.c, 0))} л) більша за загальний залишок (${fmtC(T)} л).`);
    } else if (splitIn.length && sum !== T) {
      (isText ? r.warnings : r.errors).push(`${isText ? "Сума в описі залишку на початок" : "Розбивка залишку на початок за МВО"} (${fmtC(sum)} л) не дорівнює загальному залишку (${fmtC(T)} л)` +
        (isText ? ` — різниця ${fmtC(T - sum)} л віднесена на основне джерело.` : "."));
    }
    r.warnings.push(...textWarn);
    if (anchor && T > sum) r.startBy[anchor.id] += T - sum;
    r.startSplit = splitIn.some((x) => x.c > 0) && sum > 0;
  }
  for (const s of list) r.startC += r.startBy[s.id];
  for (const k of r.keep) r.startC += k.c;
  r.empty = [e.odo_start, e.odo_end, e.consumed_manual].every((x) => isBlank(x) || Number(x) === 0) && r.startC === 0 && r.issuedC === 0;
  // --- пробіг і витрата за нормою
  const s0 = e.odo_start, s1 = e.odo_end;
  if (!isBlank(s0) && !isBlank(s1)) {
    if (Number(s1) < Number(s0)) r.errors.push("Пробіг на кінець менший за пробіг на початок — перевірте показники одометра.");
    else r.km = Math.trunc(Number(s1)) - Math.trunc(Number(s0));
  } else if (isBlank(s0) !== isBlank(s1)) {
    r.warnings.push("Не заповнено пробіг на " + (isBlank(s1) ? "кінець" : "початок") + " періоду.");
  }
  const rateC = toC(vehicle.rate);                 // норма ×100
  const wC = toC(vehicle.winter_pct);              // надбавка % ×100
  r.normC = divHalfUp(r.km * rateC * (10000 + wC), 1000000);
  const avail = r.startC + r.issuedC;
  // загальне правило: витрата за нормою округлюється до цілих літрів (половина — вгору);
  // десяті/соті лишаються лише в залишку (або коли списується весь наявний залишок із сотими)
  r.normExactC = r.normC;
  r.normC = divHalfUp(r.normC, 100) * 100;
  const target = e.manual ? toC(e.consumed_manual) : r.normC;
  if (target > avail) {
    r.capped = true; r.consC = avail;
    r.warnings.push(`За ${e.manual ? "введеною витратою" : "нормою"} потрібно ${fmtC(target)} л, а в наявності ${fmtC(avail)} л — списано ${fmtC(avail)} л, залишок 0.`);
  } else r.consC = target;
  if (e.manual && target !== r.normC && !r.capped)
    r.warnings.push(`Витрату введено вручну (${fmtC(target)} л), за нормою — ${fmtC(r.normC)} л.`);
  // --- розподіл витрати між джерелами
  const availBy = (id) => r.startBy[id] + r.issued[id];
  if (sources.legacy && !e.split) {
    // стандартний режим — точно як раніше: уся витрата в колонку основного джерела, залишок одним числом
    if (main) r.consBy[main.id] = r.consC;
    for (const s of list) r.endBy[s.id] = 0;
    if (main) r.endBy[main.id] = avail - r.consC;
    r.endC = avail - r.consC;
  } else if (e.split && e.cons) {
    let sum = 0;
    for (const s of list) {
      const c = toC(e.cons[s.id]); r.consBy[s.id] = c; sum += c;
      if (c > availBy(s.id)) r.errors.push(`З «${sourceTitle(s)}» списано ${fmtC(c)} л, а там є лише ${fmtC(availBy(s.id))} л.`);
    }
    if (sum !== r.consC) r.errors.push(`Розподіл витрати за джерелами (${fmtC(sum)} л) не дорівнює загальній витраті (${fmtC(r.consC)} л).`);
  } else {
    // як і раніше: уся витрата — на основне джерело авто; якщо там бракує — з інших джерел по черзі
    let left = r.consC;
    const rest = list.filter((s) => !main || s.id !== main.id);
    const order = [...(main ? [main] : []), ...rest.filter((s) => sourceFits(s, vehicle)), ...rest.filter((s) => !sourceFits(s, vehicle))];
    if (sources.issuedFirst) {
      // спершу списується видане за місяць, потім — залишок на початок
      for (const s of order) { if (left <= 0) break; const take = Math.min(left, Math.max(r.issued[s.id], 0)); r.consBy[s.id] += take; left -= take; }
    }
    if (free) {
      // без основного джерела: пропорційно до наявного, спершу серед джерел того ж виду пального
      const fit = order.filter((s) => sourceFits(s, vehicle));
      left = propTake(r, fit, left, availBy);
    }
    for (const s of order) { if (left <= 0) break; const take = Math.min(left, Math.max(availBy(s.id) - r.consBy[s.id], 0)); r.consBy[s.id] += take; left -= take; }
    // джерел не вистачило — списуємо з давніх талонів (їх кількість у описі зменшується)
    for (const k of r.keep) { if (left <= 0) break; const take = Math.min(left, k.left); k.left -= take; left -= take; r.keepConsC = (r.keepConsC || 0) + take; }
    if (left > 0 && anchor) r.consBy[anchor.id] += left;   // не має статися (витрата обмежена наявним)
  }
  if (!(sources.legacy && !e.split)) {
    for (const s of list) { r.endBy[s.id] = availBy(s.id) - r.consBy[s.id]; r.endC += r.endBy[s.id]; }
    for (const k of r.keep) r.endC += k.left;
  }
  // --- видано пальне іншого виду
  for (const s of list) if (r.issued[s.id] > 0 && !sourceFits(s, vehicle))
    r.warnings.push(`Видано за «${sourceTitle(s)}» (${s.fuel}), а авто на ${vehicle.fuel || "?"} — перевірте джерело.`);
  return r;
}

/** Розподіл left між джерелами arr пропорційно до ще не списаного наявного; повертає нерозподілений залишок. */
function propTake(r, arr, left, availBy) {
  const av = arr.map((s) => Math.max(availBy(s.id) - r.consBy[s.id], 0));
  const A = av.reduce((a, b) => a + b, 0);
  if (A <= 0 || left <= 0) return left;
  const want = Math.min(left, A);
  // частки — цілими літрами, якщо витрата ціла (соті лише там, де інакше не вийде)
  const unit = want % 100 === 0 ? 100 : 1;
  const t = av.map((a) => Math.min(a, divHalfUp(want * a, A * unit) * unit));
  let diff = want - t.reduce((a, b) => a + b, 0);
  const idx = av.map((_, i) => i).sort((i, j) => av[j] - av[i]);
  for (const step of unit === 100 ? [100, 1] : [1]) for (const i of idx) {
    if (!diff) break;
    const room = diff > 0 ? av[i] - t[i] : t[i];
    let d = Math.min(Math.abs(diff), room); if (step === 100) d = Math.floor(d / 100) * 100;
    if (Math.abs(diff) < step) continue;
    t[i] += diff > 0 ? d : -d; diff -= diff > 0 ? d : -d;
  }
  arr.forEach((s, i) => { r.consBy[s.id] += t[i]; });
  return left - want;
}

/** «Вид пального на залишку»: «125 А-95 с/к UKRNAFTA Мосійченко + 167 …». */
export function balanceText(r, which = "end", settings = {}) {
  const by = which === "start" ? r.startBy : r.endBy;
  const own = String(((r.entry || {})[which === "start" ? "start_text" : "end_text"]) || "").trim();
  if (own) return own;                                // довільний опис, введений вручну
  if (which === "start" && !r.startSplit) {           // загальний залишок — лише вид пального
    const m = r.sources.find((s) => s.id === r.mainId);
    return m ? sourceFor(m, r.vehicle, settings).fuel : (r.vehicle.fuel || "");
  }
  const parts = r.sources.filter((s) => by[s.id] > 0).map((s) => `${fmtC(by[s.id])} ${sourceLabelFor(s, r.vehicle, settings)}`);
  // давні талони — як написано вручну (кількість — що лишилось)
  for (const k of r.keep || []) {
    const q = which === "start" ? k.c : k.left; if (q <= 0) continue;
    parts.push(q === k.c ? k.label : k.label.replace(/(^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])(\d+(?:[.,]\d+)?)/, (_, p1) => p1 + fmtC(q)));
  }
  if (parts.length) return parts.join(" + ");
  const m = r.sources.find((s) => s.id === r.mainId);
  return m ? sourceFor(m, r.vehicle, settings).fuel : (r.vehicle.fuel || "");
}

export function totals(rows) {
  const t = { km: 0, startC: 0, issued: {}, issuedC: 0, consBy: {}, consC: 0, normC: 0, endC: 0 };
  for (const r of rows) {
    t.km += r.km; t.startC += r.startC; t.issuedC += r.issuedC; t.consC += r.consC; t.normC += r.normC; t.endC += r.endC;
    for (const k in r.issued) t.issued[k] = (t.issued[k] || 0) + r.issued[k];
    for (const k in r.consBy) t.consBy[k] = (t.consBy[k] || 0) + r.consBy[k];
  }
  return t;
}

/** Колонки реєстру: усі налаштовані джерела або лише ті, по яких у місяці був рух. */
export function registerColumns(rows, sources, mode = "all") {
  const t = totals(rows);
  const used = (m) => sources.filter((s) => (m[s.id] || 0) !== 0);
  if (mode === "used") {
    const iss = used(t.issued), cons = used(t.consBy);
    return { iss: iss.length ? iss : sources.slice(0, 1), cons: cons.length ? cons : sources.slice(0, 1) };
  }
  const act = sources.filter((s) => s.active || (t.issued[s.id] || 0) || (t.consBy[s.id] || 0));
  // стандартний режим: у «Витрачено» — три колонки, як у затвердженому зразку (без «АЗС налив»)
  if (sources.legacy) return { iss: act, cons: act.filter((s) => s.id !== "azs_smk" || (t.consBy[s.id] || 0)) };
  return { iss: act, cons: act };
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
export function monthVehicles(vehicles, month, sources = LEGACY_SOURCES) {
  const entries = (month && month.entries) || {};
  const snap = (month && month.snapshot) || {};
  const byId = Object.fromEntries(vehicles.map((v) => [v.id, v]));
  const out = [], seen = new Set();
  for (const v of vehicles) {
    const hasData = v.id in entries && !calcRow(v, entries[v.id], sources).empty;
    if (v.active || hasData) { out.push(v); seen.add(v.id); }
  }
  for (const [vid, e] of Object.entries(entries)) {
    if (!seen.has(vid) && !(vid in byId) && vid in snap) {
      const v = normalizeVehicle(snap[vid]);
      if (!calcRow(v, e, sources).empty) out.push(v);
    }
  }
  return out;
}

/** Початкові дані нового місяця з кінця попереднього (залишок — окремо за кожним джерелом). */
export function carryOver(vehicles, prevMonth, sources = LEGACY_SOURCES, prevSources = sources) {
  if (!prevMonth) return {};
  const out = {};
  for (const v of monthVehicles(vehicles, prevMonth, prevSources)) {
    const pe = prevMonth.entries[v.id];
    if (!pe) continue;
    const r = calcRow(v, pe, prevSources);
    const odo = isBlank(pe.odo_end) ? null : pe.odo_end;
    if (sources.legacy) {
      const start = {};
      for (const [id, c] of Object.entries(r.endBy)) if (c !== 0) start[id] = c / 100;
      out[v.id] = { odo_start: odo, start, iss: {} };
    } else {
      // загальний залишок + розбивка за МВО, якщо на кінець пальне було в кількох МВО
      const bm = byMvo(r, "end"), mv = Object.keys(bm).filter((k) => bm[k] > 0);
      // довільний опис залишку на кінець переходить у опис залишку на початок
      const endText = String(pe.end_text || "").trim() || (mv.length > 1 || r.keep.some((k) => k.left > 0) ? balanceText(r, "end") : "");
      out[v.id] = { odo_start: odo, fuel_start: r.endC ? r.endC / 100 : null, start: {}, iss: {},
        ...(endText ? { start_text: endText } : {}) };
    }
  }
  return out;
}

const num2 = (v) => {
  if (isBlank(v)) return null;
  const x = Number(String(v).replace(",", ".").replace(/\s/g, ""));
  return isFinite(x) ? Math.round(Math.max(x, 0) * 100) / 100 : null;
};
/** Очищення введення картки: {odo_start, odo_end, manual, consumed_manual, split, start:{}, iss:{}, cons:{}}. */
export function cleanEntry(raw = {}) {
  const out = { manual: !!raw.manual, split: !!raw.split };
  for (const f of ["odo_start", "odo_end"]) { const x = num2(raw[f]); out[f] = x === null ? null : Math.round(x); }
  out.consumed_manual = num2(raw.consumed_manual);
  if ("fuel_start" in raw) out.fuel_start = num2(raw.fuel_start);
  for (const t of ["start_text", "end_text"]) { const x = String(raw[t] || "").trim().slice(0, 300); if (x) out[t] = x; }
  for (const m of ["start", "iss", "cons", "start_mvo"]) {
    out[m] = {};
    for (const [k, v] of Object.entries(raw[m] || {})) { const x = num2(v); if (x !== null) out[m][k] = x; }
  }
  if (!out.split) delete out.cons;
  if (!Object.keys(out.start_mvo).length) delete out.start_mvo;
  return out;
}

// ---------------------------------------------------------------- акти списання
/** Групи акта: tsmk (талони + с/к) або azs (АЗС). Рядки — за видом пального, постачальником і МВО. */
export function actGroups(rows, group = null, settings = {}, mvo = null) {
  const m = new Map();
  for (const r of rows) {
    if (r.legacy) {
      // стандартний режим — як раніше: витрата авто ділиться між актами пропорційно до виданого за місяць
      const split = legacyActSplit(r);
      const main = r.sources.find((s) => s.id === r.mainId) || r.sources[0];
      const f = sourceFor(main || {}, r.vehicle, settings);
      for (const g of Object.keys(split)) {
        const q = split[g];
        if (q <= 0 || (group !== null && g !== group) || (mvo !== null && f.mvo !== mvo)) continue;
        const k = JSON.stringify([f.fuel, f.supplier, f.mvo]);
        m.set(k + "|" + g, { k, q: ((m.get(k + "|" + g) || {}).q || 0) + q });
      }
      continue;
    }
    for (const s of r.sources) {
      const q = r.consBy[s.id] || 0;
      if (q <= 0) continue;
      const f = sourceFor(s, r.vehicle, settings);
      if (group !== null && f.group !== group) continue;
      if (mvo !== null && f.mvo !== mvo) continue;
      const k = JSON.stringify([f.fuel, f.supplier, f.mvo]);
      m.set(k + "|" + f.group, { k, q: ((m.get(k + "|" + f.group) || {}).q || 0) + q });
    }
  }
  const agg = new Map();
  for (const { k, q } of m.values()) agg.set(k, (agg.get(k) || 0) + q);
  return [...agg.entries()].map(([k, q]) => { const [fuel, supplier, mvo2] = JSON.parse(k); return { fuel, supplier, mvo: mvo2, qtyC: q }; });
}
function legacyActSplit(r) {
  const out = { tsmk: 0, azs: 0 };
  if (r.consC <= 0) return out;
  const by = { tsmk: 0, azs: 0 };
  for (const s of r.sources) by[FORM_GROUP[s.form] || "tsmk"] += r.issued[s.id] || 0;
  const tot = by.tsmk + by.azs;
  if (tot <= 0) { const main = r.sources.find((s) => s.id === r.mainId); out[(main && FORM_GROUP[main.form]) || "tsmk"] = r.consC; return out; }
  const a = divHalfUp(r.consC * by.tsmk, tot);
  out.tsmk = a; out.azs = r.consC - a;
  return out;
}
/** Перелік актів місяця: один акт на групу (підписує вся комісія); рядки акта — за пальним, постачальником і МВО. */
export function actList(rows, settings = {}) {
  return Object.keys(ACT_GROUPS).map((g) => {
    const lines = actGroups(rows, g, settings);
    const mvos = [...new Set(lines.map((l) => String(l.mvo || "").trim()).filter(Boolean))];
    return { group: g, mvo: null, mvos, lines, title: ACT_GROUPS[g].title };
  });
}

export function shortName(full) {
  const p = String(full || "").trim().split(/\s+/);
  if (p.length >= 3) return `${p[0]} ${p[1][0]}.${p[2][0]}.`;
  if (p.length === 2) return `${p[0]} ${p[1][0]}.`;
  return full || "";
}

/** Джерела для місяця: поточні джерела підрозділу + збережені в місяці (щоб старі дані не зникли). */
export function monthSources(settings = {}, stored = null) {
  const cur = unitSources(settings);
  if (Array.isArray(stored) && stored.length) {
    const have = new Set(cur.map((s) => s.id));
    for (const s of stored) if (!have.has(s.id)) cur.push({ ...normalizeSource(s), active: false });
  }
  return cur;
}

/** Розкладка колонок «Додатка 6»: джерела у «Видано»/«Витрачено», «благо», номери граф із підсумками. */
export function registerLayout(rows, sources, settings = {}) {
  const cols = registerColumns(rows, sources, settings.reg_columns || "all");
  const blago = settings.show_blago !== false;
  const issTotalNo = 7 + cols.iss.length + 1;
  const consTotalNo = issTotalNo + cols.cons.length + (blago ? 1 : 0) + 1;
  return { ...cols, blago, issTotalNo, consTotalNo, total: consTotalNo + 2 };
}
