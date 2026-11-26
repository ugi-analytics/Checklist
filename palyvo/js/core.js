// Ядро розрахунків «Списання пального для МВО» — перенесено з core.py без змін у правилах.
// Літри зберігаються як цілі «копійки» (сотні долі літра), щоб округлення збігалося з Python Decimal ROUND_HALF_UP.

export const MONTHS_UA = ["січень", "лютий", "березень", "квітень", "травень", "червень",
  "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"];

// ---------------------------------------------------------------- джерела пального («рахунки» МВО)
// Джерело = вид пального + форма видачі + постачальник + МВО. Підрозділ налаштовує їх сам
// («Автопарк і реквізити»); якщо не налаштовано — чотири стандартні джерела, як було раніше.
// У стандартних джерелах вид пального, постачальник і МВО беруться з картки авто.
export const FORMS = ["с/к", "талони", "АЗС ГУНП талони", "АЗС налив", "Налив ОВА", "Субвенція", "з/ф"];
export const FORM_GROUP = { "с/к": "tsmk", "талони": "tsmk", "АЗС ГУНП талони": "azs", "АЗС налив": "azs", "Налив ОВА": "ova", "Субвенція": "azs", "з/ф": "zf" };
export const LEGACY_SOURCES = [
  { id: "azs_gunp", form: "АЗС ГУНП талони", fuel: "", supplier: "", mvo: "", label: "АЗС ГУНП талони" },
  { id: "talony", form: "талони", fuel: "", supplier: "", mvo: "", label: "Талони" },
  { id: "smk", form: "с/к", fuel: "", supplier: "", mvo: "", label: "см/к" },
  { id: "azs_smk", form: "АЗС налив", fuel: "", supplier: "", mvo: "", label: "АЗС налив" },
];
/** Коротка форма для опису залишку. */
export const FORM_SHORT = { "с/к": "с/к", "талони": "т", "АЗС ГУНП талони": "АЗС т", "АЗС налив": "АЗС", "Налив ОВА": "ОВА", "Субвенція": "субв.", "з/ф": "з/ф" };
export const SUPPLIERS = ["UKRNAFTA", "OKKO", "WOG", "BVS", "Інше"];
/** Авто без основного джерела (заправляється з різних): витрата — спершу із залишку, потім із виданого, щоб джерела вичерпувались повністю. */
export const NO_MAIN = "-";
/** Видача до цієї кількості (10 л) списується першочергово. */
export const SMALL_ISSUE_C = 1000;

export const ACT_GROUPS = {
  tsmk: { title: "Талони + см/к", file: "Талони_см-к" },
  azs: { title: "АЗС", file: "АЗС" },
  ova: { title: "Налив ОВА", file: "Налив_ОВА" },
  zf: { title: "З/ф (загальний фонд)", file: "Загальний_фонд" },
};

/** Назва джерела для шапки таблиці (без даних авто). */
export function sourceTitle(src) {
  if (src.label && String(src.label).trim()) return String(src.label).trim();
  return [src.fuel, src.form, src.supplier, src.mvo].map((x) => String(x || "").trim()).filter(Boolean).join(" ") || "Джерело";
}
/** Режим ЦЗ (лише для кабінету з profile = "cz"): чотири види актів, у кожному — окремо за МВО. */
export const CZ_ACTS = { ts: "С/к, талони", nal: "Налив", nalz: "Налив/з (загальний фонд)", ova: "ОВА" };
export const czActOf = (src) => {
  if (src.cz_act && CZ_ACTS[src.cz_act]) return src.cz_act;
  if (src.form === "Налив ОВА" || /ова/i.test(`${src.label || ""} ${src.supplier || ""}`)) return "ova";
  if (/налив\s*\/\s*з|н\/з|загальн/i.test(`${src.label || ""} ${src.supplier || ""}`)) return "nalz";
  if (src.form === "АЗС налив") return "nal";
  return "ts";
};
/** Субвенція / благодійна допомога: окремий акт із назвою джерела. Без явної галочки — BVS або «субвенц…», «благо…». */
export const autoSubv = (src) => /^bvs$/i.test(String(src.supplier || "").trim()) || /субвенц|благо/i.test(`${src.supplier || ""} ${src.label || ""}`);
export const isSubv = (src) => src.form === "Субвенція" || (src.subv === undefined || src.subv === null ? autoSubv(src) : !!src.subv);
const subvGroup = (src) => "sv:" + sourceTitle(src);
/** Реквізити джерела для конкретного авто (порожні поля стандартних джерел — з картки авто). */
export function sourceFor(src, v = {}, settings = {}) {
  const custom = Array.isArray(settings.sources) && settings.sources.length;   // свої джерела: МВО — лише з джерела
  return { fuel: src.fuel || v.fuel || "", supplier: src.supplier || "",
    mvo: src.mvo || (custom ? "" : v.mvo_text || settings.compiler) || "", form: src.form || "", group: settings.profile === "cz" ? "cz:" + czActOf(src) : custom && isSubv(src) ? subvGroup(src) : FORM_GROUP[src.form] || "tsmk" };
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
/** Групи реєстру за видом пального (налаштування «розділяти авто за видом пального»). */
export function fuelGroupOf(fuel) {
  const f = normFuel(fuel);
  if (/^А\d{2}/.test(f) || /БЕНЗ/.test(f)) return { key: "1", title: "бензин" };
  if (f === "ДП") return { key: "2", title: "дизельне паливо" };
  if (f === "ГАЗ") return { key: "3", title: "газ" };
  return { key: "4", title: f ? f.toLowerCase() : "інше" };
}
/** Рядки реєстру, розбиті на групи за пальним: [{title, rows}] (якщо налаштування вимкнене — одна група). */
export function registerGroups(rows, settings = {}) {
  if (!settings.reg_by_fuel) return [{ title: "", rows }];
  const m = new Map();
  for (const r of rows) { const g = fuelGroupOf(r.vehicle.fuel); if (!m.has(g.key)) m.set(g.key, { title: g.title, rows: [] }); m.get(g.key).rows.push(r); }
  return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([, g]) => g);
}

/** Розбивка залишків і руху за видом (форма + МВО) для рядків групи реєстру — підрядки під «Всього» (налаштування reg_breakdown). */
export function balanceBreakdown(rows, settings = {}) {
  const m = new Map();
  const add = (label, f) => { if (!m.has(label)) m.set(label, { label, startC: 0, issuedC: 0, consC: 0, endC: 0, iss: {}, cons: {} }); f(m.get(label)); };
  for (const r of rows) {
    for (const s of r.sources) {
      const st = r.startBy[s.id] || 0, is = r.issued[s.id] || 0, co = r.consBy[s.id] || 0, en = r.endBy[s.id] || 0;
      if (!st && !is && !co && !en) continue;
      const f = sourceFor(s, r.vehicle, settings);
      const label = [FORM_SHORT[f.form] || f.form, f.mvo].filter(Boolean).join(" ") || sourceTitle(s);
      add(label, (x) => { x.startC += st; x.issuedC += is; x.consC += co; x.endC += en; x.iss[s.id] = (x.iss[s.id] || 0) + is; x.cons[s.id] = (x.cons[s.id] || 0) + co; });
    }
    for (const k of r.keep || []) {
      const label = String(k.label || "").replace(/(^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])\d+(?:[.,]\d+)?\s*/, "$1").trim() + (k.mvo && !k.own ? " " + k.mvo : "");
      add(label || "давні талони", (x) => { x.startC += k.c; x.consC += k.c - k.left; x.endC += k.left; });
    }
  }
  return [...m.values()];
}

/** Чи підходить джерело авто за видом пального (джерело без виду — для всіх). */
/** Авто на двох видах пального: газ + бензин («Газ/бенз») — у реєстрі два рядки (бензин і газ). */
export function isDual(v) { const f = String((v && v.fuel) || "").toLowerCase(); return /газ|lpg|пропан/.test(f) && /бенз|[аa]-?9\d/.test(f); }
export const isGasSrc = (s) => normFuel(s && s.fuel) === "ГАЗ";
export function sourceFits(src, v) {
  const a = normFuel(src.fuel), b = normFuel(v.fuel);
  if (isDual(v)) return !a || a === "ГАЗ" || /^А\d\d$/.test(a) || /БЕНЗ/.test(a);
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
/** Реквізити й комісія, які ПОГ бере з підрозділу (МВО — сам ПОГ, тому «хто склав» не копіюється). */
export const SEED_KEYS = ["org", "dept", "unit", "appendix", "title", "edrpou", "act_approver_title", "act_approver", "act_owner", "act_order", "act_chair_full", "act_chair", "act_members"];
export const pickSeed = (st) => Object.fromEntries(SEED_KEYS.filter((k) => st && st[k] !== undefined).map((k) => [k, JSON.parse(JSON.stringify(st[k]))]));

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
/** Назва авто для показу: марка + номерний знак (у режимі ЦЗ номер — окреме поле). */
export const vehName = (v) => [String((v && v.name) || "").trim(), String((v && v.plate) || "").trim()].filter(Boolean).join(" ");
/** Водний транспорт (катери): пробіг — у мотогодинах, норма — л на 1 мотогодину. */
/** Техніка (трактор): в акті списання — у кілограмах; густина дизпалива — постійна. */
export const KG_PER_L = 0.852;
export const isMoto = (v) => !!v && v.kind === "moto";
/** літри (у сотих) → кг, до тисячних */
export const kgOf = (qC) => Math.round((qC / 100) * KG_PER_L * 1000) / 1000;
export const fmtKg = (kg) => String(kg).replace(".", ",");
export const isWater = (v) => !!v && (v.kind === "water" || v.kind === "moto");   // також техніка (трактор) — у мотогодинах
export const runUnit = (v) => (isWater(v) ? "мотогод" : "км");
/** Пробіг для показу: км — цілі; мотогодини — до сотих (кома). */
export function fmtRun(v, x) {
  if (isBlank(x)) return "";
  if (!isWater(v)) return fmtKm(x);
  return String(Math.round(Number(x) * 100) / 100).replace(".", ",");
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
  const out = { name: "", rate: 0, month_norm_km: 2640, fuel: "А-95", supplier: "", source: "smk", winter_pct: 0,
    mvo_text: "", active: true };
  for (const [k, val] of Object.entries(v)) if (val !== null && val !== undefined) out[k] = val;
  if (!out.id) out.id = newId();
  if (!out.source) out.source = "smk";
  { const byLabel = Object.fromEntries(LEGACY_SOURCES.map((l) => [l.label, l.id])); if (byLabel[out.source]) out.source = byLabel[out.source]; }
  for (const k of ["rate", "winter_pct"]) { const n = Number(String(out[k] ?? 0).replace(",", ".")); out[k] = isFinite(n) ? n : 0; }
  { const x = Number(String(out.month_norm_km || 0).replace(",", ".")); const n = isWater(out) ? Math.round(x * 10) / 10 : Math.trunc(x); out.month_norm_km = isFinite(n) ? n : 0; }
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
  // спершу джерело з тим самим видом пального (а не «будь-яке»), потім будь-яке, що підходить
  return list.find((s) => s.id === v.source) || list.find((s) => s.fuel && sourceFits(s, v)) || list.find((s) => sourceFits(s, v)) || list[0];
}

/** Форма джерела зі слів опису: «т», «тал.», «талони» → талони; «АЗС» → АЗС; «с/к», «см/к» → с/к. */
export function textForms(low) {
  low = String(low || "").toLowerCase();
  const tal = /талон|(?:^|[\s\d])т(?:ал)?\.?(?=[\s,;]|$)/.test(low);
  if (/субв|благо/.test(low)) return ["Субвенція"];
  if (/з\s*\/\s*ф|загальн/.test(low)) return ["з/ф"];
  if (/(?:^|[^а-яґєіїa-z])ова(?![а-яґєіїa-z])/.test(low)) return ["Налив ОВА"];   // окреме слово «ОВА» (не «Коваленко»)
  if (/азс/.test(low)) return tal ? ["АЗС ГУНП талони", "АЗС налив"] : ["АЗС налив", "АЗС ГУНП талони"];   // «т АЗС» — талони АЗС ГУНП
  if (tal) return ["талони", "АЗС ГУНП талони"];
  // «н», «нал.», «налив» — налив (АЗС або ОВА)
  if (/налив|(?:^|[\s\d])н(?:ал)?\.?(?=[\s,;]|$)/.test(low)) return ["АЗС налив", "Налив ОВА"];
  if (/(?:^|[\s\d])см?\/к/.test(low)) return ["с/к"];
  return null;
}

/** Джерело за назвою (міткою) у тексті опису: повне слово або його початок — «т», «та», «тал», «талон» → «талони»;
 *  мітки з «/» чи пробілом («т/н», «налив/з», «ГУНП налив») — лише повністю. */
export function labelSrc(low, cands) {
  const words = String(low || "").toLowerCase().split(/[\s,;+]+/).map((w) => w.replace(/^[\d.,]+/, "").replace(/\.+$/, "")).filter((w) => /[а-яґєіїa-z]/.test(w));
  let best = null, bs = 0;
  for (const x of cands) {
    const lab = String(x.label || "").toLowerCase().replace(/\s+/g, " ").trim(); if (!lab) continue;
    for (const w of words) {
      const sc = w === lab ? 10 : !/[\/\s]/.test(lab) && lab.startsWith(w) ? 5 + w.length / lab.length : 0;
      if (sc > bs) { bs = sc; best = x; }
    }
  }
  return best;
}

export function calcRow(vehicle, entry, sources = LEGACY_SOURCES) {
  const e = entry || {};
  const list = vehicleSources(vehicle, e, sources);
  // розбивка залишку на початок за МВО: числа (start_mvo) або довільний опис (start_text, «202 А-95 Мосійченко + 167 … Левошко»)
  const hasMap = e.start && Object.values(e.start).some((x) => !isBlank(x));
  // розбивка залишку за джерелами (start: {джерело: л}) — лише стандартний режим; у своїх джерелах старі записи
  // (з часу, коли залишок вводився в плитки джерел) читаються як загальний залишок + опис
  const useStartMap = hasMap && isBlank(e.fuel_start) && sources.legacy;
  const dual = isDual(vehicle) && !sources.legacy;
  const gasStart = dual && !isBlank(e.fuel_start_gas) ? Number(e.fuel_start_gas) || 0 : 0;
  let startTotal = !isBlank(e.fuel_start) ? e.fuel_start : hasMap && !sources.legacy ? Object.values(e.start).reduce((a, x) => a + (Number(x) || 0), 0) : null;
  if (gasStart) startTotal = (Number(startTotal) || 0) + gasStart;
  const splitIn = [];          // [{mvo, c, hint}]
  const textWarn = [];
  if (!useStartMap) {
    const nums = Object.entries(e.start_mvo || {}).filter(([, x]) => !isBlank(x));
    if (nums.length) nums.forEach(([mv, x]) => splitIn.push({ mvo: mv, c: toC(x), hint: "" }));
    else if (!sources.legacy && String(e.start_text || "").trim()) {
      const mvos = [...new Set(sources.map((x) => String(x.mvo || "").trim()).filter(Boolean))];
      // МВО пишуть у кінці групи: «185 т А-95 + 270 с/к зал. Мосійченко» — обидві частини Мосійченка;
      // частина без прізвища бере МВО найближчої наступної частини з прізвищем (якщо такої немає — попередньої)
      const parts = String(e.start_text).split("+").map((part) => {
        const m = part.match(/(?:^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])(\d+(?:[.,]\d+)?)/);   // число, але не «А-95»/«А95»
        const low = part.toLowerCase();
        const mv = mvos.find((x) => low.includes(x.toLowerCase().split(/\s+/)[0])) || "";
        // своє прізвище, якого немає серед МВО джерел (напр. «10 н Петрунь»), — частина не бере МВО сусідньої частини
        const alien = !mv && hasAlienName(part);
        return { part: part.trim(), low, c: m ? toC(m[1].replace(",", ".")) : null, mv, own: false, alien };
      });
      parts.forEach((p) => { p.own = !!p.mv; });
      // опис без числа («ДП с/к Ольховський») — увесь загальний залишок належить цьому МВО / формі
      const withNum = parts.filter((p) => p.c !== null);
      if (!withNum.length && parts.length === 1 && !isBlank(startTotal)) parts[0].c = toC(startTotal) - toC(gasStart);
      for (let k = parts.length - 1, nxt = ""; k >= 0; k--) { if (parts[k].mv) nxt = parts[k].mv; else if (!parts[k].alien) parts[k].mv = nxt; }
      for (let k = 0, prv = ""; k < parts.length; k++) { if (parts[k].mv) prv = parts[k].mv; else if (!parts[k].alien) parts[k].mv = prv; }
      for (const p of parts) {
        if (p.c === null) continue;
        const form = textForms(p.low);
        // назва джерела в тексті («7 т», «5 т/н», «10 налив/з») — прямо на це джерело
        const lsrc = labelSrc(p.low, sources.filter((x) => x.active && (!p.own || x.mvo === p.mv)));
        if (lsrc) { splitIn.push({ mvo: lsrc.mvo, c: p.c, hint: p.low, text: true, src: lsrc }); continue; }
        // давні талони — окрема частина залишку: не прив'язується до джерела, переноситься як написано
        if (form && form[0] === "талони" && !(p.own && sources.some((x) => x.mvo === p.mv && form.includes(x.form) && sourceFits(x, vehicle)))) {
          splitIn.push({ keep: true, c: p.c, text: true, label: p.part, mvo: p.mv, own: p.own }); continue;
        }
        if (!p.mv) {
          const byForm = form && sources.filter((x) => form.includes(x.form))
            .sort((a, b) => (sourceFits(b, vehicle) - sourceFits(a, vehicle)) || (form.indexOf(a.form) - form.indexOf(b.form)) || (list.includes(b) - list.includes(a)))[0];
          if (byForm) { splitIn.push({ mvo: byForm.mvo, c: p.c, hint: p.low, text: true, src: byForm }); continue; }
          textWarn.push(`У описі залишку «${p.part}» не знайдено МВО чи форму (талони, с/к, АЗС) з джерел — ця частина віднесена на основне джерело.`); continue;
        }
        splitIn.push({ mvo: p.mv, c: p.c, hint: p.low, text: true });
      }
    }
  }
  // газ (авто «Газ/бенз»): окреме поле залишку газу — на джерело газу (з якого видавали цього місяця, інакше будь-яке газове)
  if (gasStart > 0) {
    const gi = (x) => { const m = e.iss || {}; return !isBlank(m[x.id]) && Number(m[x.id]) > 0 ? 1 : 0; };
    const gsrc = [...list, ...sources.filter((x) => !list.includes(x) && x.active)].filter(isGasSrc).sort((a, b) => gi(b) - gi(a) || (list.includes(b) - list.includes(a)))[0];
    if (gsrc) { if (!list.includes(gsrc)) list.push(gsrc); splitIn.push({ mvo: gsrc.mvo, c: toC(gasStart), src: gsrc, gas: true }); }
    else splitIn.push({ keep: true, c: toC(gasStart), label: `${fmtC(toC(gasStart))} газ`, mvo: "", own: true, gas: true });
  }
  // джерело для частини залишку: та ж МВО; далі — вид пального, постачальник і форма зі слів опису
  const pickSrc = (sp) => {
    const cand = [...list, ...sources.filter((x) => !list.includes(x))].filter((x) => x.mvo === sp.mvo);
    const score = (x) => (sourceFits(x, vehicle) ? 4 : 0) + (list.includes(x) ? 1 : 0) + (x.id === vehicle.source ? 1 : 0) +
      (sp.hint && x.supplier && sp.hint.includes(x.supplier.toLowerCase()) ? 3 : 0) + (sp.hint && (textForms(sp.hint) || []).includes(x.form) ? 3 : 0) +
      (sp.hint && x.fuel && normFuel(sp.hint).includes(normFuel(x.fuel)) ? 2 : 0);
    return cand.sort((a, b) => score(b) - score(a))[0];
  };
  // джерело залишку на початок, вибране вручну («з якого джерела залишок»)
  const pickedStart = e.start_src ? sources.find((x) => x.id === e.start_src) : null;
  if (pickedStart && !list.includes(pickedStart)) list.push(pickedStart);
  for (const sp of splitIn) { sp.src = sp.src || (sp.c ? pickSrc(sp) : null); if (sp.src && !list.includes(sp.src)) list.push(sp.src); }
  // свої джерела: «різні» за замовчуванням — також коли вибране основне джерело вже не існує
  const free = !sources.legacy && (vehicle.source === NO_MAIN || !sources.some((x) => x.id === vehicle.source));
  // куди йде залишок, не описаний за МВО: основне джерело; для «різних» — джерело того ж пального, з якого видавали цього місяця,
  // і не АЗС (талони/с/к), щоб не «списувати» з АЗС ГУНП пальне, якого там не було
  const issNow = (x) => { const m = e.iss || {}; return !isBlank(m[x.id]) && Number(m[x.id]) > 0; };
  const anchorScore = (x) => (x.fuel && sourceFits(x, vehicle) ? 4 : 0) + (issNow(x) ? 2 : 0) + (FORM_GROUP[x.form] === "tsmk" ? 1 : 0) + (dual && !isGasSrc(x) ? 3 : 0);
  const anchor = free ? ([...list].filter((x) => sourceFits(x, vehicle)).sort((a, b) => anchorScore(b) - anchorScore(a))[0] || list[0])
    : mainSource(vehicle, list);
  const main = free ? null : anchor;
  const r = { vehicle, entry: e, sources: list, mainId: main ? main.id : null, km: 0, normC: 0, dual,
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
    const T = toC(startTotal);
    let sum = 0;
    for (const sp of splitIn) {
      sum += sp.c; if (!sp.c) continue;
      if (sp.keep) { r.keep.push({ label: sp.label, c: sp.c, left: sp.c, mvo: sp.mvo || "", own: sp.own, ...(sp.gas ? { gas: true } : {}) }); continue; }
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
    const startTo = pickedStart || anchor;
    if (startTo && T > sum) r.startBy[startTo.id] += T - sum;
    else if (!startTo && T > sum) {
      // немає жодного джерела цього пального — залишок не губимо, а ведемо як є (вид пального авто)
      r.keep.push({ label: `${fmtC(T - sum)} ${vehicle.fuel || ""}`.trim(), c: T - sum, left: T - sum, mvo: "", own: true });
      r.warnings.push(`Немає джерела пального ${vehicle.fuel || ""} для залишку на початок — опишіть його в полі «Вид залишку на початок», напр. «${fmtC(T - sum)} ${vehicle.fuel || ""} т».`);
    }
    if (anchor && free && !pickedStart && T > 0 && sum === 0 && vehicleMvos(sources).length > 1)
      r.warnings.push(`Залишок на початок не описано (чий і який) — віднесено на «${sourceTitle(anchor)}». Опишіть його в полі «Вид залишку на початок», напр. «${fmtC(T)} с/к Прізвище».`);
    r.startSplit = splitIn.some((x) => x.c > 0) && sum > 0;
  }
  for (const s of list) r.startC += r.startBy[s.id];
  for (const k of r.keep) r.startC += k.c;
  // «передано з ТЗ на ТЗ» (режим ЦЗ): додається до виданого (у т.ч. від'ємне — передано з цього авто);
  // у розподілі витрати — як частина залишку (на джерело, де залишок, або на основне)
  r.xferC = toC(e.xfer) || 0;
  if (r.xferC) {
    const tgt = [...list].sort((a, b) => r.startBy[b.id] - r.startBy[a.id])[0] || anchor;
    if (tgt) { r.startBy[tgt.id] += r.xferC; r.xferTo = tgt.id; } else r.keep.push({ label: "передано з ТЗ на ТЗ", c: r.xferC, left: r.xferC, mvo: "", own: true });
    r.issuedC += r.xferC;
  }
  r.empty = [e.odo_start, e.odo_end, e.consumed_manual].every((x) => isBlank(x) || Number(x) === 0) && r.startC === 0 && r.issuedC === 0;
  // --- пробіг і витрата за нормою
  const s0 = e.odo_start, s1 = e.odo_end;
  if (!isBlank(s0) && !isBlank(s1)) {
    if (Number(s1) < Number(s0)) r.errors.push("Пробіг на кінець менший за пробіг на початок — перевірте показники одометра.");
    else r.km = isWater(vehicle) ? Math.round((Number(s1) - Number(s0)) * 100) / 100 : Math.trunc(Number(s1)) - Math.trunc(Number(s0));
  } else if (isBlank(s0) !== isBlank(s1)) {
    r.warnings.push("Не заповнено пробіг на " + (isBlank(s1) ? "кінець" : "початок") + " періоду.");
  }
  const rateC = toC(vehicle.rate);                 // норма ×100
  const wC = 0;                                    // зимової надбавки немає (прибрано з порталу)
  r.normC = isWater(vehicle) ? divHalfUp(Math.round(r.km * 100) * rateC, 100)      // мотогодини × л/мотогод
    : divHalfUp(r.km * rateC * (10000 + wC), 1000000);
  const avail = r.startC + r.issuedC;
  // загальне правило: витрата за нормою округлюється до цілих літрів (половина — вгору);
  // десяті/соті лишаються лише в залишку (або коли списується весь наявний залишок із сотими)
  r.normExactC = r.normC;
  // соті лишаються, якщо вони були в залишку на початок або у виданому за місяць; інакше — до цілих літрів
  r.hasFrac = r.startC % 100 !== 0 || list.some((x) => r.issued[x.id] % 100 !== 0);
  if (!r.hasFrac) r.normC = divHalfUp(r.normC, 100) * 100;
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
    // стандартний режим — як раніше: витрата в колонку основного джерела, залишок одним числом;
    // якщо залишок на початок вказано в іншому джерелі — спершу витрата з нього (у його колонку)
    let left = r.consC;
    for (const s of list) r.endBy[s.id] = 0;
    for (const s of list) {
      if (main && s.id === main.id) continue;
      const q = Math.min(left, Math.max(r.startBy[s.id], 0)); if (!q && !r.startBy[s.id]) continue;
      r.consBy[s.id] = q; left -= q; r.endBy[s.id] = r.startBy[s.id] - q;
    }
    if (main) { r.consBy[main.id] = left; r.endBy[main.id] = avail - r.consC - list.reduce((a, s) => a + (s.id !== main.id ? r.endBy[s.id] : 0), 0); }
    r.endC = avail - r.consC;
  } else if (e.split && e.cons) {
    let sum = 0;
    for (const s of list) {
      const c = toC(e.cons[s.id]); r.consBy[s.id] = c; sum += c;
      if (c > availBy(s.id)) r.errors.push(`З «${sourceTitle(s)}» списано ${fmtC(c)} л, а там є лише ${fmtC(availBy(s.id))} л.`);
    }
    if (sum !== r.consC) r.errors.push(`Розподіл витрати за джерелами (${fmtC(sum)} л) не дорівнює загальній витраті (${fmtC(r.consC)} л).`);
  } else {
    // спершу залишок на початок, потім видане за місяць (або навпаки — налаштування «спочатку видане»);
    // у кожній черзі — так, щоб джерела вичерпувались повністю: якщо одне джерело покриває решту — лише з нього,
    // інакше спершу повністю найбільше. Основне джерело (якщо вибране) — першим; джерела іншого пального — в останню чергу.
    let left = r.consC;
    const used = { start: {}, iss: {} };
    const remOf = (x, kind) => Math.max((kind === "start" ? r.startBy[x.id] : r.issued[x.id]) - (used[kind][x.id] || 0), 0);
    const take = (x, kind, q) => { used[kind][x.id] = (used[kind][x.id] || 0) + q; r.consBy[x.id] += q; left -= q; };
    const greedy = (pool, kind) => {
      while (left > 0) {
        const c = pool.filter((x) => remOf(x, kind) > 0); if (!c.length) return;
        const enough = c.filter((x) => remOf(x, kind) >= left).sort((a, b) => remOf(a, kind) - remOf(b, kind));
        if (enough.length) { take(enough[0], kind, left); return; }
        const big = c.sort((a, b) => remOf(b, kind) - remOf(a, kind))[0]; take(big, kind, remOf(big, kind));
      }
    };
    const rest = list.filter((x) => !main || x.id !== main.id);
    const pools = [...(main ? [[main]] : []), rest.filter((x) => sourceFits(x, vehicle)), rest.filter((x) => !sourceFits(x, vehicle))];
    // правило підрозділу («спочатку видане») можна змінити для окремого авто: iss_first = "1" / "0"
    const ov = e.iss_first || vehicle.iss_first;
    const issFirst = ov === "1" ? true : ov === "0" ? false : sources.issuedFirst;
    const phases = issFirst ? ["iss", "start"] : ["start", "iss"];
    // дрібні видачі (до 10 л) списуються першочергово й повністю — щоб не лишати на залишку «0,18 л»
    const smallFirst = (pool) => {
      for (const x of pool.filter((y) => remOf(y, "iss") > 0 && remOf(y, "iss") <= SMALL_ISSUE_C).sort((a, b) => remOf(a, "iss") - remOf(b, "iss"))) {
        if (left <= 0) return; take(x, "iss", Math.min(left, remOf(x, "iss")));
      }
    };
    for (const pool of pools) { smallFirst(pool); for (const kind of phases) greedy(pool, kind); }
    // джерел не вистачило — списуємо з давніх талонів (їх кількість у описі зменшується)
    for (const k of r.keep) { if (left <= 0) break; const take = Math.min(left, k.left); k.left -= take; left -= take; r.keepConsC = (r.keepConsC || 0) + take; }
    if (left > 0 && anchor) r.consBy[anchor.id] += left;   // не має статися (витрата обмежена наявним)
    // залишок на кінець, внесений вручну з кількостями («155 А-95 с/к Іваненко + 13 т Завора»): розподіл витрати відштовхується від нього —
    // що не вказано в залишку, вважається списаним (першочергово — залишки на початок і дрібні видачі)
    if (!sources.legacy && String(e.end_text || "").trim()) applyManualEnd(r, e, list, sources, vehicle, availBy, pickSrc);
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

/** У частині опису є прізвище (слово з великої літери, не службове слово). */
const NOT_NAMES = /^(зал|залишок|залишку|налив|нал|талон|талони|резерв|бенз|бензин|дизель|газ|азс|ова|укрнафта|окко|субвенція|благо|см|ск)$/;
function hasAlienName(part) {
  return (String(part).match(/[А-ЯҐЄІЇ][а-яґєії'’ʼ]{2,}/g) || []).some((w) => !NOT_NAMES.test(w.toLowerCase()));
}
/** Розбір опису залишку («202 А-95 с/к Мосійченко + 185 т А-95») на частини: {c, mvo, own, low, part}. МВО без прізвища — від наступної частини. */
function parseBalanceParts(text, sources) {
  const mvos = [...new Set(sources.map((x) => String(x.mvo || "").trim()).filter(Boolean))];
  const parts = String(text || "").split("+").map((part) => {
    const m = part.match(/(?:^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])(\d+(?:[.,]\d+)?)/);
    const low = part.toLowerCase();
    const mvo = mvos.find((x) => low.includes(x.toLowerCase().split(/\s+/)[0])) || "";
    return { part: part.trim(), low, c: m ? toC(m[1].replace(",", ".")) : null, mvo, own: false, alien: !mvo && hasAlienName(part) };
  });
  parts.forEach((p) => { p.own = !!p.mvo; });
  for (let k = parts.length - 1, nxt = ""; k >= 0; k--) { if (parts[k].mvo) nxt = parts[k].mvo; else if (!parts[k].alien) parts[k].mvo = nxt; }
  for (let k = 0, prv = ""; k < parts.length; k++) { if (parts[k].mvo) prv = parts[k].mvo; else if (!parts[k].alien) parts[k].mvo = prv; }
  return parts;
}
/** Витрата за джерелами — так, щоб залишок на кінець збігся з внесеним вручну описом. */
function applyManualEnd(r, e, list, sources, vehicle, availBy, pickSrc) {
  const parts = parseBalanceParts(e.end_text, sources).filter((p) => p.c !== null);
  if (!parts.length) return;                                   // опис без кількостей — лише текст
  const endT = {}; const keepT = []; const unknown = [];
  for (const p of parts) {
    const form = textForms(p.low);
    const lsrc = labelSrc(p.low, list.filter((x) => !p.own || x.mvo === p.mvo));
    if (lsrc) { endT[lsrc.id] = (endT[lsrc.id] || 0) + p.c; continue; }
    if (form && form[0] === "талони" && r.keep.length && !(p.own && sources.some((x) => x.mvo === p.mvo && form.includes(x.form)))) { keepT.push(p.c); continue; }
    const src = p.mvo ? pickSrc({ mvo: p.mvo, hint: p.low }) : (form ? list.find((x) => form.includes(x.form) && sourceFits(x, vehicle)) : null);
    if (!src || !list.includes(src)) { unknown.push(p.part); continue; }
    endT[src.id] = (endT[src.id] || 0) + p.c;
  }
  // залишок за описом не може бути більшим за наявне в джерелі
  let consSum = 0;
  for (const s of list) { const av = availBy(s.id); const end = Math.min(Math.max(endT[s.id] || 0, 0), Math.max(av, 0)); r.consBy[s.id] = av - end; consSum += r.consBy[s.id]; }
  r.keepConsC = 0;
  r.keep.forEach((k, i) => { const end = Math.min(keepT[i] ?? 0, k.c); k.left = end; r.keepConsC += k.c - end; consSum += k.c - end; });
  const diff = r.consC - consSum;
  if (unknown.length) r.warnings.push(`У описі залишку на кінець не розпізнано джерело: «${unknown.join("», «")}».`);
  if (diff !== 0) {
    r.warnings.push(`Залишок на кінець за описом не збігається з розрахунком на ${fmtC(Math.abs(diff))} л — різницю ${diff > 0 ? "списано з джерел, де лишався залишок" : "повернено до залишку"} автоматично. Перевірте опис.`);
    let left = Math.abs(diff);
    const order = [...list].sort((a, b) => (diff > 0 ? (availBy(b.id) - r.consBy[b.id]) - (availBy(a.id) - r.consBy[a.id]) : r.consBy[b.id] - r.consBy[a.id]));
    for (const s of order) {
      if (left <= 0) break;
      const room = diff > 0 ? availBy(s.id) - r.consBy[s.id] : r.consBy[s.id];
      const q = Math.min(left, Math.max(room, 0)); r.consBy[s.id] += diff > 0 ? q : -q; left -= q;
    }
  }
  r.endFromText = true;
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

/** Стандартний режим: вид залишку на початок — пальне авто + джерело, якщо його вказано вручну. */
export function legacyStartVid(r) {
  if (r.frozenTxt) return r.frozenTxt.legacyStart;
  const id = (r.entry || {}).start_src;
  const src = id && r.sources.find((x) => x.id === id);
  return [r.vehicle.fuel || "", src && id !== r.mainId ? sourceTitle(src) : ""].filter(Boolean).join(" ");
}

/** «Вид пального на залишку»: «125 А-95 с/к UKRNAFTA Мосійченко + 167 …». */
export function balanceText(r, which = "end", settings = {}) {
  if (r.frozenTxt) return r.frozenTxt[which];        // закритий місяць — текст як у день закриття
  const by = which === "start" ? r.startBy : r.endBy;
  const own = String(((r.entry || {})[which === "start" ? "start_text" : "end_text"]) || "").trim();
  if (own) return own;                                // довільний опис, введений вручну
  if (which === "start" && !r.startSplit) {           // загальний залишок — лише вид пального
    const m = r.sources.find((s) => s.id === r.mainId);
    return m ? sourceFor(m, r.vehicle, settings).fuel : (r.vehicle.fuel || "");
  }
  // залишок: кількість, вид пального, форма (с/к, т, АЗС) і МВО — без постачальника; однакові — разом
  const items = [];
  for (const s of r.sources) {
    if (!(by[s.id] > 0)) continue;
    const f = sourceFor(s, r.vehicle, settings);
    const lab = [f.fuel, FORM_SHORT[f.form] || f.form, f.mvo].filter(Boolean).join(" ");
    const same = items.find((x) => x.lab === lab);
    if (same) same.c += by[s.id]; else items.push({ mvo: f.mvo, lab, c: by[s.id] });
  }
  items.forEach((x) => { x.t = `${fmtC(x.c)} ${x.lab}`; });
  // давні талони — як написано вручну (кількість — що лишилось), поруч із пальним свого МВО
  for (const k of r.keep || []) {
    const q = which === "start" ? k.c : k.left; if (q <= 0) continue;
    let t = q === k.c ? k.label : k.label.replace(/(^|[^\d\-А-ЯҐЄІЇа-яґєіїA-Za-z])(\d+(?:[.,]\d+)?)/, (_, p1) => p1 + fmtC(q));
    if (k.mvo && !k.own) t += " " + k.mvo;
    let at = -1; items.forEach((x, n) => { if (k.mvo && x.mvo === k.mvo) at = n; });
    items.splice(at >= 0 ? at + 1 : items.length, 0, { mvo: k.mvo, t });
  }
  const parts = items.map((x) => x.t);
  if (parts.length) return parts.join(" + ");
  const m = r.sources.find((s) => s.id === r.mainId);
  return m ? sourceFor(m, r.vehicle, settings).fuel : (r.vehicle.fuel || "");
}

/** Авто «Газ/бенз»: рядок розрахунку, розділений на частини «бензин» і «газ» (для двох рядків реєстру).
 *  Кожна частина — той самий вигляд, що й рядок calcRow (лише джерела свого пального). */
export function dualParts(r) {
  if (!r || !r.dual) return null;
  const mk = (gas) => {
    const srcs = r.sources.filter((x) => isGasSrc(x) === gas);
    const pick = (o) => Object.fromEntries(srcs.map((x) => [x.id, (o || {})[x.id] || 0]));
    const sum = (o) => srcs.reduce((a, x) => a + ((o || {})[x.id] || 0), 0);
    const keep = (r.keep || []).filter((k) => (!!k.gas || /газ/i.test(k.label)) === gas);
    const kS = keep.reduce((a, k) => a + k.c, 0), kE = keep.reduce((a, k) => a + Math.max(k.left, 0), 0);
    const xfer = r.xferTo && srcs.some((x) => x.id === r.xferTo) ? r.xferC : 0;
    const startBy = pick(r.startBy); if (xfer) startBy[r.xferTo] -= xfer;
    const p = { ...r, frozenTxt: null, gasPart: gas, sources: srcs, startBy, issued: pick(r.issued), consBy: pick(r.consBy), endBy: pick(r.endBy), keep,
      startC: sum(r.startBy) - xfer + kS, issuedC: sum(r.issued) + xfer, consC: sum(r.consBy) + keep.reduce((a, k) => a + (k.c - Math.max(k.left, 0)), 0), endC: sum(r.endBy) + kE,
      mainId: srcs.some((x) => x.id === r.mainId) ? r.mainId : null, startSplit: true,
      vehicle: { ...r.vehicle, fuel: gas ? "газ" : "А-95" },
      entry: { ...r.entry, start_text: gas ? "" : r.entry.start_text, end_text: gas ? "" : r.entry.end_text } };
    return p;
  };
  return [mk(false), mk(true)];
}

export function totals(rows) {
  const t = { km: 0, startC: 0, issued: {}, issuedC: 0, consBy: {}, consC: 0, normC: 0, endC: 0 };
  for (const r of rows) {
    if (isWater(r.vehicle)) t.hours = (t.hours || 0) + r.km; else t.km += r.km;
    t.startC += r.startC; t.issuedC += r.issuedC; t.consC += r.consC; t.normC += r.normC; t.endC += r.endC;
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
  const norm = (v) => vehName(v).toLowerCase().replace(/\s+/g, " ").trim();
  for (const [vid, e] of Object.entries(entries)) {
    if (!seen.has(vid) && !(vid in byId) && vid in snap) {
      const v = normalizeVehicle(snap[vid]);
      if (calcRow(v, e, sources).empty) continue;
      // авто видалили з довідника й створили знову (та сама назва/номер) — не дублюємо картку:
      // дані місяця переходять до нового запису, якщо в нього ще нічого не внесено
      const twin = vehicles.find((x) => norm(x) && norm(x) === norm(v));
      if (twin) {
        if (!(twin.id in entries) || calcRow(twin, entries[twin.id], sources).empty) entries[twin.id] = e;
        if (!seen.has(twin.id)) { out.push(twin); seen.add(twin.id); }
        continue;
      }
      out.push(v);
    }
  }
  return out;
}

/** Початкові дані нового місяця з кінця попереднього (залишок — окремо за кожним джерелом). */
export function carryOver(vehicles, prevMonth, sources = LEGACY_SOURCES, prevSources = sources) {
  if (!prevMonth) return {};
  const out = {};
  // закритий місяць — переносимо з його знімка (результати як у день закриття)
  const fr = prevMonth.frozen && Array.isArray(prevMonth.frozen.rows) ? prevMonth.frozen.rows : null;
  for (const v of fr ? fr.map((x) => x.vehicle) : monthVehicles(vehicles, prevMonth, prevSources)) {
    const pe = fr ? (fr.find((x) => x.vehicle.id === v.id) || {}).entry || {} : prevMonth.entries[v.id];
    if (!pe) continue;
    const r = fr ? fr.find((x) => x.vehicle.id === v.id) : calcRow(v, pe, prevSources);
    const odo = isBlank(pe.odo_end) ? null : pe.odo_end;
    if (sources.legacy) {
      const start = {};
      for (const [id, c] of Object.entries(r.endBy)) if (c !== 0) start[id] = c / 100;
      out[v.id] = { odo_start: odo, start, iss: {} };
    } else {
      // загальний залишок + розбивка за МВО, якщо на кінець пальне було в кількох МВО
      // опис залишку на кінець (вид пального, форма — с/к, т, АЗС — і МВО) переходить у опис залишку на початок,
      // щоб у новому місяці не загубилась форма (напр. «76 ДП т» не ставало «с/к»)
      let endText = String(pe.end_text || "").trim();
      if (!endText && r.endC > 0) {
        const t = String(balanceText(r, "end") || "").trim();
        if (t && t.toLowerCase() !== String(v.fuel || "").trim().toLowerCase()) endText = t;
      }
      const dp = dualParts(r);
      if (dp) {
        // «Газ/бенз»: залишок бензину й газу — окремими полями; опис — лише бензину
        const [b, g] = dp; let bt = String(pe.end_text || "").trim();
        if (!bt && b.endC > 0) { const t = String(balanceText(b, "end") || "").trim(); if (t && !/^а-?95$/i.test(t)) bt = t; }
        out[v.id] = { odo_start: odo, fuel_start: b.endC ? b.endC / 100 : null, ...(g.endC ? { fuel_start_gas: g.endC / 100 } : {}), start: {}, iss: {},
          ...(bt ? { start_text: bt } : {}) };
      } else
      out[v.id] = { odo_start: odo, fuel_start: r.endC ? r.endC / 100 : null, start: {}, iss: {},
        ...(endText ? { start_text: endText } : {}) };
    }
    if (pe.iss_first) out[v.id].iss_first = pe.iss_first;      // порядок списання для авто — переходить на новий місяць
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
  for (const f of ["odo_start", "odo_end"]) { const x = num2(raw[f]); out[f] = x; }    // км — цілі (рахуються цілими), мотогодини — до сотих
  out.consumed_manual = num2(raw.consumed_manual);
  if ("fuel_start" in raw) out.fuel_start = num2(raw.fuel_start);
  if (!isBlank(raw.fuel_start_gas)) { const g = num2(raw.fuel_start_gas); if (g) out.fuel_start_gas = g; }      // авто «Газ/бенз»: залишок газу
  if (raw.start_src) out.start_src = String(raw.start_src);
  if (raw.iss_first === "1" || raw.iss_first === "0") out.iss_first = raw.iss_first;   // порядок списання для цього авто (інакше, ніж у підрозділі)
  if (!isBlank(raw.xfer)) { const x = Number(String(raw.xfer).replace(",", ".").replace(/\s/g, "")); if (isFinite(x) && x !== 0) out.xfer = Math.round(x * 100) / 100; }      // «передано з ТЗ на ТЗ» (режим ЦЗ), може бути від'ємним
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
        // усередині акта — окремими пунктами за формою (с/к / талони; АЗС талони / АЗС налив), пропорційно до виданого
        const inG = r.sources.filter((s) => (FORM_GROUP[s.form] || "tsmk") === g && (r.issued[s.id] || 0) > 0);
        const tot = inG.reduce((a, s) => a + r.issued[s.id], 0);
        const parts = tot > 0 ? inG.map((s) => [s.form, q % 100 === 0 ? divHalfUp(q * r.issued[s.id], tot * 100) * 100 : divHalfUp(q * r.issued[s.id], tot)]) : [[(main && FORM_GROUP[main.form] === g ? main.form : ""), q]];
        const diff = q - parts.reduce((a, x) => a + x[1], 0); if (parts.length) parts[parts.length - 1][1] += diff;
        for (const [form, qq] of parts) {
          if (qq <= 0) continue;
          const k = JSON.stringify([f.fuel, f.supplier, f.mvo, form]);
          m.set(k + "|" + g, { k, q: ((m.get(k + "|" + g) || {}).q || 0) + qq });
        }
      }
      continue;
    }
    for (const s of r.sources) {
      const q = r.consBy[s.id] || 0;
      if (q <= 0) continue;
      const f = sourceFor(s, r.vehicle, settings);
      if (group !== null && f.group !== group) continue;
      if (mvo !== null && f.mvo !== mvo) continue;
      // окремі пункти акта за видом пального І формою (с/к окремо від талонів); техніка (трактор) — окремим пунктом у кг
      const k = JSON.stringify([f.fuel, f.supplier, f.mvo, f.form, ...(isMoto(r.vehicle) ? [vehName(r.vehicle)] : [])]);
      m.set(k + "|" + f.group, { k, q: ((m.get(k + "|" + f.group) || {}).q || 0) + q });
    }
  }
  const agg = new Map();
  for (const { k, q } of m.values()) agg.set(k, (agg.get(k) || 0) + q);
  const FORM_ORDER = FORMS;
  return [...agg.entries()].map(([k, q]) => { const [fuel, supplier, mvo2, form, moto] = JSON.parse(k);
    return { fuel, supplier, mvo: mvo2, form: form || "", qtyC: q, ...(moto ? { moto, kg: kgOf(q) } : {}) }; })
    .sort((a, b) => String(a.fuel).localeCompare(String(b.fuel), "uk") || FORM_ORDER.indexOf(a.form) - FORM_ORDER.indexOf(b.form) || String(a.mvo).localeCompare(String(b.mvo), "uk"));
}
function legacyActSplit(r) {
  const out = { tsmk: 0, azs: 0 };
  if (r.consC <= 0) return out;
  const by = { tsmk: 0, azs: 0 };
  for (const s of r.sources) by[FORM_GROUP[s.form] || "tsmk"] += r.issued[s.id] || 0;
  const tot = by.tsmk + by.azs;
  if (tot <= 0) { const main = r.sources.find((s) => s.id === r.mainId); out[(main && FORM_GROUP[main.form]) || "tsmk"] = r.consC; return out; }
  // цілими літрами, якщо витрата ціла
  const a = r.consC % 100 === 0 ? divHalfUp(r.consC * by.tsmk, tot * 100) * 100 : divHalfUp(r.consC * by.tsmk, tot);
  out.tsmk = a; out.azs = r.consC - a;
  return out;
}
/** Перелік актів місяця: група (талони+с/к / АЗС) × МВО — на кожну МВО окремий акт (підписує вся комісія і ця МВО). */
export function actList(rows, settings = {}) {
  const out = [];
  // субвенція / благо — окремий акт на кожне таке джерело (назва — від джерела)
  const sv = new Map();
  for (const r of rows) if (!r.legacy) for (const s of r.sources || []) {
    const g = sourceFor(s, r.vehicle, settings).group;
    if (g.startsWith("sv:") && !sv.has(g)) sv.set(g, sourceTitle(s));
  }
  const groups = settings.profile === "cz" ? Object.entries(CZ_ACTS).map(([k, t]) => ["cz:" + k, t, t.replace(/[^\p{L}\d]+/gu, "_")]) : [...Object.keys(ACT_GROUPS).map((g) => [g, ACT_GROUPS[g].title, ACT_GROUPS[g].file]),
    ...[...sv.entries()].map(([g, t]) => [g, `Субвенція: ${t}`, "Субвенція_" + t.replace(/[^\p{L}\d]+/gu, "_")])];
  for (const [g, title, file] of groups) {
    const lines = actGroups(rows, g, settings);
    if ((g.startsWith("sv:") || g.startsWith("cz:") || g === "zf") && !lines.length) continue;
    const mvos = [...new Set(lines.map((l) => String(l.mvo || "").trim()))];
    if (mvos.length <= 1) { out.push({ group: g, mvo: mvos[0] || null, mvos: mvos.filter(Boolean), lines, title, file, subv: g.startsWith("sv:") }); continue; }
    for (const mv of mvos) out.push({ group: g, mvo: mv, mvos: [mv].filter(Boolean), lines: lines.filter((l) => String(l.mvo || "").trim() === mv),
      title: `${title} · МВО ${mv || "—"}`, file, perMvo: true, subv: g.startsWith("sv:") });
  }
  return out;
}

export function shortName(full) {
  const p = String(full || "").trim().split(/\s+/);
  if (p.length >= 3) return `${p[0]} ${p[1][0]}.${p[2][0]}.`;
  if (p.length === 2) return `${p[0]} ${p[1][0]}.`;
  return full || "";
}

/** Джерела для місяця: поточні джерела підрозділу + збережені в місяці (щоб старі дані не зникли). */
export function monthSources(settings = {}, stored = null, asSaved = false) {
  // asSaved: джерела збереженого місяця як були тоді (для перенесення залишку — щоб зміни в автопарку
  // не змінювали форму/пальне залишку минулого місяця)
  if (asSaved && Array.isArray(stored) && stored.length) {
    const out = stored.map(normalizeSource), have = new Set(out.map((s) => s.id));
    for (const s of unitSources(settings)) if (!have.has(s.id)) out.push({ ...s, active: false });
    // стандартні джерела місяця — лише незмінені стандартні (без пального/МВО), інакше — свої
    out.legacy = out.every((s) => { const l = LEGACY_SOURCES.find((x) => x.id === s.id); return l && !s.fuel && !s.mvo && !s.supplier && s.form === l.form; })
      && (!Array.isArray(settings.sources) || !settings.sources.length || stored.length === LEGACY_SOURCES.length);
    out.issuedFirst = !out.legacy && !!settings.cons_issued_first;
    return out;
  }
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

// ---------------------------------------------------------------- закриття місяця
/** Знімок закритого місяця: джерела, налаштування, готові рядки (з текстами залишків) і акти — щоб пізніші зміни їх не чіпали. */
export function freezeMonth(rows, settings, sources) {
  const clean = (x) => JSON.parse(JSON.stringify(x));
  const fr = rows.map((r) => ({ ...clean(r), frozenTxt: { start: balanceText(r, "start", settings), end: balanceText(r, "end", settings), legacyStart: legacyStartVid(r) } }));
  return { settings: clean(settings), sources: clean(sources), flags: { legacy: !!sources.legacy, issuedFirst: !!sources.issuedFirst },
    rows: fr, acts: clean(actList(rows, settings)) };
}
/** Рядки закритого місяця з його знімка (з прапорцями джерел). */
export function frozenRows(frozen) {
  const sources = frozen.sources.map((x) => ({ ...x }));
  sources.legacy = !!frozen.flags.legacy; sources.issuedFirst = !!frozen.flags.issuedFirst;
  const rows = frozen.rows.map((r) => ({ ...r, sources: r.sources.map((x) => sources.find((y) => y.id === x.id) || x) }));
  rows.sources = sources;
  return rows;
}
