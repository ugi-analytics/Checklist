// Вивантаження автопарку всіх підрозділів у форматі щомісячного зрізу, який читає модуль хабу
// «Пробіг і пальне автопарку» (fleet_mileage/core.py): 12 колонок, шапка у 2 рядки,
// реквізити авто — лише в першому рядку авто, далі рядки з історією показників «Пробіг» / «Пробіг станом на».
// Дата зрізу береться модулем з назви файлу: «…на_01.11.2026.xlsx».
import { parseYm, ymKey } from "./core.js?v=20261029g";

export const FLEET_HEADERS = ["№ п/п", "Код регіона Д.Н.З.", "Державний номерний знак", "Марка", "Модель", "Рік випуску",
  "Вид палива", "Норми витрати Л/100 км за містом", "Орган", "Служба/Підрозділ", "Пробіг", "Пробіг станом на"];
export const FLEET_WIDTHS = [7, 10, 14, 16, 18, 9, 16, 14, 30, 30, 12, 14];

// «Peugeot 301 1825» → марка Peugeot, модель 301, номер 1825 (номер — останнє слово з цифрою)
export function splitVehicleName(name) {
  const p = String(name || "").trim().split(/\s+/).filter(Boolean);
  let plate = "";
  if (p.length > 1 && /\d/.test(p[p.length - 1])) plate = p.pop();
  return { brand: p[0] || "", model: p.slice(1).join(" "), plate };
}

// Вид пального у словах реєстру, щоб модуль хабу розпізнав групу (Бензин / Дизель / Газ)
export function fuelLabel(f) {
  const s = String(f || "").trim();
  if (!s) return "";
  if (/^а-?\d{2}/i.test(s) || /бенз/i.test(s)) return /бенз/i.test(s) ? s : `Бензин ${s.toUpperCase().replace(/^А(\d)/, "А-$1")}`;
  if (/^дп$/i.test(s) || /дизел/i.test(s)) return /дизел/i.test(s) ? s : "Дизельне паливо";
  if (/газ|lpg|cng/i.test(s)) return /газ/i.test(s) && s.length > 3 ? s : "Газ";
  return s;
}

// Норма, за якою портал рахує витрату (і готує документи для бухгалтерії): базова + зимова надбавка, %
const effRate = (v) => { const r = num(v.rate); if (r === null) return null; const w = num(v.winter_pct) || 0; return Math.round(r * (1 + w / 100) * 1000) / 1000; };
const firstDay = (ym) => { const [y, m] = parseYm(ym); return new Date(y, m - 1, 1); };
export const nextYm = (ym) => { const [y, m] = parseYm(ym); return m === 12 ? ymKey(y + 1, 1) : ymKey(y, m + 1); };
const num = (v) => (v === null || v === undefined || v === "" || !Number.isFinite(Number(v)) ? null : Number(v));
const pad = (n) => String(n).padStart(2, "0");
export const dmy = (d) => `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;

/**
 * items: [{org, unit, vehicles: [...], entries: {vehicleId: {odo_start, odo_end}}}]
 * kind: "end"   — зріз на 01 числа наступного місяця: показник на початок (01.ММ) і на кінець (01.ММ+1);
 *       "start" — зріз на 01.ММ: лише показник на початок місяця (для першого місяця, коли попереднього файлу немає).
 * Повертає {rows, date, stats}.
 */
export function buildFleetRows(items, ym, kind = "end") {
  const d0 = firstDay(ym), d1 = firstDay(nextYm(ym));
  const rows = [FLEET_HEADERS.slice(), FLEET_HEADERS.map((_, i) => i + 1)];
  let n = 0, withReadings = 0, noReadings = 0;
  for (const it of items) {
    for (const v of it.vehicles) {
      const e = (it.entries || {})[v.id] || {};
      const { brand, model, plate } = splitVehicleName(v.name);
      const readings = [];
      const s = num(e.odo_start), f = num(e.odo_end);
      if (s !== null) readings.push([s, d0]);
      if (kind === "end" && f !== null) readings.push([f, d1]);
      n += 1; readings.length ? withReadings++ : noReadings++;
      const meta = [n, v.region || "", plate || "БН", brand, model, num(v.year), fuelLabel(v.fuel), effRate(v), it.org || "", it.unit || ""];
      if (!readings.length) { rows.push([...meta, null, null]); continue; }
      // Орган і служба — у кожному рядку показника (як у реєстрі): модуль хабу бере їх з останнього показника
      readings.forEach(([km, d], i) => rows.push(i === 0 ? [...meta, km, d] : [null, null, null, null, null, null, null, null, it.org || "", it.unit || "", km, d]));
    }
  }
  return { rows, date: kind === "end" ? d1 : d0, stats: { cars: n, withReadings, noReadings } };
}
