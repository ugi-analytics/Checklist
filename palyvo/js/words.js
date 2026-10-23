// Числа прописом українською (перенесено з words_ua.py).
const UNITS_M = ["", "один", "два", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"];
const UNITS_F = ["", "одна", "дві", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"];
const UNITS_FA = ["", "одну", "дві", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"];
const TEENS = ["десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п'ятнадцять", "шістнадцять",
  "сімнадцять", "вісімнадцять", "дев'ятнадцять"];
const TENS = ["", "", "двадцять", "тридцять", "сорок", "п'ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев'яносто"];
const HUNDREDS = ["", "сто", "двісті", "триста", "чотириста", "п'ятсот", "шістсот", "сімсот", "вісімсот", "дев'ятсот"];

export function plural(n, one, few, many) {
  n = Math.abs(n) % 100;
  if (n >= 11 && n <= 19) return many;
  n %= 10;
  if (n === 1) return one;
  if (n >= 2 && n <= 4) return few;
  return many;
}
function triad(n, units) {
  const out = []; const h = Math.floor(n / 100), rest = n % 100;
  if (h) out.push(HUNDREDS[h]);
  if (rest >= 10 && rest <= 19) out.push(TEENS[rest - 10]);
  else { const t = Math.floor(rest / 10), u = rest % 10; if (t) out.push(TENS[t]); if (u) out.push(units[u]); }
  return out;
}
export function intWords(n, gender = "m", acc = false) {
  const units = { m: UNITS_M, f: UNITS_F, fa: UNITS_FA }[gender];
  if (n === 0) return "нуль";
  const parts = [];
  const millions = Math.floor(n / 1e6); let rest = n % 1e6;
  const thousands = Math.floor(rest / 1000); rest = rest % 1000;
  if (millions) parts.push(...triad(millions, UNITS_M), plural(millions, "мільйон", "мільйони", "мільйонів"));
  if (thousands) parts.push(...triad(thousands, acc ? UNITS_FA : UNITS_F), plural(thousands, acc ? "тисячу" : "тисяча", "тисячі", "тисяч"));
  if (rest) parts.push(...triad(rest, units));
  return parts.join(" ");
}
/** Копійки → «шістсот тридцять одну цілу сімдесят сім сотих літра». */
export function litresWords(c) {
  const whole = Math.floor(c / 100), frac = c % 100;
  if (frac === 0) return `${intWords(whole, "m", true)} ${plural(whole, "літр", "літри", "літрів")}`;
  const w = `${intWords(whole, "fa", true)} ${plural(whole, "цілу", "цілих", "цілих")}`;
  let f, one, many;
  if (frac % 10 === 0) { f = frac / 10; one = "десяту"; many = "десятих"; } else { f = frac; one = "соту"; many = "сотих"; }
  return `${w} ${intWords(f, "fa")} ${plural(f, one, many, many)} літра`;
}
