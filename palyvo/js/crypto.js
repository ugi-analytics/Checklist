// Шифрування даних у браузері (Web Crypto API). У базу потрапляє лише зашифроване.
//
// Схема:
//  • Кожен підрозділ має випадковий ключ даних DK (AES-256-GCM). Ним шифруються автопарк, реквізити й місяці.
//  • DK зберігається в базі лише «загорнутим» двічі:
//      crypto.pass  — ключем, отриманим із ключа підрозділу (фрази) через PBKDF2-SHA256 (600 000 ітерацій);
//      crypto.admin — відкритим ключем адміністратора (RSA-OAEP 3072).
//  • Закритий ключ адміністратора лежить у базі загорнутим ключем адміністратора (фразою) — бачить його лише адмін.
//  • Кожен зашифрований блок прив'язаний до місця (AAD «unit:…», «month:…:РРРР-ММ»), тож блоки не можна підмінити місцями.
const enc = new TextEncoder(), dec = new TextDecoder();
const subtle = () => {
  if (!globalThis.crypto || !crypto.subtle) throw new Error("Браузер не підтримує шифрування (потрібен сучасний Chrome/Edge/Firefox і адреса https).");
  return crypto.subtle;
};
export const b64 = (buf) => { const b = new Uint8Array(buf); let s = ""; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); };
export const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const rnd = (n) => crypto.getRandomValues(new Uint8Array(n));
const ITER = 600000;
const GCM = (iv, aad) => (aad ? { name: "AES-GCM", iv, additionalData: enc.encode(aad) } : { name: "AES-GCM", iv });
const RSA = { name: "RSA-OAEP", hash: "SHA-256" };

export class WrongKey extends Error { constructor(msg = "Невірний ключ.") { super(msg); this.code = "wrong-key"; } }

async function kek(pass, salt, iter) {
  const base = await subtle().importKey("raw", enc.encode(String(pass).normalize("NFC")), "PBKDF2", false, ["deriveKey"]);
  return subtle().deriveKey({ name: "PBKDF2", hash: "SHA-256", salt, iterations: iter }, base, { name: "AES-GCM", length: 256 }, false, ["wrapKey", "unwrapKey"]);
}

// ------------------------------------------------------------- ключ даних підрозділу
export const newDataKey = () => subtle().generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt", "decrypt"]);

export async function wrapWithPass(dk, pass, aad) {
  const salt = rnd(16), iv = rnd(12);
  const w = await subtle().wrapKey("raw", dk, await kek(pass, salt, ITER), GCM(iv, aad));
  return { salt: b64(salt), iter: ITER, iv: b64(iv), w: b64(w) };
}
export async function unwrapWithPass(obj, pass, aad, extractable = false) {
  try {
    return await subtle().unwrapKey("raw", unb64(obj.w), await kek(pass, unb64(obj.salt), obj.iter || ITER), GCM(unb64(obj.iv), aad),
      { name: "AES-GCM", length: 256 }, extractable, ["encrypt", "decrypt"]);
  } catch (e) { if (e && e.name === "OperationError") throw new WrongKey("Невірний ключ підрозділу."); throw e; }
}

// ------------------------------------------------------------- ключ адміністратора
export async function newAdminKeys(pass) {
  const kp = await subtle().generateKey({ ...RSA, modulusLength: 3072, publicExponent: new Uint8Array([1, 0, 1]) }, true, ["wrapKey", "unwrapKey"]);
  const pub = b64(await subtle().exportKey("spki", kp.publicKey));
  const salt = rnd(16), iv = rnd(12);
  const w = await subtle().wrapKey("pkcs8", kp.privateKey, await kek(pass, salt, ITER), GCM(iv, "admin-private"));
  return { pub: { spki: pub, alg: "RSA-OAEP-3072-SHA256" }, priv: { salt: b64(salt), iter: ITER, iv: b64(iv), w: b64(w) } };
}
export async function unlockAdmin(priv, pass, extractable = false) {
  try {
    return await subtle().unwrapKey("pkcs8", unb64(priv.w), await kek(pass, unb64(priv.salt), priv.iter || ITER), GCM(unb64(priv.iv), "admin-private"),
      RSA, extractable, ["unwrapKey"]);
  } catch (e) { if (e && e.name === "OperationError") throw new WrongKey("Невірний ключ адміністратора."); throw e; }
}
export async function rewrapAdmin(priv, oldPass, newPass) {
  const pk = await unlockAdmin(priv, oldPass, true);
  const salt = rnd(16), iv = rnd(12);
  const w = await subtle().wrapKey("pkcs8", pk, await kek(newPass, salt, ITER), GCM(iv, "admin-private"));
  return { salt: b64(salt), iter: ITER, iv: b64(iv), w: b64(w) };
}
export async function wrapForAdmin(dk, pub) {
  const pk = await subtle().importKey("spki", unb64(pub.spki), RSA, false, ["wrapKey"]);
  return { w: b64(await subtle().wrapKey("raw", dk, pk, RSA)) };
}
export async function unwrapByAdmin(obj, adminPriv, extractable = false) {
  try {
    return await subtle().unwrapKey("raw", unb64(obj.w), adminPriv, RSA, { name: "AES-GCM", length: 256 }, extractable, ["encrypt", "decrypt"]);
  } catch (e) { if (e && e.name === "OperationError") throw new WrongKey("Ключ адміністратора не підходить до цього підрозділу."); throw e; }
}

// ------------------------------------------------------------- ключ диспетчера (спільний для диспетчерів)
// Диспетчер бачить лише довідник авто всіх підрозділів (без показників за місяці). Підрозділ публікує довідник окремим блоком,
// зашифрованим випадковим ключем CK; CK загорнуто відкритим ключем диспетчера. Закритий — у базі, загорнутий фразою диспетчера.
export async function newDispKeys(pass) {
  const kp = await subtle().generateKey({ ...RSA, modulusLength: 3072, publicExponent: new Uint8Array([1, 0, 1]) }, true, ["wrapKey", "unwrapKey"]);
  const spki = b64(await subtle().exportKey("spki", kp.publicKey));
  const salt = rnd(16), iv = rnd(12);
  const w = await subtle().wrapKey("pkcs8", kp.privateKey, await kek(pass, salt, ITER), GCM(iv, "disp-private"));
  const kid = b64(await subtle().digest("SHA-256", enc.encode(spki))).slice(0, 16);
  return { pub: { spki, kid, alg: "RSA-OAEP-3072-SHA256" }, priv: { kid, salt: b64(salt), iter: ITER, iv: b64(iv), w: b64(w) } };
}
export async function unlockDisp(priv, pass) {
  try {
    return await subtle().unwrapKey("pkcs8", unb64(priv.w), await kek(pass, unb64(priv.salt), priv.iter || ITER), GCM(unb64(priv.iv), "disp-private"),
      RSA, false, ["unwrapKey"]);
  } catch (e) { if (e && e.name === "OperationError") throw new WrongKey("Невірний ключ диспетчера."); throw e; }
}
/** Довідник авто підрозділу для диспетчера: {enc, k, kid}. */
export async function sealForDisp(obj, pub, aad) {
  const ck = await newDataKey();
  const encBlob = await encJson(ck, obj, aad);
  const k = await wrapForAdmin(ck, pub);
  return { enc: encBlob, k, kid: pub.kid };
}
export async function openForDisp(doc, priv, aad) {
  const ck = await subtle().unwrapKey("raw", unb64(doc.k.w), priv, RSA, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  return decJson(ck, doc.enc, aad);
}
export async function sha256(text) { return b64(await subtle().digest("SHA-256", enc.encode(text))); }
export const aadFleet = (unitId) => `fleet:${unitId}`;

// ------------------------------------------------------------- дані
export async function encJson(dk, obj, aad) {
  const iv = rnd(12);
  const ct = await subtle().encrypt(GCM(iv, aad), dk, enc.encode(JSON.stringify(obj)));
  return { v: 1, iv: b64(iv), ct: b64(ct) };
}
export async function decJson(dk, blob, aad) {
  try { return JSON.parse(dec.decode(await subtle().decrypt(GCM(unb64(blob.iv), aad), dk, unb64(blob.ct)))); }
  catch (e) { if (e && e.name === "OperationError") throw new Error("Дані пошкоджено або зашифровано іншим ключем."); throw e; }
}
export const aadUnit = (unitId) => `unit:${unitId}`;
export const aadMonth = (unitId, ym) => `month:${unitId}:${ym}`;
export const aadPass = (unitId) => `unitkey:${unitId}`;

// ------------------------------------------------------------- запам'ятовування ключів на цьому комп'ютері
// Ключ підрозділу зберігається в IndexedDB браузера як НЕЕКСПОРТОВАНИЙ CryptoKey: скрипт може ним шифрувати,
// але сам ключ дістати й переслати неможливо. Прив'язано до користувача (uid) і підрозділу.
const DB = "palyvo-keys";
function idb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore("k");
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
}
async function tx(mode, fn) {
  const db = await idb();
  return new Promise((res, rej) => {
    const t = db.transaction("k", mode); const st = t.objectStore("k"); const r = fn(st);
    t.oncomplete = () => { db.close(); res(r && r.result); }; t.onerror = () => { db.close(); rej(t.error); };
  });
}
export const keyCache = {
  async get(id) { try { return (await tx("readonly", (s) => s.get(id))) || null; } catch { return null; } },
  async put(id, key) { try { await tx("readwrite", (s) => s.put(key, id)); return true; } catch { return false; } },
  async del(id) { try { await tx("readwrite", (s) => s.delete(id)); } catch { /* */ } },
  async clearPrefix(prefix) {
    try { await tx("readwrite", (s) => { const r = s.openCursor(); r.onsuccess = () => { const c = r.result; if (!c) return; if (String(c.key).startsWith(prefix)) c.delete(); c.continue(); }; return r; }); } catch { /* */ }
  },
};
// перетворити ключ на неекспортований (для зберігання)
export async function lockDown(dk) {
  if (!dk.extractable) return dk;
  const raw = await subtle().exportKey("raw", dk);
  return subtle().importKey("raw", raw, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
}

// перевірка фрази-ключа
export function passProblem(p, what = "Ключ") {
  if (String(p).length < 10) return `${what} має містити щонайменше 10 символів (краще — фраза з кількох слів).`;
  if (/^\d+$/.test(p)) return `${what} не може складатися лише з цифр.`;
  return null;
}
