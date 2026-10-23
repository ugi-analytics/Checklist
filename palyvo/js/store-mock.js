// Локальна заглушка сховища (той самий інтерфейс, що й store-firebase.js) — лише для перевірки інтерфейсу без інтернету.
// Вмикається адресою index.html?demo=1. Дані живуть у пам'яті вкладки (sessionStorage).
const KEY = "palyvo-demo-db";
const nowIso = () => new Date().toLocaleString("sv-SE").replace("T", " ");
const rid = () => Math.random().toString(36).slice(2, 12);

export function createStore() {
  let S;
  try { S = JSON.parse(sessionStorage.getItem(KEY)); } catch { S = null; }
  if (!S) S = { auth: {}, users: {}, units: {}, months: {}, audit: [], setup: false, session: null };
  const save = () => { try { sessionStorage.setItem(KEY, JSON.stringify(S)); } catch { /* */ } };
  const listeners = [];
  const emit = () => { const u = S.session ? { uid: S.session, email: S.auth[S.session]?.email } : null; listeners.forEach((f) => f(u)); };
  const err = (code) => { const e = new Error(code); e.code = code; throw e; };
  const uidByEmail = (email) => Object.keys(S.auth).find((k) => S.auth[k].email === email.trim().toLowerCase());
  return {
    kind: "demo",
    errUa: (e) => ({ "auth/invalid-credential": "Невірний e-mail або пароль.", "auth/email-already-in-use": "Цей e-mail уже зареєстровано. Увійдіть або відновіть пароль." }[e?.code] || e?.message || String(e)),
    onAuth(cb) { listeners.push(cb); setTimeout(emit, 0); return () => {}; },
    currentUser() { return S.session ? { uid: S.session, email: S.auth[S.session].email } : null; },
    async signIn(email, pw) { const uid = uidByEmail(email); if (!uid || S.auth[uid].pw !== pw) err("auth/invalid-credential"); S.session = uid; save(); emit(); },
    async signUp(email, pw) { if (uidByEmail(email)) err("auth/email-already-in-use"); const uid = "u" + rid(); S.auth[uid] = { email: email.trim().toLowerCase(), pw }; S.session = uid; save(); emit(); },
    async signOut() { S.session = null; save(); emit(); },
    async resetPassword() { /* демо: лист не надсилається */ },
    async changePassword(o, n) { if (S.auth[S.session].pw !== o) err("auth/invalid-credential"); S.auth[S.session].pw = n; save(); },
    async setupDone() { return S.setup; },
    async bootstrapAdmin(profile, unitName, unitSettings) {
      const id = "un" + rid(); S.units[id] = { name: unitName, settings: unitSettings, vehicles: [], created_at: nowIso() };
      S.users[S.session] = { ...profile, email: S.auth[S.session].email, role: "admin", status: "active", unit_id: id, created_at: nowIso() };
      S.setup = true; save(); return id;
    },
    async getMe() { return S.session && S.users[S.session] ? { uid: S.session, ...S.users[S.session] } : null; },
    async createMyProfile(p) { S.users[S.session] = { ...p, email: S.auth[S.session].email, role: "user", status: "pending", created_at: nowIso() }; save(); },
    async touchLogin() { if (S.users[S.session]) { S.users[S.session].last_login = nowIso(); save(); } },
    async listUsers() { return Object.entries(S.users).map(([uid, u]) => ({ uid, ...u })); },
    async updateUser(uid, f) { Object.assign(S.users[uid], f); save(); },
    async deleteUser(uid) { delete S.users[uid]; save(); },
    async listUnits() { return Object.entries(S.units).map(([id, u]) => ({ id, ...u })).sort((a, b) => a.name.localeCompare(b.name, "uk")); },
    async getUnit(id) { return S.units[id] ? { id, ...JSON.parse(JSON.stringify(S.units[id])) } : null; },
    async createUnit(name, settings) { const id = "un" + rid(); S.units[id] = { name, settings, vehicles: [], created_at: nowIso() }; save(); return id; },
    async updateUnit(id, f) { Object.assign(S.units[id], JSON.parse(JSON.stringify(f))); save(); },
    async deleteUnit(id) { delete S.units[id]; for (const k of Object.keys(S.months)) if (k.startsWith(id + "/")) delete S.months[k]; save(); },
    async getMonth(u, ym) { const m = S.months[`${u}/${ym}`]; return m ? JSON.parse(JSON.stringify(m)) : null; },
    async saveMonth(u, ym, d) { S.months[`${u}/${ym}`] = JSON.parse(JSON.stringify({ ...d, updated_at: nowIso(), updated_by: S.session })); save(); },
    async listMonths(u) { return Object.entries(S.months).filter(([k]) => k.startsWith(u + "/")).map(([k, v]) => ({ ym: k.split("/")[1], ...v })).sort((a, b) => a.ym.localeCompare(b.ym)); },
    async log(action, details = "", unitId = null) { S.audit.unshift({ ts: nowIso(), uid: S.session, email: S.auth[S.session]?.email, unit_id: unitId, action, details }); save(); },
    async listAudit(n = 200) { return S.audit.slice(0, n); },
  };
}
