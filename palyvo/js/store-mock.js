// Локальна заглушка сховища (той самий інтерфейс, що й store-firebase.js) — лише для перевірки інтерфейсу без інтернету.
// Вмикається адресою index.html?demo=1. Дані живуть у пам'яті вкладки (sessionStorage).
// Двофакторний вхід імітується: приймається будь-який 6-значний код; «лист підтвердження» вважається відкритим після «Я перейшов за посиланням».
const KEY = "palyvo-demo-db";
const nowIso = () => new Date().toLocaleString("sv-SE").replace("T", " ");
const rid = () => Math.random().toString(36).slice(2, 12);
const clone = (x) => JSON.parse(JSON.stringify(x));

export function createStore() {
  let S;
  try { S = JSON.parse(sessionStorage.getItem(KEY)); } catch { S = null; }
  if (!S) S = { auth: {}, users: {}, units: {}, months: {}, audit: [], setup: false, session: null, mfaSession: false, meta: {} };
  const save = () => { try { sessionStorage.setItem(KEY, JSON.stringify(S)); } catch { /* */ } };
  const listeners = [];
  const emit = () => { const u = S.session ? { uid: S.session, email: S.auth[S.session]?.email } : null; listeners.forEach((f) => f(u)); };
  const err = (code, msg) => { const e = new Error(msg || code); e.code = code; throw e; };
  const uidByEmail = (email) => Object.keys(S.auth).find((k) => S.auth[k].email === email.trim().toLowerCase());
  let pendingMfa = null;
  return {
    kind: "demo",
    errUa: (e) => ({ "auth/invalid-credential": "Невірний e-mail або пароль.", "auth/email-already-in-use": "Цей e-mail уже зареєстровано. Увійдіть або відновіть пароль.",
      "auth/invalid-verification-code": "Невірний код. Введіть 6 цифр, які зараз показує застосунок на телефоні." }[e?.code] || e?.message || String(e)),
    onAuth(cb) { listeners.push(cb); setTimeout(emit, 0); return () => {}; },
    currentUser() { return S.session ? { uid: S.session, email: S.auth[S.session].email } : null; },
    async signIn(email, pw) {
      const uid = uidByEmail(email); if (!uid || S.auth[uid].pw !== pw) err("auth/invalid-credential");
      if (S.auth[uid].mfa) { pendingMfa = uid; err("mfa-required"); }
      S.session = uid; S.mfaSession = false; save(); emit();
    },
    async finishMfaSignIn(code) {
      if (!pendingMfa) err("x", "Сесія входу застаріла. Введіть e-mail і пароль ще раз.");
      // демо: будь-які 6 цифр, крім 000000 (ним перевіряється «невірний код»)
      if (!/^\d{6}$/.test(String(code).trim()) || String(code).trim() === "000000") err("auth/invalid-verification-code");
      S.session = pendingMfa; S.mfaSession = true; pendingMfa = null; save(); emit();
    },
    mfaRequired: true,
    async mfaState() { const a = S.auth[S.session]; return a ? { enrolled: !!a.mfa, session: !!S.mfaSession, emailVerified: !!a.verified } : null; },
    async sendVerify() { /* демо: лист не надсилається */ },
    async reloadUser() { S.auth[S.session].verified = true; save(); return true; },
    async reauth(pw) { if (S.auth[S.session].pw !== pw) err("auth/invalid-credential"); },
    async startTotp() { return { uri: `otpauth://totp/Palyvo%20MVO:${encodeURIComponent(S.auth[S.session].email)}?secret=JBSWY3DPEHPK3PXPJBSWY3DP&issuer=Palyvo%20MVO`, key: "JBSWY3DPEHPK3PXPJBSWY3DP" }; },
    async finishTotp(code) { if (!/^\d{6}$/.test(String(code).trim())) err("auth/invalid-verification-code"); S.auth[S.session].mfa = true; S.mfaSession = true; save(); },
    async signUp(email, pw) { if (uidByEmail(email)) err("auth/email-already-in-use"); const uid = "u" + rid(); S.auth[uid] = { email: email.trim().toLowerCase(), pw }; S.session = uid; S.mfaSession = false; save(); emit(); },
    async signOut() { S.session = null; S.mfaSession = false; save(); emit(); },
    async resetPassword() { /* демо: лист не надсилається */ },
    async changePassword(o, n) { if (S.auth[S.session].pw !== o) err("auth/invalid-credential"); S.auth[S.session].pw = n; save(); },
    async setupDone() { return S.setup; },
    async bootstrapAdmin(profile, unitName, unitSettings, adminKeys) {
      const id = "un" + rid(); S.units[id] = { name: unitName, settings: unitSettings, n_vehicles: 0, created_at: nowIso() };
      S.users[S.session] = { ...profile, email: S.auth[S.session].email, role: "admin", status: "active", unit_id: id, created_at: nowIso() };
      S.meta.admin_pub = adminKeys.pub; S.meta.admin_priv = adminKeys.priv;
      S.setup = true; save(); return id;
    },
    async getMe() { return S.session && S.users[S.session] ? { uid: S.session, ...S.users[S.session] } : null; },
    async createMyProfile(p) { S.users[S.session] = { ...p, email: S.auth[S.session].email, role: "user", status: "pending", created_at: nowIso() }; save(); },
    async heartbeat() { if (S.users[S.session]) { S.users[S.session].last_seen = Date.now(); save(); } },
    seenMs: (v) => Number(v) || 0,
    async touchLogin() { if (S.users[S.session]) { S.users[S.session].last_login = nowIso(); save(); } },
    async listUsers() { return Object.entries(S.users).map(([uid, u]) => ({ uid, ...u })); },
    async updateUser(uid, f) { Object.assign(S.users[uid], f); save(); },
    async deleteUser(uid) { delete S.users[uid]; save(); },
    async listUnits() { return Object.entries(S.units).map(([id, u]) => ({ id, ...clone(u) })).sort((a, b) => a.name.localeCompare(b.name, "uk")); },
    async getUnit(id) { return S.units[id] ? { id, ...clone(S.units[id]) } : null; },
    async createUnit(name, settings) { const id = "un" + rid(); S.units[id] = { name, settings, n_vehicles: 0, created_at: nowIso() }; save(); return id; },
    async updateUnit(id, f) { Object.assign(S.units[id], clone(f)); save(); },
    async deleteUnit(id) { delete S.units[id]; for (const k of Object.keys(S.months)) if (k.startsWith(id + "/")) delete S.months[k]; save(); },
    async getAdminPub() { if (!S.meta.admin_pub) err("x", "Адміністратор порталу ще не завершив налаштування шифрування (ключ адміністратора не створено). Зверніться до нього."); return clone(S.meta.admin_pub); },
    async hasAdminKey() { return !!S.meta.admin_pub; },
    async createAdminKeys(k) { if (S.meta.admin_pub) err("permission-denied", "Ключ адміністратора вже існує."); S.meta.admin_pub = clone(k.pub); S.meta.admin_priv = clone(k.priv); save(); },
    async getAdminPriv() { if (!S.meta.admin_priv) err("x", "Ключ адміністратора ще не створено."); return clone(S.meta.admin_priv); },
    async setAdminPriv(p) { S.meta.admin_priv = clone(p); save(); },
    async initUnitCrypto(id, c, e, n) {
      const u = S.units[id]; if (!u) err("x", "Підрозділ не знайдено.");
      if (u.crypto) err("already-init", "Ключ цього підрозділу вже створено.");
      u.crypto = clone(c); u.enc = clone(e); u.n_vehicles = n; delete u.settings; delete u.vehicles; save();
    },
    async saveUnitData(id, e, n) { const u = S.units[id]; u.enc = clone(e); u.n_vehicles = n; delete u.settings; delete u.vehicles; save(); },
    async setUnitPass(id, p) { S.units[id].crypto.pass = clone(p); save(); },
    async getMonth(u, ym) { const m = S.months[`${u}/${ym}`]; return m ? clone(m) : null; },
    async saveMonth(u, ym, e) { S.months[`${u}/${ym}`] = clone({ enc: e, updated_at: nowIso(), updated_by: S.session }); save(); },
    async listMonths(u) { return Object.entries(S.months).filter(([k]) => k.startsWith(u + "/")).map(([k, v]) => ({ ym: k.split("/")[1], ...clone(v) })).sort((a, b) => a.ym.localeCompare(b.ym)); },
    async log(action, details = "", unitId = null) { S.audit.unshift({ ts: nowIso(), uid: S.session, email: S.auth[S.session]?.email, unit_id: unitId, action, details }); save(); },
    async listAudit(n = 200) { return S.audit.slice(0, n); },
  };
}
