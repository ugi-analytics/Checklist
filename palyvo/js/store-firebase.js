// Дані порталу у Firebase: Authentication (e-mail + пароль) і Firestore.
// Структура Firestore:
//   meta/setup                       — позначка, що адміністратора створено
//   users/{uid}                      — профіль: ПІБ, посада, підрозділ, роль, статус
//   units/{unitId}                   — підрозділ: назва, реквізити (settings), автопарк (vehicles[])
//   units/{unitId}/months/{РРРР-ММ}  — дані місяця: entries, snapshot
//   audit/{id}                       — журнал дій
// Хто що бачить — визначають правила firestore.rules (кожен підрозділ — лише своє).
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut,
  sendPasswordResetEmail, updatePassword, reauthenticateWithCredential, EmailAuthProvider, setPersistence,
  browserLocalPersistence, sendEmailVerification, multiFactor, TotpMultiFactorGenerator, getMultiFactorResolver,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, collection, writeBatch, query, orderBy, limit,
  addDoc, deleteField, runTransaction, serverTimestamp,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const nowIso = () => new Date().toLocaleString("sv-SE").replace("T", " ");

export function createStore(config) {
  const app = initializeApp(config);
  const auth = getAuth(app);
  setPersistence(auth, browserLocalPersistence).catch(() => {});
  const db = getFirestore(app);

  let mfaResolver = null, totpSecret = null;
  const errUa = (e) => {
    const c = (e && e.code) || "";
    const map = {
      "auth/invalid-credential": "Невірний e-mail або пароль.", "auth/wrong-password": "Невірний e-mail або пароль.",
      "auth/user-not-found": "Невірний e-mail або пароль.", "auth/invalid-email": "Некоректний e-mail.",
      "auth/email-already-in-use": "Цей e-mail уже зареєстровано. Увійдіть або відновіть пароль.",
      "auth/weak-password": "Пароль занадто простий (мінімум 8 символів, літери й цифри).",
      "auth/too-many-requests": "Забагато спроб. Зачекайте кілька хвилин.",
      "auth/network-request-failed": "Немає з'єднання з інтернетом.",
      "permission-denied": "Немає доступу (перевірте, чи підтверджено ваш обліковий запис і чи виконано вхід із кодом із телефона).",
      "auth/invalid-verification-code": "Невірний код. Введіть 6 цифр, які зараз показує застосунок на телефоні.",
      "auth/totp-challenge-timeout": "Час на введення коду минув. Увійдіть ще раз.",
      "auth/requires-recent-login": "Для цієї дії потрібно ще раз ввести пароль.",
      "auth/unverified-email": "Спочатку підтвердіть e-mail за посиланням у листі.",
      "auth/operation-not-allowed": "Двофакторний вхід ще не увімкнено в налаштуваннях Firebase (див. інструкцію адміністратора).",
      "auth/admin-restricted-operation": "Двофакторний вхід ще не увімкнено в налаштуваннях Firebase (див. інструкцію адміністратора).",
      "auth/missing-multi-factor-session": "Сесія налаштування застаріла. Почніть налаштування ще раз.",
    };
    return map[c] || (e && e.message) || String(e);
  };

  return {
    kind: "firebase",
    errUa,
    // ---------------------------------------------------------------- вхід
    onAuth(cb) { return onAuthStateChanged(auth, (u) => cb(u ? { uid: u.uid, email: u.email } : null)); },
    currentUser() { const u = auth.currentUser; return u ? { uid: u.uid, email: u.email } : null; },
    async signIn(email, pw) {
      try { await signInWithEmailAndPassword(auth, email.trim(), pw); }
      catch (e) {
        if (e && e.code === "auth/multi-factor-auth-required") {
          mfaResolver = getMultiFactorResolver(auth, e);
          const err = new Error("mfa-required"); err.code = "mfa-required"; throw err;
        }
        throw e;
      }
    },
    // другий крок входу: 6-значний код із застосунку-автентифікатора
    async finishMfaSignIn(code) {
      if (!mfaResolver) { const e = new Error("Сесія входу застаріла. Введіть e-mail і пароль ще раз."); e.code = "x"; throw e; }
      // якщо телефон прив'язували кілька разів — пробуємо код з кожним записом (новіші першими)
      const hints = mfaResolver.hints.filter((h) => h.factorId === TotpMultiFactorGenerator.FACTOR_ID)
        .sort((x, y) => (Date.parse(y.enrollmentTime) || 0) - (Date.parse(x.enrollmentTime) || 0));
      const c = String(code).replace(/\D/g, ""); let last = null;
      for (const h of (hints.length ? hints : mfaResolver.hints)) {
        try { await mfaResolver.resolveSignIn(TotpMultiFactorGenerator.assertionForSignIn(h.uid, c)); mfaResolver = null; return; }
        catch (e) { last = e; if (e && e.code !== "auth/invalid-verification-code") throw e; }
      }
      throw last;
    },
    mfaRequired: true,
    async mfaState() {
      const u = auth.currentUser; if (!u) return null;
      // свіжі дані з сервера: на іншому комп'ютері збережений вхід міг лишитися ще з часу ДО прив'язки телефона
      try { await u.reload(); } catch { /* офлайн — беремо збережене */ }
      const t = await u.getIdTokenResult();
      return { enrolled: multiFactor(u).enrolledFactors.length > 0, session: !!(t.claims.firebase && t.claims.firebase.sign_in_second_factor),
        emailVerified: !!u.emailVerified };
    },
    sendVerify: () => sendEmailVerification(auth.currentUser, { url: location.origin + location.pathname }),
    async reloadUser() { const u = auth.currentUser; await u.reload(); await u.getIdToken(true); return auth.currentUser.emailVerified; },
    reauth: (pw) => reauthenticateWithCredential(auth.currentUser, EmailAuthProvider.credential(auth.currentUser.email, pw)),
    async startTotp() {
      const u = auth.currentUser;
      totpSecret = await TotpMultiFactorGenerator.generateSecret(await multiFactor(u).getSession());
      return { uri: totpSecret.generateQrCodeUrl(u.email, "Palyvo MVO"), key: totpSecret.secretKey };
    },
    async finishTotp(code) {
      const u = auth.currentUser;
      await multiFactor(u).enroll(TotpMultiFactorGenerator.assertionForEnrollment(totpSecret, String(code).replace(/\D/g, "")), "Телефон");
      totpSecret = null;
      await u.getIdToken(true).catch(() => {});
    },
    signUp: (email, pw) => createUserWithEmailAndPassword(auth, email.trim(), pw),
    signOut: () => signOut(auth),
    resetPassword: (email) => sendPasswordResetEmail(auth, email.trim()),
    async changePassword(oldPw, newPw) {
      const u = auth.currentUser;
      await reauthenticateWithCredential(u, EmailAuthProvider.credential(u.email, oldPw));
      await updatePassword(u, newPw);
    },
    // ---------------------------------------------------------------- налаштування
    async setupDone() { return (await getDoc(doc(db, "meta", "setup"))).exists(); },
    async bootstrapAdmin(profile, unitName, unitSettings, adminKeys) {
      const uid = auth.currentUser.uid;
      const unitRef = doc(collection(db, "units"));
      const b = writeBatch(db);
      b.set(unitRef, { name: unitName, settings: unitSettings, n_vehicles: 0, created_at: nowIso() });
      b.set(doc(db, "meta", "admin_pub"), { ...adminKeys.pub, at: nowIso() });
      b.set(doc(db, "meta", "admin_priv"), { ...adminKeys.priv, at: nowIso() });
      b.set(doc(db, "users", uid), { ...profile, email: auth.currentUser.email, role: "admin", status: "active",
        unit_id: unitRef.id, created_at: nowIso(), last_login: nowIso() });
      b.set(doc(db, "meta", "setup"), { done: true, at: nowIso(), by: uid });
      await b.commit();
      return unitRef.id;
    },
    // ---------------------------------------------------------------- користувачі
    async getMe() { const u = auth.currentUser; if (!u) return null; const s = await getDoc(doc(db, "users", u.uid)); return s.exists() ? { uid: u.uid, ...s.data() } : null; },
    async createMyProfile(profile) {
      const u = auth.currentUser;
      await setDoc(doc(db, "users", u.uid), { ...profile, email: u.email, role: "user", status: "pending", created_at: nowIso() });
    },
    touchLogin: () => updateDoc(doc(db, "users", auth.currentUser.uid), { last_login: nowIso() }).catch(() => {}),
    // «я на сайті» — час сервера, щоб годинники різних комп'ютерів не заважали
    heartbeat: () => (auth.currentUser ? updateDoc(doc(db, "users", auth.currentUser.uid), { last_seen: serverTimestamp() }).catch(() => {}) : Promise.resolve()),
    seenMs: (v) => (v && typeof v.toMillis === "function" ? v.toMillis() : 0),
    async listUsers() { return (await getDocs(collection(db, "users"))).docs.map((d) => ({ uid: d.id, ...d.data() })); },
    updateUser: (uid, fields) => updateDoc(doc(db, "users", uid), fields),
    deleteUser: (uid) => deleteDoc(doc(db, "users", uid)),
    // ---------------------------------------------------------------- підрозділи
    async listUnits() { return (await getDocs(collection(db, "units"))).docs.map((d) => ({ id: d.id, ...d.data() }))
      .sort((a, b) => a.name.localeCompare(b.name, "uk")); },
    async getUnit(id) { const s = await getDoc(doc(db, "units", id)); return s.exists() ? { id, ...s.data() } : null; },
    async createUnit(name, settings) { const r = doc(collection(db, "units")); await setDoc(r, { name, settings, n_vehicles: 0, created_at: nowIso() }); return r.id; },
    // ---------------------------------------------------------------- шифрування
    async getAdminPub() { const s = await getDoc(doc(db, "meta", "admin_pub")); if (!s.exists()) throw new Error("Адміністратор порталу ще не завершив налаштування шифрування (ключ адміністратора не створено). Зверніться до нього."); return s.data(); },
    async hasAdminKey() { return (await getDoc(doc(db, "meta", "admin_pub"))).exists(); },
    // ключ адміністратора, якщо портал налаштовано попередньою версією (без шифрування)
    async createAdminKeys(keys) {
      const b = writeBatch(db);
      b.set(doc(db, "meta", "admin_pub"), { ...keys.pub, at: nowIso() });
      b.set(doc(db, "meta", "admin_priv"), { ...keys.priv, at: nowIso() });
      await b.commit();
    },
    async getAdminPriv() { const s = await getDoc(doc(db, "meta", "admin_priv")); if (!s.exists()) throw new Error("Ключ адміністратора ще не створено."); return s.data(); },
    setAdminPriv: (priv) => setDoc(doc(db, "meta", "admin_priv"), { ...priv, at: nowIso() }),
    // ---------------------------------------------------------------- диспетчер: ключ і довідники авто
    async getDispPub() { const s = await getDoc(doc(db, "meta", "disp_pub")); return s.exists() ? s.data() : null; },
    async getDispPriv() { const s = await getDoc(doc(db, "meta", "disp_priv")); return s.exists() ? s.data() : null; },
    async createDispKeys(keys) {
      const b = writeBatch(db);
      b.set(doc(db, "meta", "disp_pub"), { ...keys.pub, at: nowIso() });
      b.set(doc(db, "meta", "disp_priv"), { ...keys.priv, at: nowIso() });
      await b.commit();
    },
    async resetDispKeys() { await deleteDoc(doc(db, "meta", "disp_priv")); await deleteDoc(doc(db, "meta", "disp_pub")); },
    async getFleetCat(unitId) { const s = await getDoc(doc(db, "fleet", unitId)); return s.exists() ? s.data() : null; },
    async listFleetCats() { return (await getDocs(collection(db, "fleet"))).docs.map((d) => ({ unit_id: d.id, ...d.data() })); },
    saveFleetCat: (unitId, data) => setDoc(doc(db, "fleet", unitId), { ...data, updated_at: nowIso() }),
    // перше створення ключа підрозділу — у транзакції, щоб двоє колег одночасно не створили різні ключі
    async initUnitCrypto(id, cryptoObj, encBlob, n) {
      const ref = doc(db, "units", id);
      return runTransaction(db, async (t) => {
        const s = await t.get(ref);
        if (!s.exists()) throw new Error("Підрозділ не знайдено.");
        if (s.data().crypto) { const e = new Error("Ключ цього підрозділу вже створено."); e.code = "already-init"; throw e; }
        t.update(ref, { crypto: cryptoObj, enc: encBlob, n_vehicles: n, settings: deleteField(), vehicles: deleteField() });
      });
    },
    saveUnitData: (id, encBlob, n) => updateDoc(doc(db, "units", id), { enc: encBlob, n_vehicles: n, settings: deleteField(), vehicles: deleteField() }),
    setUnitPass: (id, passObj) => updateDoc(doc(db, "units", id), { "crypto.pass": passObj }),
    updateUnit: (id, fields) => updateDoc(doc(db, "units", id), fields),
    async deleteUnit(id) {
      const ms = await getDocs(collection(db, "units", id, "months"));
      for (const d of ms.docs) await deleteDoc(d.ref);
      await deleteDoc(doc(db, "fleet", id)).catch(() => {});
      await deleteDoc(doc(db, "units", id));
    },
    // ---------------------------------------------------------------- місяці
    async getMonth(unitId, ym) { const s = await getDoc(doc(db, "units", unitId, "months", ym)); return s.exists() ? s.data() : null; },
    saveMonth: (unitId, ym, encBlob, extra = {}) => setDoc(doc(db, "units", unitId, "months", ym), { ...extra, enc: encBlob, updated_at: nowIso(),
      updated_by: auth.currentUser.uid }),
    async listMonths(unitId) { return (await getDocs(collection(db, "units", unitId, "months"))).docs.map((d) => ({ ym: d.id, ...d.data() }))
      .sort((a, b) => a.ym.localeCompare(b.ym)); },
    // ---------------------------------------------------------------- журнал
    log(action, details = "", unitId = null) {
      const u = auth.currentUser; if (!u) return Promise.resolve();
      return addDoc(collection(db, "audit"), { ts: nowIso(), uid: u.uid, email: u.email, unit_id: unitId, action, details: String(details).slice(0, 1000) }).catch(() => {});
    },
    async listAudit(n = 200) { return (await getDocs(query(collection(db, "audit"), orderBy("ts", "desc"), limit(n)))).docs.map((d) => d.data()); },
  };
}
