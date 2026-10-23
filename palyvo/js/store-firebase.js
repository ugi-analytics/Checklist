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
  browserLocalPersistence,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, collection, writeBatch, query, orderBy, limit,
  addDoc,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const nowIso = () => new Date().toLocaleString("sv-SE").replace("T", " ");

export function createStore(config) {
  const app = initializeApp(config);
  const auth = getAuth(app);
  setPersistence(auth, browserLocalPersistence).catch(() => {});
  const db = getFirestore(app);

  const errUa = (e) => {
    const c = (e && e.code) || "";
    const map = {
      "auth/invalid-credential": "Невірний e-mail або пароль.", "auth/wrong-password": "Невірний e-mail або пароль.",
      "auth/user-not-found": "Невірний e-mail або пароль.", "auth/invalid-email": "Некоректний e-mail.",
      "auth/email-already-in-use": "Цей e-mail уже зареєстровано. Увійдіть або відновіть пароль.",
      "auth/weak-password": "Пароль занадто простий (мінімум 8 символів, літери й цифри).",
      "auth/too-many-requests": "Забагато спроб. Зачекайте кілька хвилин.",
      "auth/network-request-failed": "Немає з'єднання з інтернетом.",
      "permission-denied": "Немає доступу (перевірте, чи підтверджено ваш обліковий запис).",
    };
    return map[c] || (e && e.message) || String(e);
  };

  return {
    kind: "firebase",
    errUa,
    // ---------------------------------------------------------------- вхід
    onAuth(cb) { return onAuthStateChanged(auth, (u) => cb(u ? { uid: u.uid, email: u.email } : null)); },
    currentUser() { const u = auth.currentUser; return u ? { uid: u.uid, email: u.email } : null; },
    signIn: (email, pw) => signInWithEmailAndPassword(auth, email.trim(), pw),
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
    async bootstrapAdmin(profile, unitName, unitSettings) {
      const uid = auth.currentUser.uid;
      const unitRef = doc(collection(db, "units"));
      const b = writeBatch(db);
      b.set(unitRef, { name: unitName, settings: unitSettings, vehicles: [], created_at: nowIso() });
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
    async listUsers() { return (await getDocs(collection(db, "users"))).docs.map((d) => ({ uid: d.id, ...d.data() })); },
    updateUser: (uid, fields) => updateDoc(doc(db, "users", uid), fields),
    deleteUser: (uid) => deleteDoc(doc(db, "users", uid)),
    // ---------------------------------------------------------------- підрозділи
    async listUnits() { return (await getDocs(collection(db, "units"))).docs.map((d) => ({ id: d.id, ...d.data() }))
      .sort((a, b) => a.name.localeCompare(b.name, "uk")); },
    async getUnit(id) { const s = await getDoc(doc(db, "units", id)); return s.exists() ? { id, ...s.data() } : null; },
    async createUnit(name, settings) { const r = doc(collection(db, "units")); await setDoc(r, { name, settings, vehicles: [], created_at: nowIso() }); return r.id; },
    updateUnit: (id, fields) => updateDoc(doc(db, "units", id), fields),
    async deleteUnit(id) {
      const ms = await getDocs(collection(db, "units", id, "months"));
      for (const d of ms.docs) await deleteDoc(d.ref);
      await deleteDoc(doc(db, "units", id));
    },
    // ---------------------------------------------------------------- місяці
    async getMonth(unitId, ym) { const s = await getDoc(doc(db, "units", unitId, "months", ym)); return s.exists() ? s.data() : null; },
    saveMonth: (unitId, ym, data) => setDoc(doc(db, "units", unitId, "months", ym), { ...data, updated_at: nowIso(),
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
