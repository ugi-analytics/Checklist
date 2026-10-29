// Обробка посилань із листів Firebase (підтвердження e-mail, новий пароль, скасування прив'язки телефона).
// Навмисно НІЧОГО не виконує без натискання кнопки людиною: поштові «сканери посилань» (зокрема в службовій пошті)
// автоматично відкривають посилання з листів — і без цього скасовували б прив'язку телефона до входу.
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, applyActionCode, checkActionCode, verifyPasswordResetCode, confirmPasswordReset }
  from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { FIREBASE_CONFIG } from "./config.js?v=20261029b";

const box = document.getElementById("box");
const q = new URLSearchParams(location.search);
const mode = q.get("mode"), code = q.get("oobCode");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const ttl = (ico, t) => `<h2 class="display ttl"><span class="ico">${ico}</span>${t}</h2>`;
const toPortal = `<p style="margin-top:16px"><a class="btn sky" href="https://palyvo-mvo.web.app/">Перейти до порталу →</a></p>`;
const errUa = (e) => ({
  "auth/invalid-action-code": "Посилання вже використане або недійсне. Якщо потрібно — запросіть новий лист на порталі.",
  "auth/expired-action-code": "Термін дії посилання минув. Запросіть новий лист на порталі.",
  "auth/weak-password": "Пароль занадто простий.", "auth/network-request-failed": "Немає з'єднання з інтернетом.",
}[e && e.code] || (e && e.message) || String(e));
const show = (html) => { box.innerHTML = html; };
const fail = (e) => show(ttl("⚠️", "Не вдалося") + `<div class="flash err">${esc(errUa(e))}</div>` + toPortal);

const auth = getAuth(initializeApp(FIREBASE_CONFIG));
auth.languageCode = "uk";

function confirmStep(icon, title, text, btn, run) {
  show(ttl(icon, title) + `<div class="note">${text}</div><form class="js-go"><button class="btn primary">${btn}</button></form>` + toPortal);
  box.querySelector(".js-go").addEventListener("submit", async (e) => {
    e.preventDefault(); e.target.querySelector("button").disabled = true;
    try { await run(); } catch (err) { fail(err); }
  });
}

if (!mode || !code) {
  show(ttl("🔗", "Посилання неповне") + `<p class="muted">Відкрийте посилання з листа повністю (скопіюйте його цілком у адресний рядок).</p>` + toPortal);
} else if (mode === "verifyEmail") {
  confirmStep("✉️", "Підтвердження e-mail", "Натисніть кнопку, щоб підтвердити свою адресу електронної пошти для порталу.", "Підтвердити e-mail",
    async () => { await applyActionCode(auth, code);
      show(ttl("✅", "E-mail підтверджено") + `<p>Поверніться на вкладку порталу й натисніть «Я підтвердив» (або оновіть сторінку).</p>` + toPortal); });
} else if (mode === "resetPassword") {
  verifyPasswordResetCode(auth, code).then((email) => {
    show(ttl("🔒", "Новий пароль") + `<p class="muted">Обліковий запис: <b>${esc(email)}</b></p>
      <form class="js-pw"><label class="f">Новий пароль (від 8 символів, літери й цифри)<input name="p1" type="password" required minlength="8" autocomplete="new-password"></label>
      <label class="f">Повторіть пароль<input name="p2" type="password" required minlength="8" autocomplete="new-password"></label>
      <div class="js-err"></div><button class="btn primary">Зберегти пароль</button></form>`);
    box.querySelector(".js-pw").addEventListener("submit", async (e) => {
      e.preventDefault(); const f = e.target; const err = box.querySelector(".js-err");
      if (f.p1.value !== f.p2.value) { err.innerHTML = `<div class="flash err">Паролі не збігаються.</div>`; return; }
      if (/^\d+$/.test(f.p1.value) || /^[^\d]+$/.test(f.p1.value)) { err.innerHTML = `<div class="flash err">Пароль має містити і літери, і цифри.</div>`; return; }
      try { await confirmPasswordReset(auth, code, f.p1.value); show(ttl("✅", "Пароль змінено") + `<p>Увійдіть на порталі з новим паролем (далі — код із телефона).</p>` + toPortal); }
      catch (e2) { fail(e2); }
    });
  }).catch(fail);
} else if (mode === "revertSecondFactorAddition") {
  confirmStep("🛡️", "Скасувати прив'язку телефона?",
    `Цей лист надсилається щоразу, коли до облікового запису прив'язують телефон для входу з кодом.<br>
     <b>Якщо телефон прив'язували ви самі — нічого не робіть</b>, просто закрийте цю сторінку.<br>
     Натискайте кнопку нижче лише якщо прив'язку зробив хтось сторонній: телефон буде відв'язано, і варто одразу змінити пароль.`,
    "Так, скасувати прив'язку — це був не я",
    async () => { await checkActionCode(auth, code); await applyActionCode(auth, code);
      show(ttl("✅", "Прив'язку скасовано") + `<div class="note">Змініть пароль (кнопка «Забули пароль?» на сторінці входу) і повідомте адміністратора порталу.</div>` + toPortal); });
} else if (mode === "recoverEmail") {
  confirmStep("✉️", "Повернути попередній e-mail?", "Натисніть кнопку, лише якщо зміну адреси пошти робили не ви.", "Так, повернути",
    async () => { await applyActionCode(auth, code); show(ttl("✅", "Адресу повернуто") + `<p>Рекомендуємо змінити пароль.</p>` + toPortal); });
} else {
  show(ttl("🔗", "Невідома дія") + `<p class="muted">Це посилання не підтримується порталом.</p>` + toPortal);
}
