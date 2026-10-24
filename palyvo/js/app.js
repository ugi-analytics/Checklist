// Портал «Списання пального для МВО» — односторінковий застосунок (GitHub Pages + Firebase).
import * as C from "./core.js?v=20261024a";
import * as U from "./ui.js?v=20261024a";
import { litresWords } from "./words.js?v=20261024a";
import { buildRegister, buildAct } from "./docs.js?v=20261024a";
import { FIREBASE_CONFIG, ADMIN_EMAIL } from "./config.js?v=20261024a";
import * as X from "./crypto.js?v=20261024a";
import { qrSvg } from "../vendor/qrcode.js?v=20261024a";

const $app = document.getElementById("app");
const esc = U.esc;
const params = new URLSearchParams(location.search);
const DEMO = params.has("demo");
let store = null;

const state = { authUser: null, me: null, unitId: null, units: [], flash: null, setupDone: null,
  mfa: null, adminPriv: null, keys: {}, adminKeyReady: null };
// ключі в пам'яті вкладки скидаються при виході
function forgetSessionKeys() { state.keys = {}; state.adminPriv = null; }
const keyCacheId = (unitId) => `${DEMO ? "demo" : FIREBASE_CONFIG.projectId}:${state.authUser?.uid}:${unitId}`;

// ============================================================ утиліти
function flash(type, text) { state.flash = [type, text]; }
function go(hash) { if (location.hash === hash) render(); else location.hash = hash; }
function route() {
  const h = location.hash.replace(/^#\/?/, "");
  const [path, qs] = h.split("?");
  return { path: path || "", q: new URLSearchParams(qs || "") };
}
function busy(on) { document.body.classList.toggle("busy", !!on); }
async function download(doc, filename) {
  const blob = await window.docx.Packer.toBlob(doc);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
function pwProblem(pw) {
  if (pw.length < 8) return "Пароль має містити щонайменше 8 символів.";
  if (/^\d+$/.test(pw) || /^[^\d]+$/.test(pw)) return "Пароль має містити і літери, і цифри.";
  return null;
}
const isAdmin = () => state.me && state.me.role === "admin" && state.me.status === "active";
const years = () => { const y = new Date().getFullYear(); const a = []; for (let i = 2024; i <= y + 1; i++) a.push(i); return a; };
function monthPicker(ym, base) {
  const [y, m] = C.parseYm(ym);
  return `<form class="row js-month" data-base="${base}">
    <label class="f">Звітний місяць<select name="m">${C.MONTHS_UA.map((n, i) => `<option value="${i + 1}" ${i + 1 === m ? "selected" : ""}>${n[0].toUpperCase() + n.slice(1)}</option>`).join("")}</select></label>
    <label class="f">Рік<select name="y">${years().map((yy) => `<option ${yy === y ? "selected" : ""}>${yy}</option>`).join("")}</select></label></form>`;
}
function bindMonthPicker() {
  document.querySelectorAll(".js-month").forEach((f) => f.addEventListener("change", () => {
    go(`${f.dataset.base}?ym=${C.ymKey(f.y.value, f.m.value)}`);
  }));
}
function currentYm(q) {
  const ym = q.get("ym");
  if (ym && /^\d{4}-\d{2}$/.test(ym)) { sessionStorage.setItem("ym", ym); return ym; }
  return sessionStorage.getItem("ym") || C.defaultMonth();
}

// ============================================================ каркас
function shell(content, { nav = true } = {}) {
  const me = state.me; const p = route().path;
  const unit = state.units.find((u) => u.id === state.unitId);
  const unitChip = "";  // назву підрозділу в шапці не показуємо (за проханням); адміністратор перемикає підрозділ списком нижче
  const switcher = isAdmin() && state.units.length > 1 ? `<select class="js-unit-switch" title="Працювати з підрозділом">
      ${state.units.map((u) => `<option value="${u.id}" ${u.id === state.unitId ? "selected" : ""}>${esc(u.name)}</option>`).join("")}</select>` : "";
  const tab = (h, icon, label, extra = "") => `<a href="#/${h}" class="tile ${p === h || p.startsWith(h + "/") || (h === "entry" && !p) ? "on" : ""}">
    <span class="ti">${icon}</span><span class="tl">${label}${extra}</span><span class="chev">›</span></a>`;
  const pending = isAdmin() && state.pendingCount ? `<span class="badge">${state.pendingCount}</span>` : "";
  content = String(content).replace(/<h2 class="display">((?:\p{Extended_Pictographic}|[✓✍])\uFE0F?)\s*/gu, '<h2 class="display ttl"><span class="ico">$1</span>');
  $app.innerHTML = `
  <header class="hero"><div class="wrap">
    <div class="topbar"><div class="logo"><img src="img/hub_icon_small.png" alt=""></div>
      <div class="brand"><div class="bname"><span class="display">UGI-Analytics</span>
        <span class="online ${navigator.onLine ? "" : "off"}"><i></i><b>${navigator.onLine ? "ON-LINE" : "OFF-LINE"}</b></span></div>
        <small>Управління головної інспекції ГУНП в Полтавській області</small></div>
      <div class="sp"></div>
      ${state.authUser ? `<div class="who"><b>${esc(me?.full_name || state.authUser.email)}</b>${isAdmin() ? " · адміністратор" : ""}<br>
        ${me && me.status === "active" ? `<a href="#/password">змінити пароль</a> · ` : ""}<a href="#" class="js-logout">вийти</a></div>` : ""}
    </div>
    <div class="title-row"><h1 class="display"><span>Списання пального</span> <span>для <em>МВО</em></span></h1>
      <div class="chips">${unitChip}${switcher}
      ${DEMO ? `<span class="chip" style="border-color:#f87171;color:#fecaca">ДЕМО-РЕЖИМ · дані не зберігаються</span>` : ""}</div></div>
    ${nav && me && me.status === "active" ? `<nav class="tiles">
      ${tab("entry", "💧", "Введення даних")}${tab("report", "📋", "Звіт і акти")}${tab("fleet", "🚙", "Автопарк і реквізити")}${tab("archive", "📁", "Архів")}
      ${isAdmin() ? tab("admin", "🛡️", "Адміністрування", pending) : ""}</nav>` : ""}
    <div class="motto">Контроль <i>/</i> Прозорість <i>/</i> Ефективність</div>
  </div></header>
  <main class="wrap">${state.flash ? `<div class="flash ${state.flash[0]}">${esc(state.flash[1])}</div>` : ""}${content}</main>`;
  state.flash = null;
  $app.querySelector(".js-logout")?.addEventListener("click", async (e) => { e.preventDefault(); forgetSessionKeys(); await store.signOut(); go("#/login"); });
  $app.querySelector(".js-unit-switch")?.addEventListener("change", (e) => {
    state.unitId = e.target.value; sessionStorage.setItem("unitCtx", state.unitId); render();
  });
}
function authCard(inner, width = 460) { return `<div class="auth" style="max-width:${width}px"><div class="panel">${inner}</div></div>`; }

// ============================================================ сторінки входу
function pageNotConfigured() {
  shell(authCard(`<h2 class="display">⚙️ Портал ще не налаштовано</h2>
    <p>Не заповнено параметри Firebase у файлі <code>js/config.js</code>.</p>
    <p class="muted">Для перегляду інтерфейсу без бази відкрийте <a href="?demo=1">демо-режим</a>.</p>`), { nav: false });
}

function pageLogin() {
  shell(authCard(`<h2 class="display">🔑 Вхід до порталу</h2>
    <form class="js-login" autocomplete="on">
      <label class="f">E-mail<input name="email" type="email" required autofocus autocomplete="username"></label>
      <label class="f">Пароль<input name="pw" type="password" required autocomplete="current-password"></label>
      <button class="btn primary">Увійти</button>
    </form>
    <p class="muted" style="margin-top:16px"><a href="#/forgot">Забули пароль?</a></p>
    <p class="muted">Немає облікового запису? <a href="#/register">Подати заявку на реєстрацію</a> — після підтвердження адміністратором ви зможете працювати.</p>`), { nav: false });
  $app.querySelector(".js-login").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target; busy(true);
    try { await store.signIn(f.email.value, f.pw.value); }
    catch (err) { if (err.code === "mfa-required") return pageMfaCode(); flash("err", store.errUa(err)); pageLogin(); }
    finally { busy(false); }
  });
}

// другий крок входу — код із телефона
function pageMfaCode() {
  shell(authCard(`<h2 class="display">🔐 Код із телефона</h2>
    <p class="muted">Відкрийте на телефоні застосунок-автентифікатор (Google Authenticator / Microsoft Authenticator) і введіть
      6-значний код для «Palyvo MVO».</p>
    <form class="js-mfa" autocomplete="off"><label class="f">Код<input name="code" inputmode="numeric" pattern="[0-9 ]{6,7}" maxlength="7" required autofocus
      autocomplete="one-time-code" style="font-size:24px;letter-spacing:6px;text-align:center"></label>
      <button class="btn primary">Увійти</button></form>
    <p class="muted" style="margin-top:14px"><a href="#/login" class="js-back">← Інший обліковий запис</a></p>`), { nav: false });
  $app.querySelector(".js-mfa").addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try { await store.finishMfaSignIn(e.target.code.value); }
    catch (err) { flash("err", store.errUa(err)); if (String(err.message).includes("застаріла")) pageLogin(); else pageMfaCode(); }
    finally { busy(false); }
  });
  $app.querySelector(".js-back").addEventListener("click", (e) => { e.preventDefault(); pageLogin(); });
}

function pageForgot() {
  shell(authCard(`<h2 class="display">✉️ Відновлення пароля</h2>
    <p class="muted">Вкажіть e-mail, з яким реєструвалися, — надішлемо лист із посиланням для встановлення нового пароля.</p>
    <form class="js-forgot"><label class="f">E-mail<input name="email" type="email" required></label>
      <button class="btn primary">Надіслати лист</button></form>
    <p class="muted" style="margin-top:14px"><a href="#/login">← До входу</a></p>`), { nav: false });
  $app.querySelector(".js-forgot").addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try { await store.resetPassword(e.target.email.value); flash("ok", "Якщо такий e-mail зареєстровано — лист надіслано. Перевірте також «Спам»."); go("#/login"); }
    catch (err) { flash("err", store.errUa(err)); pageForgot(); } finally { busy(false); }
  });
}

function profileFields(p = {}) {
  return `<label class="f">Підрозділ *<input name="unit" required value="${esc(p.requested_unit || "")}" placeholder="напр. Миргородський РВП ГУНП"></label>
    <label class="f">Прізвище, ім'я, по батькові *<input name="full_name" required value="${esc(p.full_name || "")}"></label>
    <div class="two"><label class="f">Посада *<input name="position" required value="${esc(p.position || "")}"></label>
      <label class="f">Телефон<input name="phone" value="${esc(p.phone || "")}"></label></div>`;
}
const readProfile = (f) => ({ requested_unit: f.unit.value.trim(), full_name: f.full_name.value.trim(), position: f.position.value.trim(), phone: f.phone.value.trim() });

function pageRegister() {
  shell(authCard(`<h2 class="display">📝 Заявка на реєстрацію</h2>
    <p class="muted">Після підтвердження адміністратором ви отримаєте доступ до шаблону свого підрозділу.</p>
    <form class="js-reg" autocomplete="off">${profileFields()}
      <label class="f">E-mail * (буде логіном; на нього приходить лист для відновлення пароля)<input name="email" type="email" required autocomplete="username"></label>
      <div class="two"><label class="f">Пароль * (від 8 символів, літери й цифри)<input name="pw" type="password" required minlength="8" autocomplete="new-password"></label>
        <label class="f">Повторіть пароль *<input name="pw2" type="password" required minlength="8" autocomplete="new-password"></label></div>
      <button class="btn primary">Надіслати заявку</button></form>
    <p class="muted" style="margin-top:14px"><a href="#/login">← Повернутися до входу</a></p>`, 620), { nav: false });
  $app.querySelector(".js-reg").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.pw.value !== f.pw2.value) { flash("err", "Паролі не збігаються."); return pageRegister(); }
    const pp = pwProblem(f.pw.value); if (pp) { flash("err", pp); return pageRegister(); }
    const prof = readProfile(f); busy(true);
    try {
      state.registering = true;
      await store.signUp(f.email.value, f.pw.value);
      await store.createMyProfile(prof);
      await store.sendVerify().catch(() => {});
      state.registering = false;
      await onAuthChanged(store.currentUser());
    } catch (err) { state.registering = false; flash("err", store.errUa(err)); pageRegister(); } finally { busy(false); }
  });
}

function pagePending() {
  shell(authCard(`<h2 class="display">⏳ Заявку надіслано</h2>
    <p>Адміністратор перевірить її та прив'яже ваш обліковий запис до підрозділу <b>${esc(state.me.requested_unit)}</b>.
      Після підтвердження оновіть сторінку — відкриється робоча частина порталу.</p>
    ${state.mfa && !state.mfa.emailVerified ? `<div class="flash info">✉️ Ми надіслали лист на <b>${esc(state.authUser.email)}</b> — перейдіть за посиланням у ньому,
      щоб підтвердити e-mail (перевірте й «Спам»). Без цього не вдасться налаштувати вхід із кодом із телефона.
      <br><button class="btn sm js-resend" style="margin-top:8px">Надіслати лист ще раз</button></div>` : ""}
    <button class="btn sky js-reload">🔄 Перевірити ще раз</button>`), { nav: false });
  $app.querySelector(".js-reload").addEventListener("click", async () => { await store.reloadUser().catch(() => {}); onAuthChanged(store.currentUser()); });
  $app.querySelector(".js-resend")?.addEventListener("click", async () => {
    try { await store.sendVerify(); flash("ok", "Лист надіслано ще раз."); } catch (err) { flash("err", store.errUa(err)); } pagePending();
  });
}

function pageBlocked() {
  shell(authCard(`<h2 class="display">⛔ Доступ закрито</h2>
    <p>Обліковий запис ${state.me.status === "rejected" ? "відхилено" : "заблоковано"}. Зверніться до адміністратора порталу.</p>`), { nav: false });
}

function pageNoProfile() {
  shell(authCard(`<h2 class="display">👤 Профіль не знайдено</h2>
    <p class="muted">Ваш обліковий запис видалено або заявку ще не створено. Можете подати заявку повторно.</p>
    <form class="js-reapply">${profileFields()}<button class="btn primary">Подати заявку</button></form>`, 620), { nav: false });
  $app.querySelector(".js-reapply").addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try { await store.createMyProfile(readProfile(e.target)); await onAuthChanged(store.currentUser()); }
    catch (err) { flash("err", store.errUa(err)); pageNoProfile(); } finally { busy(false); }
  });
}

// Перший запуск: 1) обліковий запис адміністратора → 2) підтвердження e-mail → 3) підрозділ і ключ адміністратора
function pageSetup() {
  const lockedEmail = ADMIN_EMAIL && !DEMO ? ADMIN_EMAIL : "";
  const steps = (n) => `<p class="muted">Крок ${n} з 3 · далі — вхід із кодом із телефона</p>`;
  if (!state.authUser) {
    shell(authCard(`<h2 class="display">🚀 Перший запуск порталу</h2>${steps(1)}
      <p class="muted">Створіть обліковий запис адміністратора області${lockedEmail ? ` (лише для <b>${esc(lockedEmail)}</b>)` : ""}.</p>
      <form class="js-setup1" autocomplete="off">
        <label class="f">E-mail адміністратора<input name="email" type="email" required value="${esc(lockedEmail)}" ${lockedEmail ? "readonly" : ""}></label>
        <div class="two"><label class="f">Пароль (від 8 символів, літери й цифри)<input name="pw" type="password" required minlength="8"></label>
          <label class="f">Повторіть пароль<input name="pw2" type="password" required minlength="8"></label></div>
        <button class="btn primary">Далі</button></form>`, 640), { nav: false });
    $app.querySelector(".js-setup1").addEventListener("submit", async (e) => {
      e.preventDefault(); const f = e.target;
      if (f.pw.value !== f.pw2.value) { flash("err", "Паролі не збігаються."); return pageSetup(); }
      const pp = pwProblem(f.pw.value); if (pp) { flash("err", pp); return pageSetup(); }
      busy(true);
      try {
        try { await store.signUp(f.email.value, f.pw.value); await store.sendVerify().catch(() => {}); }
        catch (err) { if (err.code === "auth/email-already-in-use") await store.signIn(f.email.value, f.pw.value); else throw err; }
      } catch (err) { if (err.code === "mfa-required") return pageMfaCode(); flash("err", store.errUa(err)); pageSetup(); } finally { busy(false); }
    });
    return;
  }
  if (!state.mfa || !state.mfa.emailVerified) {
    shell(authCard(`<h2 class="display">✉️ Підтвердіть e-mail</h2>${steps(2)}
      <p>Ми надіслали лист на <b>${esc(state.authUser.email)}</b>. Відкрийте його та перейдіть за посиланням (перевірте й «Спам»),
        потім поверніться сюди й натисніть «Я підтвердив».</p>
      <div class="row"><button class="btn primary js-done">✓ Я підтвердив</button><button class="btn js-resend">Надіслати лист ще раз</button></div>
      <p class="muted" style="margin-top:14px"><a href="#" class="js-out">Вийти</a></p>`, 560), { nav: false });
    $app.querySelector(".js-done").addEventListener("click", async () => {
      busy(true);
      try { const ok = await store.reloadUser(); state.mfa = await store.mfaState(); if (!ok) flash("err", "E-mail ще не підтверджено — перейдіть за посиланням у листі."); }
      catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
      pageSetup();
    });
    $app.querySelector(".js-resend").addEventListener("click", async () => {
      try { await store.sendVerify(); flash("ok", "Лист надіслано ще раз."); } catch (err) { flash("err", store.errUa(err)); } pageSetup();
    });
    $app.querySelector(".js-out").addEventListener("click", async (e) => { e.preventDefault(); await store.signOut(); });
    return;
  }
  shell(authCard(`<h2 class="display">🛡️ Підрозділ і ключ адміністратора</h2>${steps(3)}
    <form class="js-setup" autocomplete="off">
      <label class="f">Ваш підрозділ<input name="unit" required value="УГІ ГУНП"></label>
      <div class="two"><label class="f">ПІБ<input name="full_name" required></label><label class="f">Посада<input name="position"></label></div>
      <div class="note">🔐 <b>Ключ адміністратора</b> — окрема фраза (не пароль від входу). Нею ви відкриваєте зведення
        по області й відновлюєте ключі підрозділів, якщо ті їх забудуть. Її не знає ніхто, крім вас, і її <b>неможливо відновити</b>:
        запишіть на папері й зберігайте як ключ від сейфа.</div>
      <div class="two"><label class="f">Ключ адміністратора (від 10 символів, краще фраза)<input name="ak" type="password" required minlength="10"></label>
        <label class="f">Повторіть ключ адміністратора<input name="ak2" type="password" required minlength="10"></label></div>
      <label style="display:flex;gap:8px;align-items:center;margin:10px 0"><input type="checkbox" name="ok" required> Я записав ключ адміністратора й зберігатиму його окремо від пароля</label>
      <button class="btn primary">Створити портал</button></form>`, 680), { nav: false });
  $app.querySelector(".js-setup").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.ak.value !== f.ak2.value) { flash("err", "Ключі адміністратора не збігаються."); return pageSetup(); }
    const kp = X.passProblem(f.ak.value, "Ключ адміністратора"); if (kp) { flash("err", kp); return pageSetup(); }
    busy(true);
    try {
      const unitName = f.unit.value.trim();
      const settings = { ...C.DEFAULT_SETTINGS, dept: unitName, unit: unitName, act_owner: unitName, compiler: C.shortName(f.full_name.value) };
      const keys = await X.newAdminKeys(f.ak.value);
      await store.reloadUser().catch(() => {});
      await store.bootstrapAdmin({ full_name: f.full_name.value.trim(), position: f.position.value.trim(), phone: "", requested_unit: unitName },
        unitName, settings, keys);
      state.setupDone = true;
      flash("ok", "Портал створено. Залишилось налаштувати вхід із кодом із телефона.");
      await onAuthChanged(store.currentUser());
    } catch (err) { flash("err", store.errUa(err)); pageSetup(); } finally { busy(false); }
  });
}

// Обов'язковий двофакторний вхід: підтвердження e-mail → прив'язка застосунку-автентифікатора на телефоні
function pageMfaSetup(stage = {}) {
  const m = state.mfa || {};
  const out = `<p class="muted" style="margin-top:14px"><a href="#" class="js-out">Вийти</a></p>`;
  const bindOut = () => $app.querySelector(".js-out")?.addEventListener("click", async (e) => { e.preventDefault(); forgetSessionKeys(); await store.signOut(); });
  if (!m.emailVerified) {
    shell(authCard(`<h2 class="display">🔐 Захищений вхід · крок 1 з 2</h2>
      <p>Для роботи з порталом потрібен вхід із кодом із телефона. Спочатку підтвердіть e-mail <b>${esc(state.authUser.email)}</b>:
        натисніть «Надіслати лист», відкрийте лист і перейдіть за посиланням, потім натисніть «Я підтвердив».</p>
      <div class="row"><button class="btn sky js-send">✉️ Надіслати лист</button><button class="btn primary js-done">✓ Я підтвердив</button></div>${out}`, 560), { nav: false });
    $app.querySelector(".js-send").addEventListener("click", async () => {
      try { await store.sendVerify(); flash("ok", "Лист надіслано. Перевірте також «Спам»."); } catch (err) { flash("err", store.errUa(err)); } pageMfaSetup();
    });
    $app.querySelector(".js-done").addEventListener("click", async () => {
      busy(true);
      try { const ok = await store.reloadUser(); state.mfa = await store.mfaState(); if (!ok) flash("err", "E-mail ще не підтверджено — перейдіть за посиланням у листі."); }
      catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
      render();
    });
    return bindOut();
  }
  if (m.enrolled && !m.session) {
    shell(authCard(`<h2 class="display">🔐 Увійдіть із кодом</h2>
      <p>Вхід із кодом із телефона для цього облікового запису вже налаштовано. Увійдіть ще раз — після пароля портал попросить 6-значний код.</p>
      <button class="btn primary js-relogin">Увійти ще раз</button>`, 520), { nav: false });
    $app.querySelector(".js-relogin").addEventListener("click", async () => { forgetSessionKeys(); await store.signOut(); });
    return;
  }
  if (stage.needPw) {
    shell(authCard(`<h2 class="display">🔐 Підтвердіть пароль</h2>
      <p class="muted">З міркувань безпеки перед налаштуванням ще раз введіть пароль від порталу.</p>
      <form class="js-re"><label class="f">Пароль<input name="pw" type="password" required autofocus></label><button class="btn primary">Далі</button></form>${out}`), { nav: false });
    $app.querySelector(".js-re").addEventListener("submit", async (e) => {
      e.preventDefault(); busy(true);
      try { await store.reauth(e.target.pw.value); pageMfaSetup({ start: true }); } catch (err) { flash("err", store.errUa(err)); pageMfaSetup({ needPw: true }); } finally { busy(false); }
    });
    return bindOut();
  }
  if (!stage.secret) {
    shell(authCard(`<h2 class="display">🔐 Захищений вхід · крок 2 з 2</h2>
      <p>Встановіть на службовий або особистий смартфон безкоштовний застосунок <b>Google Authenticator</b> або <b>Microsoft Authenticator</b>
        (Google Play / App Store). Він щохвилини показує новий 6-значний код, який портал питатиме після пароля.</p>
      <p class="muted">Інтернет на телефоні для кодів не потрібен.</p>
      <button class="btn primary js-start">Застосунок встановлено — показати QR-код</button>${out}`, 560), { nav: false });
    const start = async () => {
      busy(true);
      try { const sec = await store.startTotp(); pageMfaSetup({ secret: sec }); }
      catch (err) { if (err.code === "auth/requires-recent-login") return pageMfaSetup({ needPw: true }); flash("err", store.errUa(err)); pageMfaSetup(); }
      finally { busy(false); }
    };
    $app.querySelector(".js-start").addEventListener("click", start);
    if (stage.start) start();
    return bindOut();
  }
  const keyGroups = String(stage.secret.key).replace(/(.{4})/g, "$1 ").trim();
  shell(authCard(`<h2 class="display">🔐 Прив'яжіть телефон</h2>
    <ol style="line-height:1.7;padding-left:20px"><li>У застосунку натисніть «+» → «Сканувати QR-код» і наведіть камеру на код нижче.</li>
      <li>Введіть 6 цифр, які з'явилися в застосунку, і натисніть «Підтвердити».</li></ol>
    <div style="text-align:center;margin:12px 0"><div style="display:inline-block;padding:8px;background:#fff;border-radius:12px">${qrSvg(stage.secret.uri, 5)}</div></div>
    <p class="muted" style="text-align:center">Не сканується? Додайте вручну — назва «Palyvo MVO», ключ:<br><code style="font-size:15px;user-select:all">${esc(keyGroups)}</code></p>
    <form class="js-enroll" autocomplete="off"><label class="f">Код із застосунку<input name="code" inputmode="numeric" maxlength="7" required autofocus
      autocomplete="one-time-code" style="font-size:24px;letter-spacing:6px;text-align:center"></label>
      <button class="btn primary">Підтвердити</button></form>${out}`, 560), { nav: false });
  $app.querySelector(".js-enroll").addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try {
      await store.finishTotp(e.target.code.value);
      state.mfa = await store.mfaState();
      store.log("mfa_enroll", "");
      if (state.mfa.session) { flash("ok", "Вхід із кодом налаштовано ✓"); await onAuthChanged(store.currentUser()); }
      else { flash("ok", "Вхід із кодом налаштовано ✓ Увійдіть ще раз: пароль, потім код із телефона."); forgetSessionKeys(); await store.signOut(); }
    } catch (err) { flash("err", store.errUa(err)); pageMfaSetup(stage); } finally { busy(false); }
  });
  bindOut();
}

// База ще працює за старими правилами доступу (без розділу для ключа адміністратора)
function pageRulesOld() {
  shell(authCard(`<h2 class="display">⚠️ Правила бази не оновлено</h2>
    <p>Вхід із кодом працює, але база Firestore ще працює за <b>старими правилами доступу</b>, тому портал не може створити ключ адміністратора.</p>
    <ol style="line-height:1.7;padding-left:20px"><li>Відкрийте console.firebase.google.com → проєкт <b>palyvo-mvo</b> → <b>Firestore Database</b> → вкладка <b>Rules</b>.</li>
      <li>Видаліть увесь текст і вставте вміст файлу <b>Правила_Firestore.txt</b> (найновішого).</li>
      <li>Натисніть <b>Publish</b>, зачекайте хвилину й натисніть кнопку нижче.</li></ol>
    <button class="btn primary js-recheck">🔄 Перевірити ще раз</button>`, 620));
  $app.querySelector(".js-recheck").addEventListener("click", () => onAuthChanged(store.currentUser()));
}

// Ключ адміністратора, якщо портал створено ще версією без шифрування
function pageAdminKeyCreate() {
  shell(authCard(`<h2 class="display">🛡️ Створіть ключ адміністратора</h2>
    <p>Портал переведено на шифрування даних. Залишилось створити <b>ключ адміністратора</b> — окрему фразу (не пароль від входу).
      Нею ви відкриваєте зведення по області й встановлюєте новий ключ підрозділу, якщо той його забуде.</p>
    <div class="note">Ключ адміністратора знаєте лише ви, і його <mark>неможливо відновити</mark>: запишіть на папері
      й зберігайте окремо від пароля, як ключ від сейфа.</div>
    <form class="js-akcreate" autocomplete="off">
      <div class="two"><label class="f">Ключ адміністратора (від 10 символів, краще фраза)<input name="ak" type="password" required minlength="10"></label>
        <label class="f">Повторіть ключ<input name="ak2" type="password" required minlength="10"></label></div>
      <label style="display:flex;gap:8px;align-items:center;margin:10px 0"><input type="checkbox" name="ok" required> Я записав ключ адміністратора й зберігатиму його окремо від пароля</label>
      <button class="btn primary">Створити ключ</button></form>`, 640));
  $app.querySelector(".js-akcreate").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.ak.value !== f.ak2.value) { flash("err", "Ключі не збігаються."); return pageAdminKeyCreate(); }
    const kp = X.passProblem(f.ak.value, "Ключ адміністратора"); if (kp) { flash("err", kp); return pageAdminKeyCreate(); }
    busy(true);
    try {
      await store.createAdminKeys(await X.newAdminKeys(f.ak.value));
      state.adminKeyReady = true; store.log("admin_key_create", "");
      flash("ok", "Ключ адміністратора створено. Тепер створіть ключ свого підрозділу.");
      render();
    } catch (err) { flash("err", store.errUa(err)); pageAdminKeyCreate(); } finally { busy(false); }
  });
}

function pagePassword() {
  shell(authCard(`<h2 class="display">🔒 Зміна пароля</h2>
    <form class="js-pw"><label class="f">Поточний пароль<input name="old" type="password" required></label>
      <label class="f">Новий пароль (від 8 символів, літери й цифри)<input name="n1" type="password" required minlength="8"></label>
      <label class="f">Повторіть новий пароль<input name="n2" type="password" required minlength="8"></label>
      <button class="btn primary">Зберегти</button></form>
    ${isAdmin() ? `<hr style="border-color:rgba(148,163,184,.25);margin:22px 0"><h2 class="display">🛡️ Зміна ключа адміністратора</h2>
      <form class="js-akchange" autocomplete="off"><label class="f">Поточний ключ адміністратора<input name="old" type="password" required></label>
        <label class="f">Новий ключ (від 10 символів)<input name="n1" type="password" required minlength="10"></label>
        <label class="f">Повторіть новий ключ<input name="n2" type="password" required minlength="10"></label>
        <button class="btn sky">Змінити ключ адміністратора</button></form>` : ""}`));
  $app.querySelector(".js-akchange")?.addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.n1.value !== f.n2.value) { flash("err", "Нові ключі не збігаються."); return pagePassword(); }
    const kp = X.passProblem(f.n1.value, "Ключ адміністратора"); if (kp) { flash("err", kp); return pagePassword(); }
    busy(true);
    try { await store.setAdminPriv(await X.rewrapAdmin(await store.getAdminPriv(), f.old.value, f.n1.value)); store.log("admin_key_change", ""); flash("ok", "Ключ адміністратора змінено. Запишіть новий ключ."); go("#/entry"); }
    catch (err) { flash("err", store.errUa(err)); pagePassword(); } finally { busy(false); }
  });
  $app.querySelector(".js-pw").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.n1.value !== f.n2.value) { flash("err", "Нові паролі не збігаються."); return pagePassword(); }
    const pp = pwProblem(f.n1.value); if (pp) { flash("err", pp); return pagePassword(); }
    busy(true);
    try { await store.changePassword(f.old.value, f.n1.value); flash("ok", "Пароль змінено."); go("#/entry"); }
    catch (err) { flash("err", store.errUa(err)); pagePassword(); } finally { busy(false); }
  });
}

// ============================================================ дані підрозділу (з розшифруванням)
class NeedKey extends Error { constructor(mode, unit) { super("need-key"); this.code = "need-key"; this.mode = mode; this.unit = unit; } }

// ключ даних підрозділу: з пам'яті вкладки → з цього комп'ютера → ключем адміністратора (якщо відкрито)
async function getUnitKey(u) {
  if (state.keys[u.id]) return state.keys[u.id];
  const cached = await X.keyCache.get(keyCacheId(u.id));
  if (cached) { state.keys[u.id] = cached; return cached; }
  if (isAdmin() && state.adminPriv && u.crypto && u.crypto.admin) {
    const dk = await X.unwrapByAdmin(u.crypto.admin, state.adminPriv);
    state.keys[u.id] = dk; return dk;
  }
  return null;
}
async function decUnit(u, dk) {
  const data = u.enc ? await X.decJson(dk, u.enc, X.aadUnit(u.id)) : { settings: u.settings || {}, vehicles: [] };
  u.settings = { ...C.DEFAULT_SETTINGS, ...(data.settings || {}) };
  u.vehicles = (data.vehicles || []).map(C.normalizeVehicle);
  u.dk = dk;
  return u;
}
async function loadUnit(id = state.unitId) {
  const u = await store.getUnit(id);
  if (!u) return null;
  if (!u.crypto) throw new NeedKey("init", u);
  const dk = await getUnitKey(u);
  if (!dk) throw new NeedKey("unlock", u);
  return decUnit(u, dk);
}
async function readMonth(unit, ym) {
  const m = await store.getMonth(unit.id, ym);
  if (!m) return null;
  return m.enc ? { ...m, ...(await X.decJson(unit.dk, m.enc, X.aadMonth(unit.id, ym))) } : m;
}
async function writeMonth(unit, ym, data) {
  await store.saveMonth(unit.id, ym, await X.encJson(unit.dk, data, X.aadMonth(unit.id, ym)));
}
async function listMonthsDec(unit) {
  const ms = await store.listMonths(unit.id);
  return Promise.all(ms.map(async (m) => (m.enc ? { ...m, ...(await X.decJson(unit.dk, m.enc, X.aadMonth(unit.id, m.ym))) } : m)));
}
async function writeUnit(unit, settings, vehicles) {
  await store.saveUnitData(unit.id, await X.encJson(unit.dk, { settings, vehicles }, X.aadUnit(unit.id)), vehicles.filter((v) => v.active !== false).length);
}
async function rememberKey(unitId, dk, remember) {
  const locked = await X.lockDown(dk);
  state.keys[unitId] = locked;
  if (remember) await X.keyCache.put(keyCacheId(unitId), locked);
}
async function monthState(unit, ym) {
  const saved = await readMonth(unit, ym);
  let entries, carried = null;
  if (saved) entries = saved.entries || {};
  else {
    const prev = await readMonth(unit, C.prevYm(ym));
    entries = prev ? C.carryOver(unit.vehicles, { entries: prev.entries || {}, snapshot: prev.snapshot || {} }) : {};
    if (prev) carried = C.prevYm(ym);
  }
  const vehicles = C.monthVehicles(unit.vehicles, { entries, snapshot: (saved && saved.snapshot) || {} });
  return { saved, entries, carried, vehicles };
}
async function rowsFor(unit, ym) {
  const st = await monthState(unit, ym);
  return st.vehicles.map((v) => C.calcRow(v, st.entries[v.id]));
}

// ------------------------------------------------------------ ключ підрозділу: створення / введення
const rememberBox = `<label style="display:flex;gap:8px;align-items:center;margin:10px 0"><input type="checkbox" name="remember" checked>
  Запам'ятати ключ на цьому комп'ютері (не ставте на чужому чи спільному комп'ютері)</label>`;
function adminKeyForm(note = "") {
  return `<form class="js-akey panel" style="background:rgba(2,6,23,.45)" autocomplete="off"><h3>🔐 Ключ адміністратора</h3>
    <p class="muted">${note || "Потрібен, щоб відкривати дані підрозділів. Діє до закриття вкладки або виходу."}</p>
    <div class="row"><label class="f" style="flex:1;min-width:260px">Ключ адміністратора<input name="ak" type="password" required></label>
      <button class="btn primary">Відкрити</button></div></form>`;
}
function bindAdminKeyForm(after) {
  $app.querySelector(".js-akey")?.addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try { state.adminPriv = await X.unlockAdmin(await store.getAdminPriv(), e.target.ak.value); store.log("admin_key_open", ""); flash("ok", "Ключ адміністратора відкрито."); }
    catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
    after ? after() : render();
  });
}
function pageUnitKey(mode, u) {
  const mine = state.me.unit_id === u.id;
  if (mode === "init") {
    shell(`<div class="auth" style="max-width:680px"><div class="panel"><h2 class="display">🔑 Створіть ключ підрозділу</h2>
      <div class="note"><p><b class="hl">${esc(u.name)}</b> <i>(автопарк, реквізити, показники за місяці)</i> шифруються ще у вашому браузері — у базі
        вони лежать як незрозумілий набір символів. Відкрити їх можна лише <b>ключем підрозділу</b>.</p>
      <ul><li>Придумайте фразу з кількох слів (від 10 символів), наприклад <i>«Лубни-пальне-синій-кіт»</i>.</li>
        <li>Передайте її колегам свого підрозділу особисто чи службовим порядком — <mark>не через пошту чи месенджери</mark>.</li>
        <li>Кожен колега введе її один раз на своєму комп'ютері.</li>
        <li>Якщо ключ забудуть усі — адміністратор порталу зможе встановити новий, дані не загубляться.</li></ul></div>
      ${mine ? "" : `<div class="flash info">Ви створюєте ключ для чужого підрозділу — не забудьте передати його цьому підрозділу.</div>`}
      <form class="js-kinit" autocomplete="off"><div class="two">
        <label class="f">Ключ підрозділу<input name="k1" type="password" required minlength="10"></label>
        <label class="f">Повторіть ключ<input name="k2" type="password" required minlength="10"></label></div>${rememberBox}
        <button class="btn primary">Створити ключ і продовжити</button></form></div></div>`);
    $app.querySelector(".js-kinit").addEventListener("submit", async (e) => {
      e.preventDefault(); const f = e.target;
      if (f.k1.value !== f.k2.value) { flash("err", "Ключі не збігаються."); return pageUnitKey(mode, u); }
      const kp = X.passProblem(f.k1.value, "Ключ підрозділу"); if (kp) { flash("err", kp); return pageUnitKey(mode, u); }
      busy(true);
      try {
        const pub = await store.getAdminPub();
        const dk = await X.newDataKey();
        const cryptoObj = { v: 1, pass: await X.wrapWithPass(dk, f.k1.value, X.aadPass(u.id)), admin: await X.wrapForAdmin(dk, pub) };
        const enc = await X.encJson(dk, { settings: { ...C.DEFAULT_SETTINGS, ...(u.settings || {}) }, vehicles: u.vehicles || [] }, X.aadUnit(u.id));
        await store.initUnitCrypto(u.id, cryptoObj, enc, (u.vehicles || []).length);
        await rememberKey(u.id, dk, f.remember.checked);
        store.log("unit_key_create", "", u.id);
        flash("ok", "Ключ підрозділу створено. Дані підрозділу тепер шифруються.");
        render();
      } catch (err) {
        if (err.code === "already-init") { flash("info", "Колега вже створив ключ цього підрозділу — введіть його."); return render(); }
        flash("err", store.errUa(err)); pageUnitKey(mode, u);
      } finally { busy(false); }
    });
    return;
  }
  shell(`<div class="auth" style="max-width:620px"><div class="panel"><h2 class="display">🔑 Ключ підрозділу</h2>
    <div class="note"><p>Дані підрозділу <b class="hl">${esc(u.name)}</b> зашифровано. Введіть ключ підрозділу — його знають ваші колеги, які працюють із порталом.</p></div>
    <form class="js-kopen" autocomplete="off"><label class="f">Ключ підрозділу<input name="k" type="password" required autofocus></label>${rememberBox}
      <button class="btn primary">Відкрити</button></form>
    <p class="muted" style="margin-top:12px">Ніхто не пам'ятає ключ? Зверніться до адміністратора порталу — він встановить новий ключ, дані збережуться.</p></div>
    ${isAdmin() ? adminKeyForm("Як адміністратор ви можете відкрити будь-який підрозділ своїм ключем адміністратора.") : ""}</div>`);
  $app.querySelector(".js-kopen").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target; busy(true);
    try {
      const dk = await X.unwrapWithPass(u.crypto.pass, f.k.value, X.aadPass(u.id));
      await rememberKey(u.id, dk, f.remember.checked);
      render();
    } catch (err) { flash("err", store.errUa(err)); pageUnitKey(mode, u); } finally { busy(false); }
  });
  bindAdminKeyForm();
}

// ============================================================ введення даних
async function pageEntry(q) {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  if (!unit.vehicles.length) return go("#/fleet?new=1");
  const ym = currentYm(q);
  const st = await monthState(unit, ym);
  const entries = st.entries;
  const issue = C.ISSUE_SOURCES;
  const inp = (name, val, ph, step = "0.01") => `<input type="number" inputmode="decimal" min="0" step="${step}" name="${name}" placeholder="${ph}" value="${val ?? ""}">`;
  const cards = st.vehicles.map((v) => {
    const e = entries[v.id] || {};
    const r = C.calcRow(v, e); const s = C.statusOf(r, e);
    return `<div class="frc ${s}" data-vid="${v.id}"><span class="saving">збережено ✓</span><div class="grid">
      <div class="js-id">${U.idHtml(v, s)}</div>
      <div><div class="fr-h">Пробіг і залишок на початок</div><div class="f3">
        <div><label>Початок, км</label>${inp("odo_start", e.odo_start, "км", "1")}</div>
        <div><label>Кінець, км</label>${inp("odo_end", e.odo_end, "км", "1")}</div>
        <div><label>Залишок, л</label>${inp("fuel_start", e.fuel_start, "л")}</div></div>
        <div class="js-km">${r.empty ? "" : U.kmHtml(v, r)}</div></div>
      <div><div class="fr-h">Видано палива, л</div><div class="f4">
        ${issue.map(([k, l]) => `<div><label title="${l}">${l}</label>${inp(k, e[k], "—")}</div>`).join("")}</div>
        <div class="notes"><div class="fr-note js-note">${U.noteHtml(v, r, e)}</div>
          <div class="manual"><label style="display:flex;gap:6px;align-items:center;margin:0;color:#cbd5e1"><input type="checkbox" name="manual" ${e.manual ? "checked" : ""}> ввести витрату вручну</label>
            <input type="number" inputmode="decimal" min="0" step="0.01" name="consumed_manual" placeholder="витрачено, л" value="${e.consumed_manual ?? ""}" ${e.manual ? "" : 'style="display:none"'}></div></div></div>
      <div class="js-tank">${U.tankHtml(r, s)}</div></div>
      <div class="js-msgs">${U.msgsHtml(r)}</div></div>`;
  }).join("");
  const info = st.carried ? `<div class="flash info" style="margin:0">↪️ Новий місяць: пробіг і залишок на початок перенесено з кінця ${C.monthTitle(st.carried)}. Введіть пробіг на кінець і видане пальне.</div>`
    : st.saved ? `<div class="flash ok" style="margin:0">💾 Дані за ${C.monthTitle(ym)} збережено · зміни зберігаються автоматично${st.saved.updated_at ? " · останнє збереження " + esc(st.saved.updated_at) : ""}</div>`
      : `<div class="flash info" style="margin:0">📝 Новий місяць без попередніх даних — заповніть початкові показники.</div>`;
  const hasPrev = !!(await store.getMonth(unit.id, C.prevYm(ym)));  // лише факт наявності — без розшифрування
  shell(`<div class="row" style="justify-content:space-between;margin-top:14px">${monthPicker(ym, "#/entry")}<div style="flex:1;min-width:280px">${info}</div></div>
    <div id="kpi"></div>
    ${hasPrev ? `<button class="btn sm js-prev">↪️ Підтягнути пробіг і залишок на початок з ${C.monthTitle(C.prevYm(ym))}</button>` : ""}
    <div style="margin-top:14px">${cards}</div>
    <div class="stickybar"><button class="btn primary js-reg">📥 Реєстр «Додаток 6» (Word)</button>
      <a class="btn sky" href="#/report?ym=${ym}">🧾 Акти списання та перегляд</a>
      <span class="muted js-savestate">Дані зберігаються автоматично після кожної зміни.</span></div>`);
  bindMonthPicker();

  const cardEls = [...$app.querySelectorAll(".frc")];
  const vmap = Object.fromEntries(st.vehicles.map((v) => [v.id, v]));
  const readCard = (card) => { const e = {}; card.querySelectorAll("input[name]").forEach((i) => { e[i.name] = i.type === "checkbox" ? i.checked : i.value; }); return C.cleanEntry(e); };
  const refreshKpi = () => {
    const rows = st.vehicles.map((v) => C.calcRow(v, entries[v.id]));
    const sts = rows.map((r, i) => C.statusOf(r, entries[st.vehicles[i].id] || {}));
    $app.querySelector("#kpi").innerHTML = U.kpiHtml(rows, sts, C.MONTHS_UA[C.parseYm(ym)[1] - 1]);
  };
  const refreshCard = (card) => {
    const v = vmap[card.dataset.vid]; const e = entries[v.id] || {};
    const r = C.calcRow(v, e); const s = C.statusOf(r, e);
    card.className = "frc " + s;
    card.querySelector(".js-id").innerHTML = U.idHtml(v, s);
    card.querySelector(".js-km").innerHTML = r.empty ? "" : U.kmHtml(v, r);
    card.querySelector(".js-note").innerHTML = U.noteHtml(v, r, e);
    card.querySelector(".js-tank").innerHTML = U.tankHtml(r, s);
    card.querySelector(".js-msgs").innerHTML = U.msgsHtml(r);
  };
  let timer = null; let dirtyCards = new Set();
  const saveNow = async () => {
    const touched = [...dirtyCards]; dirtyCards = new Set();
    const ss = $app.querySelector(".js-savestate");
    try {
      const snapshot = Object.fromEntries(st.vehicles.map((v) => [v.id, v]));
      await writeMonth(unit, ym, { entries, snapshot });
      touched.forEach((c) => { const s = c.querySelector(".saving"); s.classList.add("on"); setTimeout(() => s.classList.remove("on"), 1200); });
      if (ss) ss.textContent = "Збережено ✓ " + new Date().toLocaleTimeString("uk-UA");
    } catch (err) { if (ss) ss.textContent = "⚠️ Не збережено: " + store.errUa(err); alert("Не вдалося зберегти: " + store.errUa(err)); }
  };
  const schedule = (card, ms) => { dirtyCards.add(card); clearTimeout(timer); timer = setTimeout(saveNow, ms); };
  cardEls.forEach((card) => {
    card.querySelectorAll("input").forEach((inp) => {
      const mark = () => { if (inp.type !== "checkbox") inp.classList.toggle("filled", inp.value !== ""); };
      mark();
      const onChange = (ms) => {
        mark();
        if (inp.name === "manual") card.querySelector("[name=consumed_manual]").style.display = inp.checked ? "" : "none";
        entries[card.dataset.vid] = readCard(card); refreshCard(card); refreshKpi(); schedule(card, ms);
      };
      inp.addEventListener("input", () => onChange(800));
      inp.addEventListener("change", () => onChange(200));
    });
  });
  refreshKpi();
  $app.querySelector(".js-prev")?.addEventListener("click", async () => {
    if (!confirm("Перезаписати «Пробіг на початок» і «Залишок на початок» значеннями на кінець попереднього місяця?")) return;
    const prev = await readMonth(unit, C.prevYm(ym));
    const co = C.carryOver(unit.vehicles, { entries: prev.entries || {}, snapshot: prev.snapshot || {} });
    for (const card of cardEls) {
      const e = co[card.dataset.vid]; if (!e) continue;
      card.querySelector("[name=odo_start]").value = e.odo_start ?? ""; card.querySelector("[name=fuel_start]").value = e.fuel_start ?? "";
      card.querySelectorAll("input").forEach((i) => { if (i.type !== "checkbox") i.classList.toggle("filled", i.value !== ""); });
      entries[card.dataset.vid] = readCard(card); refreshCard(card); dirtyCards.add(card);
    }
    refreshKpi(); await saveNow();
  });
  $app.querySelector(".js-reg").addEventListener("click", async () => {
    clearTimeout(timer); if (dirtyCards.size) await saveNow();
    const rows = st.vehicles.map((v) => C.calcRow(v, entries[v.id]));
    await download(buildRegister(window.docx, unit.settings, ym, rows), `Додаток_6_Реєстр_пального_${ym}.docx`);
    store.log("doc_register", ym, unit.id);
  });
  window.onbeforeunload = () => (dirtyCards.size ? "Є незбережені зміни" : undefined);
}

// ============================================================ звіт і акти
async function pageReport(q) {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  const ym = currentYm(q);
  const rows = await rowsFor(unit, ym);
  const nErr = rows.reduce((a, r) => a + r.errors.length, 0);
  const acts = Object.entries(C.ACT_GROUPS).map(([g, meta]) => {
    const lines = C.actGroups(rows, g); const tot = lines.reduce((a, l) => a + l.qtyC, 0);
    return `<div class="act"><h3>Акт «${meta.title}»</h3><div class="big">${C.fmtC(tot)} л</div>
      ${lines.length ? `<ul>${lines.map((l) => `<li>${esc(l.fuel)} ${esc(l.supplier)} — ${C.fmtC(l.qtyC)} л (МВО ${esc(l.mvo)})</li>`).join("")}</ul>
        <p class="muted">«${litresWords(tot)}»</p><button class="btn sky js-act" data-g="${g}">📥 Завантажити акт (Word)</button>`
        : `<p class="muted">Немає пального до списання.</p>`}</div>`;
  }).join("");
  shell(`<div style="margin-top:14px">${monthPicker(ym, "#/report")}</div>
    <div class="panel"><h2>${esc(unit.settings.title)} за ${C.monthTitle(ym)}</h2>
      ${nErr ? `<div class="flash err">У даних є помилки (${nErr}) — перевірте картки на вкладці «Введення даних».</div>` : ""}
      ${rows.length ? `<div class="tblwrap">${U.previewHtml(rows, unit.settings.unit)}</div>
        <p class="muted">* дані графи № 12 повинні відповідати даним граф № 11, 12 звіту про одержане і видане пальне.<br>
          * дані графи № 17 повинні відповідати кількості пального в акті списання.<br>
          Підпис особи, яка склала реєстр: ${esc(unit.settings.compiler || "— (вкажіть у «Автопарк і реквізити»)")}</p>
        <button class="btn primary js-reg">📥 Завантажити реєстр «Додаток 6» (Word)</button>` : `<p class="muted">Немає даних за цей місяць.</p>`}</div>
    <div class="panel"><h2>🧾 Акти списання пального</h2>
      <p class="muted">Два окремі акти: «Талони + см/к» (графи 9, 10) та «АЗС» (АЗС ГУНП талони, АЗС налив — графи 8, 11).
        Витрата авто розноситься між ними пропорційно до виданого за місяць; якщо нічого не видано — за основним джерелом списання.
        Номер, дати, ціна й сума лишаються порожніми.</p><div class="acts">${acts}</div></div>`);
  bindMonthPicker();
  $app.querySelector(".js-reg")?.addEventListener("click", async () => {
    await download(buildRegister(window.docx, unit.settings, ym, rows), `Додаток_6_Реєстр_пального_${ym}.docx`); store.log("doc_register", ym, unit.id);
  });
  $app.querySelectorAll(".js-act").forEach((b) => b.addEventListener("click", async () => {
    const g = b.dataset.g;
    await download(buildAct(window.docx, unit.settings, ym, rows, g), `Акт_списання_${C.ACT_GROUPS[g].file}_${ym}.docx`);
    store.log("doc_act", `${g} ${ym}`, unit.id);
  }));
}

// ============================================================ архів
async function pageArchive() {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  const months = (await listMonthsDec(unit)).reverse();
  const recs = months.map((m) => {
    const vehicles = C.monthVehicles(unit.vehicles, { entries: m.entries || {}, snapshot: m.snapshot || {} });
    const rows = vehicles.map((v) => C.calcRow(v, (m.entries || {})[v.id]));
    const acts = Object.fromEntries(Object.keys(C.ACT_GROUPS).map((g) => [g, C.actGroups(rows, g).reduce((a, l) => a + l.qtyC, 0)]));
    return { ym: m.ym, rows, t: C.totals(rows), acts };
  });
  shell(`<div class="panel"><h2>🗂️ Архів місяців</h2>${recs.length ? `<table class="list"><thead><tr><th>Місяць</th><th class="num">Авто</th>
    <th class="num">Пробіг, км</th><th class="num">Видано, л</th><th class="num">Витрачено, л</th><th class="num">Залишок, л</th><th>Документи</th></tr></thead><tbody>
    ${recs.map((r) => `<tr><td><a href="#/entry?ym=${r.ym}">${C.monthTitle(r.ym)}</a></td><td class="num">${r.rows.length}</td><td class="num">${C.fmtKm(r.t.km)}</td>
      <td class="num">${C.fmtC(r.t.issuedC)}</td><td class="num">${C.fmtC(r.t.consC)}</td><td class="num">${C.fmtC(r.t.endC)}</td>
      <td><button class="btn sm js-doc" data-ym="${r.ym}" data-k="reg">Реєстр</button>
        ${Object.entries(C.ACT_GROUPS).map(([g, meta]) => `<button class="btn sm js-doc ${r.acts[g] > 0 ? "" : "disabled"}" data-ym="${r.ym}" data-k="${g}">Акт «${meta.title}»</button>`).join(" ")}</td></tr>`).join("")}
    </tbody></table>` : `<p class="muted">Збережених місяців ще немає.</p>`}</div>`);
  $app.querySelectorAll(".js-doc").forEach((b) => b.addEventListener("click", async () => {
    const r = recs.find((x) => x.ym === b.dataset.ym); const k = b.dataset.k;
    if (k === "reg") await download(buildRegister(window.docx, unit.settings, r.ym, r.rows), `Додаток_6_Реєстр_пального_${r.ym}.docx`);
    else await download(buildAct(window.docx, unit.settings, r.ym, r.rows, k), `Акт_списання_${C.ACT_GROUPS[k].file}_${r.ym}.docx`);
  }));
}

// ============================================================ автопарк і реквізити
async function pageFleet() {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  const onboarding = !unit.vehicles.length;
  const s = unit.settings;
  const f = (k, label, cls = "") => `<label class="f ${cls}">${label}<input data-k="${k}" value="${esc(s[k] ?? "")}"></label>`;
  shell(`${onboarding ? `<div class="panel" style="border-color:rgba(250,204,21,.5);background:linear-gradient(120deg,rgba(8,145,178,.25),rgba(55,48,163,.25))">
      <h2 class="display">👋 Вітаємо! Створіть автопарк свого підрозділу</h2>
      <p>Це перший вхід для підрозділу <b>${esc(unit.name)}</b> — автопарк ще порожній. Додайте службові авто (марка, модель і номер,
        норма витрати, вид пального, АЗС), перевірте реквізити бланків і склад комісії для акта списання, а потім натисніть
        <b>«Зберегти шаблон підрозділу»</b>. Після цього відкриється введення даних за місяць.</p>
      <p class="muted">Шаблон створюється один раз: колеги з вашого підрозділу, які зареєструються пізніше, одразу побачать цей автопарк.</p></div>` : ""}
    <div class="panel"><h2>🚗 Автопарк підрозділу</h2>
      <p class="muted">Шаблон вашого автопарку: усі місяці беруть авто звідси. Авто, яке тимчасово не звітує, краще не видаляти, а зняти «Активне» —
        історія місяців збережеться. «Марка, ДНЗ» пишіть як у звіті, номер — останнім словом (напр. <i>Peugeot 301 1825</i>).</p>
      <div class="tblwrap"><table class="edit" id="veh"><thead><tr><th style="min-width:200px">Марка, ДНЗ</th><th>Норма, л/100 км</th><th>Надбавка, %</th>
        <th>Норма пробігу, км/міс</th><th>Вид пального</th><th>АЗС</th><th>Джерело списання</th><th style="min-width:170px">МВО (графа 19)</th><th>Активне</th><th></th></tr></thead><tbody></tbody></table></div>
      <button class="btn sm js-addveh" style="margin-top:10px">➕ Додати авто</button>
      <datalist id="fuels"><option value="А-95"><option value="А-92"><option value="ДП"><option value="ГАЗ"></datalist></div>
    <div class="panel"><h2>📄 Реквізити бланків</h2><div class="form-grid">
      ${f("org", "Установа (шапка ліворуч)")}${f("dept", "Підрозділ (другий рядок шапки)")}${f("unit", "Графа 2 «Підрозділ» у реєстрі")}
      ${f("title", "Назва реєстру (без «за … року»)", "w2")}${f("appendix", "Номер додатка")}${f("compiler", "Особа, яка склала реєстр (ПІБ для підпису)")}</div></div>
    <div class="panel"><h2>🧾 Акт списання (типова форма № З-2)</h2><div class="form-grid">
      ${f("edrpou", "Код за ЄДРПОУ")}${f("act_approver_title", "«Затверджую» — посада")}${f("act_approver", "«Затверджую» — ПІБ")}
      ${f("act_owner", "За ким закріплені ТЗ (у тексті акта)")}${f("act_order", "Ким призначена комісія (наказ)", "w2")}
      <label class="f w2">Голова комісії — посада, звання, ПІБ (родовий відмінок)<textarea data-k="act_chair_full" rows="2">${esc(s.act_chair_full)}</textarea></label>
      ${f("act_chair", "Голова комісії — ПІБ для підпису")}</div>
      <h3 style="margin-top:16px">Члени комісії</h3>
      <table class="edit" id="mem"><thead><tr><th>Посада, звання, ПІБ (родовий відмінок)</th><th style="width:220px">ПІБ для підпису</th><th style="width:40px"></th></tr></thead><tbody></tbody></table>
      <button class="btn sm js-addmem" style="margin-top:10px">➕ Додати члена комісії</button></div>
    <div class="panel"><h2>🔐 Ключ підрозділу</h2>
      <p class="muted">Дані підрозділу зашифровано ключем підрозділу. Змінюйте ключ, якщо його міг дізнатися сторонній або з підрозділу
        звільнився працівник (доступ до порталу йому також закриває адміністратор), — після зміни повідомте новий ключ колегам. Кнопка «Забути» прибирає ключ із цього комп'ютера (наступного разу його треба буде ввести).</p>
      <form class="js-kchange row" autocomplete="off">
        <label class="f">Поточний ключ<input name="old" type="password" required></label>
        <label class="f">Новий ключ (від 10 символів)<input name="n1" type="password" required minlength="10"></label>
        <label class="f">Повторіть новий ключ<input name="n2" type="password" required minlength="10"></label>
        <button class="btn sky">Змінити ключ</button></form>
      <button class="btn sm danger js-kforget" style="margin-top:10px">Забути ключ на цьому комп'ютері</button></div>
    <div class="stickybar"><button class="btn primary js-save">💾 Зберегти шаблон підрозділу</button><span class="muted js-st"></span></div>`);
  $app.querySelector(".js-kchange").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.n1.value !== f.n2.value) { flash("err", "Нові ключі не збігаються."); return render(); }
    const kp = X.passProblem(f.n1.value, "Ключ підрозділу"); if (kp) { flash("err", kp); return render(); }
    busy(true);
    try {
      const fresh = await store.getUnit(unit.id);
      const dk = await X.unwrapWithPass(fresh.crypto.pass, f.old.value, X.aadPass(unit.id), true);
      await store.setUnitPass(unit.id, await X.wrapWithPass(dk, f.n1.value, X.aadPass(unit.id)));
      store.log("unit_key_change", "", unit.id);
      flash("ok", "Ключ підрозділу змінено. Колеги, які вже ввели ключ на своїх комп'ютерах, працюють далі; на новому комп'ютері знадобиться новий ключ.");
    } catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
    render();
  });
  $app.querySelector(".js-kforget").addEventListener("click", async () => {
    if (!confirm("Прибрати ключ підрозділу з цього комп'ютера? Наступного разу його доведеться ввести.")) return;
    await X.keyCache.del(keyCacheId(unit.id)); delete state.keys[unit.id];
    flash("ok", "Ключ прибрано з цього комп'ютера."); render();
  });

  const vbody = $app.querySelector("#veh tbody"), mbody = $app.querySelector("#mem tbody");
  const addVeh = (v) => {
    const tr = document.createElement("tr"); tr.dataset.id = v.id || C.newId();
    const sup = [...new Set([...C.SUPPLIERS, v.supplier].filter(Boolean))];
    tr.innerHTML = `<td><input data-f="name" value="${esc(v.name)}" placeholder="Марка модель номер"></td>
      <td><input data-f="rate" type="number" step="0.1" min="0" value="${v.rate ?? ""}"></td>
      <td><input data-f="winter_pct" type="number" step="1" min="0" max="100" value="${v.winter_pct ?? 0}"></td>
      <td><input data-f="month_norm_km" type="number" step="10" min="0" value="${v.month_norm_km ?? 2640}"></td>
      <td><input data-f="fuel" list="fuels" value="${esc(v.fuel ?? "А-95")}"></td>
      <td><select data-f="supplier">${sup.map((x) => `<option ${x === (v.supplier || "UKRNAFTA") ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></td>
      <td><select data-f="source">${C.CONSUME_SOURCES.map(([k, l]) => `<option value="${k}" ${k === (v.source || "smk") ? "selected" : ""}>${l}</option>`).join("")}</select></td>
      <td><input data-f="mvo_text" value="${esc(v.mvo_text)}" placeholder="за замовчуванням — хто склав"></td>
      <td class="c"><input data-f="active" type="checkbox" ${v.active === false ? "" : "checked"}></td>
      <td><button class="btn sm danger js-del" title="Видалити з довідника">🗑</button></td>`;
    tr.querySelector(".js-del").addEventListener("click", () => { if (confirm("Видалити авто з довідника?")) tr.remove(); });
    vbody.appendChild(tr);
  };
  const addMem = (m) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-f="full" value="${esc(m.full)}" placeholder="начальника … капітана поліції Прізвища І.Б."></td>
      <td><input data-f="short" value="${esc(m.short)}" placeholder="Прізвище І.Б."></td><td><button class="btn sm danger">🗑</button></td>`;
    tr.querySelector("button").addEventListener("click", () => tr.remove());
    mbody.appendChild(tr);
  };
  unit.vehicles.forEach(addVeh); if (!unit.vehicles.length) addVeh({});
  (s.act_members || []).forEach(addMem);
  $app.querySelector(".js-addveh").addEventListener("click", () => addVeh({}));
  $app.querySelector(".js-addmem").addEventListener("click", () => addMem({}));
  $app.querySelector(".js-save").addEventListener("click", async () => {
    const settings = { ...s };
    $app.querySelectorAll("[data-k]").forEach((i) => { settings[i.dataset.k] = i.value.trim(); });
    settings.act_members = [...mbody.querySelectorAll("tr")].map((tr) => ({ full: tr.querySelector("[data-f=full]").value.trim(),
      short: tr.querySelector("[data-f=short]").value.trim() })).filter((m) => m.full || m.short);
    const vehicles = [...vbody.querySelectorAll("tr")].map((tr) => {
      const v = { id: tr.dataset.id };
      tr.querySelectorAll("[data-f]").forEach((i) => { v[i.dataset.f] = i.type === "checkbox" ? i.checked : i.value; });
      if (!String(v.mvo_text || "").trim()) v.mvo_text = settings.compiler || "";
      return C.normalizeVehicle(v);
    }).filter((v) => v.name);
    const st = $app.querySelector(".js-st"); st.textContent = "Збереження…"; busy(true);
    try {
      await writeUnit(unit, settings, vehicles);
      store.log("fleet_save", `${vehicles.length} авто`, unit.id);
      flash("ok", "Шаблон підрозділу збережено.");
      go(onboarding && vehicles.length ? "#/entry" : "#/fleet");
    } catch (err) { st.textContent = "⚠️ " + store.errUa(err); } finally { busy(false); }
  });
}

// ============================================================ адміністрування
async function pageAdmin(sub, q) {
  if (!isAdmin()) return go("#/entry");
  const tab = sub || "requests";
  const [users, units] = await Promise.all([store.listUsers(), store.listUnits()]);
  state.units = units; state.pendingCount = users.filter((u) => u.status === "pending").length;
  const unitName = (id) => units.find((u) => u.id === id)?.name || "—";
  const usersIn = (id) => users.filter((u) => u.unit_id === id).length;
  const tabs = `<nav class="tabs" style="margin-top:16px">
    ${[["requests", "Заявки", state.pendingCount], ["users", "Користувачі"], ["units", `Підрозділи (${units.length})`], ["summary", "📊 Зведення по області"], ["audit", "Журнал дій"]]
      .map(([k, l, b]) => `<a href="#/admin/${k}" class="${tab === k ? "on" : ""}">${l}${b ? `<span class="badge">${b}</span>` : ""}</a>`).join("")}</nav>`;
  let body = "";
  const unitOptions = (sel) => units.map((u) => `<option value="${u.id}" ${u.id === sel ? "selected" : ""}>${esc(u.name)}</option>`).join("");

  if (tab === "requests") {
    const pend = users.filter((u) => u.status === "pending");
    body = `<div class="panel"><h2>Заявки на реєстрацію</h2>${pend.length ? "" : `<p class="muted">Нових заявок немає.</p>`}
      ${pend.map((u) => {
        const match = units.find((x) => x.name.toLowerCase() === String(u.requested_unit || "").toLowerCase());
        return `<div class="panel" style="background:rgba(2,6,23,.45)" data-uid="${u.uid}">
          <div><b>${esc(u.full_name)}</b> · ${esc(u.position)}${u.phone ? " · " + esc(u.phone) : ""}<br>
            <span class="muted">${esc(u.email)} · вказаний підрозділ: <b>${esc(u.requested_unit)}</b> · подано ${esc(u.created_at || "")}</span></div>
          <div class="row" style="margin-top:10px">
            <label class="f">Прив'язати до підрозділу<select class="js-unitsel"><option value="new" ${match ? "" : "selected"}>➕ Створити новий підрозділ</option>${unitOptions(match?.id)}</select></label>
            <label class="f js-newname" ${match ? 'style="display:none"' : ""}>Назва нового підрозділу<input value="${esc(u.requested_unit)}"></label>
            <label class="f">Роль<select class="js-role"><option value="user">Користувач підрозділу</option><option value="admin">Адміністратор</option></select></label>
            <button class="btn primary js-approve">✓ Підтвердити</button></div>
          <div style="margin-top:8px"><button class="btn sm danger js-reject">✕ Відхилити</button> <button class="btn sm danger js-udel">🗑 Видалити заявку</button></div></div>`;
      }).join("")}</div>`;
  } else if (tab === "users") {
    const list = users.filter((u) => u.status !== "pending").sort((a, b) => unitName(a.unit_id).localeCompare(unitName(b.unit_id), "uk"));
    body = `<div class="panel"><h2>Користувачі</h2><table class="list"><thead><tr><th>ПІБ / посада</th><th>E-mail</th><th>Підрозділ і роль</th><th>Статус</th><th>Останній вхід</th><th>Дії</th></tr></thead><tbody>
      ${list.map((u) => `<tr data-uid="${u.uid}"><td><b>${esc(u.full_name)}</b><br><span class="muted">${esc(u.position)}${u.phone ? " · " + esc(u.phone) : ""}</span></td>
        <td>${esc(u.email)}</td>
        <td><div class="row" style="gap:6px"><select class="js-unitsel">${unitOptions(u.unit_id)}</select>
          <select class="js-role"><option value="user" ${u.role === "user" ? "selected" : ""}>користувач</option><option value="admin" ${u.role === "admin" ? "selected" : ""}>адміністратор</option></select>
          <button class="btn sm js-usave">Зберегти</button></div></td>
        <td>${u.status === "active" ? `<span class="pill ok">активний</span>` : u.status === "blocked" ? `<span class="pill err">заблокований</span>` : `<span class="pill mut">${esc(u.status)}</span>`}</td>
        <td class="muted">${esc(u.last_login || "—")}</td>
        <td><button class="btn sm js-ureset">✉️ Лист для скидання пароля</button>
          ${u.uid !== state.authUser.uid ? `${u.status === "active" ? `<button class="btn sm danger js-ublock">Заблокувати</button>` : `<button class="btn sm js-uunblock">Розблокувати</button>`}
          <button class="btn sm danger js-udel">🗑 Видалити</button>` : ""}</td></tr>`).join("")}</tbody></table>
      <p class="muted">Видалення прибирає профіль і доступ до даних. Якщо людина знову увійде зі своїм e-mail, портал запропонує подати нову заявку.</p></div>`;
  } else if (tab === "units") {
    body = `<div class="panel"><h2>Підрозділи</h2><div class="row" style="margin-bottom:14px">
      <label class="f">Новий підрозділ<input class="js-newunit" placeholder="напр. Кременчуцьке РУП ГУНП" style="min-width:320px"></label><button class="btn sky js-ucreate">➕ Створити</button></div>
      ${state.adminPriv ? "" : adminKeyForm("Відкрийте ключ адміністратора, щоб встановлювати новий ключ підрозділу, який його забув.")}
      <table class="list"><thead><tr><th>Назва</th><th class="num">Користувачів</th><th class="num">Авто</th><th>Шифрування</th><th>Дії</th></tr></thead><tbody>
      ${units.map((u) => `<tr data-id="${u.id}"><td><div class="row" style="gap:6px"><input class="js-uname" value="${esc(u.name)}" style="min-width:300px"><button class="btn sm js-urename">Перейменувати</button></div></td>
        <td class="num">${usersIn(u.id)}</td><td class="num">${u.n_vehicles ?? 0}</td>
        <td>${u.crypto ? `<span class="pill ok">🔒 увімкнено</span>` : `<span class="pill mut">ключ ще не створено</span>`}</td>
        <td><button class="btn sm js-open">Відкрити як підрозділ →</button> ${usersIn(u.id) ? "" : `<button class="btn sm danger js-unitdel">🗑 Видалити</button>`}
          ${u.crypto && state.adminPriv ? `<button class="btn sm js-rkshow">🔑 Новий ключ підрозділу</button>
          <form class="js-rk row" hidden autocomplete="off" style="margin-top:8px"><input name="n1" type="password" placeholder="новий ключ (від 10 символів)" required minlength="10">
            <input name="n2" type="password" placeholder="повторіть" required minlength="10"><button class="btn sm primary">Встановити</button></form>` : ""}</td></tr>`).join("")}
      </tbody></table></div>`;
  } else if (tab === "summary" && !state.adminPriv) {
    body = adminKeyForm("Зведення по області формується з даних усіх підрозділів, а вони зашифровані. Введіть ключ адміністратора.");
  } else if (tab === "summary") {
    const ym = currentYm(q);
    const rowsSum = await Promise.all(units.map(async (u) => {
      if (!u.crypto) return { u, n: 0, done: 0, err: 0, t: null };
      const unit = await decUnit(u, await getUnitKey(u));
      const m = await readMonth(unit, ym);
      if (!m) return { u, n: unit.vehicles.filter((v) => v.active).length, done: 0, err: 0, t: null };
      const vehicles = C.monthVehicles(unit.vehicles, { entries: m.entries || {}, snapshot: m.snapshot || {} });
      const rows = vehicles.map((v) => C.calcRow(v, (m.entries || {})[v.id]));
      const sts = rows.map((r, i) => C.statusOf(r, (m.entries || {})[vehicles[i].id] || {}));
      return { u, n: rows.length, done: sts.filter((s) => s === "ok" || s === "warn").length, err: sts.filter((s) => s === "err").length, t: C.totals(rows), updated: m.updated_at };
    }));
    state.lastSummary = { ym, rowsSum };
    const sum = (k) => rowsSum.reduce((a, r) => a + (r.t ? r.t[k] : 0), 0);
    body = `<div class="panel"><div class="row" style="margin-bottom:12px">${monthPicker(ym, "#/admin/summary")}<button class="btn sm js-sumdoc">📥 Зведення у Word</button></div>
      <h2>Зведення по області за ${C.monthTitle(ym)}</h2>
      <p class="muted">Повністю заповнили: <b>${rowsSum.filter((r) => r.n && r.done === r.n).length}</b> з ${rowsSum.length} підрозділів ·
        видано <b>${C.fmtC(sum("issuedC"))} л</b> · витрачено <b>${C.fmtC(sum("consC"))} л</b> · залишок у баках <b>${C.fmtC(sum("endC"))} л</b></p>
      <table class="list"><thead><tr><th>Підрозділ</th><th>Заповнено</th><th class="num">Пробіг, км</th><th class="num">Видано, л</th><th class="num">Витрачено, л</th><th class="num">Залишок, л</th><th>Оновлено</th></tr></thead><tbody>
      ${rowsSum.map((r) => `<tr><td>${esc(r.u.name)}</td><td>${r.n === 0 ? `<span class="pill mut">немає авто</span>` : r.done === r.n ? `<span class="pill ok">${r.done}/${r.n}</span>`
        : r.done ? `<span class="pill warn">${r.done}/${r.n}</span>` : `<span class="pill err">0/${r.n}</span>`}${r.err ? ` <span class="pill err">помилок: ${r.err}</span>` : ""}</td>
        <td class="num">${r.t ? C.fmtKm(r.t.km) : "0"}</td><td class="num">${r.t ? C.fmtC(r.t.issuedC) : "0"}</td><td class="num">${r.t ? C.fmtC(r.t.consC) : "0"}</td>
        <td class="num">${r.t ? C.fmtC(r.t.endC) : "0"}</td><td class="muted">${esc(r.updated || "—")}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (tab === "audit") {
    const audit = await store.listAudit(200);
    body = `<div class="panel"><h2>Журнал дій (останні 200)</h2><table class="list"><thead><tr><th>Час</th><th>Користувач</th><th>Підрозділ</th><th>Дія</th><th>Деталі</th></tr></thead><tbody>
      ${audit.map((a) => `<tr><td class="muted">${esc(a.ts)}</td><td>${esc(a.email || "")}</td><td>${esc(a.unit_id ? unitName(a.unit_id) : "—")}</td><td>${esc(a.action)}</td><td class="muted">${esc(a.details)}</td></tr>`).join("")}</tbody></table></div>`;
  }
  shell(tabs + body);
  bindMonthPicker();
  bindAdminKeyForm();
  const act = async (fn, okMsg) => { busy(true); try { await fn(); if (okMsg) flash("ok", okMsg); } catch (err) { flash("err", store.errUa(err)); } finally { busy(false); render(); } };

  $app.querySelectorAll(".js-unitsel").forEach((s) => s.addEventListener("change", () => {
    const nn = s.closest("[data-uid]").querySelector(".js-newname"); if (nn) nn.style.display = s.value === "new" ? "" : "none";
  }));
  $app.querySelectorAll(".js-approve").forEach((b) => b.addEventListener("click", () => act(async () => {
    const box = b.closest("[data-uid]"); const uid = box.dataset.uid; const u = users.find((x) => x.uid === uid);
    let unitId = box.querySelector(".js-unitsel").value;
    if (unitId === "new") {
      const name = box.querySelector(".js-newname input").value.trim(); if (!name) throw new Error("Вкажіть назву нового підрозділу.");
      const ex = units.find((x) => x.name.toLowerCase() === name.toLowerCase());
      unitId = ex ? ex.id : await store.createUnit(name, { ...C.DEFAULT_SETTINGS, dept: name, unit: name, act_owner: name, compiler: C.shortName(u.full_name) });
    }
    await store.updateUser(uid, { status: "active", unit_id: unitId, role: box.querySelector(".js-role").value });
    store.log("approve", `${u.full_name} → ${unitName(unitId) === "—" ? "новий підрозділ" : unitName(unitId)}`, unitId);
  }, "Заявку підтверджено.")));
  $app.querySelectorAll(".js-reject").forEach((b) => b.addEventListener("click", () => {
    if (!confirm("Відхилити заявку?")) return;
    act(() => store.updateUser(b.closest("[data-uid]").dataset.uid, { status: "rejected" }), "Заявку відхилено.");
  }));
  $app.querySelectorAll(".js-udel").forEach((b) => b.addEventListener("click", () => {
    if (!confirm("Видалити користувача (заявку) остаточно? Дані підрозділу залишаться.")) return;
    const uid = b.closest("[data-uid]").dataset.uid; const u = users.find((x) => x.uid === uid);
    act(async () => { await store.deleteUser(uid); store.log("user_delete", `${u.full_name} (${u.email})`); }, "Користувача видалено.");
  }));
  $app.querySelectorAll(".js-usave").forEach((b) => b.addEventListener("click", () => {
    const tr = b.closest("[data-uid]"); const uid = tr.dataset.uid; const role = tr.querySelector(".js-role").value;
    if (uid === state.authUser.uid && role !== "admin") { flash("err", "Не можна зняти роль адміністратора із себе."); return render(); }
    act(() => store.updateUser(uid, { unit_id: tr.querySelector(".js-unitsel").value, role }), "Користувача оновлено.");
  }));
  $app.querySelectorAll(".js-ublock").forEach((b) => b.addEventListener("click", () => { if (confirm("Заблокувати користувача?")) act(() => store.updateUser(b.closest("[data-uid]").dataset.uid, { status: "blocked" }), "Заблоковано."); }));
  $app.querySelectorAll(".js-uunblock").forEach((b) => b.addEventListener("click", () => act(() => store.updateUser(b.closest("[data-uid]").dataset.uid, { status: "active" }), "Розблоковано.")));
  $app.querySelectorAll(".js-ureset").forEach((b) => b.addEventListener("click", () => {
    const u = users.find((x) => x.uid === b.closest("[data-uid]").dataset.uid);
    act(() => store.resetPassword(u.email), `Лист для встановлення нового пароля надіслано на ${u.email}.`);
  }));
  $app.querySelector(".js-ucreate")?.addEventListener("click", () => {
    const name = $app.querySelector(".js-newunit").value.trim(); if (!name) return;
    if (units.some((u) => u.name.toLowerCase() === name.toLowerCase())) { flash("err", "Така назва вже існує."); return render(); }
    act(() => store.createUnit(name, { ...C.DEFAULT_SETTINGS, dept: name, unit: name, act_owner: name }), `Підрозділ «${name}» створено.`);
  });
  $app.querySelectorAll(".js-urename").forEach((b) => b.addEventListener("click", () => {
    const tr = b.closest("[data-id]"); act(() => store.updateUnit(tr.dataset.id, { name: tr.querySelector(".js-uname").value.trim() }), "Назву змінено.");
  }));
  $app.querySelectorAll(".js-rkshow").forEach((b) => b.addEventListener("click", () => { const f = b.parentElement.querySelector(".js-rk"); f.hidden = !f.hidden; }));
  $app.querySelectorAll(".js-rk").forEach((f) => f.addEventListener("submit", (e) => {
    e.preventDefault();
    if (f.n1.value !== f.n2.value) { flash("err", "Ключі не збігаються."); return render(); }
    const kp = X.passProblem(f.n1.value, "Ключ підрозділу"); if (kp) { flash("err", kp); return render(); }
    const id = f.closest("[data-id]").dataset.id;
    act(async () => {
      const u = await store.getUnit(id);
      const dk = await X.unwrapByAdmin(u.crypto.admin, state.adminPriv, true);
      await store.setUnitPass(id, await X.wrapWithPass(dk, f.n1.value, X.aadPass(id)));
      store.log("unit_key_reset", unitName(id), id);
    }, `Новий ключ підрозділу «${unitName(id)}» встановлено. Передайте його підрозділу особисто.`);
  }));
  $app.querySelectorAll(".js-open").forEach((b) => b.addEventListener("click", () => {
    state.unitId = b.closest("[data-id]").dataset.id; sessionStorage.setItem("unitCtx", state.unitId); go("#/entry");
  }));
  $app.querySelectorAll(".js-unitdel").forEach((b) => b.addEventListener("click", () => {
    if (!confirm("Видалити підрозділ разом з автопарком і всіма місяцями? Це незворотно.")) return;
    const id = b.closest("[data-id]").dataset.id;
    act(async () => { await store.deleteUnit(id); store.log("unit_delete", unitName(id)); if (state.unitId === id) state.unitId = state.me.unit_id; }, "Підрозділ видалено.");
  }));
  $app.querySelector(".js-sumdoc")?.addEventListener("click", async () => {
    const D = window.docx; const { ym, rowsSum } = state.lastSummary;
    const cellP = (t, bold = false) => new D.TableCell({ children: [new D.Paragraph({ children: [new D.TextRun({ text: String(t), bold, font: "Times New Roman", size: 22 })] })] });
    const rowsDoc = [new D.TableRow({ tableHeader: true, children: ["№", "Підрозділ", "Авто", "Заповнено", "Видано, л", "Витрачено, л", "Залишок, л"].map((h) => cellP(h, true)) }),
      ...rowsSum.map((r, i) => new D.TableRow({ children: [i + 1, r.u.name, r.n, `${r.done}/${r.n}`, r.t ? C.fmtC(r.t.issuedC) : "0", r.t ? C.fmtC(r.t.consC) : "0", r.t ? C.fmtC(r.t.endC) : "0"].map((x) => cellP(x)) }))];
    const doc = new D.Document({ sections: [{ children: [
      new D.Paragraph({ children: [new D.TextRun({ text: `Зведення про видане та витрачене пальне по підрозділах за ${C.monthTitle(ym)}`, bold: true, font: "Times New Roman", size: 26 })], spacing: { after: 200 } }),
      new D.Table({ rows: rowsDoc, width: { size: 100, type: D.WidthType.PERCENTAGE } })] }] });
    await download(doc, `Зведення_пального_${ym}.docx`);
  });
}

// ============================================================ маршрутизатор
async function render() {
  const { path, q } = route();
  if (!store) return pageNotConfigured();
  if (state.setupDone === null) state.setupDone = await store.setupDone().catch(() => true);
  if (!state.authUser) {
    if (!state.setupDone) return pageSetup();
    if (path === "register") return pageRegister();
    if (path === "forgot") return pageForgot();
    return pageLogin();
  }
  if (!state.me) return state.setupDone ? pageNoProfile() : pageSetup();
  if (state.me.status === "pending") return pagePending();
  if (state.me.status !== "active") return pageBlocked();
  if (store.mfaRequired && !(state.mfa && state.mfa.session)) return pageMfaSetup();
  if (isAdmin() && state.adminKeyReady === "rules") return pageRulesOld();
  if (isAdmin() && state.adminKeyReady === false) return pageAdminKeyCreate();
  busy(true);
  try {
    const [p1, p2] = path.split("/");
    if (p1 === "password") return pagePassword();
    if (p1 === "report") return await pageReport(q);
    if (p1 === "fleet") return await pageFleet(q);
    if (p1 === "archive") return await pageArchive(q);
    if (p1 === "admin") return await pageAdmin(p2, q);
    return await pageEntry(q);
  } catch (err) {
    if (err && err.code === "need-key") return pageUnitKey(err.mode, err.unit);
    console.error(err);
    shell(`<div class="panel"><h2>Помилка</h2><p>${esc(store.errUa(err))}</p><button class="btn" onclick="location.reload()">Оновити сторінку</button></div>`);
  } finally { busy(false); }
}

async function onAuthChanged(user) {
  if (state.registering) return;
  state.authUser = user; state.me = null; state.mfa = null;
  if (!user) forgetSessionKeys();
  if (user) {
    try { state.me = await store.getMe(); } catch (err) { state.me = null; }
    try { state.mfa = await store.mfaState(); } catch (err) { state.mfa = null; }
    if (state.me && state.me.status === "active" && (!store.mfaRequired || (state.mfa && state.mfa.session))) {
      store.touchLogin();
      if (state.me.role === "admin") {
        state.adminKeyReady = await store.hasAdminKey().catch((e) => (e && e.code === "permission-denied" ? "rules" : true));
        state.units = await store.listUnits().catch(() => []);
        state.pendingCount = (await store.listUsers().catch(() => [])).filter((u) => u.status === "pending").length;
        const ctx = sessionStorage.getItem("unitCtx");
        state.unitId = ctx && state.units.some((u) => u.id === ctx) ? ctx : state.me.unit_id;
      } else {
        state.unitId = state.me.unit_id;
        const u = await store.getUnit(state.me.unit_id).catch(() => null);
        state.units = u ? [u] : [];
      }
    }
  }
  const pth = route().path;
  if (!user && !["login", "register", "forgot"].includes(pth)) { location.hash = "#/login"; return; }
  if (user && ["login", "register", "forgot", ""].includes(pth)) { location.hash = "#/entry"; return; }
  render();
}

// ============================================================ старт
(async () => {
  try {
    if (DEMO) store = (await import("./store-mock.js?v=20261024a")).createStore();
    else if (FIREBASE_CONFIG && FIREBASE_CONFIG.apiKey) store = (await import("./store-firebase.js?v=20261024a")).createStore(FIREBASE_CONFIG);
  } catch (err) {
    $app.innerHTML = `<div class="wrap"><div class="panel"><h2>Не вдалося завантажити портал</h2><p>${esc(err.message)}</p></div></div>`;
    return;
  }
  if (!store) return render();
  window.addEventListener("hashchange", () => { window.onbeforeunload = null; render(); });
  const net = () => { const el = document.querySelector(".online"); if (!el) return; el.classList.toggle("off", !navigator.onLine); el.querySelector("b").textContent = navigator.onLine ? "ON-LINE" : "OFF-LINE"; };
  window.addEventListener("online", net); window.addEventListener("offline", net);
  store.onAuth(onAuthChanged);
})();
