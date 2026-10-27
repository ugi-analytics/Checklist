// Портал «Списання пального для МВО» — односторінковий застосунок (GitHub Pages + Firebase).
import * as C from "./core.js?v=20261027s";
import * as U from "./ui.js?v=20261027s";
import { litresWords } from "./words.js?v=20261027s";
import { buildRegister, buildAct } from "./docs.js?v=20261027s";
import { FIREBASE_CONFIG, ADMIN_EMAIL } from "./config.js?v=20261027s";
import * as X from "./crypto.js?v=20261027s";
import { qrSvg } from "../vendor/qrcode.js?v=20261027s";
import { buildXlsx } from "./xlsx.js?v=20261027s";
import { buildFleetRows, FLEET_WIDTHS, nextYm, dmy } from "./fleet_export.js?v=20261027s";

const $app = document.getElementById("app");
const esc = U.esc;
const params = new URLSearchParams(location.search);
const DEMO = params.has("demo");
let store = null;

const state = { authUser: null, me: null, unitId: null, units: [], flash: null, setupDone: null,
  mfa: null, adminPriv: null, keys: {}, adminKeyReady: null };
// ключі в пам'яті вкладки скидаються при виході
function forgetSessionKeys() { state.keys = {}; state.adminPriv = null; state.dispPriv = null; state.dispPub = undefined; state.dispCat = null; }
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
const firstOf = (k) => { const [y, m] = C.parseYm(k); return new Date(y, m - 1, 1); };
function downloadBytes(bytes, filename, mime) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([bytes], { type: mime })); a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
function pwProblem(pw) {
  if (pw.length < 8) return "Пароль має містити щонайменше 8 символів.";
  if (/^\d+$/.test(pw) || /^[^\d]+$/.test(pw)) return "Пароль має містити і літери, і цифри.";
  return null;
}
const isAdmin = () => state.me && state.me.role === "admin" && state.me.status === "active";
const isDisp = () => state.me && state.me.role === "dispatcher" && state.me.status === "active";
const ROLE_UA = { user: "користувач", admin: "адміністратор", dispatcher: "диспетчер" };
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
  const unitChip = "";  // назва підрозділу — у віконці користувача праворуч угорі
  const whoUnit = unit && me && me.status === "active" ? `<div class="who-unit">${esc(unit.name)}</div>` : "<br>";
  const switcher = isAdmin() && state.units.length > 1 ? `<select class="js-unit-switch" title="Працювати з підрозділом">
      ${state.units.map((u) => `<option value="${u.id}" ${u.id === state.unitId ? "selected" : ""}>${esc(u.name)}</option>`).join("")}</select>` : "";
  const tab = (h, icon, label, extra = "") => `<a href="#/${h}" class="tile ${p === h || p.startsWith(h + "/") || (h === "entry" && !p) ? "on" : ""}">
    <span class="ti">${icon}</span><span class="tl">${label}${extra}</span><span class="chev">›</span></a>`;
  const pending = isAdmin() && state.pendingCount ? `<span class="badge">${state.pendingCount}</span>` : "";
  const stats = isAdmin() && state.userStats ? `<div class="who-stats">${statsHtml()}</div>` : "";
  content = String(content).replace(/<h2 class="display">((?:\p{Extended_Pictographic}|[✓✍])\uFE0F?)\s*/gu, '<h2 class="display ttl"><span class="ico">$1</span>');
  $app.innerHTML = `
  <header class="hero"><div class="wrap">
    <div class="topbar"><div class="logo"><img src="img/hub_icon_small.png" alt=""></div>
      <div class="brand"><div class="bname"><span class="display">UGI-Analytics</span>
        <span class="online ${navigator.onLine ? "" : "off"}"><i></i><b>${navigator.onLine ? "ON-LINE" : "OFF-LINE"}</b></span></div>
        <small>Управління головної інспекції ГУНП в Полтавській області</small></div>
      <div class="sp"></div>
      ${state.authUser ? `<div class="who"><b>${esc(me?.full_name || state.authUser.email)}</b>${isAdmin() ? " · адміністратор" : isDisp() ? " · диспетчер" : ""}${isDisp() ? "<br>" : whoUnit}${stats}
        ${me && me.status === "active" ? `<a href="#/password">змінити пароль</a> · ` : ""}<a href="#" class="js-logout">вийти</a></div>` : ""}
    </div>
    <div class="title-row"><h1 class="display"><span>Списання пального</span> <span>для <em>МВО</em></span></h1>
      <div class="chips">${unitChip}${switcher}
      ${DEMO ? `<span class="chip" style="border-color:#f87171;color:#fecaca">ДЕМО-РЕЖИМ · дані не зберігаються</span>` : ""}</div></div>
    ${nav && isDisp() ? `<nav class="tiles">${tab("disp", "🧮", "Диспетчер · перевірка розрахунку")}</nav>` : ""}
    ${nav && me && me.status === "active" && !isDisp() ? `<nav class="tiles">
      ${tab("entry", "💧", "Введення даних")}${tab("report", "📋", "Звіт і акти")}${tab("fleet", "🚙", "Автопарк і реквізити")}${tab("archive", "📁", "Архів")}
      ${isAdmin() ? tab("admin", "🛡️", "Адміністрування", pending) : ""}</nav>` : ""}
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
      <div class="note">Вкажіть <b>особисту пошту</b> (Gmail, Ukr.net тощо), <mark>не службову @police.gov.ua</mark>: захист службової пошти
        автоматично відкриває посилання з листів і через це скасовує прив'язку телефона для входу з кодом.</div>
      <div class="two"><label class="f">Пароль * (від 8 символів, літери й цифри)<input name="pw" type="password" required minlength="8" autocomplete="new-password"></label>
        <label class="f">Повторіть пароль *<input name="pw2" type="password" required minlength="8" autocomplete="new-password"></label></div>
      <button class="btn primary">Надіслати заявку</button></form>
    <p class="muted" style="margin-top:14px"><a href="#/login">← Повернутися до входу</a></p>`, 620), { nav: false });
  $app.querySelector(".js-reg").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.pw.value !== f.pw2.value) { flash("err", "Паролі не збігаються."); return pageRegister(); }
    const pp = pwProblem(f.pw.value); if (pp) { flash("err", pp); return pageRegister(); }
    if (/police\.gov\.ua$/i.test(f.email.value.trim())) { flash("err", "Службова пошта @police.gov.ua не підходить для порталу — вкажіть особисту (Gmail, Ukr.net тощо)."); return pageRegister(); }
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
    <div class="note">Під час першого входу після підтвердження портал попросить прив'язати телефон (застосунок Google Authenticator або Microsoft Authenticator) — це робиться один раз.</div>
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
    // телефон уже прив'язано, а цей браузер пам'ятає старий вхід без коду — виходимо й просимо увійти з кодом
    flash("info", "Вхід із кодом для цього облікового запису вже налаштовано. Введіть пароль — далі портал попросить 6-значний код із телефона.");
    forgetSessionKeys(); store.signOut();
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
      <div class="note">Це робиться <b>один раз</b> — під час першого входу. Надалі після пароля портал питатиме лише 6-значний код із застосунку.</div>
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
  const unit = await decUnit(u, dk);
  publishCatalog(unit);
  return unit;
}
// довідник авто для диспетчера: лише назва, норма, надбавка, пальне (без показників за місяці);
// публікується, коли змінився довідник або ключ диспетчера
async function publishCatalog(unit) {
  try {
    if (isDisp()) return;
    if (state.dispPub === undefined) state.dispPub = await store.getDispPub();
    const pub = state.dispPub; if (!pub) return;
    const cat = { unit: unit.name, org: unit.settings.org || "", vehicles: unit.vehicles.filter((v) => v.active !== false)
      .map((v) => ({ id: v.id, name: v.name, rate: v.rate, winter_pct: v.winter_pct, month_norm_km: v.month_norm_km, fuel: v.fuel })) };
    const h = await X.sha256(pub.kid + "|" + JSON.stringify(cat));
    const cur = await store.getFleetCat(unit.id);
    if (cur && cur.h === h) return;
    const sealed = await X.sealForDisp(cat, pub, X.aadFleet(unit.id));
    await store.saveFleetCat(unit.id, { ...sealed, h, unit_name: unit.name, n: cat.vehicles.length });
  } catch (err) { console.warn("Довідник для диспетчера не оновлено:", err && err.message); }
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
  await publishCatalog({ ...unit, settings, vehicles: vehicles.map(C.normalizeVehicle) });
}
async function rememberKey(unitId, dk, remember) {
  const locked = await X.lockDown(dk);
  state.keys[unitId] = locked;
  if (remember) await X.keyCache.put(keyCacheId(unitId), locked);
}
async function monthState(unit, ym) {
  const saved = await readMonth(unit, ym);
  let entries, carried = null;
  const sources = C.monthSources(unit.settings, saved && saved.sources);
  if (saved) entries = saved.entries || {};
  else {
    const prev = await readMonth(unit, C.prevYm(ym));
    const prevSources = C.monthSources(unit.settings, prev && prev.sources);
    entries = prev ? C.carryOver(unit.vehicles, { entries: prev.entries || {}, snapshot: prev.snapshot || {} }, sources, prevSources) : {};
    if (prev) carried = C.prevYm(ym);
  }
  const vehicles = C.monthVehicles(unit.vehicles, { entries, snapshot: (saved && saved.snapshot) || {} }, sources);
  return { saved, entries, carried, vehicles, sources };
}
async function rowsFor(unit, ym) {
  const st = await monthState(unit, ym);
  const rows = st.vehicles.map((v) => C.calcRow(v, st.entries[v.id], st.sources));
  rows.sources = st.sources;
  return rows;
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
  const SRC = st.sources;
  const simple = !!SRC.legacy;          // стандартні джерела: картка як раніше
  const inp = (name, val, ph, step = "0.01", cls = "") => `<input type="number" inputmode="decimal" min="0" step="${step}" name="${name}" placeholder="${ph}" value="${val ?? ""}"${cls ? ` class="${cls}"` : ""}>`;
  const lv = (c) => (c ? C.cToNum(c) : "");
  const mainTitle = (r) => { const m = r.sources.find((x) => x.id === r.mainId); return m ? C.sourceTitle(m) : ""; };
  const tiles = (v, r, e) => {
    if (simple) {
      return `<div class="f4">${r.sources.map((x) => `<div><label title="${esc(C.sourceTitle(x))}">${esc(C.sourceTitle(x))}</label>${inp("iss." + x.id, lv(r.issued[x.id]), "—")}</div>`).join("")}</div>`;
    }
    // інші активні джерела (інший вид пального) — сховані, відкриваються кнопкою
    const extra = SRC.filter((x) => x.active && !r.sources.some((y) => y.id === x.id));
    return `<div class="srcs">${[...r.sources, ...extra].map((x) => `<div class="srct${x.id === r.mainId ? " main" : ""}${r.sources.includes(x) ? "" : " extra"}" data-sid="${x.id}">
        <div class="st" title="${esc(C.sourceTitle(x))}">${esc(C.sourceTitle(x))}${x.id === r.mainId ? ` <span class="mb">основне</span>` : ""}</div>
        <div class="sr"><div><label>Видано</label>${inp("iss." + x.id, lv(r.issued[x.id]), "—")}</div>
          <div class="sp"><label>Початок</label><b class="js-s">${C.fmtC(r.startBy[x.id] || 0)}</b></div></div>
        <div class="sc"><span>Витрачено <b class="js-c">${C.fmtC(r.consBy[x.id] || 0)}</b>${inp("cons." + x.id, e.split && e.cons ? e.cons[x.id] : lv(r.consBy[x.id]), "0", "0.01", "js-ci")}</span>
          <span>Кінець <b class="js-e">${C.fmtC(r.endBy[x.id] || 0)}</b></span></div></div>`).join("")}</div>
      ${extra.length ? `<button type="button" class="btn sm js-more" style="margin-top:6px">＋ інші джерела (${extra.length}) — з іншим видом пального</button>` : ""}
      ${endTxt(r, e)}`;
  };
  // довільний опис залишку (напр. БКС: «125 А-95 с/к UKRNAFTA Мосійченко + 197 … Левошко») — для авто, де кілька МВО
  const multiMvo = (r) => !simple && new Set([...C.vehicleMvos(r.sources), ...C.vehicleMvos(SRC.filter((x) => x.active))]).size > 1;
  const txt = (name, val, ph, rows = 2) => `<textarea name="${name}" rows="${rows}" placeholder="${esc(ph)}" class="vtxt">${esc(val || "")}</textarea>`;
  const startTxt = (r, e) => (multiMvo(r) || e.start_text ? `<div class="mvo" style="margin-top:6px"><label title="Що стоїть у графі «вид» залишку на початок у «Додатку 6» — пишіть як потрібно">Вид залишку на початок / МВО (довільно)</label>
      ${txt("start_text", e.start_text, "напр. 125 А-95 с/к UKRNAFTA Мосійченко + 197 А-95 BVS Левошко", 3)}</div>` : "");
  const endTxt = (r, e) => (multiMvo(r) || e.end_text ? `<div class="mvo endt"><label title="Якщо порожньо — рахується автоматично за джерелами">Вид залишку на кінець / МВО (довільно; якщо порожньо — автоматично)</label>
      ${txt("end_text", e.end_text, C.balanceText({ ...r, entry: { ...e, end_text: "" } }, "end", unit.settings))}</div>` : "");
  const cards = st.vehicles.map((v) => {
    const e = entries[v.id] || {};
    const r = C.calcRow(v, e, SRC); const s = C.statusOf(r, e);
    return `<div class="frc ${s}${e.split ? " split" : ""}" data-vid="${v.id}"><span class="saving">збережено ✓</span><div class="grid${simple ? "" : " ext"}">
      <div class="js-id">${U.idHtml(v, s, mainTitle(r))}</div>
      <div><div class="fr-h">Пробіг і залишок на початок</div>${simple ? `<div class="f3">
        <div><label>Початок, км</label>${inp("odo_start", e.odo_start, "км", "1")}</div>
        <div><label>Кінець, км</label>${inp("odo_end", e.odo_end, "км", "1")}</div>
        <div><label>Залишок, л</label>${inp("start." + r.mainId, lv(r.startC), "л")}</div></div>` : `<div class="f3">
        <div><label>Початок, км</label>${inp("odo_start", e.odo_start, "км", "1")}</div>
        <div><label>Кінець, км</label>${inp("odo_end", e.odo_end, "км", "1")}</div>
        <div><label>Залишок, л</label>${inp("fuel_start", lv(r.startC), "л")}</div></div>${startTxt(r, e)}`}
        <div class="js-km">${r.empty ? "" : U.kmHtml(v, r)}</div></div>
      <div><div class="fr-h">${simple ? "Видано палива, л" : "Пальне за джерелами, л"}</div>${tiles(v, r, e)}
        <div class="notes"><div class="fr-note js-note">${U.noteHtml(v, r, e)}</div>
          <div class="manual"><label style="display:flex;gap:6px;align-items:center;margin:0;color:#cbd5e1"><input type="checkbox" name="manual" ${e.manual ? "checked" : ""}> ввести витрату вручну</label>
            <input type="number" inputmode="decimal" min="0" step="0.01" name="consumed_manual" placeholder="витрачено, л" value="${e.consumed_manual ?? ""}" ${e.manual ? "" : 'style="display:none"'}>
            ${simple || r.sources.length < 2 ? "" : `<label style="display:flex;gap:6px;align-items:center;margin:0;color:#cbd5e1" title="Вказати, скільки витрачено з кожного джерела (МВО)"><input type="checkbox" name="split" ${e.split ? "checked" : ""}> розподілити витрату вручну</label>`}</div></div></div>
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
  const readCard = (card) => {
    const e = { start: {}, iss: {}, cons: {}, start_mvo: {} };
    card.querySelectorAll("input[name], textarea[name]").forEach((i) => {
      const val = i.type === "checkbox" ? i.checked : i.value;
      const dot = i.name.indexOf(".");
      if (dot > 0) e[i.name.slice(0, dot)][i.name.slice(dot + 1)] = val; else e[i.name] = val;
    });
    return C.cleanEntry(e);
  };
  const calc = (v) => C.calcRow(v, entries[v.id], SRC);
  const refreshKpi = () => {
    const rows = st.vehicles.map(calc);
    const sts = rows.map((r, i) => C.statusOf(r, entries[st.vehicles[i].id] || {}));
    $app.querySelector("#kpi").innerHTML = U.kpiHtml(rows, sts, C.MONTHS_UA[C.parseYm(ym)[1] - 1]);
  };
  const refreshCard = (card) => {
    const v = vmap[card.dataset.vid]; const e = entries[v.id] || {};
    const r = calc(v); const s = C.statusOf(r, e);
    card.className = "frc " + s + (e.split ? " split" : "") + (card.classList.contains("showall") ? " showall" : "");
    card.querySelector(".js-id").innerHTML = U.idHtml(v, s, mainTitle(r));
    card.querySelector(".js-km").innerHTML = r.empty ? "" : U.kmHtml(v, r);
    card.querySelector(".js-note").innerHTML = U.noteHtml(v, r, e);
    card.querySelector(".js-tank").innerHTML = U.tankHtml(r, s);
    card.querySelector(".js-msgs").innerHTML = U.msgsHtml(r);
    const et = card.querySelector("[name=end_text]"); if (et) et.placeholder = C.balanceText({ ...r, entry: { ...e, end_text: "" } }, "end", unit.settings);
    card.querySelectorAll(".srct").forEach((t) => {
      const id = t.dataset.sid;
      t.querySelector(".js-c").textContent = C.fmtC(r.consBy[id] || 0);
      t.querySelector(".js-e").textContent = C.fmtC(r.endBy[id] || 0);
      t.querySelector(".js-s").textContent = C.fmtC(r.startBy[id] || 0);
      t.classList.toggle("neg", (r.endBy[id] || 0) < 0);
    });
  };
  let timer = null; let dirtyCards = new Set();
  const saveNow = async () => {
    const touched = [...dirtyCards]; dirtyCards = new Set();
    const ss = $app.querySelector(".js-savestate");
    try {
      const snapshot = Object.fromEntries(st.vehicles.map((v) => [v.id, v]));
      await writeMonth(unit, ym, { entries, snapshot, sources: SRC.map((x) => ({ ...x })) });
      touched.forEach((c) => { const s = c.querySelector(".saving"); s.classList.add("on"); setTimeout(() => s.classList.remove("on"), 1200); });
      if (ss) ss.textContent = "Збережено ✓ " + new Date().toLocaleTimeString("uk-UA");
    } catch (err) { if (ss) ss.textContent = "⚠️ Не збережено: " + store.errUa(err); alert("Не вдалося зберегти: " + store.errUa(err)); }
  };
  const schedule = (card, ms) => { dirtyCards.add(card); clearTimeout(timer); timer = setTimeout(saveNow, ms); };
  cardEls.forEach((card) => {
    card.querySelector(".js-more")?.addEventListener("click", (ev) => {
      const on = card.classList.toggle("showall");
      ev.target.textContent = on ? "－ сховати інші джерела" : ev.target.dataset.t;
    });
    const mb = card.querySelector(".js-more"); if (mb) mb.dataset.t = mb.textContent;
    card.querySelectorAll("input, textarea").forEach((inp) => {
      const mark = () => { if (inp.type !== "checkbox") inp.classList.toggle("filled", inp.value !== ""); };
      mark();
      const onChange = (ms) => {
        mark();
        if (inp.name === "manual") card.querySelector("[name=consumed_manual]").style.display = inp.checked ? "" : "none";
        if (inp.name === "split" && inp.checked) {
          // ручний розподіл стартує з поточного автоматичного
          const r = calc(vmap[card.dataset.vid]);
          card.querySelectorAll(".js-ci").forEach((ci) => { const id = ci.name.slice(5); ci.value = r.consBy[id] ? C.cToNum(r.consBy[id]) : ""; });
        }
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
    const co = C.carryOver(unit.vehicles, { entries: prev.entries || {}, snapshot: prev.snapshot || {} }, SRC, C.monthSources(unit.settings, prev.sources));
    for (const card of cardEls) {
      const e = co[card.dataset.vid]; if (!e) continue;
      card.querySelector("[name=odo_start]").value = e.odo_start ?? "";
      card.querySelectorAll("input[name^='start.']").forEach((i) => { const id = i.name.slice(6); i.value = e.start && e.start[id] ? e.start[id] : ""; });
      const fs = card.querySelector("[name=fuel_start]"); if (fs) fs.value = e.fuel_start ?? "";
      const stx = card.querySelector("[name=start_text]"); if (stx) stx.value = e.start_text || "";
      if (simple) { const i = card.querySelector("input[name^='start.']"); if (i) i.value = Object.values(e.start || {}).reduce((a, b) => a + b, 0) || ""; }
      card.querySelectorAll("input, textarea").forEach((i) => { if (i.type !== "checkbox") i.classList.toggle("filled", i.value !== ""); });
      entries[card.dataset.vid] = readCard(card); refreshCard(card); dirtyCards.add(card);
    }
    refreshKpi(); await saveNow();
  });
  $app.querySelector(".js-reg").addEventListener("click", async () => {
    clearTimeout(timer); if (dirtyCards.size) await saveNow();
    const rows = st.vehicles.map(calc);
    await download(buildRegister(window.docx, unit.settings, ym, rows, SRC), `Додаток_6_Реєстр_пального_${ym}.docx`);
    store.log("doc_register", ym, unit.id);
  });
  window.onbeforeunload = () => (dirtyCards.size ? "Є незбережені зміни" : undefined);
}

// ============================================================ диспетчер
// Бачить лише довідник авто всіх підрозділів і рахує витрату на своєму екрані: нічого не зберігається і не змінюється.
async function pageDisp(q) {
  const pub = await store.getDispPub();
  const cacheId = pub ? `${DEMO ? "demo" : FIREBASE_CONFIG.projectId}:${state.authUser.uid}:disp:${pub.kid}` : "";
  if (!pub) {
    shell(`<div class="auth" style="max-width:640px"><div class="panel"><h2 class="display">🔑 Створіть ключ диспетчера</h2>
      <div class="note"><p>Довідники авто підрозділів зашифровані. Ключ диспетчера — це фраза, якою відкриваються <b>лише довідники авто</b> (назва, норма, надбавка, пальне).
        Показників підрозділів за місяці диспетчер не бачить.</p><p>Ключ знаєте лише ви (та інші диспетчери, яким ви його передасте). Адміністратор його не бачить; якщо ключ забуто — він скидає його, і ви створюєте новий.</p></div>
      <form class="js-dnew" autocomplete="off"><label class="f">Ключ диспетчера (від 10 символів, краще фраза)<input name="k1" type="password" required></label>
        <label class="f">Повторіть<input name="k2" type="password" required></label>${rememberBox}<button class="btn primary">Створити ключ</button></form></div></div>`);
    $app.querySelector(".js-dnew").addEventListener("submit", async (e) => {
      e.preventDefault(); const f = e.target;
      const prob = X.passProblem(f.k1.value, "Ключ диспетчера"); if (prob) { flash("err", prob); return render(); }
      if (f.k1.value !== f.k2.value) { flash("err", "Ключі не збігаються."); return render(); }
      busy(true);
      try {
        const keys = await X.newDispKeys(f.k1.value);
        await store.createDispKeys(keys);
        const priv = await X.unlockDisp(keys.priv, f.k1.value); state.dispPriv = priv;
        if (f.remember.checked) await X.keyCache.put(`${DEMO ? "demo" : FIREBASE_CONFIG.projectId}:${state.authUser.uid}:disp:${keys.pub.kid}`, priv);
        store.log("disp_key_create", ""); flash("ok", "Ключ диспетчера створено. Довідники авто з'являться, коли підрозділи відкриють портал (або адміністратор відкриє «Зведення по області»).");
      } catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
      render();
    });
    return;
  }
  if (!state.dispPriv) state.dispPriv = await X.keyCache.get(cacheId);
  if (!state.dispPriv) {
    shell(`<div class="auth" style="max-width:560px"><div class="panel"><h2 class="display">🔑 Ключ диспетчера</h2>
      <p class="muted">Довідники авто зашифровані. Введіть ключ диспетчера.</p>
      <form class="js-dopen" autocomplete="off"><label class="f">Ключ диспетчера<input name="k" type="password" required autofocus></label>${rememberBox}
        <button class="btn primary">Відкрити</button></form>
      <p class="muted" style="margin-top:12px">Ніхто не пам'ятає ключ? Адміністратор може його скинути — тоді створите новий.</p></div></div>`);
    $app.querySelector(".js-dopen").addEventListener("submit", async (e) => {
      e.preventDefault(); const f = e.target; busy(true);
      try {
        const priv = await X.unlockDisp(await store.getDispPriv(), f.k.value); state.dispPriv = priv;
        if (f.remember.checked) await X.keyCache.put(cacheId, priv);
      } catch (err) { flash("err", store.errUa(err)); } finally { busy(false); }
      render();
    });
    return;
  }
  // довідники всіх підрозділів
  if (!state.dispCat || q.get("reload")) {
    const docs = await store.listFleetCats();
    const cats = [];
    let stale = 0;
    for (const d of docs) {
      if (d.kid !== pub.kid) { stale++; continue; }
      try { const c = await X.openForDisp(d, state.dispPriv, X.aadFleet(d.unit_id)); cats.push({ unit_id: d.unit_id, unit: c.unit || d.unit_name, org: c.org, vehicles: c.vehicles || [], updated: d.updated_at }); }
      catch { stale++; }
    }
    cats.sort((a, b) => a.unit.localeCompare(b.unit, "uk"));
    state.dispCat = { cats, stale };
  }
  const { cats, stale } = state.dispCat;
  const all = cats.flatMap((c) => c.vehicles.map((v) => ({ ...v, unit: c.unit, unit_id: c.unit_id })));
  const sel = q.get("v") ? all.find((v) => `${v.unit_id}:${v.id}` === q.get("v")) : null;
  const rowHtml = (v) => `<tr class="js-pick ${sel && sel === v ? "on" : ""}" data-k="${esc(v.unit_id + ":" + v.id)}" data-s="${esc((v.name + " " + v.unit + " " + (v.fuel || "")).toLowerCase())}">
    <td><b>${esc(v.name)}</b></td><td class="muted">${esc(v.unit)}</td><td>${esc(v.fuel || "")}</td><td class="num">${esc(String(v.rate ?? ""))}</td><td class="num">${v.winter_pct ? esc(String(v.winter_pct)) + "%" : "—"}</td></tr>`;
  const calc = sel ? `<div class="panel disp-calc"><h2>🧮 ${esc(sel.name)} <span class="muted" style="font-size:.9rem">· ${esc(sel.unit)}</span></h2>
    <p class="muted">Розрахунок лише на цьому екрані — нікуди не зберігається і не змінює даних підрозділу.</p>
    <form class="js-dcalc grid-calc" autocomplete="off">
      <label class="f">Пробіг на початок, км<input name="odo_start" type="number" min="0" step="1"></label>
      <label class="f">Пробіг на кінець, км<input name="odo_end" type="number" min="0" step="1"></label>
      <label class="f">Норма, л/100 км<input name="rate" type="number" min="0" step="0.01" value="${esc(String(sel.rate ?? ""))}"></label>
      <label class="f">Зимова надбавка, %<input name="winter_pct" type="number" min="0" max="100" step="1" value="${esc(String(sel.winter_pct ?? 0))}"></label>
      <label class="f">Залишок на початок, л<input name="fuel_start" type="number" min="0" step="0.01"></label>
      <label class="f">Видано, л<input name="issued" type="number" min="0" step="0.01"></label>
    </form>
    <div class="js-dres disp-res"></div>
    <button class="btn sm js-dclear">Очистити</button></div>` : "";
  shell(`<div class="panel" style="margin-top:14px"><h2 class="display">🧮 Перевірка розрахунку</h2>
      <div class="row" style="justify-content:space-between"><input class="js-dq" placeholder="🔍 номер, марка або підрозділ" style="min-width:300px;flex:1;max-width:520px">
        <span class="muted">Авто: <b>${all.length}</b> · підрозділів: <b>${cats.length}</b>${stale ? ` · ще не оновлено: ${stale}` : ""} <a href="#/disp?reload=1" class="btn sm">⟳ Оновити</a></span></div>
      ${all.length ? "" : `<div class="flash info" style="margin-top:10px">Довідників авто ще немає. Вони з'являються, коли хтось із підрозділу відкриває портал (або адміністратор відкриває «Зведення по області»).</div>`}
    </div>
    ${calc}
    <div class="panel"><div class="tblwrap" style="max-height:${sel ? 360 : 640}px;overflow:auto"><table class="list js-dlist"><thead><tr><th>Авто</th><th>Підрозділ</th><th>Пальне</th><th class="num">Норма</th><th class="num">Надбавка</th></tr></thead>
      <tbody>${all.map(rowHtml).join("")}</tbody></table></div></div>`);
  const qi = $app.querySelector(".js-dq");
  const filter = () => { const t = qi.value.trim().toLowerCase(); $app.querySelectorAll(".js-pick").forEach((tr) => { tr.style.display = !t || tr.dataset.s.includes(t) ? "" : "none"; }); };
  qi.value = sessionStorage.getItem("dispQ") || ""; filter();
  qi.addEventListener("input", () => { sessionStorage.setItem("dispQ", qi.value); filter(); });
  $app.querySelectorAll(".js-pick").forEach((tr) => tr.addEventListener("click", () => go(`#/disp?v=${encodeURIComponent(tr.dataset.k)}`)));
  if (sel) {
    const f = $app.querySelector(".js-dcalc"), res = $app.querySelector(".js-dres");
    const run = () => {
      const val = (n) => f[n].value;
      const v = { ...sel, rate: val("rate"), winter_pct: val("winter_pct"), source: "smk" };
      const r = C.calcRow(v, { odo_start: val("odo_start"), odo_end: val("odo_end"), fuel_start: val("fuel_start"), smk: val("issued") }, C.unitSources({}));
      const coef = 1 + (Number(val("winter_pct")) || 0) / 100;
      res.innerHTML = r.empty && !r.km ? `<p class="muted">Введіть пробіг на початок і кінець.</p>` : `<div class="disp-kv">
        <div><span>Пробіг</span><b>${C.fmtKm(r.km)} км</b></div>
        <div><span>Норма з надбавкою</span><b>${C.fmtC(Math.round(C.toC(val("rate")) * coef))} л/100 км</b></div>
        <div><span>За нормою (точно)</span><b>${C.fmtC(r.normExactC)} л</b></div>
        <div class="hl"><span>Витрата (округлено до цілих)</span><b>${C.fmtC(r.consC)} л</b></div>
        <div><span>Залишок на початок + видано</span><b>${C.fmtC(r.startC)} + ${C.fmtC(r.issuedC)} л</b></div>
        <div class="hl"><span>Залишок на кінець</span><b>${C.fmtC(r.endC)} л</b></div></div>${U.msgsHtml(r)}
        <p class="muted" style="margin-top:6px">${C.fmtKm(r.km)} км × ${esc(val("rate") || "0")} л/100 км${coef !== 1 ? ` × ${coef.toFixed(2).replace(".", ",")}` : ""} = ${C.fmtC(r.normExactC)} л → ${C.fmtC(r.normC)} л</p>`;
    };
    f.addEventListener("input", run); run();
    $app.querySelector(".js-dclear").addEventListener("click", () => { ["odo_start", "odo_end", "fuel_start", "issued"].forEach((n) => { f[n].value = ""; }); f.rate.value = sel.rate ?? ""; f.winter_pct.value = sel.winter_pct ?? 0; run(); });
    f.odo_start.focus();
  }
}

// ============================================================ звіт і акти
async function pageReport(q) {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  const ym = currentYm(q);
  const rows = await rowsFor(unit, ym);
  const S = unit.settings;
  const nErr = rows.reduce((a, r) => a + r.errors.length, 0);
  const lay = C.registerLayout(rows, rows.sources, S);
  const actList = C.actList(rows, S);
  const acts = actList.map((a, i) => {
    const tot = a.lines.reduce((x, l) => x + l.qtyC, 0);
    return `<div class="act"><h3>Акт «${esc(a.title)}»</h3><div class="big">${C.fmtC(tot)} л</div>
      ${a.lines.length ? `<ul>${a.lines.map((l) => `<li>${esc(l.fuel)} ${esc(l.supplier)} — ${C.fmtC(l.qtyC)} л (МВО ${esc(l.mvo)})</li>`).join("")}</ul>
        <p class="muted">«${litresWords(tot)}»</p><button class="btn sky js-act" data-i="${i}">📥 Завантажити акт (Word)</button>`
        : `<p class="muted">Немає пального до списання.</p>`}</div>`;
  }).join("");
  const multi = actList.some((a) => a.perMvo);
  shell(`<div style="margin-top:14px">${monthPicker(ym, "#/report")}</div>
    <div class="panel"><h2>${esc(S.title)} за ${C.monthTitle(ym)}</h2>
      ${nErr ? `<div class="flash err">У даних є помилки (${nErr}) — перевірте картки на вкладці «Введення даних».</div>` : ""}
      ${rows.length ? `<div class="tblwrap">${U.previewHtml(rows, S, rows.sources)}</div>
        <p class="muted">* дані графи № ${lay.issTotalNo} повинні відповідати даним граф № 11, 12 звіту про одержане і видане пальне.<br>
          * дані графи № ${lay.consTotalNo} повинні відповідати кількості пального в акті списання.<br>
          Підпис особи, яка склала реєстр: ${esc(S.compiler || "— (вкажіть у «Автопарк і реквізити»)")}</p>
        <button class="btn primary js-reg">📥 Завантажити реєстр «Додаток 6» (Word)</button>` : `<p class="muted">Немає даних за цей місяць.</p>`}</div>
    <div class="panel"><h2>🧾 Акти списання пального</h2>
      <p class="muted">Окремі акти: «Талони + см/к» (талони, смарт-картки) та «АЗС» (АЗС ГУНП талони, АЗС налив)${multi ? " — і окремо на кожну МВО (підписує вся комісія та ця МВО)" : ""}.
        ${rows.sources.legacy ? "Витрата авто розноситься між ними пропорційно до виданого за місяць; якщо нічого не видано — за основним джерелом списання."
          : "Кількість у акті — це витрачене з відповідних джерел (графи «Витрачено»)."}
        Номер, дати, ціна й сума лишаються порожніми.</p><div class="acts">${acts}</div></div>`);
  bindMonthPicker();
  $app.querySelector(".js-reg")?.addEventListener("click", async () => {
    await download(buildRegister(window.docx, S, ym, rows, rows.sources), `Додаток_6_Реєстр_пального_${ym}.docx`); store.log("doc_register", ym, unit.id);
  });
  $app.querySelectorAll(".js-act").forEach((b) => b.addEventListener("click", async () => {
    const a = actList[Number(b.dataset.i)];
    await download(buildAct(window.docx, S, ym, a), actFileName(a, ym));
    store.log("doc_act", `${a.group}${a.mvo ? " " + a.mvo : ""} ${ym}`, unit.id);
  }));
}
const actFileName = (a, ym) => `Акт_списання_${C.ACT_GROUPS[a.group].file}${a.perMvo && a.mvo ? "_" + String(a.mvo).replace(/[^\p{L}\d]+/gu, "_") : ""}_${ym}.docx`;

// ============================================================ архів
async function pageArchive() {
  const unit = await loadUnit();
  if (!unit) return shell(`<div class="panel">Підрозділ не знайдено.</div>`);
  const months = (await listMonthsDec(unit)).reverse();
  const recs = months.map((m) => {
    const sources = C.monthSources(unit.settings, m.sources);
    const vehicles = C.monthVehicles(unit.vehicles, { entries: m.entries || {}, snapshot: m.snapshot || {} }, sources);
    const rows = vehicles.map((v) => C.calcRow(v, (m.entries || {})[v.id], sources));
    rows.sources = sources;
    const acts = C.actList(rows, unit.settings).filter((a) => a.lines.length);
    return { ym: m.ym, rows, t: C.totals(rows), acts };
  });
  shell(`<div class="panel"><h2>🗂️ Архів місяців</h2>${recs.length ? `<table class="list"><thead><tr><th>Місяць</th><th class="num">Авто</th>
    <th class="num">Пробіг, км</th><th class="num">Видано, л</th><th class="num">Витрачено, л</th><th class="num">Залишок, л</th><th>Документи</th></tr></thead><tbody>
    ${recs.map((r) => `<tr><td><a href="#/entry?ym=${r.ym}">${C.monthTitle(r.ym)}</a></td><td class="num">${r.rows.length}</td><td class="num">${C.fmtKm(r.t.km)}</td>
      <td class="num">${C.fmtC(r.t.issuedC)}</td><td class="num">${C.fmtC(r.t.consC)}</td><td class="num">${C.fmtC(r.t.endC)}</td>
      <td><button class="btn sm js-doc" data-ym="${r.ym}" data-k="reg">Реєстр</button>
        ${r.acts.map((a, i) => `<button class="btn sm js-doc" data-ym="${r.ym}" data-k="${i}">Акт «${esc(a.title)}»</button>`).join(" ")}</td></tr>`).join("")}
    </tbody></table>` : `<p class="muted">Збережених місяців ще немає.</p>`}</div>`);
  $app.querySelectorAll(".js-doc").forEach((b) => b.addEventListener("click", async () => {
    const r = recs.find((x) => x.ym === b.dataset.ym); const k = b.dataset.k;
    if (k === "reg") await download(buildRegister(window.docx, unit.settings, r.ym, r.rows, r.rows.sources), `Додаток_6_Реєстр_пального_${r.ym}.docx`);
    else { const a = r.acts[Number(k)]; await download(buildAct(window.docx, unit.settings, r.ym, a), actFileName(a, r.ym)); }
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
        <th>Норма пробігу, км/міс</th><th>Вид пального</th><th>АЗС</th><th title="Куди списується витрата в першу чергу">Основне джерело</th><th style="min-width:170px">МВО</th><th>Активне</th><th></th></tr></thead><tbody></tbody></table></div>
      <button class="btn sm js-addveh" style="margin-top:10px">➕ Додати авто</button>
      <datalist id="fuels"><option value="А-95"><option value="А-92"><option value="ДП"><option value="ГАЗ"></datalist>
      <datalist id="suppliers">${C.SUPPLIERS.map((x) => `<option value="${x}">`).join("")}</datalist></div>
    <div class="panel"><h2>⛽ Джерела пального</h2>
      <p class="muted">Звідки підрозділ отримує пальне: кожне джерело — окрема колонка «Видано» й «Витрачено» в «Додатку 6» і окремий залишок у баку.
        Невеликим підрозділам досить стандартних джерел. Якщо пальне видають кілька МВО, різні АЗС чи постачальники — увімкніть свої джерела.</p>
      <div class="row" style="gap:18px;margin-bottom:10px">
        <label style="display:flex;gap:8px;align-items:center"><input type="radio" name="srcmode" value="std" ${s.sources && s.sources.length ? "" : "checked"}> Стандартні (АЗС ГУНП талони, Талони, см/к, АЗС налив)</label>
        <label style="display:flex;gap:8px;align-items:center"><input type="radio" name="srcmode" value="custom" ${s.sources && s.sources.length ? "checked" : ""}> Свої джерела (кілька МВО, АЗС, постачальників)</label></div>
      <div class="js-srcbox">
        <div class="tblwrap"><table class="edit" id="srcs"><thead><tr><th>Вид пального</th><th>Форма</th><th>Постачальник</th><th>МВО</th>
          <th style="min-width:220px">Назва колонки в реєстрі</th><th>Активне</th><th></th></tr></thead><tbody></tbody></table></div>
        <button class="btn sm js-addsrc" style="margin-top:10px">➕ Додати джерело</button>
        <p class="muted" style="margin-top:8px">Порожній «Вид пального» — джерело для всіх авто. Порожня «Назва колонки» складається сама: «А-95 с/к UKRNAFTA Мосійченко».
          Акти списання: с/к і талони — акт «Талони + см/к», АЗС ГУНП талони й АЗС налив — акт «АЗС»; якщо МВО кілька — окремий акт на кожну.</p></div>
      <h3 style="margin-top:14px">«Додаток 6»</h3>
      <div class="row" style="gap:18px">
        <label class="f">Колонки джерел<select data-k="reg_columns"><option value="all" ${s.reg_columns === "used" ? "" : "selected"}>усі налаштовані джерела</option>
          <option value="used" ${s.reg_columns === "used" ? "selected" : ""}>лише ті, по яких у місяці був рух</option></select></label>
        <label style="display:flex;gap:8px;align-items:center"><input type="checkbox" class="js-blago" ${s.show_blago === false ? "" : "checked"}> колонка «благо» у «Витрачено»</label>
        <label style="display:flex;gap:8px;align-items:center" title="Інакше — спочатку з основного джерела авто (залишок + видане), потім з інших"><input type="checkbox" class="js-isfirst" ${s.cons_issued_first ? "checked" : ""}> списувати спочатку видане за місяць, потім залишок</label></div></div>
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
      <td><select data-f="source" data-cur="${esc(v.source || "smk")}"></select></td>
      <td><div class="mvopick"><input data-f="mvo_text" value="${esc(v.mvo_text)}" placeholder="за замовчуванням — хто склав" title="Можна кілька МВО через кому"><button type="button" class="btn sm js-mvo" title="Вибрати одну або кілька МВО">▾</button><div class="mvomenu"></div></div></td>
      <td class="c"><input data-f="active" type="checkbox" ${v.active === false ? "" : "checked"}></td>
      <td><button class="btn sm danger js-del" title="Видалити з довідника">🗑</button></td>`;
    tr.querySelector(".js-del").addEventListener("click", () => { if (confirm("Видалити авто з довідника?")) tr.remove(); });
    bindMvoPick(tr);
    vbody.appendChild(tr);
  };
  // вибір кількох МВО для авто: прізвища з інших авто, джерел і «хто склав»
  const mvoParts = (t) => String(t || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  const knownMvos = () => {
    const out = new Set();
    vbody.querySelectorAll("[data-f=mvo_text]").forEach((i) => mvoParts(i.value).forEach((x) => out.add(x)));
    const comp = $app.querySelector("[data-k=compiler]"); mvoParts(comp && comp.value).forEach((x) => out.add(x));
    [...sbody.querySelectorAll("[data-f=mvo]")].forEach((i) => { const x = i.value.trim(); if (x && ![...out].some((y) => y.startsWith(x))) out.add(x); });
    return [...out].sort((a, b) => a.localeCompare(b, "uk"));
  };
  function bindMvoPick(tr) {
    const inpM = tr.querySelector("[data-f=mvo_text]"), menu = tr.querySelector(".mvomenu");
    tr.querySelector(".js-mvo").addEventListener("click", () => {
      if (menu.classList.toggle("on")) {
        const cur = mvoParts(inpM.value);
        menu.innerHTML = knownMvos().map((x) => `<label><input type="checkbox" value="${esc(x)}" ${cur.includes(x) ? "checked" : ""}> ${esc(x)}</label>`).join("") ||
          `<span class="muted">Впишіть прізвище в поле</span>`;
        menu.querySelectorAll("input").forEach((c) => c.addEventListener("change", () => {
          const now = mvoParts(inpM.value).filter((x) => x !== c.value);
          if (c.checked) now.push(c.value);
          inpM.value = now.join(", ");
        }));
      }
    });
    document.addEventListener("click", (ev) => { if (!tr.contains(ev.target)) menu.classList.remove("on"); });
  }
  const addMem = (m) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-f="full" value="${esc(m.full)}" placeholder="начальника … капітана поліції Прізвища І.Б."></td>
      <td><input data-f="short" value="${esc(m.short)}" placeholder="Прізвище І.Б."></td><td><button class="btn sm danger">🗑</button></td>`;
    tr.querySelector("button").addEventListener("click", () => tr.remove());
    mbody.appendChild(tr);
  };
  // --- джерела пального
  const sbody = $app.querySelector("#srcs tbody");
  const addSrc = (x) => {
    const tr = document.createElement("tr"); tr.dataset.id = x.id || C.normalizeSource({}).id;
    tr.innerHTML = `<td><input data-f="fuel" list="fuels" value="${esc(x.fuel)}" placeholder="будь-який" style="width:90px"></td>
      <td><select data-f="form">${C.FORMS.map((fm) => `<option ${fm === (x.form || "с/к") ? "selected" : ""}>${esc(fm)}</option>`).join("")}</select></td>
      <td><input data-f="supplier" list="suppliers" value="${esc(x.supplier)}" placeholder="з картки авто" style="width:120px"></td>
      <td><input data-f="mvo" value="${esc(x.mvo)}" placeholder="з картки авто" style="width:140px"></td>
      <td><input data-f="label" value="${esc(C.LEGACY_SOURCES.some((l) => l.id === x.id && l.label === x.label) ? "" : x.label)}" placeholder="${esc(C.sourceTitle({ ...x, label: "" }))}"></td>
      <td class="c"><input data-f="active" type="checkbox" ${x.active === false ? "" : "checked"}></td>
      <td><button class="btn sm danger js-del" title="Прибрати джерело">🗑</button></td>`;
    tr.querySelector(".js-del").addEventListener("click", () => { if (confirm("Прибрати джерело? Дані минулих місяців збережуться.")) { tr.remove(); refreshSrcSelects(); } });
    tr.querySelectorAll("input,select").forEach((i) => i.addEventListener("input", () => {
      tr.querySelector("[data-f=label]").placeholder = C.sourceTitle({ ...readSrc(tr), label: "" }); refreshSrcSelects(); }));
    sbody.appendChild(tr);
  };
  const readSrc = (tr) => { const x = { id: tr.dataset.id }; tr.querySelectorAll("[data-f]").forEach((i) => { x[i.dataset.f] = i.type === "checkbox" ? i.checked : i.value.trim(); }); return C.normalizeSource(x); };
  const custom = () => $app.querySelector("input[name=srcmode]:checked").value === "custom";
  const currentSources = () => (custom() ? [...sbody.querySelectorAll("tr")].map(readSrc) : C.unitSources({}));
  function refreshSrcSelects() {
    const list = currentSources().filter((x) => x.active);
    vbody.querySelectorAll("select[data-f=source]").forEach((sel) => {
      if (!sel.dataset.b) { sel.dataset.b = 1; sel.addEventListener("change", () => { sel.dataset.cur = sel.value; }); }
      const cur = sel.dataset.cur;          // вибір користувача (а не те, що браузер вибрав за замовчуванням)
      sel.innerHTML = list.map((x) => `<option value="${x.id}" ${x.id === cur ? "selected" : ""}>${esc(C.sourceTitle(x))}</option>`).join("") +
        (custom() && list.length ? `<option value="${C.NO_MAIN}" ${cur === C.NO_MAIN ? "selected" : ""}>— різні (без основного) —</option>` : "");
    });
  }
  (s.sources && s.sources.length ? C.unitSources(s) : C.LEGACY_SOURCES.map((l) => ({ ...l }))).forEach(addSrc);
  const toggleBox = () => { $app.querySelector(".js-srcbox").style.display = custom() ? "" : "none"; refreshSrcSelects(); };
  $app.querySelectorAll("input[name=srcmode]").forEach((r) => r.addEventListener("change", toggleBox));
  $app.querySelector(".js-addsrc").addEventListener("click", () => { addSrc({ form: "с/к" }); refreshSrcSelects(); });

  unit.vehicles.forEach(addVeh); if (!unit.vehicles.length) addVeh({});
  toggleBox();
  (s.act_members || []).forEach(addMem);
  $app.querySelector(".js-addveh").addEventListener("click", () => { addVeh({}); refreshSrcSelects(); });
  $app.querySelector(".js-addmem").addEventListener("click", () => addMem({}));
  $app.querySelector(".js-save").addEventListener("click", async () => {
    const settings = { ...s };
    $app.querySelectorAll("[data-k]").forEach((i) => { settings[i.dataset.k] = i.value.trim(); });
    settings.show_blago = $app.querySelector(".js-blago").checked;
    delete settings.start_by_mvo;
    settings.cons_issued_first = $app.querySelector(".js-isfirst").checked;
    if (custom()) {
      const list = [...sbody.querySelectorAll("tr")].map(readSrc);
      if (!list.length) { flash("err", "Додайте хоча б одне джерело пального або оберіть стандартні."); return render(); }
      settings.sources = list;
    } else delete settings.sources;
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
  state.units = units; setUserStats(users);
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
            <label class="f">Роль<select class="js-role"><option value="user">Користувач підрозділу</option><option value="dispatcher">Диспетчер (перевірка розрахунків, уся область)</option><option value="admin">Адміністратор</option></select></label>
            <button class="btn primary js-approve">✓ Підтвердити</button></div>
          <div style="margin-top:8px"><button class="btn sm danger js-reject">✕ Відхилити</button> <button class="btn sm danger js-udel">🗑 Видалити заявку</button></div></div>`;
      }).join("")}</div>`;
  } else if (tab === "users") {
    const list = users.filter((u) => u.status !== "pending").sort((a, b) => unitName(a.unit_id).localeCompare(unitName(b.unit_id), "uk"));
    body = `<div class="panel"><h2>Користувачі</h2><table class="list"><thead><tr><th>ПІБ / посада</th><th>E-mail</th><th>Підрозділ і роль</th><th>Статус</th><th>Останній вхід</th><th>Дії</th></tr></thead><tbody>
      ${list.map((u) => `<tr data-uid="${u.uid}"><td><b>${esc(u.full_name)}</b><br><span class="muted">${esc(u.position)}${u.phone ? " · " + esc(u.phone) : ""}</span></td>
        <td>${esc(u.email)}</td>
        <td><div class="row" style="gap:6px"><select class="js-unitsel">${unitOptions(u.unit_id)}</select>
          <select class="js-role">${Object.entries(ROLE_UA).map(([k, l]) => `<option value="${k}" ${u.role === k ? "selected" : ""}>${l}</option>`).join("")}</select>
          <button class="btn sm js-usave">Зберегти</button></div></td>
        <td>${u.status === "active" ? `<span class="pill ok">активний</span>` : u.status === "blocked" ? `<span class="pill err">заблокований</span>` : `<span class="pill mut">${esc(u.status)}</span>`}</td>
        <td class="muted">${esc(u.last_login || "—")}</td>
        <td><button class="btn sm js-ureset">✉️ Лист для скидання пароля</button>
          ${u.uid !== state.authUser.uid ? `${u.status === "active" ? `<button class="btn sm danger js-ublock">Заблокувати</button>` : `<button class="btn sm js-uunblock">Розблокувати</button>`}
          <button class="btn sm danger js-udel">🗑 Видалити</button>` : ""}</td></tr>`).join("")}</tbody></table>
      <p class="muted">Видалення прибирає профіль і доступ до даних. Якщо людина знову увійде зі своїм e-mail, портал запропонує подати нову заявку.</p>
      <h3 style="margin-top:18px">🧮 Диспетчер</h3>
      <p class="muted">Диспетчер бачить лише довідник авто всіх підрозділів (назва, норма, надбавка, пальне) і рахує витрату на своєму екрані — показників підрозділів він не бачить і нічого в них не змінює.
        Ключ диспетчера створює перший диспетчер; якщо його забули — скиньте, і диспетчер створить новий (довідники підрозділів оновляться, коли їх відкриють).</p>
      ${(await store.getDispPub().catch(() => null)) ? `<button class="btn sm danger js-dispreset">Скинути ключ диспетчера</button>` : `<span class="muted">Ключ диспетчера ще не створено.</span>`}</div>`;
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
      publishCatalog(unit);
      const st = await monthState(unit, ym);
      // для вивантаження в Excel (модуль хабу «Пробіг і пальне»): збережений місяць або перенесений з попереднього показник на початок
      const exp = { org: unit.settings.org, unit: u.name, vehicles: st.vehicles, entries: st.entries };
      const m = st.saved;
      if (!m) return { u, n: unit.vehicles.filter((v) => v.active).length, done: 0, err: 0, t: null, exp };
      const vehicles = C.monthVehicles(unit.vehicles, { entries: m.entries || {}, snapshot: m.snapshot || {} }, st.sources);
      const rows = vehicles.map((v) => C.calcRow(v, (m.entries || {})[v.id], st.sources));
      const sts = rows.map((r, i) => C.statusOf(r, (m.entries || {})[vehicles[i].id] || {}));
      return { u, n: rows.length, done: sts.filter((s) => s === "ok" || s === "warn").length, err: sts.filter((s) => s === "err").length, t: C.totals(rows), updated: m.updated_at, exp };
    }));
    state.lastSummary = { ym, rowsSum };
    const sum = (k) => rowsSum.reduce((a, r) => a + (r.t ? r.t[k] : 0), 0);
    body = `<div class="panel"><div class="row" style="margin-bottom:12px">${monthPicker(ym, "#/admin/summary")}<button class="btn sm js-sumdoc">📥 Зведення у Word</button>
        <button class="btn sm sky js-fleetxl" data-kind="end" title="Показники на початок і на кінець місяця — для модуля «Пробіг і пальне автопарку» в хабі">📊 Excel для модуля пробігу · на ${dmy(firstOf(nextYm(ym)))}</button>
        <button class="btn sm js-fleetxl" data-kind="start" title="Лише показники на початок місяця — коли файлу за попередній місяць ще немає">Зріз на початок · ${dmy(firstOf(ym))}</button></div>
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
    const role = box.querySelector(".js-role").value;
    if (role === "dispatcher") {
      // диспетчер не належить підрозділу і не бачить його даних
      await store.updateUser(uid, { status: "active", unit_id: "", role });
      store.log("approve", `${u.full_name} → диспетчер`); return;
    }
    if (unitId === "new") {
      const name = box.querySelector(".js-newname input").value.trim(); if (!name) throw new Error("Вкажіть назву нового підрозділу.");
      const ex = units.find((x) => x.name.toLowerCase() === name.toLowerCase());
      unitId = ex ? ex.id : await store.createUnit(name, { ...C.DEFAULT_SETTINGS, dept: name, unit: name, act_owner: name, compiler: C.shortName(u.full_name) });
    }
    await store.updateUser(uid, { status: "active", unit_id: unitId, role });
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
    act(() => store.updateUser(uid, { unit_id: role === "dispatcher" ? "" : tr.querySelector(".js-unitsel").value, role }), "Користувача оновлено.");
  }));
  $app.querySelector(".js-dispreset")?.addEventListener("click", () => {
    if (!confirm("Скинути ключ диспетчера? Диспетчер створить новий, а довідники авто оновляться, коли підрозділи відкриють портал (або ви відкриєте «Зведення по області»).")) return;
    act(async () => { await store.resetDispKeys(); state.dispPub = undefined; store.log("disp_key_reset", ""); }, "Ключ диспетчера скинуто.");
  });
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
  $app.querySelectorAll(".js-fleetxl").forEach((b) => b.addEventListener("click", () => {
    const { ym, rowsSum } = state.lastSummary;
    const items = rowsSum.filter((r) => r.exp).map((r) => r.exp);
    const skipped = rowsSum.filter((r) => !r.exp).map((r) => r.u.name);
    const { rows, date, stats } = buildFleetRows(items, ym, b.dataset.kind);
    downloadBytes(buildXlsx(rows, { sheet: "Портал_МВО", headerRows: 2, widths: FLEET_WIDTHS }), `Пробіг_МВО_на_${dmy(date)}.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    store.log("fleet_xlsx", `${b.dataset.kind} ${ym}: ${stats.cars} авто`);
    if (skipped.length) { flash("info", `У файл не потрапили підрозділи без ключа шифрування: ${skipped.join(", ")}.`); render(); }
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

// ============================================================ присутність на сайті й лічильник користувачів
const BEAT_MS = 5 * 60 * 1000, ONLINE_MS = 7 * 60 * 1000;
let beatTimer = null, statsTimer = null, lastBeat = 0;
function beat() { if (document.visibilityState !== "visible") return; lastBeat = Date.now(); store.heartbeat(); }
function onVisible() { if (document.visibilityState === "visible" && Date.now() - lastBeat > BEAT_MS) { beat(); if (isAdmin()) refreshStats(); } }
function startPresence() {
  stopPresence(); beat();
  beatTimer = setInterval(beat, BEAT_MS);
  document.addEventListener("visibilitychange", onVisible);
  if (state.me && state.me.role === "admin") statsTimer = setInterval(() => { if (document.visibilityState === "visible") refreshStats(); }, BEAT_MS);
}
function stopPresence() { clearInterval(beatTimer); clearInterval(statsTimer); beatTimer = statsTimer = null; document.removeEventListener("visibilitychange", onVisible); }
function setUserStats(users) {
  const now = Date.now();
  const active = users.filter((u) => u.status === "active");
  const online = active.filter((u) => now - store.seenMs(u.last_seen) < ONLINE_MS || u.uid === state.authUser?.uid);
  state.pendingCount = users.filter((u) => u.status === "pending").length;
  state.userStats = { active: active.length, pending: state.pendingCount, online: online.map((u) => ({ name: u.full_name || u.email, unit: u.unit_id })) };
}
async function refreshStats() {
  try { setUserStats(await store.listUsers()); } catch { return; }
  const el = document.querySelector(".who-stats"); if (el) el.innerHTML = statsHtml();
}
function statsHtml() {
  const s = state.userStats; if (!s) return "";
  const uname = (id) => state.units.find((u) => u.id === id)?.name || "";
  const list = s.online.map((o) => `<li><b>${esc(o.name)}</b>${uname(o.unit) ? `<span>${esc(uname(o.unit))}</span>` : ""}</li>`).join("");
  return `<span title="Активних користувачів">👥 ${s.active}</span>${s.pending ? ` · <a href="#/admin/requests">заявок: ${s.pending}</a>` : ""}
    · <span class="on">🟢 ${s.online.length} онлайн</span><div class="pop"><div class="ph">Зараз на сайті</div><ul>${list}</ul></div>`;
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
    if (isDisp()) return await pageDisp(q);
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
  if (!user) { forgetSessionKeys(); stopPresence(); }
  if (user) {
    try { state.me = await store.getMe(); } catch (err) { state.me = null; }
    try { state.mfa = await store.mfaState(); } catch (err) { state.mfa = null; }
    if (state.me && state.me.status === "active" && (!store.mfaRequired || (state.mfa && state.mfa.session))) {
      store.touchLogin(); startPresence();
      if (state.me.role === "admin") {
        state.adminKeyReady = await store.hasAdminKey().catch((e) => (e && e.code === "permission-denied" ? "rules" : true));
        state.units = await store.listUnits().catch(() => []);
        setUserStats(await store.listUsers().catch(() => []));
        const ctx = sessionStorage.getItem("unitCtx");
        state.unitId = ctx && state.units.some((u) => u.id === ctx) ? ctx : state.me.unit_id;
      } else if (state.me.role === "dispatcher") {
        state.unitId = null; state.units = [];
      } else {
        state.unitId = state.me.unit_id;
        const u = await store.getUnit(state.me.unit_id).catch(() => null);
        state.units = u ? [u] : [];
      }
    }
  }
  const pth = route().path;
  if (!user && !["login", "register", "forgot"].includes(pth)) { location.hash = "#/login"; return; }
  if (user && ["login", "register", "forgot", ""].includes(pth)) { location.hash = state.me && state.me.role === "dispatcher" ? "#/disp" : "#/entry"; return; }
  render();
}

// ============================================================ старт
(async () => {
  try {
    if (DEMO) store = (await import("./store-mock.js?v=20261027s")).createStore();
    else if (FIREBASE_CONFIG && FIREBASE_CONFIG.apiKey) store = (await import("./store-firebase.js?v=20261027s")).createStore(FIREBASE_CONFIG);
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
