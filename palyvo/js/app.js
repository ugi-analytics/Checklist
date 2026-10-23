// Портал «Списання пального для МВО» — односторінковий застосунок (GitHub Pages + Firebase).
import * as C from "./core.js";
import * as U from "./ui.js";
import { litresWords } from "./words.js";
import { buildRegister, buildAct } from "./docs.js";
import { FIREBASE_CONFIG, ADMIN_EMAIL } from "./config.js";

const $app = document.getElementById("app");
const esc = U.esc;
const params = new URLSearchParams(location.search);
const DEMO = params.has("demo");
let store = null;

const state = { authUser: null, me: null, unitId: null, units: [], flash: null, setupDone: null };

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
  const unitChip = unit ? `<span class="chip unit-chip">${esc(unit.name)}</span>` : "";
  const switcher = isAdmin() && state.units.length > 1 ? `<select class="js-unit-switch" title="Працювати з підрозділом">
      ${state.units.map((u) => `<option value="${u.id}" ${u.id === state.unitId ? "selected" : ""}>${esc(u.name)}</option>`).join("")}</select>` : "";
  const tab = (h, label, extra = "") => `<a href="#/${h}" class="${p === h || p.startsWith(h + "/") ? "on" : ""}">${label}${extra}</a>`;
  const pending = isAdmin() && state.pendingCount ? `<span class="badge">${state.pendingCount}</span>` : "";
  $app.innerHTML = `
  <header class="hero"><div class="wrap">
    <div class="topbar"><img src="img/hub_icon_small.png" alt="">
      <div class="brand display grad">UGI-Analytics<small>Управління головної інспекції ГУНП в Полтавській області</small></div>
      <div class="sp"></div>
      ${state.authUser ? `<div class="who"><b>${esc(me?.full_name || state.authUser.email)}</b>${isAdmin() ? " · адміністратор" : ""}<br>
        ${me && me.status === "active" ? `<a href="#/password">змінити пароль</a> · ` : ""}<a href="#" class="js-logout">вийти</a></div>` : ""}
    </div>
    <div class="title-row"><h1 class="display grad">Списання пального для МВО</h1>${unitChip}${switcher}
      ${DEMO ? `<span class="chip" style="border-color:#f87171;color:#fecaca">ДЕМО-РЕЖИМ · дані не зберігаються</span>` : ""}</div>
    ${nav && me && me.status === "active" ? `<nav class="tabs">
      ${tab("entry", "✍️ Введення даних")}${tab("report", "📄 Звіт і акти")}${tab("fleet", "🚗 Автопарк і реквізити")}${tab("archive", "🗂️ Архів")}
      ${isAdmin() ? tab("admin", "🛡️ Адміністрування", pending) : ""}</nav>` : ""}
  </div></header>
  <main class="wrap">${state.flash ? `<div class="flash ${state.flash[0]}">${esc(state.flash[1])}</div>` : ""}${content}</main>`;
  state.flash = null;
  $app.querySelector(".js-logout")?.addEventListener("click", async (e) => { e.preventDefault(); await store.signOut(); go("#/login"); });
  $app.querySelector(".js-unit-switch")?.addEventListener("change", (e) => {
    state.unitId = e.target.value; sessionStorage.setItem("unitCtx", state.unitId); render();
  });
}
function authCard(inner, width = 460) { return `<div class="auth" style="max-width:${width}px"><div class="panel">${inner}</div></div>`; }

// ============================================================ сторінки входу
function pageNotConfigured() {
  shell(authCard(`<h2 class="display">Портал ще не налаштовано</h2>
    <p>Не заповнено параметри Firebase у файлі <code>js/config.js</code>.</p>
    <p class="muted">Для перегляду інтерфейсу без бази відкрийте <a href="?demo=1">демо-режим</a>.</p>`), { nav: false });
}

function pageLogin() {
  shell(authCard(`<h2 class="display">Вхід до порталу</h2>
    <form class="js-login" autocomplete="on">
      <label class="f">E-mail<input name="email" type="email" required autofocus autocomplete="username"></label>
      <label class="f">Пароль<input name="pw" type="password" required autocomplete="current-password"></label>
      <button class="btn primary">Увійти</button>
    </form>
    <p class="muted" style="margin-top:16px"><a href="#/forgot">Забули пароль?</a></p>
    <p class="muted">Немає облікового запису? <a href="#/register">Подати заявку на реєстрацію</a> — після підтвердження адміністратором ви зможете працювати.</p>`), { nav: false });
  $app.querySelector(".js-login").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target; busy(true);
    try { await store.signIn(f.email.value, f.pw.value); } catch (err) { flash("err", store.errUa(err)); pageLogin(); } finally { busy(false); }
  });
}

function pageForgot() {
  shell(authCard(`<h2 class="display">Відновлення пароля</h2>
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
  shell(authCard(`<h2 class="display">Заявка на реєстрацію</h2>
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
      state.registering = false;
      await onAuthChanged(store.currentUser());
    } catch (err) { state.registering = false; flash("err", store.errUa(err)); pageRegister(); } finally { busy(false); }
  });
}

function pagePending() {
  shell(authCard(`<h2 class="display">Заявку надіслано ✓</h2>
    <p>Адміністратор перевірить її та прив'яже ваш обліковий запис до підрозділу <b>${esc(state.me.requested_unit)}</b>.
      Після підтвердження оновіть сторінку — відкриється робоча частина порталу.</p>
    <button class="btn sky js-reload">🔄 Перевірити ще раз</button>`), { nav: false });
  $app.querySelector(".js-reload").addEventListener("click", () => onAuthChanged(store.currentUser()));
}

function pageBlocked() {
  shell(authCard(`<h2 class="display">Доступ закрито</h2>
    <p>Обліковий запис ${state.me.status === "rejected" ? "відхилено" : "заблоковано"}. Зверніться до адміністратора порталу.</p>`), { nav: false });
}

function pageNoProfile() {
  shell(authCard(`<h2 class="display">Профіль не знайдено</h2>
    <p class="muted">Ваш обліковий запис видалено або заявку ще не створено. Можете подати заявку повторно.</p>
    <form class="js-reapply">${profileFields()}<button class="btn primary">Подати заявку</button></form>`, 620), { nav: false });
  $app.querySelector(".js-reapply").addEventListener("submit", async (e) => {
    e.preventDefault(); busy(true);
    try { await store.createMyProfile(readProfile(e.target)); await onAuthChanged(store.currentUser()); }
    catch (err) { flash("err", store.errUa(err)); pageNoProfile(); } finally { busy(false); }
  });
}

function pageSetup() {
  const lockedEmail = ADMIN_EMAIL && !DEMO ? ADMIN_EMAIL : "";
  shell(authCard(`<h2 class="display">Перший запуск порталу</h2>
    <p class="muted">Створіть обліковий запис адміністратора області${lockedEmail ? ` (лише для <b>${esc(lockedEmail)}</b>)` : ""}.</p>
    <form class="js-setup" autocomplete="off">
      <label class="f">Ваш підрозділ<input name="unit" required value="УГІ ГУНП"></label>
      <div class="two"><label class="f">ПІБ<input name="full_name" required></label><label class="f">Посада<input name="position"></label></div>
      <label class="f">E-mail адміністратора<input name="email" type="email" required value="${esc(lockedEmail)}" ${lockedEmail ? "readonly" : ""}></label>
      <div class="two"><label class="f">Пароль (від 8 символів, літери й цифри)<input name="pw" type="password" required minlength="8"></label>
        <label class="f">Повторіть пароль<input name="pw2" type="password" required minlength="8"></label></div>
      <button class="btn primary">Створити й увійти</button></form>`, 640), { nav: false });
  $app.querySelector(".js-setup").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.pw.value !== f.pw2.value) { flash("err", "Паролі не збігаються."); return pageSetup(); }
    const pp = pwProblem(f.pw.value); if (pp) { flash("err", pp); return pageSetup(); }
    busy(true);
    try {
      state.registering = true;
      try { await store.signUp(f.email.value, f.pw.value); }
      catch (err) { if (err.code === "auth/email-already-in-use") await store.signIn(f.email.value, f.pw.value); else throw err; }
      const unitName = f.unit.value.trim();
      const settings = { ...C.DEFAULT_SETTINGS, dept: unitName, unit: unitName, act_owner: unitName, compiler: C.shortName(f.full_name.value) };
      await store.bootstrapAdmin({ full_name: f.full_name.value.trim(), position: f.position.value.trim(), phone: "", requested_unit: unitName },
        unitName, settings);
      state.registering = false; state.setupDone = true;
      flash("ok", "Портал налаштовано. Створіть автопарк свого підрозділу.");
      await onAuthChanged(store.currentUser());
    } catch (err) { state.registering = false; flash("err", store.errUa(err)); pageSetup(); } finally { busy(false); }
  });
}

function pagePassword() {
  shell(authCard(`<h2 class="display">Зміна пароля</h2>
    <form class="js-pw"><label class="f">Поточний пароль<input name="old" type="password" required></label>
      <label class="f">Новий пароль (від 8 символів, літери й цифри)<input name="n1" type="password" required minlength="8"></label>
      <label class="f">Повторіть новий пароль<input name="n2" type="password" required minlength="8"></label>
      <button class="btn primary">Зберегти</button></form>`));
  $app.querySelector(".js-pw").addEventListener("submit", async (e) => {
    e.preventDefault(); const f = e.target;
    if (f.n1.value !== f.n2.value) { flash("err", "Нові паролі не збігаються."); return pagePassword(); }
    const pp = pwProblem(f.n1.value); if (pp) { flash("err", pp); return pagePassword(); }
    busy(true);
    try { await store.changePassword(f.old.value, f.n1.value); flash("ok", "Пароль змінено."); go("#/entry"); }
    catch (err) { flash("err", store.errUa(err)); pagePassword(); } finally { busy(false); }
  });
}

// ============================================================ дані підрозділу
async function loadUnit() {
  const u = await store.getUnit(state.unitId);
  if (!u) return null;
  u.settings = { ...C.DEFAULT_SETTINGS, ...(u.settings || {}) };
  u.vehicles = (u.vehicles || []).map(C.normalizeVehicle);
  return u;
}
async function monthState(unit, ym) {
  const saved = await store.getMonth(unit.id, ym);
  let entries, carried = null;
  if (saved) entries = saved.entries || {};
  else {
    const prev = await store.getMonth(unit.id, C.prevYm(ym));
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
  const hasPrev = !!(await store.getMonth(unit.id, C.prevYm(ym)));
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
      await store.saveMonth(unit.id, ym, { entries, snapshot });
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
    const prev = await store.getMonth(unit.id, C.prevYm(ym));
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
  const months = (await store.listMonths(unit.id)).reverse();
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
    <div class="stickybar"><button class="btn primary js-save">💾 Зберегти шаблон підрозділу</button><span class="muted js-st"></span></div>`);

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
      await store.updateUnit(unit.id, { settings, vehicles });
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
      <table class="list"><thead><tr><th>Назва</th><th class="num">Користувачів</th><th class="num">Авто</th><th>Дії</th></tr></thead><tbody>
      ${units.map((u) => `<tr data-id="${u.id}"><td><div class="row" style="gap:6px"><input class="js-uname" value="${esc(u.name)}" style="min-width:300px"><button class="btn sm js-urename">Перейменувати</button></div></td>
        <td class="num">${usersIn(u.id)}</td><td class="num">${(u.vehicles || []).filter((v) => v.active !== false).length}</td>
        <td><button class="btn sm js-open">Відкрити як підрозділ →</button> ${usersIn(u.id) ? "" : `<button class="btn sm danger js-unitdel">🗑 Видалити</button>`}</td></tr>`).join("")}
      </tbody></table></div>`;
  } else if (tab === "summary") {
    const ym = currentYm(q);
    const rowsSum = await Promise.all(units.map(async (u) => {
      const unit = { ...u, settings: { ...C.DEFAULT_SETTINGS, ...(u.settings || {}) }, vehicles: (u.vehicles || []).map(C.normalizeVehicle) };
      const m = await store.getMonth(u.id, ym);
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
    console.error(err);
    shell(`<div class="panel"><h2>Помилка</h2><p>${esc(store.errUa(err))}</p><button class="btn" onclick="location.reload()">Оновити сторінку</button></div>`);
  } finally { busy(false); }
}

async function onAuthChanged(user) {
  if (state.registering) return;
  state.authUser = user; state.me = null;
  if (user) {
    try { state.me = await store.getMe(); } catch (err) { state.me = null; }
    if (state.me && state.me.status === "active") {
      store.touchLogin();
      if (state.me.role === "admin") {
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
    if (DEMO) store = (await import("./store-mock.js")).createStore();
    else if (FIREBASE_CONFIG && FIREBASE_CONFIG.apiKey) store = (await import("./store-firebase.js")).createStore(FIREBASE_CONFIG);
  } catch (err) {
    $app.innerHTML = `<div class="wrap"><div class="panel"><h2>Не вдалося завантажити портал</h2><p>${esc(err.message)}</p></div></div>`;
    return;
  }
  if (!store) return render();
  window.addEventListener("hashchange", () => { window.onbeforeunload = null; render(); });
  store.onAuth(onAuthChanged);
})();
