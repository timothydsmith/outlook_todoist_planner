"use strict";
const TOKEN = document.querySelector('meta[name="planner-token"]').content;
const SNAP = 15;
const HOUR_PX = 48;

const state = {
  view: "week", anchor: startOfDay(new Date()), events: [], tasks: [], errors: [], settings: null,
  updated: null, loading: false, demo: false, nodateOpen: false, nodateFilter: "", adExpanded: false,
  firstPaint: true, dragging: false, pendingRender: false,
};
let loadSeq = 0;
const $ = (s, el = document) => el.querySelector(s);

/* ---------- utilities ---------- */
function pad(n) { return String(n).padStart(2, "0"); }
function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function parseYmd(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function parseLocal(s) {
  const [a, b] = s.split("T"); const [y, m, d] = a.split("-").map(Number);
  const [hh, mm] = (b || "00:00").split(":").map(Number); return new Date(y, m - 1, d, hh, mm);
}
function hhmm(min) { return `${pad(Math.floor(min / 60))}:${pad(min % 60)}`; }
function atMinutes(day, min) { return new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, min); }
function esc(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
function plain(s) { return String(s || "").replace(/\[([^\]]+)\]\((?:https?:)?[^)]*\)/g, "$1").replace(/(\*\*|__)(.+?)\1/g, "$2"); }
function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

const F = {
  time: new Intl.DateTimeFormat(undefined, { hour: "numeric", minute: "2-digit" }),
  hour: new Intl.DateTimeFormat(undefined, { hour: "numeric" }),
  wd: new Intl.DateTimeFormat(undefined, { weekday: "short" }),
  mon: new Intl.DateTimeFormat(undefined, { month: "short" }),
  dayShort: new Intl.DateTimeFormat(undefined, { weekday: "short", day: "numeric", month: "short" }),
  dayLong: new Intl.DateTimeFormat(undefined, { weekday: "long", day: "numeric", month: "long" }),
  dayFull: new Intl.DateTimeFormat(undefined, { weekday: "long", day: "numeric", month: "long", year: "numeric" }),
  range: new Intl.DateTimeFormat(undefined, { day: "numeric", month: "long", year: "numeric" }),
};
function timeRange(s, e) {
  try { return F.time.formatRange(s, e); } catch { return `${F.time.format(s)} – ${F.time.format(e)}`; }
}
function dateRange(a, b) {
  try { return F.range.formatRange(a, b); } catch { return `${F.range.format(a)} – ${F.range.format(b)}`; }
}
function prioClass(p) { return "pc" + (5 - clamp(p || 1, 1, 4)); }
const ICON = {
  check: '<svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 5.2l2 2L8 3"/></svg>',
  recur: '<svg class="recur" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-label="Recurring"><path d="M2 5a4 4 0 0 1 7-2.3M10 7a4 4 0 0 1-7 2.3"/><path d="M9.3 1v2h-2M2.7 11V9h2"/></svg>',
  chev: '<svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 2.5L8 6l-3.5 3.5"/></svg>',
};

async function api(path, body) {
  const opts = { method: body === undefined ? "GET" : "POST", headers: { "X-Planner-Token": TOKEN } };
  if (body !== undefined) { opts.headers["Content-Type"] = "application/json"; opts.body = JSON.stringify(body); }
  let res;
  try { res = await fetch(path, opts); }
  catch { throw new Error("The Planner app isn't running. Start it again, then reload this window."); }
  let data = {};
  try { data = await res.json(); } catch { /* empty body */ }
  if (!res.ok) throw new Error(data.error || `Something went wrong (HTTP ${res.status}).`);
  return data;
}

/* ---------- range ---------- */
function visibleDays() {
  if (state.view === "day") return [state.anchor];
  const ws = state.settings ? Number(state.settings.week_start) : 1;
  const start = addDays(state.anchor, -((state.anchor.getDay() - ws + 7) % 7));
  let days = Array.from({ length: 7 }, (_, i) => addDays(start, i));
  if (state.settings && !state.settings.show_weekends) days = days.filter(d => d.getDay() !== 0 && d.getDay() !== 6);
  return days;
}
function step(dir) {
  state.anchor = addDays(state.anchor, dir * (state.view === "day" ? 1 : 7));
  if (state.view === "day" && state.settings && !state.settings.show_weekends) {
    while (state.anchor.getDay() === 0 || state.anchor.getDay() === 6) state.anchor = addDays(state.anchor, dir);
  }
  render(); load();
}
function setView(v) { state.view = v; render(); load(); }
function goToday() { state.anchor = startOfDay(new Date()); render(); load(); scrollToStart(); }

/* ---------- loading ---------- */
async function load(force = false) {
  const days = visibleDays();
  const start = ymd(days[0]), end = ymd(addDays(days[days.length - 1], 1));
  const seq = ++loadSeq;
  state.loading = true; renderStatus();
  try {
    const d = await api(`/api/data?start=${start}&end=${end}&force=${force ? 1 : 0}`);
    if (seq !== loadSeq) return;
    state.events = d.events; state.tasks = d.tasks; state.errors = d.errors; state.demo = d.demo;
    state.updated = new Date();
  } catch (e) {
    if (seq !== loadSeq) return;
    state.errors = [{ source: "app", message: e.message }];
  } finally {
    if (seq === loadSeq) { state.loading = false; render(); }
  }
}

/* ---------- building the calendar ---------- */
function collect(days) {
  const keys = days.map(ymd), set = new Set(keys);
  const allDay = new Map(keys.map(k => [k, []])), timed = new Map(keys.map(k => [k, []]));
  for (const ev of state.events) {
    if (ev.allDay) {
      for (let d = parseYmd(ev.start), end = parseYmd(ev.end); d < end; d = addDays(d, 1)) {
        const k = ymd(d); if (set.has(k)) allDay.get(k).push({ kind: "event", ev });
      }
      continue;
    }
    const s = parseLocal(ev.start), e = parseLocal(ev.end);
    if (e - s >= 864e5) {
      for (let d = startOfDay(s); d < e; d = addDays(d, 1)) { const k = ymd(d); if (set.has(k)) allDay.get(k).push({ kind: "event", ev }); }
      continue;
    }
    let d = startOfDay(s);
    for (let guard = 0; guard < 3; guard++) {
      const next = addDays(d, 1), k = ymd(d);
      const segS = s > d ? s : d, segE = e < next ? e : next;
      if (set.has(k)) timed.get(k).push({ kind: "event", ev, s: (segS - d) / 6e4, e: (segE - d) / 6e4, cutTop: s < d, cutBottom: e > next });
      if (e <= next) break;
      d = next;
    }
  }
  const defMin = Number(state.settings?.default_task_minutes || 30);
  for (const t of state.tasks) {
    if (!t.due || !set.has(t.due.date)) continue;
    if (!t.due.time) { allDay.get(t.due.date).push({ kind: "task", t }); continue; }
    const [hh, mm] = t.due.time.split(":").map(Number), s = hh * 60 + mm;
    timed.get(t.due.date).push({ kind: "task", t, s, e: Math.min(1440, s + (t.duration || defMin)) });
  }
  for (const list of allDay.values()) list.sort((a, b) =>
    (a.kind === b.kind ? 0 : a.kind === "event" ? -1 : 1) || ((b.t?.priority || 0) - (a.t?.priority || 0)));
  return { allDay, timed };
}
function layout(items) {
  items.sort((a, b) => a.s - b.s || (b.e - b.s) - (a.e - a.s));
  let cluster = [], end = -1;
  const flush = () => { const cols = []; for (const it of cluster) { let i = 0; while (i < cols.length && cols[i] > it.s) i++; cols[i] = Math.max(it.e, it.s + 22); it.col = i; } cluster.forEach(it => it.ncol = cols.length); };
  for (const it of items) { if (cluster.length && it.s >= end) { flush(); cluster = []; end = -1; } cluster.push(it); end = Math.max(end, it.e, it.s + 22); }
  if (cluster.length) flush();
  return items;
}

function taskInner(t, extraMeta) {
  return `<button class="check" data-complete="${esc(t.id)}" aria-label="Complete: ${esc(plain(t.content))}" title="Complete">${ICON.check}</button>` +
    `<span class="t">${esc(plain(t.content))}${t.due?.recurring ? ICON.recur : ""}</span>${extraMeta || ""}`;
}
// An Outlook category colour: the block's fill (or tint and border when tentative or free), with dark or
// white text, whichever reads better on it.
function catLook(ev) {
  const m = /^#([0-9a-f]{6})$/i.exec(ev.color || "");
  if (!m) return { cls: "", style: "" };
  const lin = c => (c /= 255) <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  const n = parseInt(m[1], 16), L = 0.2126 * lin(n >> 16) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  return { cls: ` cat${L < 0.23 ? " cat-dark" : ""}`, style: `--cat:#${m[1]};` };
}
function blockHTML(it) {
  const top = it.s / 60 * HOUR_PX, h = Math.max((it.e - it.s) / 60 * HOUR_PX - 2, 18);
  const w = 100 / it.ncol, left = it.col * w;
  const style = `top:${top}px;height:${h}px;left:calc(${left}% + 2px);width:calc(${w}% - 5px)`;
  const size = h < 34 ? "short" : h >= 70 ? "tall" : "";
  if (it.kind === "event") {
    const ev = it.ev, s = parseLocal(ev.start);
    const meta = size === "short" ? F.time.format(s) : esc([timeRange(s, parseLocal(ev.end)), ev.location].filter(Boolean).join(", "));
    const look = catLook(ev);
    const cls = `blk ev busy-${esc(ev.busy)}${look.cls} ${size} ${ev.cancelled ? "cancelled" : ""} ${it.cutTop ? "cut-top" : ""} ${it.cutBottom ? "cut-bottom" : ""}`;
    return `<div class="${cls}" style="${look.style}${style}" data-ev="${esc(ev.id)}" tabindex="0" role="button"><span class="t">${esc(ev.title)}</span><span class="m">${meta}</span></div>`;
  }
  const t = it.t, day = parseYmd(t.due.date), s = atMinutes(day, it.s), e = atMinutes(day, it.e);
  const meta = `<span class="m">${size === "short" ? F.time.format(s) : esc(timeRange(s, e) + (t.project ? ", " + t.project : ""))}</span>`;
  return `<div class="blk task ${prioClass(t.priority)} ${size}" style="${style}" data-task="${esc(t.id)}" tabindex="0" role="button">` +
    taskInner(t, meta) + `${t.due.recurring ? "" : '<span class="resize" aria-hidden="true"></span>'}</div>`;
}
function chipHTML(it) {
  if (it.kind === "event") {
    const ev = it.ev, look = catLook(ev);
    return `<div class="chip ev busy-${esc(ev.busy)}${look.cls} ${ev.cancelled ? "cancelled" : ""}" style="${look.style}" data-ev="${esc(ev.id)}" tabindex="0" role="button"><span class="t">${esc(ev.title)}</span></div>`;
  }
  return `<div class="chip task ${prioClass(it.t.priority)}" data-task="${esc(it.t.id)}" tabindex="0" role="button">${taskInner(it.t)}</div>`;
}

/* ---------- render ---------- */
function render() {
  if (state.dragging) { state.pendingRender = true; return; }
  const days = visibleDays();
  document.documentElement.style.setProperty("--ncols", days.length);
  const first = days[0], last = days[days.length - 1];
  $("#range").textContent = state.view === "day" ? F.dayFull.format(first) : dateRange(first, last);
  document.querySelectorAll(".seg button").forEach(b => b.setAttribute("aria-pressed", String(b.dataset.view === state.view)));
  renderStatus(); renderBanner(); renderSide(); renderCal(days);
  if (state.firstPaint && state.settings) { state.firstPaint = false; scrollToStart(); }
}
function renderStatus() {
  const el = $("#status");
  const btn = $('[data-act="refresh"]');
  btn.classList.toggle("spin", state.loading);
  el.textContent = state.loading ? "Refreshing…" : state.updated ? `Updated ${F.time.format(state.updated)}` : "";
}
function renderBanner() {
  const el = $("#banner"), s = state.settings;
  const rows = [];
  const names = { calendar: "Outlook", outlook: "Outlook", todoist: "Todoist", app: "Planner", timezones: "Time zones" };
  for (const e of state.errors) rows.push(`<p><span class="src">${names[e.source] || "Problem"}</span><span>${esc(e.message)}</span>` +
    (e.source === "app" || e.source === "timezones" ? "" : ` <button data-act="settings">Open settings</button>`) + `</p>`);
  if (s && !state.demo) {
    if (!s.has_token) rows.push(`<p><span class="src">Todoist</span><span>Add your API token to see your tasks.</span> <button data-act="settings">Open settings</button></p>`);
    if (s.calendar_source === "ics" && !s.ics_url) rows.push(`<p><span class="src">Outlook</span><span>Add your published calendar link to see your meetings.</span> <button data-act="settings">Open settings</button></p>`);
  }
  if (state.demo) rows.push(`<p><span class="src">Demo</span><span>You're looking at sample data. Start the app without <code>--demo</code> to use your own calendar and tasks.</span></p>`);
  el.innerHTML = rows.join("");
  el.hidden = rows.length === 0;
  el.classList.toggle("demo", state.demo && state.errors.length === 0);
}
function renderSide() {
  const side = $("#side"), today = ymd(new Date());
  const overdue = state.tasks.filter(t => t.due && t.due.date < today)
    .sort((a, b) => a.due.date.localeCompare(b.due.date) || b.priority - a.priority);
  const nodate = state.tasks.filter(t => !t.due).sort((a, b) => b.priority - a.priority || a.content.localeCompare(b.content));
  const focused = document.activeElement && document.activeElement.classList.contains("filter");
  side.innerHTML =
    `<section><h2>Overdue <span class="count">${overdue.length || ""}</span></h2>` +
    (overdue.length ? `<ul>${overdue.map(t => itemHTML(t, true)).join("")}</ul>` : `<p class="empty">Nothing overdue.</p>`) + `</section>` +
    `<section><h2><button class="disclose" data-act="nodate" aria-expanded="${state.nodateOpen}">${ICON.chev}No date</button><span class="count">${nodate.length || ""}</span></h2>` +
    (state.nodateOpen ? `<input class="filter" type="search" placeholder="Filter these tasks" value="${esc(state.nodateFilter)}" aria-label="Filter tasks with no date"><ul id="nodateList"></ul>` : "") +
    `</section>` +
    `<p class="hint">Drag a task onto the calendar to give it a time, or onto the all-day row to give it a date. Double-click an empty slot to add a task there.</p>`;
  if (state.nodateOpen) {
    renderNodate(nodate);
    if (focused) { const f = $(".filter", side); f.focus(); f.setSelectionRange(f.value.length, f.value.length); }
  }
}
function renderNodate(list) {
  const q = state.nodateFilter.trim().toLowerCase();
  const items = (list || state.tasks.filter(t => !t.due)).filter(t => !q || t.content.toLowerCase().includes(q) || t.project.toLowerCase().includes(q));
  const ul = $("#nodateList"); if (!ul) return;
  ul.innerHTML = items.slice(0, 200).map(t => itemHTML(t, false)).join("") || `<li class="empty">${q ? "No matching tasks." : "Every task has a date."}</li>`;
}
function itemHTML(t, showDate) {
  const proj = t.project ? `<span><span class="dot" style="background:${esc(t.projectColor)}"></span>${esc(t.project)}</span>` : "";
  const date = showDate ? `<span class="late">${esc(F.dayShort.format(parseYmd(t.due.date)))}</span>` : "";
  return `<li class="item ${prioClass(t.priority)}" data-task="${esc(t.id)}" tabindex="0" role="button">${taskInner(t)}<span class="meta">${date}${proj}</span></li>`;
}
function renderCal(days) {
  const today = startOfDay(new Date()), tk = ymd(today);
  const { allDay, timed } = collect(days);
  const maxChips = Math.max(0, ...[...allDay.values()].map(l => l.length));
  const limit = state.adExpanded ? Infinity : 3;
  $("#calHead").innerHTML =
    `<div class="row days"><div></div>${days.map(d => {
      const k = ymd(d), cls = k === tk ? "today" : d < today ? "past" : "";
      const sub = (d.getDate() === 1 || d === days[0]) ? `<b>${esc(F.wd.format(d))}</b>${esc(F.mon.format(d))}` : `<b>${esc(F.wd.format(d))}</b>`;
      return `<button class="dayhead ${cls}" data-goto="${k}" title="Open ${esc(F.dayLong.format(d))}"><span class="num">${d.getDate()}</span><span class="wd">${sub}</span></button>`;
    }).join("")}</div>` +
    `<div class="row allday"><div class="gut">All day</div>${days.map(d => {
      const k = ymd(d), list = allDay.get(k), wk = d.getDay() === 0 || d.getDay() === 6 ? "wknd" : "";
      const shown = list.slice(0, limit).map(chipHTML).join("");
      const more = list.length > limit ? `<button class="more" data-act="adexpand">${list.length - limit} more</button>` :
        (state.adExpanded && maxChips > 3 && k === ymd(days[0]) ? `<button class="more" data-act="adexpand">Show less</button>` : "");
      return `<div class="ad-cell ${wk}" data-date="${k}">${shown}${more}</div>`;
    }).join("")}</div>`;

  let hours = "";
  for (let h = 1; h < 24; h++) hours += `<span style="top:${h * HOUR_PX}px">${esc(F.hour.format(new Date(2000, 0, 1, h)))}</span>`;
  const nowMin = (Date.now() - today) / 6e4;
  $("#grid").innerHTML = `<div class="hours">${hours}</div>` + days.map(d => {
    const k = ymd(d), cls = [k === tk ? "today" : "", d.getDay() === 0 || d.getDay() === 6 ? "wknd" : ""].join(" ");
    const blocks = layout(timed.get(k)).map(blockHTML).join("");
    const now = k === tk ? `<div class="now" style="top:${nowMin / 60 * HOUR_PX}px"></div>` : "";
    return `<div class="col ${cls}" data-date="${k}">${blocks}${now}</div>`;
  }).join("");
}
function scrollToStart() {
  const h = Number(state.settings?.day_start_hour ?? 7);
  $("#scroller").scrollTop = Math.max(0, h * HOUR_PX - 6);
}
function tickNow() {
  const el = $(".now"); if (!el) return;
  const today = startOfDay(new Date());
  if (!$(`.col[data-date="${ymd(today)}"]`)) return;
  el.style.top = `${(Date.now() - today) / 6e4 / 60 * HOUR_PX}px`;
}

/* ---------- lookups ---------- */
const taskById = id => state.tasks.find(t => t.id === id);
const eventById = id => state.events.find(e => e.id === id);

/* ---------- toast ---------- */
let toastTimer = null;
function toast(msg, opts = {}) {
  const el = $("#toast");
  el.className = "toast" + (opts.error ? " err" : "");
  el.innerHTML = `<span>${esc(msg)}</span>` + (opts.action ? `<button type="button">${esc(opts.action)}</button>` : "");
  el.hidden = false;
  if (opts.action) $("button", el).onclick = () => { el.hidden = true; opts.onAction(); };
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, opts.error ? 9000 : 5000);
}

/* ---------- task actions ---------- */
async function completeTask(id) {
  const t = taskById(id); if (!t) return;
  document.querySelectorAll(`[data-task="${CSS.escape(id)}"]`).forEach(el => el.classList.add("done"));
  closePop();
  const idx = state.tasks.indexOf(t);
  setTimeout(() => { state.tasks = state.tasks.filter(x => x !== t); render(); }, 450);
  try {
    await api("/api/task/close", { id });
    toast(`Completed "${plain(t.content).slice(0, 60)}"`, {
      action: "Undo", onAction: async () => {
        try { await api("/api/task/reopen", { id }); state.tasks.splice(Math.min(idx, state.tasks.length), 0, t); render(); load(); }
        catch (e) { toast(e.message, { error: true }); }
      }
    });
  } catch (e) {
    setTimeout(() => { if (!state.tasks.includes(t)) state.tasks.splice(idx, 0, t); render(); }, 500);
    toast(e.message, { error: true });
  }
}
function describeWhen(date, time) {
  const d = parseYmd(date);
  return time ? `${F.dayShort.format(d)}, ${F.time.format(atMinutes(d, toMin(time)))}` : F.dayShort.format(d);
}
function toMin(time) { const [h, m] = time.split(":").map(Number); return h * 60 + m; }
async function scheduleTask(t, date, time, duration) {
  const before = { due: t.due ? { ...t.due } : null, duration: t.duration };
  t.due = { ...(t.due || {}), date, time: time || null, recurring: false, string: "" };
  if (time && duration) t.duration = duration;
  render();
  try {
    await api("/api/task/schedule", { id: t.id, date, time: time || null, duration: time ? duration : null });
    toast(`Scheduled for ${describeWhen(date, time)}`);
    load();
  } catch (e) {
    t.due = before.due; t.duration = before.duration; render();
    toast(e.message, { error: true });
  }
}
async function createTask(content, date, time) {
  const duration = Number(state.settings?.default_task_minutes || 30);
  try {
    await api("/api/task/create", { content, date, time, duration: time ? duration : null });
    toast(`Added "${content.slice(0, 60)}" to your Todoist inbox`);
    load();
  } catch (e) { toast(e.message, { error: true }); }
}

/* ---------- popover ---------- */
const pop = $("#pop");
function openPop(html, anchor) {
  pop.innerHTML = `<button class="close" data-act="closepop" aria-label="Close">×</button>` + html;
  pop.hidden = false;
  const r = anchor.getBoundingClientRect(), pw = pop.offsetWidth, ph = pop.offsetHeight;
  let x = r.right + 8;
  if (x + pw > innerWidth - 8) x = r.left - pw - 8;
  if (x < 8) x = clamp(r.left, 8, innerWidth - pw - 8);
  let y = clamp(r.top, 8, Math.max(8, innerHeight - ph - 8));
  pop.style.left = `${x}px`; pop.style.top = `${y}px`;
  const focusable = pop.querySelector(".actions .btn"); if (focusable) focusable.focus({ preventScroll: true });
}
function closePop() { pop.hidden = true; pop.innerHTML = ""; }
const BUSY_LABEL = { busy: "Busy", tentative: "Tentative", free: "Free", oof: "Away", elsewhere: "Working elsewhere" };
function showEvent(ev, anchor) {
  let when;
  if (ev.allDay) {
    const s = parseYmd(ev.start), e = addDays(parseYmd(ev.end), -1);
    when = +s === +e ? `${F.dayLong.format(s)}, all day` : `${F.dayShort.format(s)} to ${F.dayShort.format(e)}, all day`;
  } else {
    const s = parseLocal(ev.start), e = parseLocal(ev.end);
    when = ymd(s) === ymd(e) || (e - startOfDay(e) === 0 && e - s <= 864e5)
      ? `${F.dayLong.format(s)}, ${timeRange(s, e)}` : `${F.dayShort.format(s)} ${F.time.format(s)} to ${F.dayShort.format(e)} ${F.time.format(e)}`;
  }
  const join = ev.url ? `<div class="actions"><a class="btn primary" href="${esc(ev.url)}" target="_blank" rel="noopener">Join meeting</a></div>` : "";
  const look = catLook(ev), cats = ev.categories || [];
  const catLine = cats.length ? `<p class="cats">${look.style ? `<span class="swatch" style="${look.style}" aria-hidden="true"></span>` : ""}${esc(cats.join(", "))}</p>` : "";
  openPop(`<h3>${esc(ev.title)}</h3><p>${esc(when)}</p>${ev.location ? `<p>${esc(ev.location)}</p>` : ""}` +
    `<p>${esc(BUSY_LABEL[ev.busy] || "Busy")}${ev.cancelled ? ", cancelled" : ""}</p>${catLine}${join}` +
    `<p class="note">From Outlook. Edit it in Outlook.</p>`, anchor);
}
const PRIO_LABEL = { 4: "Priority 1", 3: "Priority 2", 2: "Priority 3", 1: "No priority" };
function showTask(t, anchor) {
  let when = "No date";
  if (t.due) {
    const d = parseYmd(t.due.date);
    if (t.due.time) {
      const s = atMinutes(d, toMin(t.due.time)), dur = t.duration || Number(state.settings?.default_task_minutes || 30);
      when = `${F.dayLong.format(d)}, ${timeRange(s, new Date(+s + dur * 6e4))}` + (t.duration ? "" : " (no length set)");
    } else when = `${F.dayLong.format(d)}, no time`;
  }
  const bits = [t.project ? `<span class="dot" style="background:${esc(t.projectColor)}"></span>${esc(t.project)}` : "", PRIO_LABEL[t.priority] || ""]
    .filter(Boolean).join(", ");
  const recur = t.due?.recurring ? `<p>Repeats: ${esc(t.due.string || "recurring")}</p>` : "";
  const unset = t.due?.time && !t.due.recurring ? `<button class="btn" data-act="untime" data-id="${esc(t.id)}">Remove time</button>` : "";
  openPop(`<h3>${esc(plain(t.content))}</h3><p>${esc(when)}</p><p>${bits}</p>${recur}` +
    (t.labels.length ? `<p>${t.labels.map(l => "@" + esc(l)).join(" ")}</p>` : "") +
    (t.description ? `<p class="desc">${esc(t.description)}</p>` : "") +
    `<div class="actions"><button class="btn primary" data-act="complete" data-id="${esc(t.id)}">Complete</button>${unset}` +
    `<a class="btn" href="${esc(t.url)}" target="_blank" rel="noopener">Open in Todoist</a></div>` +
    (t.inOutlook ? `<p class="note">Also in your Outlook calendar.</p>` : "") +
    (t.due?.recurring ? `<p class="note">Recurring tasks are rescheduled in Todoist so their pattern stays intact.</p>` : ""), anchor);
}

/* ---------- drag to schedule / resize ---------- */
let drag = null, suppressClick = false, autoScroll = null;
const scroller = $("#scroller");

function onPointerDown(e) {
  if (e.button !== 0 || e.target.closest(".check")) return;
  const el = e.target.closest("[data-task]");
  if (!el || el.closest("#pop")) return;
  const t = taskById(el.dataset.task); if (!t) return;
  const resize = !!e.target.closest(".resize");
  const r = el.getBoundingClientRect();
  drag = { el, t, x0: e.clientX, y0: e.clientY, mode: resize ? "resize" : "move", started: false, target: null,
    grabMin: el.classList.contains("blk") && !resize ? ((e.clientY - r.top) / HOUR_PX) * 60 : 0 };
  el.setPointerCapture(e.pointerId);
  el.addEventListener("pointermove", onPointerMove);
  el.addEventListener("pointerup", onPointerUp);
  el.addEventListener("pointercancel", endDrag);
}
function startDrag() {
  const { t, el } = drag;
  if (t.due?.recurring) {
    toast("Recurring tasks keep their pattern in Todoist, so change their date or time there.");
    endDrag(); return false;
  }
  drag.started = true; state.dragging = true; closePop();
  document.body.classList.add("dragging");
  if (drag.mode === "move") {
    el.classList.add("lifted");
    const g = document.createElement("div");
    g.className = `ghost ${prioClass(t.priority)}`; g.textContent = plain(t.content);
    document.body.appendChild(g); drag.ghost = g;
  }
  return true;
}
function onPointerMove(e) {
  if (!drag) return;
  if (!drag.started) {
    if (Math.hypot(e.clientX - drag.x0, e.clientY - drag.y0) < 5) return;
    if (!startDrag()) return;
  }
  drag.lastX = e.clientX; drag.lastY = e.clientY;
  updateDrag();
  const r = scroller.getBoundingClientRect();
  const overX = e.clientX >= r.left && e.clientX <= r.right;
  const edge = !overX ? 0 : (e.clientY >= r.top && e.clientY < r.top + 28) ? -1 : (e.clientY > r.bottom - 28 && e.clientY < r.bottom + 40) ? 1 : 0;
  if (edge && !autoScroll) autoScroll = setInterval(() => { scroller.scrollTop += edge * 14; updateDrag(); }, 30);
  if (!edge && autoScroll) { clearInterval(autoScroll); autoScroll = null; }
}
function clearDropMarks() {
  document.querySelectorAll(".preview").forEach(p => p.remove());
  document.querySelectorAll(".ad-cell.drop").forEach(c => c.classList.remove("drop"));
}
function updateDrag() {
  const { t } = drag, x = drag.lastX, y = drag.lastY;
  const defMin = Number(state.settings?.default_task_minutes || 30);
  if (drag.mode === "resize") {
    const col = drag.el.closest(".col"); if (!col) return;
    const r = col.getBoundingClientRect(), start = toMin(t.due.time);
    const endMin = clamp(Math.round(((y - r.top) / HOUR_PX * 60) / SNAP) * SNAP, start + SNAP, 1440);
    drag.duration = endMin - start;
    drag.el.style.height = `${Math.max(drag.duration / 60 * HOUR_PX - 2, 18)}px`;
    const m = drag.el.querySelector(".m");
    if (m) m.textContent = timeRange(atMinutes(parseYmd(t.due.date), start), atMinutes(parseYmd(t.due.date), endMin));
    return;
  }
  if (drag.ghost) { drag.ghost.style.left = `${x + 12}px`; drag.ghost.style.top = `${y + 10}px`; }
  const under = document.elementFromPoint(x, y);
  const col = under && under.closest(".col"), cell = under && under.closest(".ad-cell");
  clearDropMarks();
  drag.target = null;
  if (col) {
    const r = col.getBoundingClientRect(), dur = t.duration || defMin;
    const min = clamp(Math.round((((y - r.top) / HOUR_PX) * 60 - drag.grabMin) / SNAP) * SNAP, 0, 1440 - SNAP);
    const p = document.createElement("div");
    p.className = "preview";
    p.style.top = `${min / 60 * HOUR_PX}px`;
    p.style.height = `${Math.max(Math.min(dur, 1440 - min) / 60 * HOUR_PX - 2, 18)}px`;
    p.textContent = F.time.format(atMinutes(parseYmd(col.dataset.date), min));
    col.appendChild(p);
    if (drag.ghost) drag.ghost.hidden = true;
    drag.target = { date: col.dataset.date, time: hhmm(min), duration: dur };
  } else if (cell) {
    cell.classList.add("drop");
    if (drag.ghost) drag.ghost.hidden = false;
    drag.target = { date: cell.dataset.date, time: null, duration: null };
  } else if (drag.ghost) drag.ghost.hidden = false;
}
function onPointerUp() {
  if (!drag) return;
  const d = drag;
  endDrag();
  if (!d.started) return;
  suppressClick = true; setTimeout(() => { suppressClick = false; }, 0);
  if (d.mode === "resize") {
    if (d.duration && d.duration !== d.t.duration) scheduleTask(d.t, d.t.due.date, d.t.due.time, d.duration);
    else render();
    return;
  }
  if (!d.target) return;
  const same = d.t.due && d.t.due.date === d.target.date && (d.t.due.time || null) === d.target.time;
  if (!same) scheduleTask(d.t, d.target.date, d.target.time, d.target.duration);
}
function endDrag() {
  if (!drag) return;
  const { el, ghost } = drag;
  el.removeEventListener("pointermove", onPointerMove);
  el.removeEventListener("pointerup", onPointerUp);
  el.removeEventListener("pointercancel", endDrag);
  el.classList.remove("lifted");
  if (ghost) ghost.remove();
  clearDropMarks();
  if (autoScroll) { clearInterval(autoScroll); autoScroll = null; }
  document.body.classList.remove("dragging");
  const wasDragging = state.dragging;
  drag = null; state.dragging = false;
  if (wasDragging && state.pendingRender) { state.pendingRender = false; render(); }
}

/* ---------- quick add by double-click ---------- */
function quickAdd(container, date, time) {
  document.querySelectorAll(".new-task").forEach(n => n.remove());
  const box = document.createElement("div");
  box.className = "new-task";
  if (time) box.style.top = `${toMin(time) / 60 * HOUR_PX}px`;
  const label = time ? `New task at ${F.time.format(atMinutes(parseYmd(date), toMin(time)))}` : `New task for ${F.dayShort.format(parseYmd(date))}`;
  box.innerHTML = `<input type="text" placeholder="${esc(label)}" aria-label="${esc(label)}" maxlength="500">`;
  container.appendChild(box);
  const input = $("input", box);
  input.focus();
  let done = false;
  const finish = (save) => {
    if (done) return; done = true;
    const v = input.value.trim(); box.remove();
    if (save && v) createTask(v, date, time);
  };
  input.addEventListener("keydown", ev => { if (ev.key === "Enter") finish(true); if (ev.key === "Escape") finish(false); ev.stopPropagation(); });
  input.addEventListener("blur", () => finish(true));
}

/* ---------- settings ---------- */
const dlg = $("#settings"), form = $("#settingsForm");
(function fillHours() {
  const sel = form.elements.day_start_hour;
  for (let h = 0; h <= 12; h++) sel.add(new Option(F.hour.format(new Date(2000, 0, 1, h)), h));
})();
function openSettings(intro = false) {
  const s = state.settings || {};
  if (state.demo) { toast("Settings are off in demo mode. Start the app without --demo to connect your accounts."); return; }
  $("#settingsIntro").hidden = !intro;
  $("#settingsError").hidden = true;
  $("#appVersion").textContent = s.version ? `Planner ${s.version}` : "";
  form.elements.todoist_token.value = "";
  form.elements.todoist_token.placeholder = s.has_token ? `Saved (ends ${s.token_hint.replace("…", "")}). Paste a new one to replace it.` : "Paste your token";
  form.elements.todoist_filter.value = s.todoist_filter || "";
  form.elements.ics_url.value = s.ics_url || "";
  for (const r of form.elements.calendar_source) r.checked = r.value === (s.calendar_source || "ics");
  $("#desktopOption").hidden = s.platform !== "win32" && s.calendar_source !== "outlook_desktop";
  form.elements.week_start.value = String(s.week_start ?? 1);
  form.elements.day_start_hour.value = String(s.day_start_hour ?? 7);
  form.elements.default_task_minutes.value = String(s.default_task_minutes ?? 30);
  form.elements.refresh_minutes.value = String(s.refresh_minutes ?? 5);
  form.elements.show_weekends.checked = s.show_weekends !== false;
  form.elements.outlook_writeback.checked = !!s.outlook_writeback;
  form.elements.outlook_show_as.value = s.outlook_show_as || "busy";
  const daysSel = form.elements.outlook_writeback_days, days = String(s.outlook_writeback_days ?? 14);
  if (![...daysSel.options].some(o => o.value === days)) daysSel.add(new Option(`${days} days ahead`, days));
  daysSel.value = days;
  $("#writebackSet").hidden = $("#desktopOption").hidden;
  syncIcsField();
  dlg.showModal();
}
function syncIcsField() {
  const ics = form.elements.calendar_source.value === "ics";
  $("#icsField").hidden = !ics; $("#icsHelp").hidden = !ics;
  const desktop = form.elements.calendar_source.value === "outlook_desktop", f = form.elements;
  f.outlook_writeback.disabled = !desktop;
  f.outlook_show_as.disabled = f.outlook_writeback_days.disabled = !desktop || !f.outlook_writeback.checked;
  $("#writebackNeeds").hidden = desktop;
}
form.addEventListener("change", e => { if (e.target.name === "calendar_source" || e.target.name === "outlook_writeback") syncIcsField(); });
form.addEventListener("submit", async e => {
  if (e.submitter && e.submitter.value === "cancel") return;
  e.preventDefault();
  const f = form.elements, src = f.calendar_source.value;
  const url = f.ics_url.value.trim();
  const fail = msg => { const err = $("#settingsError"); err.textContent = msg; err.hidden = false; err.scrollIntoView({ block: "nearest" }); };
  if (src === "ics" && url && !/^(https?|webcal):\/\//i.test(url)) return fail("The ICS link should start with https://. Copy it from Outlook's Publish a calendar section.");
  if (src === "ics" && /\.html?(\?|#|$)/i.test(url)) return fail("That's the HTML link. Outlook shows two links when you publish; copy the ICS one (it ends in .ics).");
  const body = {
    todoist_token: f.todoist_token.value.trim(), todoist_filter: f.todoist_filter.value.trim(),
    calendar_source: src, ics_url: url, week_start: Number(f.week_start.value), day_start_hour: Number(f.day_start_hour.value),
    default_task_minutes: Number(f.default_task_minutes.value), refresh_minutes: Number(f.refresh_minutes.value),
    show_weekends: f.show_weekends.checked,
    outlook_writeback: f.outlook_writeback.checked && src === "outlook_desktop",
    outlook_show_as: f.outlook_show_as.value, outlook_writeback_days: Number(f.outlook_writeback_days.value),
  };
  const s = state.settings || {};
  if (s.outlook_writeback && s.calendar_source === "outlook_desktop" && !body.outlook_writeback &&
      !confirm("Turning this off removes the upcoming task events Planner added to your Outlook calendar (they go to Deleted Items). Continue?")) return;
  const btn = $("#saveSettings"); btn.disabled = true;
  try {
    state.settings = await api("/api/settings", body);
    dlg.close(); restartTimers(); render(); scrollToStart(); load(true);
    toast("Settings saved");
  } catch (err) {
    const el = $("#settingsError"); el.textContent = err.message; el.hidden = false;
  } finally { btn.disabled = false; }
});

/* ---------- events ---------- */
document.addEventListener("pointerdown", onPointerDown);
document.addEventListener("click", e => {
  if (suppressClick) return;
  const t = e.target;
  const complete = t.closest("[data-complete]");
  if (complete) { e.stopPropagation(); completeTask(complete.dataset.complete); return; }
  const act = t.closest("[data-act]");
  if (act) {
    const a = act.dataset.act;
    if (a === "prev") step(-1); else if (a === "next") step(1); else if (a === "today") goToday();
    else if (a === "refresh") load(true); else if (a === "settings") { closePop(); openSettings(); }
    else if (a === "closepop") closePop();
    else if (a === "complete") completeTask(act.dataset.id);
    else if (a === "untime") { const tk = taskById(act.dataset.id); closePop(); if (tk) scheduleTask(tk, tk.due.date, null, null); }
    else if (a === "nodate") { state.nodateOpen = !state.nodateOpen; renderSide(); if (state.nodateOpen) $(".filter")?.focus(); }
    else if (a === "adexpand") { state.adExpanded = !state.adExpanded; render(); }
    else if (a === "side") $("#side").classList.toggle("open");
    return;
  }
  const view = t.closest("[data-view]");
  if (view) { setView(view.dataset.view); return; }
  const go = t.closest("[data-goto]");
  if (go) { state.anchor = parseYmd(go.dataset.goto); setView("day"); return; }
  const tEl = t.closest("[data-task]");
  if (tEl && !tEl.closest("#pop")) { const tk = taskById(tEl.dataset.task); if (tk) showTask(tk, tEl); return; }
  const eEl = t.closest("[data-ev]");
  if (eEl) { const ev = eventById(eEl.dataset.ev); if (ev) showEvent(ev, eEl); return; }
  if (!t.closest("#pop")) closePop();
  if (!t.closest("#side") && !t.closest('[data-act="side"]')) $("#side").classList.remove("open");
});
document.addEventListener("dblclick", e => {
  if (e.target.closest("[data-task],[data-ev],.new-task")) return;
  const col = e.target.closest(".col");
  if (col) {
    const r = col.getBoundingClientRect();
    const min = clamp(Math.floor(((e.clientY - r.top) / HOUR_PX * 60) / 30) * 30, 0, 1410);
    quickAdd(col, col.dataset.date, hhmm(min)); return;
  }
  const cell = e.target.closest(".ad-cell");
  if (cell) quickAdd(cell, cell.dataset.date, null);
});
document.addEventListener("input", e => {
  if (e.target.classList.contains("filter")) { state.nodateFilter = e.target.value; renderNodate(); }
});
document.addEventListener("keydown", e => {
  if (e.key === "Escape") { closePop(); $("#side").classList.remove("open"); return; }
  if (dlg.open || e.target.closest("input,textarea,select") || e.metaKey || e.ctrlKey || e.altKey) return;
  const el = e.target.closest && e.target.closest("[data-task],[data-ev]");
  if (el && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); el.click(); return; }
  const k = e.key.toLowerCase();
  if (e.key === "ArrowLeft") step(-1); else if (e.key === "ArrowRight") step(1);
  else if (k === "t") goToday(); else if (k === "d") setView("day"); else if (k === "w") setView("week");
  else if (k === "r") load(true);
});

/* ---------- timers + lifecycle ---------- */
let refreshTimer = null;
function restartTimers() {
  clearInterval(refreshTimer);
  const mins = Number(state.settings?.refresh_minutes || 5);
  refreshTimer = setInterval(() => { if (!state.dragging && !document.hidden) load(); }, mins * 6e4);
}
setInterval(tickNow, 30e3);
setInterval(() => { api("/api/ping", {}).catch(() => {}); }, 20e3);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) { api("/api/ping", {}).catch(() => {}); if (!state.updated || Date.now() - state.updated > 6e4) load(); }
});
addEventListener("pagehide", () => { try { navigator.sendBeacon(`/api/bye?t=${encodeURIComponent(TOKEN)}`); } catch { } });
addEventListener("resize", () => { if (!pop.hidden) closePop(); });
let lastDay = ymd(new Date());
setInterval(() => { const d = ymd(new Date()); if (d !== lastDay) { lastDay = d; render(); load(); } }, 60e3);

(async function init() {
  if (innerWidth < 820) state.view = "day";
  try { state.settings = await api("/api/settings"); state.demo = !!state.settings.demo; }
  catch (e) { state.errors = [{ source: "app", message: e.message }]; }
  if (state.view === "day" && state.settings && !state.settings.show_weekends) {
    while (state.anchor.getDay() === 0 || state.anchor.getDay() === 6) state.anchor = addDays(state.anchor, 1);
  }
  render(); restartTimers();
  const s = state.settings;
  if (s && !s.demo && !s.has_token) openSettings(true);
  load();
})();
