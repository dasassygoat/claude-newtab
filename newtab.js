'use strict';
const $ = (id) => document.getElementById(id);
let settings = null;
let cache = null;
let graphAuth = undefined;   // null = configured but not signed in

const CLAUDE_LINKS = [
  ['New chat', '/new'], ['Chats', '/recents'], ['Projects', '/projects'], ['Claude Code', '/code'], ['Artifacts', '/artifacts'],
];

function send(type, extra) {
  return new Promise((resolve) => chrome.runtime.sendMessage(Object.assign({ type }, extra || {}), (r) => resolve(r || { ok: false, error: chrome.runtime.lastError && chrome.runtime.lastError.message })));
}

// ---------- Clock / greeting ----------
function tickClock() {
  const now = new Date();
  $('clock').textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  $('date').textContent = now.toLocaleDateString([], { weekday: 'long', month: 'long', day: 'numeric' });
  const h = now.getHours();
  const part = h < 5 ? 'Up late' : h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
  $('greeting').textContent = settings && settings.name ? `${part}, ${settings.name}` : part;
  if (cache) renderAgenda();   // keep "now" highlight and past dimming fresh
}

// ---------- Claude prompt ----------
function claudeUrl(path, q) {
  const base = (settings && settings.claudeBase || 'https://claude.ai').replace(/\/+$/, '');
  return q ? `${base}/new?q=${encodeURIComponent(q)}` : base + path;
}
function openClaude(q, newTab) {
  const url = claudeUrl('/new', q);
  if (newTab) window.open(url, '_blank'); else location.href = url;
}
function autosize(el) { el.style.height = 'auto'; el.style.height = Math.min(el.scrollHeight, 220) + 'px'; }

function planDayPrompt() {
  const today = (cache && cache.events || []).filter((e) => sameDay(e.start, Date.now()) || (e.allDay && e.start <= Date.now() && e.end > Date.now()));
  const lines = today.map((e) => e.allDay ? `- All day: ${e.title}` : `- ${fmtTime(e.start)}–${fmtTime(e.end)}: ${e.title}${e.location ? ` (${e.location})` : ''}`);
  const dateStr = new Date().toLocaleDateString([], { weekday: 'long', month: 'long', day: 'numeric' });
  return `Help me plan my day. Today is ${dateStr}, and it's ${fmtTime(Date.now())} now.\n\nMy calendar:\n${lines.length ? lines.join('\n') : '- (nothing scheduled)'}\n\nAsk me what else I need to get done, then suggest a realistic schedule with focus blocks around my meetings.`;
}

function renderChips() {
  const box = $('chips');
  box.innerHTML = '';
  for (const p of settings.prompts || []) {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'chip'; b.textContent = p.label;
    b.addEventListener('click', (ev) => {
      if (p.text === '__PLAN_DAY__') { openClaude(planDayPrompt(), ev.ctrlKey || ev.metaKey); return; }
      if (/[:?]\s*$/.test(p.text)) {          // a prefix: put it in the box to complete
        const t = $('prompt'); t.value = p.text; autosize(t); t.focus(); t.setSelectionRange(t.value.length, t.value.length);
      } else openClaude(p.text, ev.ctrlKey || ev.metaKey);
    });
    box.appendChild(b);
  }
  const nav = $('claudeLinks');
  nav.innerHTML = '';
  for (const [label, path] of CLAUDE_LINKS) {
    const a = document.createElement('a'); a.href = claudeUrl(path); a.textContent = label; nav.appendChild(a);
  }
}

// ---------- Agenda ----------
function sameDay(ms, ref) { const a = new Date(ms), b = new Date(ref); return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function fmtTime(ms) { return new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).replace(' ', ' '); }
function dayLabel(d, i) {
  if (i === 0) return 'Today';
  if (i === 1) return 'Tomorrow';
  return d.toLocaleDateString([], { weekday: 'long' });
}

function renderAgenda() {
  const box = $('agenda');
  const errs = $('agendaErrors');
  const hasSources = (settings.feeds && settings.feeds.some((f) => f.url)) || (settings.graph && settings.graph.clientId);
  box.innerHTML = ''; errs.textContent = '';
  if (!hasSources) {
    box.innerHTML = `<div class="setup">No calendars yet. <a href="options.html">Open settings</a> to sign in with Microsoft 365 or add an iCal (.ics) feed from Outlook or Google Calendar.</div>`;
    return;
  }
  const needsSignIn = graphAuth === null || ((cache && cache.errors) || []).some((e) => /sign in again|Not signed in/i.test(e));
  if (settings.graph && settings.graph.clientId && needsSignIn) {
    const div = document.createElement('div');
    div.className = 'setup';
    div.innerHTML = `Microsoft 365 is configured but not signed in yet. <button type="button" class="chip" id="graphSignInBtn">Sign in with Microsoft</button> <span class="small" id="graphSignInStatus"></span>`;
    div.querySelector('#graphSignInBtn').addEventListener('click', async () => {
      const st = div.querySelector('#graphSignInStatus');
      st.textContent = 'Opening Microsoft sign-in…';
      const r = await send('graphSignIn');
      st.textContent = r.ok ? `Signed in as ${r.auth.email}` : `Sign-in failed: ${r.error}`;
      if (r.ok) { graphAuth = r.auth; const res = await send('refresh'); if (res.cache) cache = res.cache; renderAgenda(); }
    });
    box.appendChild(div);
  }
  const now = Date.now();
  const days = Math.max(1, settings.daysAhead || 3);
  const events = (cache && cache.events) || [];
  const frag = document.createDocumentFragment();
  let any = false;
  for (let i = 0; i < days; i++) {
    const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + i);
    const dayStart = d.getTime(), dayEnd = dayStart + 86400000;
    const list = events.filter((e) => e.start < dayEnd && e.end > dayStart)
      .sort((a, b) => (a.allDay !== b.allDay ? (a.allDay ? -1 : 1) : a.start - b.start));
    if (!list.length && i > 0) continue;       // only always show today
    any = any || list.length > 0;
    const sec = document.createElement('div');
    sec.className = 'day' + (i === 0 ? ' today' : '');
    sec.innerHTML = `<h3>${dayLabel(d, i)} <span class="sub">${d.toLocaleDateString([], { month: 'short', day: 'numeric' })}</span></h3>`;
    if (!list.length) {
      sec.insertAdjacentHTML('beforeend', `<div class="empty">Nothing scheduled — a clear day.</div>`);
    }
    for (const e of list) {
      const el = document.createElement('a');
      el.className = 'ev';
      if (e.url) { el.href = e.url; el.target = '_blank'; el.rel = 'noopener'; }
      if (!e.allDay && e.end <= now) el.classList.add('past');
      if (e.start <= now && e.end > now && !e.allDay) el.classList.add('now');
      const time = e.allDay ? 'All day'
        : (e.start < dayStart ? '…' : fmtTime(e.start)) + (e.end > dayEnd ? ' →' : ` – ${fmtTime(e.end)}`);
      const meta = [];
      if (e.location && !/^https?:\/\//.test(e.location)) meta.push(escapeHtml(e.location));
      if (e.organizer) meta.push(escapeHtml(e.organizer));
      if (e.source) meta.push(escapeHtml(e.source));
      const join = e.joinUrl ? `<a class="join" href="${escapeAttr(e.joinUrl)}" target="_blank" rel="noopener">Join</a>` : '';
      el.innerHTML = `<div class="time">${time}</div><div><div class="title"><span class="dot" style="background:${escapeAttr(e.color || '#888')}"></span><span class="t" title="${escapeAttr(e.title)}">${escapeHtml(e.title)}</span>${join}</div>${meta.length ? `<div class="meta">${meta.map((m) => `<span>${m}</span>`).join('')}</div>` : ''}</div>`;
      sec.appendChild(el);
    }
    frag.appendChild(sec);
  }
  box.appendChild(frag);
  if (cache && cache.errors && cache.errors.length) errs.textContent = cache.errors.join(' · ');
  $('updated').textContent = cache && cache.fetchedAt ? `updated ${fmtTime(cache.fetchedAt)}` : '';
}

function escapeHtml(s) { return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function escapeAttr(s) { return escapeHtml(s); }

// ---------- Links ----------
function renderLinks() {
  const box = $('links');
  box.innerHTML = '';
  for (const l of settings.links || []) {
    if (!l.url) continue;
    const a = document.createElement('a');
    a.href = l.url;
    let host = ''; try { host = new URL(l.url).hostname; } catch (e) { /* ignore */ }
    const img = document.createElement('img');
    img.alt = ''; img.src = `https://www.google.com/s2/favicons?domain=${encodeURIComponent(host)}&sz=32`;
    img.addEventListener('error', () => img.remove());
    a.appendChild(img);
    a.appendChild(document.createTextNode(l.label || host));
    box.appendChild(a);
  }
  if (!box.children.length) box.innerHTML = '<div class="empty">Add links in settings.</div>';
}

// ---------- Boot ----------
async function load() {
  const s = await send('getSettings');
  settings = s.settings || {};
  renderChips(); renderLinks(); tickClock();
  if (settings.graph && settings.graph.clientId) { const g = await send('graphStatus'); graphAuth = g.auth || null; }
  const r = await send('getEvents');
  cache = r.cache || { events: [], errors: [] };
  renderAgenda();
}

$('askForm').addEventListener('submit', (ev) => {
  ev.preventDefault();
  const q = $('prompt').value.trim();
  openClaude(q, false);
});
$('prompt').addEventListener('keydown', (ev) => {
  if (ev.key === 'Enter' && !ev.shiftKey) {
    ev.preventDefault();
    const q = $('prompt').value.trim();
    openClaude(q, ev.ctrlKey || ev.metaKey);
  }
});
$('prompt').addEventListener('input', (ev) => autosize(ev.target));
document.addEventListener('keydown', (ev) => {
  if (ev.key === '/' && document.activeElement !== $('prompt') && !/INPUT|TEXTAREA/.test(document.activeElement.tagName)) { ev.preventDefault(); $('prompt').focus(); }
});
$('refresh').addEventListener('click', async () => {
  $('refresh').classList.add('spin');
  const r = await send('refresh');
  if (r.cache) cache = r.cache;
  renderAgenda();
  $('refresh').classList.remove('spin');
});
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && changes.cache) { cache = changes.cache.newValue; renderAgenda(); }
  if (area === 'sync' && changes.settings) load();
});

load();
setInterval(tickClock, 1000);
window.addEventListener('load', () => $('prompt').focus());
