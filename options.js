'use strict';
const $ = (id) => document.getElementById(id);
function send(type, extra) {
  return new Promise((resolve) => chrome.runtime.sendMessage(Object.assign({ type }, extra || {}), (r) => resolve(r || { ok: false, error: chrome.runtime.lastError && chrome.runtime.lastError.message })));
}
let settings = null;

function setStatus(el, msg, cls) { el.textContent = msg || ''; el.className = 'status ' + (cls || ''); }

function feedRow(feed) {
  const div = document.createElement('div');
  div.className = 'feed row3';
  div.innerHTML = `
    <div><label>Name</label><input type="text" class="f-name" placeholder="Work"></div>
    <div><label>ICS URL</label><input type="url" class="f-url" placeholder="https://outlook.office365.com/owa/calendar/.../calendar.ics"></div>
    <div><label>Color</label><input type="color" class="f-color" style="width:100%;height:36px;border:1px solid var(--line);border-radius:8px;background:var(--bg)"></div>
    <div><button type="button" class="btn secondary small f-remove" title="Remove">✕</button></div>
    <div class="status f-status" style="grid-column:1/-1"></div>`;
  div.querySelector('.f-name').value = feed.name || '';
  div.querySelector('.f-url').value = feed.url || '';
  div.querySelector('.f-color').value = feed.color || '#6b8afd';
  div.querySelector('.f-remove').addEventListener('click', () => div.remove());
  return div;
}

function readFeeds() {
  return [...document.querySelectorAll('#feeds .feed')].map((d) => ({
    name: d.querySelector('.f-name').value.trim(), url: d.querySelector('.f-url').value.trim(), color: d.querySelector('.f-color').value,
  })).filter((f) => f.url);
}
function parseLines(text, urlMode) {
  return text.split('\n').map((l) => l.trim()).filter(Boolean).map((l) => {
    const i = l.indexOf('|');
    const label = i >= 0 ? l.slice(0, i).trim() : l, val = i >= 0 ? l.slice(i + 1).trim() : l;
    return urlMode ? { label, url: val } : { label, text: val };
  });
}

async function requestFeedPermissions(feeds) {
  const origins = [];
  for (const f of feeds) { try { origins.push(new URL(f.url).origin + '/*'); } catch (e) { /* ignore */ } }
  if (!origins.length) return true;
  const has = await chrome.permissions.contains({ origins });
  if (has) return true;
  return chrome.permissions.request({ origins });
}

async function refreshGraphStatus() {
  const r = await send('graphStatus');
  const a = r.auth;
  $('graphSignedIn').hidden = !a;
  if (a) { $('graphName').textContent = a.name || ''; $('graphEmail').textContent = a.email || ''; }
}

async function load() {
  const r = await send('getSettings');
  settings = r.settings;
  $('name').value = settings.name || '';
  $('daysAhead').value = settings.daysAhead || 3;
  $('claudeBase').value = settings.claudeBase || 'https://claude.ai';
  $('clientId').value = settings.graph.clientId || '';
  $('tenant').value = settings.graph.tenant || 'organizations';
  $('includeAllCalendars').checked = settings.graph.includeAllCalendars !== false;
  $('prompts').value = (settings.prompts || []).map((p) => `${p.label} | ${p.text}`).join('\n');
  $('links').value = (settings.links || []).map((l) => `${l.label} | ${l.url}`).join('\n');
  $('feeds').innerHTML = '';
  for (const f of settings.feeds || []) $('feeds').appendChild(feedRow(f));
  $('redirectUri').textContent = chrome.identity.getRedirectURL();
  await refreshGraphStatus();
}

function collect() {
  return {
    name: $('name').value.trim(),
    daysAhead: Math.min(14, Math.max(1, parseInt($('daysAhead').value, 10) || 3)),
    claudeBase: $('claudeBase').value.trim() || 'https://claude.ai',
    graph: { clientId: $('clientId').value.trim(), tenant: $('tenant').value.trim() || 'organizations', includeAllCalendars: $('includeAllCalendars').checked },
    prompts: parseLines($('prompts').value, false),
    links: parseLines($('links').value, true),
    feeds: readFeeds(),
  };
}

$('addFeed').addEventListener('click', () => $('feeds').appendChild(feedRow({})));

$('save').addEventListener('click', async () => {
  const s = collect();
  const granted = await requestFeedPermissions(s.feeds);     // must run inside the click gesture
  if (!granted) { setStatus($('saveStatus'), 'Permission to fetch the feed host was declined; feed will not load.', 'err'); }
  const r = await send('saveSettings', { settings: s });
  settings = s;
  setStatus($('saveStatus'), r.ok ? 'Saved. Refreshing calendars…' : `Error: ${r.error}`, r.ok ? 'ok' : 'err');
  if (r.ok) {
    const res = await send('refresh');
    const errs = (res.cache && res.cache.errors) || [];
    const n = (res.cache && res.cache.events || []).length;
    setStatus($('saveStatus'), errs.length ? `Saved. ${n} events loaded; problems: ${errs.join(' · ')}` : `Saved. ${n} events loaded.`, errs.length ? 'err' : 'ok');
  }
});

$('graphSignIn').addEventListener('click', async () => {
  const s = collect();
  await send('saveSettings', { settings: s });
  setStatus($('graphStatus'), 'Opening Microsoft sign-in…');
  const r = await send('graphSignIn');
  setStatus($('graphStatus'), r.ok ? `Signed in as ${r.auth.email}` : `Sign-in failed: ${r.error}`, r.ok ? 'ok' : 'err');
  await refreshGraphStatus();
});
$('graphSignOut').addEventListener('click', async () => { await send('graphSignOut'); setStatus($('graphStatus'), 'Signed out.'); await refreshGraphStatus(); });

load();
