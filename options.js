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

// ---------- Background ----------
const GRADIENTS = {
  dusk: 'linear-gradient(135deg, #1f1c2c 0%, #928dab 100%)',
  ocean: 'linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%)',
  sunset: 'linear-gradient(135deg, #2b1d16 0%, #7a3b2e 45%, #d97757 100%)',
  forest: 'linear-gradient(135deg, #0b2a1e 0%, #1f4d3a 60%, #3c7a5a 100%)',
  slate: 'linear-gradient(160deg, #1b1f24 0%, #2d3440 100%)',
  aurora: 'linear-gradient(135deg, #0b1026 0%, #123a5c 40%, #1c8b7a 80%, #6fd3a6 100%)',
  peach: 'linear-gradient(135deg, #fbe9dc 0%, #f6c9b4 50%, #e9a58a 100%)',
  mist: 'linear-gradient(160deg, #eef1f5 0%, #d8dee8 100%)',
};
const BG_DEFAULT = { type: 'default', color: '#1f1c2c', gradient: 'dusk', imageUrl: '', dim: 0.35, frost: true, tone: 'auto', nasaKey: '', apodHd: true };
const DAILY_TYPES = ['apod', 'commons', 'rotate'];
function rotateSource() { const day = Math.round(new Date().setHours(0, 0, 0, 0) / 86400000); return day % 2 === 0 ? 'apod' : 'commons'; }
let hasUploadedImage = false;

function bgType() { const r = document.querySelector('input[name=bgType]:checked'); return r ? r.value : 'default'; }
function collectBg() {
  return {
    type: bgType(), color: $('bgColor').value, gradient: document.querySelector('.swatch.selected')?.dataset.name || 'dusk',
    imageUrl: $('bgImageUrl').value.trim(), dim: parseFloat($('bgDim').value), frost: $('bgFrost').checked, tone: $('bgTone').value,
    nasaKey: $('nasaKey').value.trim(), apodHd: $('apodHd').checked,
  };
}
async function updateBgUI() {
  const type = bgType();
  $('bgColorRow').hidden = type !== 'color';
  $('bgGradientRow').hidden = type !== 'gradient';
  $('bgImageRow').hidden = type !== 'image';
  $('bgApodRow').hidden = type !== 'apod' && type !== 'rotate';
  $('bgCommonsRow').hidden = type !== 'commons';
  $('bgDailyRow').hidden = !DAILY_TYPES.includes(type);
  $('bgExtras').hidden = type === 'default';
  $('bgDimVal').textContent = Math.round(parseFloat($('bgDim').value) * 100) + '%';
  $('bgDimDailyVal').textContent = $('bgDimVal').textContent;
  const pv = $('bgPreview'); const bg = collectBg(); pv.querySelector('span').textContent = 'Preview';
  pv.style.backgroundImage = ''; pv.style.backgroundColor = ''; pv.style.setProperty('--pdim', '0');
  if (bg.type === 'color') pv.style.backgroundColor = bg.color;
  else if (bg.type === 'gradient') pv.style.backgroundImage = GRADIENTS[bg.gradient];
  else if (DAILY_TYPES.includes(bg.type)) {
    pv.style.setProperty('--pdim', String(bg.dim));
    const r = await send('getDaily', { source: bg.type === 'rotate' ? rotateSource() : bg.type });
    if (r.ok && r.daily && r.daily.imageUrl) { pv.style.backgroundImage = `url("${r.daily.imageUrl.replace(/"/g, '%22')}")`; pv.querySelector('span').textContent = r.daily.title; }
    else pv.querySelector('span').textContent = `Could not load: ${r.error || (r.daily && r.daily.error) || 'unknown error'}`;
    return;
  }
  else if (bg.type === 'image') {
    let src = bg.imageUrl;
    if (!src) { try { src = (await chrome.storage.local.get('bgImage')).bgImage || ''; } catch (e) { src = ''; } }
    if (src) { pv.style.backgroundImage = `url("${src.replace(/"/g, '%22')}")`; pv.style.setProperty('--pdim', String(bg.dim)); }
  }
}
function initBg(bg) {
  bg = Object.assign({}, BG_DEFAULT, bg || {});
  const radio = document.querySelector(`input[name=bgType][value="${bg.type}"]`) || document.querySelector('input[name=bgType][value=default]');
  radio.checked = true;
  $('bgColor').value = bg.color; $('bgImageUrl').value = bg.imageUrl; $('bgDim').value = bg.dim; $('bgDimDaily').value = bg.dim; $('nasaKey').value = bg.nasaKey || ''; $('apodHd').checked = bg.apodHd !== false; $('bgFrost').checked = bg.frost !== false; $('bgTone').value = bg.tone;
  const sw = $('bgSwatches'); sw.innerHTML = '';
  for (const [name, css] of Object.entries(GRADIENTS)) {
    const d = document.createElement('div'); d.className = 'swatch' + (name === bg.gradient ? ' selected' : ''); d.dataset.name = name; d.style.backgroundImage = css; d.textContent = name; d.title = name;
    d.addEventListener('click', () => { sw.querySelectorAll('.swatch').forEach((x) => x.classList.remove('selected')); d.classList.add('selected'); updateBgUI(); });
    sw.appendChild(d);
  }
  chrome.storage.local.get('bgImage').then((r) => { hasUploadedImage = !!r.bgImage; $('bgFileStatus').textContent = hasUploadedImage ? 'An uploaded picture is stored.' : ''; });
  updateBgUI();
}
document.querySelectorAll('input[name=bgType]').forEach((r) => r.addEventListener('change', updateBgUI));
['bgColor', 'bgImageUrl', 'bgDim', 'bgFrost', 'bgTone', 'nasaKey', 'apodHd'].forEach((id) => $(id).addEventListener('input', updateBgUI));
$('bgDimDaily').addEventListener('input', () => { $('bgDim').value = $('bgDimDaily').value; updateBgUI(); });

$('bgFile').addEventListener('change', async () => {
  const file = $('bgFile').files[0]; if (!file) return;
  setStatus($('bgFileStatus'), 'Processing…');
  try {
    const bmp = await createImageBitmap(file);
    const scale = Math.min(1, 2560 / bmp.width, 1600 / bmp.height);
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    let dataUrl = c.toDataURL('image/jpeg', 0.86);
    if (dataUrl.length > 6 * 1024 * 1024) dataUrl = c.toDataURL('image/jpeg', 0.7);
    await chrome.storage.local.set({ bgImage: dataUrl });
    hasUploadedImage = true;
    $('bgImageUrl').value = '';
    setStatus($('bgFileStatus'), `Stored (${Math.round(dataUrl.length / 1024)} KB). Click Save.`, 'ok');
  } catch (e) { setStatus($('bgFileStatus'), `Could not read image: ${e.message}`, 'err'); }
  updateBgUI();
});
$('bgFileClear').addEventListener('click', async () => { await chrome.storage.local.remove('bgImage'); hasUploadedImage = false; $('bgFile').value = ''; setStatus($('bgFileStatus'), 'Removed.'); updateBgUI(); });

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
  $('showPast').checked = !!settings.showPast;
  $('clientId').value = settings.graph.clientId || '';
  $('tenant').value = settings.graph.tenant || 'organizations';
  $('includeAllCalendars').checked = settings.graph.includeAllCalendars !== false;
  $('prompts').value = (settings.prompts || []).map((p) => `${p.label} | ${p.text}`).join('\n');
  $('links').value = (settings.links || []).map((l) => `${l.label} | ${l.url}`).join('\n');
  $('feeds').innerHTML = '';
  for (const f of settings.feeds || []) $('feeds').appendChild(feedRow(f));
  initBg(settings.bg);
  $('redirectUri').textContent = chrome.identity.getRedirectURL();
  await refreshGraphStatus();
}

function collect() {
  return {
    name: $('name').value.trim(),
    daysAhead: Math.min(14, Math.max(1, parseInt($('daysAhead').value, 10) || 3)),
    claudeBase: $('claudeBase').value.trim() || 'https://claude.ai',
    showPast: $('showPast').checked,
    graph: { clientId: $('clientId').value.trim(), tenant: $('tenant').value.trim() || 'organizations', includeAllCalendars: $('includeAllCalendars').checked },
    prompts: parseLines($('prompts').value, false),
    links: parseLines($('links').value, true),
    feeds: readFeeds(),
    bg: collectBg(),
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
