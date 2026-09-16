// Service worker: fetches calendar data (ICS feeds + Microsoft Graph), caches it,
// and refreshes on a timer so the new tab page paints instantly.
importScripts('ics.js');

const REFRESH_MINUTES = 10;
const GRAPH = 'https://graph.microsoft.com/v1.0';

const DEFAULT_SETTINGS = {
  name: '',
  claudeBase: 'https://claude.ai',
  daysAhead: 3,
  prompts: [
    { label: 'Plan my day', text: '__PLAN_DAY__' },
    { label: 'Draft an email', text: 'Help me draft a concise, friendly email about: ' },
    { label: 'Explain this', text: 'Explain this clearly and briefly: ' },
    { label: 'Summarize', text: 'Summarize the key points of the following: ' },
  ],
  links: [
    { label: 'Outlook', url: 'https://outlook.office.com/mail/' },
    { label: 'Calendar', url: 'https://outlook.office.com/calendar/' },
    { label: 'Teams', url: 'https://teams.microsoft.com/' },
    { label: 'HubSpot', url: 'https://app.hubspot.com/' },
  ],
  feeds: [],                      // [{name, url, color}]
  graph: { clientId: '', tenant: 'organizations', includeAllCalendars: true },   // set your own in Settings
};

chrome.runtime.onInstalled.addListener(() => { scheduleAlarm(); refresh().catch(() => {}); });
chrome.runtime.onStartup.addListener(() => { scheduleAlarm(); refresh().catch(() => {}); });
chrome.alarms.onAlarm.addListener((a) => { if (a.name === 'refresh') refresh().catch(() => {}); });
chrome.action.onClicked.addListener(() => chrome.runtime.openOptionsPage());

function scheduleAlarm() {
  chrome.alarms.get('refresh', (a) => { if (!a) chrome.alarms.create('refresh', { periodInMinutes: REFRESH_MINUTES }); });
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      switch (msg && msg.type) {
        case 'getSettings': return sendResponse({ ok: true, settings: await getSettings() });
        case 'saveSettings': await saveSettings(msg.settings); refresh().catch(() => {}); return sendResponse({ ok: true });
        case 'getEvents': {
          const cache = await getCache();
          const stale = !cache.fetchedAt || Date.now() - cache.fetchedAt > REFRESH_MINUTES * 60000;
          if (stale) refresh().catch(() => {});
          return sendResponse({ ok: true, cache });
        }
        case 'refresh': return sendResponse({ ok: true, cache: await refresh() });
        case 'graphSignIn': return sendResponse(await graphSignIn());
        case 'graphSignOut': await chrome.storage.local.remove(['graphAuth']); await refresh().catch(() => {}); return sendResponse({ ok: true });
        case 'graphStatus': return sendResponse({ ok: true, auth: summarizeAuth((await chrome.storage.local.get('graphAuth')).graphAuth) });
        case 'testFeed': return sendResponse(await testFeed(msg.url));
        case 'getDaily': return sendResponse({ ok: true, daily: await getDaily(msg.source, await getSettings(), !!msg.force) });
        default: return sendResponse({ ok: false, error: 'unknown message' });
      }
    } catch (e) {
      sendResponse({ ok: false, error: String(e && e.message || e) });
    }
  })();
  return true;
});

async function getSettings() {
  const { settings } = await chrome.storage.sync.get('settings');
  const s = Object.assign({}, DEFAULT_SETTINGS, settings || {});
  const storedGraph = Object.fromEntries(Object.entries((settings && settings.graph) || {}).filter(([, v]) => v !== '' && v != null));
  s.graph = Object.assign({}, DEFAULT_SETTINGS.graph, storedGraph);
  if (!s.claudeBase) s.claudeBase = DEFAULT_SETTINGS.claudeBase;
  return s;
}
async function saveSettings(settings) { await chrome.storage.sync.set({ settings }); }
async function getCache() { return (await chrome.storage.local.get('cache')).cache || { events: [], errors: [], fetchedAt: 0 }; }

// ---------- Refresh ----------
let refreshing = null;
function refresh() {
  if (refreshing) return refreshing;
  refreshing = doRefresh().finally(() => { refreshing = null; });
  return refreshing;
}

async function doRefresh() {
  const settings = await getSettings();
  const now = new Date();
  const winStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const winEnd = winStart + (Math.max(1, settings.daysAhead) + 1) * 86400000;
  const errors = [];
  const events = [];

  const jobs = settings.feeds.filter((f) => f.url).map(async (feed, i) => {
    try {
      const res = await fetch(feed.url, { cache: 'no-store', credentials: 'omit' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('Not an iCalendar feed');
      for (const ev of ICS.parseICS(text, winStart, winEnd)) {
        events.push(Object.assign(ev, { source: feed.name || `Feed ${i + 1}`, color: feed.color || '#6b8afd', id: `${feed.url}|${ev.uid}|${ev.start}` }));
      }
    } catch (e) {
      errors.push(`${feed.name || feed.url}: ${e.message}`);
    }
  });

  if (settings.graph.clientId) {
    jobs.push(fetchGraphEvents(settings, winStart, winEnd).then((evs) => events.push(...evs)).catch((e) => errors.push(`Microsoft 365: ${e.message}`)));
  }

  if (settings.bg && (settings.bg.type === 'apod' || settings.bg.type === 'commons')) {
    jobs.push(getDaily(settings.bg.type, settings, false).catch(() => {}));   // prefetch today's picture
  }
  await Promise.all(jobs);
  events.sort((a, b) => (a.allDay !== b.allDay ? (a.allDay ? -1 : 1) : a.start - b.start));
  const cache = { events, errors, fetchedAt: Date.now(), winStart, winEnd };
  await chrome.storage.local.set({ cache });
  return cache;
}

async function testFeed(url) {
  const res = await fetch(url, { cache: 'no-store', credentials: 'omit' });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const text = await res.text();
  if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('Not an iCalendar feed');
  const count = (text.match(/BEGIN:VEVENT/g) || []).length;
  const name = (/X-WR-CALNAME:(.*)/.exec(text) || [])[1];
  return { ok: true, count, name: name ? name.trim() : '' };
}

// ---------- Microsoft Graph (OAuth 2.0 + PKCE via chrome.identity) ----------
const SCOPES = 'openid profile offline_access https://graph.microsoft.com/Calendars.Read';

function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
async function pkcePair() {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  return { verifier, challenge };
}
function decodeJwt(t) {
  try { return JSON.parse(atob(t.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))); } catch (e) { return {}; }
}
function summarizeAuth(auth) {
  if (!auth) return null;
  return { name: auth.name, email: auth.email, expiresAt: auth.expiresAt };
}

async function authorize(settings, { interactive, prompt, loginHint }) {
  const { clientId, tenant } = settings.graph;
  const redirect = chrome.identity.getRedirectURL();
  const { verifier, challenge } = await pkcePair();
  const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
  const params = {
    client_id: clientId, response_type: 'code', redirect_uri: redirect, response_mode: 'query',
    scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state, prompt,
  };
  if (loginHint) params.login_hint = loginHint;
  const authUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant || 'organizations')}/oauth2/v2.0/authorize?` + new URLSearchParams(params);
  const resultUrl = await chrome.identity.launchWebAuthFlow({ url: authUrl, interactive });
  if (!resultUrl) throw new Error('No response from sign-in');
  const u = new URL(resultUrl);
  if (u.searchParams.get('error')) throw new Error(`${u.searchParams.get('error')}: ${u.searchParams.get('error_description')}`);
  if (u.searchParams.get('state') !== state) throw new Error('State mismatch');
  const code = u.searchParams.get('code');
  const auth = await tokenRequest(settings, { grant_type: 'authorization_code', code, redirect_uri: redirect, code_verifier: verifier });
  auth.rtIssuedAt = Date.now();          // SPA refresh tokens live exactly 24h from this moment
  await chrome.storage.local.set({ graphAuth: auth });
  return auth;
}

async function graphSignIn() {
  const settings = await getSettings();
  if (!settings.graph.clientId) return { ok: false, error: 'Enter an Application (client) ID first.' };
  const auth = await authorize(settings, { interactive: true, prompt: 'select_account' });
  refresh().catch(() => {});
  return { ok: true, auth: summarizeAuth(auth) };
}

// Re-authenticate without UI using the browser's existing Microsoft session (prompt=none).
async function silentSignIn(settings, prev) {
  return authorize(settings, { interactive: false, prompt: 'none', loginHint: prev && prev.email });
}

async function tokenRequest(settings, params) {
  const { clientId, tenant } = settings.graph;
  const res = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenant || 'organizations')}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(Object.assign({ client_id: clientId, scope: SCOPES }, params)),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.error || `HTTP ${res.status}`);
  const claims = decodeJwt(data.id_token || '');
  const prev = (await chrome.storage.local.get('graphAuth')).graphAuth || {};
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || prev.refreshToken,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
    name: claims.name || prev.name || '', email: claims.preferred_username || prev.email || '',
  };
}

const RT_RENEW_AFTER = 20 * 3600 * 1000;   // renew silently once the SPA refresh token is 20h old

async function graphToken(settings) {
  let auth = (await chrome.storage.local.get('graphAuth')).graphAuth;
  if (!auth) throw new Error('Not signed in (open settings to sign in)');
  const rtAge = auth.rtIssuedAt ? Date.now() - auth.rtIssuedAt : Infinity;
  if (rtAge > RT_RENEW_AFTER) {
    try { return (await silentSignIn(settings, auth)).accessToken; } catch (e) { /* fall through to normal path */ }
  }
  if (Date.now() < auth.expiresAt) return auth.accessToken;
  try {
    if (!auth.refreshToken) throw new Error('no refresh token');
    auth = await tokenRequest(settings, { grant_type: 'refresh_token', refresh_token: auth.refreshToken });
    auth.rtIssuedAt = (await chrome.storage.local.get('graphAuth')).graphAuth.rtIssuedAt;
    await chrome.storage.local.set({ graphAuth: auth });
    return auth.accessToken;
  } catch (e) {
    try {
      return (await silentSignIn(settings, auth)).accessToken;
    } catch (e2) {
      throw new Error(`Session expired, sign in again (${e2.message})`);
    }
  }
}

async function graphGet(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.timezone="UTC"' } });
  const data = await res.json();
  if (!res.ok) throw new Error((data.error && data.error.message) || `HTTP ${res.status}`);
  return data;
}

async function fetchGraphEvents(settings, winStart, winEnd) {
  const token = await graphToken(settings);
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const q = new URLSearchParams({
    startDateTime: new Date(winStart).toISOString(), endDateTime: new Date(winEnd).toISOString(),
    $orderby: 'start/dateTime', $top: '100',
    $select: 'id,subject,start,end,isAllDay,location,webLink,showAs,isCancelled,organizer,onlineMeeting,onlineMeetingUrl,bodyPreview',
  });
  let calendars = [{ id: null, name: 'Calendar', hexColor: '' }];
  if (settings.graph.includeAllCalendars) {
    try {
      const list = await graphGet(token, `${GRAPH}/me/calendars?$select=id,name,hexColor,isDefaultCalendar`);
      if (list.value && list.value.length) calendars = list.value;
    } catch (e) { /* fall back to default calendar */ }
  }
  const all = [];
  await Promise.all(calendars.map(async (cal) => {
    const base = cal.id ? `${GRAPH}/me/calendars/${cal.id}/calendarView` : `${GRAPH}/me/calendarView`;
    let url = `${base}?${q}`;
    for (let page = 0; page < 5 && url; page++) {
      const data = await graphGet(token, url);
      for (const ev of data.value || []) {
        if (ev.isCancelled) continue;
        const start = ev.isAllDay ? localMidnightFromDateString(ev.start.dateTime) : Date.parse(ev.start.dateTime + 'Z');
        const end = ev.isAllDay ? localMidnightFromDateString(ev.end.dateTime) : Date.parse(ev.end.dateTime + 'Z');
        all.push({
          id: `graph|${ev.id}`, uid: ev.id, title: ev.subject || '(No title)', start, end, allDay: !!ev.isAllDay,
          location: (ev.location && ev.location.displayName) || '',
          description: ev.bodyPreview || '', url: ev.webLink || '',
          joinUrl: (ev.onlineMeeting && ev.onlineMeeting.joinUrl) || ev.onlineMeetingUrl || '',
          busy: ev.showAs !== 'free', source: cal.name || 'Calendar', color: cal.hexColor || '#0f6cbd',
          organizer: ev.organizer && ev.organizer.emailAddress ? ev.organizer.emailAddress.name : '',
        });
      }
      url = data['@odata.nextLink'] || null;
    }
  }));
  void tz;
  return all;
}

function localMidnightFromDateString(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]).getTime() : Date.parse(s);
}

// ---------- Picture of the day (NASA APOD, Wikimedia Commons) ----------
const DAILY_TTL = 6 * 3600 * 1000;
const pad2 = (n) => String(n).padStart(2, '0');
const localDateStr = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const utcDateStr = (d) => d.toISOString().slice(0, 10);
const stripHtml = (h) => String(h || '').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#0?39;/g, "'").replace(/\s+/g, ' ').trim();
const stripWiki = (w) => String(w || '')
  .replace(/\[\[(?:[^\]|]*\|)?([^\]]*)\]\]/g, '$1').replace(/\{\{[^}]*\}\}/g, '').replace(/'{2,3}/g, '')
  .replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();

async function getDaily(source, settings, force) {
  const key = `daily_${source}`;
  const cached = (await chrome.storage.local.get(key))[key];
  const today = localDateStr(new Date());
  if (!force && cached && cached.fetchedFor === today && Date.now() - cached.fetchedAt < DAILY_TTL && cached.imageUrl) return cached;
  try {
    const data = source === 'apod' ? await fetchApod(settings) : await fetchCommonsPotd();
    const rec = Object.assign(data, { source, fetchedFor: today, fetchedAt: Date.now() });
    await chrome.storage.local.set({ [key]: rec });
    return rec;
  } catch (e) {
    if (cached && cached.imageUrl) return Object.assign({}, cached, { error: e.message });
    throw e;
  }
}

async function fetchApod(settings) {
  const apiKey = (settings.bg && settings.bg.nasaKey) || 'DEMO_KEY';
  const hd = !(settings.bg && settings.bg.apodHd === false);   // HD by default
  let lastErr = 'no image found';
  for (let back = 0; back < 7; back++) {
    const d = new Date(); d.setDate(d.getDate() - back);
    const date = localDateStr(d);
    const res = await fetch(`https://api.nasa.gov/planetary/apod?api_key=${encodeURIComponent(apiKey)}&date=${date}`, { cache: 'no-store' });
    if (res.status === 429 || res.status === 403) throw new Error('NASA API rate limit reached (add your own free key in settings)');
    if (res.status === 404 || res.status === 400) { lastErr = `no APOD for ${date}`; continue; }
    if (!res.ok) throw new Error(`NASA API HTTP ${res.status}`);
    const j = await res.json();
    if (j.media_type !== 'image' || !j.url) { lastErr = `APOD for ${date} is a ${j.media_type}`; continue; }
    const ymd = date.replace(/-/g, '').slice(2);
    return {
      date, imageUrl: hd && j.hdurl ? j.hdurl : j.url, title: j.title || 'Astronomy Picture of the Day',
      text: j.explanation || '', credit: j.copyright ? `© ${String(j.copyright).trim()}` : 'NASA',
      sourceUrl: `https://apod.nasa.gov/apod/ap${ymd}.html`, sourceName: 'NASA APOD',
    };
  }
  throw new Error(lastErr);
}

async function commonsApi(params) {
  const url = 'https://commons.wikimedia.org/w/api.php?' + new URLSearchParams(Object.assign({ format: 'json', formatversion: '2' }, params));
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Commons API HTTP ${res.status}`);
  return res.json();
}

async function fetchCommonsPotd() {
  let lastErr = 'no picture found';
  for (let back = 0; back < 4; back++) {
    const d = new Date(); d.setUTCDate(d.getUTCDate() - back);
    const date = utcDateStr(d);
    const q1 = await commonsApi({ action: 'query', prop: 'images', titles: `Template:Potd/${date}` });
    const page = q1.query && q1.query.pages && q1.query.pages[0];
    const file = page && page.images && page.images[0] && page.images[0].title;
    if (!file) { lastErr = `no Commons POTD for ${date}`; continue; }
    const q2 = await commonsApi({ action: 'query', titles: file, prop: 'imageinfo', iiprop: 'url|extmetadata|mime', iiurlwidth: '2560' });
    const ii = q2.query && q2.query.pages && q2.query.pages[0] && q2.query.pages[0].imageinfo && q2.query.pages[0].imageinfo[0];
    if (!ii || !/^image\//.test(ii.mime || '')) { lastErr = `Commons POTD for ${date} is not an image`; continue; }
    const meta = ii.extmetadata || {};
    let caption = '';
    try {
      const q3 = await commonsApi({ action: 'query', prop: 'revisions', rvprop: 'content', rvslots: 'main', titles: `Template:Potd/${date} (en)` });
      const content = q3.query.pages[0].revisions[0].slots.main.content;
      const m = /\|\s*1\s*=\s*([\s\S]*?)(?=\|\s*\d+\s*=|\}\})/.exec(content);
      if (m) caption = stripWiki(m[1]);
    } catch (e) { /* caption optional */ }
    const title = file.replace(/^File:/, '').replace(/\.[a-z0-9]+$/i, '').replace(/_/g, ' ');
    const artist = stripHtml(meta.Artist && meta.Artist.value);
    const license = stripHtml(meta.LicenseShortName && meta.LicenseShortName.value);
    return {
      date, imageUrl: ii.thumburl || ii.url, title,
      text: caption || stripHtml(meta.ImageDescription && meta.ImageDescription.value).slice(0, 600),
      credit: [artist, license].filter(Boolean).join(' · '),
      sourceUrl: ii.descriptionurl || `https://commons.wikimedia.org/wiki/${encodeURIComponent(file)}`, sourceName: 'Wikimedia Commons',
    };
  }
  throw new Error(lastErr);
}
