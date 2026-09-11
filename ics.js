// Minimal iCalendar (RFC 5545) parser with recurrence expansion.
// Shared by background.js (importScripts) and nothing else; kept dependency-free.
(function (root) {
  'use strict';

  // Windows time zone names (as emitted by Outlook) -> IANA.
  const WINDOWS_TZ = {
    'Eastern Standard Time': 'America/New_York', 'Central Standard Time': 'America/Chicago',
    'Mountain Standard Time': 'America/Denver', 'US Mountain Standard Time': 'America/Phoenix',
    'Pacific Standard Time': 'America/Los_Angeles', 'Alaskan Standard Time': 'America/Anchorage',
    'Hawaiian Standard Time': 'Pacific/Honolulu', 'Atlantic Standard Time': 'America/Halifax',
    'Newfoundland Standard Time': 'America/St_Johns', 'Canada Central Standard Time': 'America/Regina',
    'Central America Standard Time': 'America/Guatemala', 'SA Pacific Standard Time': 'America/Bogota',
    'Venezuela Standard Time': 'America/Caracas', 'SA Western Standard Time': 'America/La_Paz',
    'Argentina Standard Time': 'America/Buenos_Aires', 'E. South America Standard Time': 'America/Sao_Paulo',
    'UTC': 'UTC', 'GMT Standard Time': 'Europe/London', 'Greenwich Standard Time': 'Atlantic/Reykjavik',
    'W. Europe Standard Time': 'Europe/Berlin', 'Central Europe Standard Time': 'Europe/Budapest',
    'Romance Standard Time': 'Europe/Paris', 'Central European Standard Time': 'Europe/Warsaw',
    'E. Europe Standard Time': 'Europe/Chisinau', 'FLE Standard Time': 'Europe/Kiev',
    'GTB Standard Time': 'Europe/Bucharest', 'Russian Standard Time': 'Europe/Moscow',
    'Turkey Standard Time': 'Europe/Istanbul', 'Israel Standard Time': 'Asia/Jerusalem',
    'Arabian Standard Time': 'Asia/Dubai', 'Arab Standard Time': 'Asia/Riyadh',
    'India Standard Time': 'Asia/Kolkata', 'China Standard Time': 'Asia/Shanghai',
    'Singapore Standard Time': 'Asia/Singapore', 'Tokyo Standard Time': 'Asia/Tokyo',
    'Korea Standard Time': 'Asia/Seoul', 'AUS Eastern Standard Time': 'Australia/Sydney',
    'E. Australia Standard Time': 'Australia/Brisbane', 'AUS Central Standard Time': 'Australia/Darwin',
    'W. Australia Standard Time': 'Australia/Perth', 'New Zealand Standard Time': 'Pacific/Auckland',
    'South Africa Standard Time': 'Africa/Johannesburg', 'Egypt Standard Time': 'Africa/Cairo',
    'Mexico Standard Time': 'America/Mexico_City', 'Central Standard Time (Mexico)': 'America/Mexico_City',
  };

  function ianaZone(tzid) {
    if (!tzid) return null;
    tzid = tzid.replace(/^\/+/, '').replace(/^"|"$/g, '');
    if (WINDOWS_TZ[tzid]) return WINDOWS_TZ[tzid];
    try { new Intl.DateTimeFormat('en-US', { timeZone: tzid }); return tzid; } catch (e) { return null; }
  }

  const dtfCache = {};
  function tzParts(ms, zone) {
    const f = dtfCache[zone] || (dtfCache[zone] = new Intl.DateTimeFormat('en-US', {
      timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }));
    const o = {};
    for (const p of f.formatToParts(new Date(ms))) if (p.type !== 'literal') o[p.type] = parseInt(p.value, 10);
    if (o.hour === 24) o.hour = 0;
    return o;
  }
  // Wall-clock (y,m,d,h,mi,s) in `zone` -> epoch ms.
  function zonedToUtc(y, m, d, h, mi, s, zone) {
    if (!zone) return new Date(y, m - 1, d, h, mi, s).getTime();          // floating: local
    if (zone === 'UTC') return Date.UTC(y, m - 1, d, h, mi, s);
    let guess = Date.UTC(y, m - 1, d, h, mi, s);
    for (let i = 0; i < 2; i++) {
      const p = tzParts(guess, zone);
      const asUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
      guess -= (asUtc - Date.UTC(y, m - 1, d, h, mi, s));
    }
    return guess;
  }

  // Parse a DATE or DATE-TIME value. Returns {ms, allDay, zone, y,m,d,h,mi,s}
  function parseDateValue(value, params) {
    const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(value.trim());
    if (!m) return null;
    const y = +m[1], mo = +m[2], d = +m[3];
    if (!m[4] || (params.VALUE || '').toUpperCase() === 'DATE') {
      return { ms: new Date(y, mo - 1, d).getTime(), allDay: true, zone: null, y, m: mo, d, h: 0, mi: 0, s: 0 };
    }
    const h = +m[4], mi = +m[5], s = +(m[6] || 0);
    const zone = m[7] ? 'UTC' : ianaZone(params.TZID);
    return { ms: zonedToUtc(y, mo, d, h, mi, s, zone), allDay: false, zone, y, m: mo, d, h, mi, s };
  }

  function parseDuration(v) {
    const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(v.trim());
    if (!m) return 0;
    const ms = ((+m[2] || 0) * 7 * 86400 + (+m[3] || 0) * 86400 + (+m[4] || 0) * 3600 + (+m[5] || 0) * 60 + (+m[6] || 0)) * 1000;
    return m[1] === '-' ? -ms : ms;
  }

  function unescapeText(s) {
    return s.replace(/\\n/gi, '\n').replace(/\\,/g, ',').replace(/\;/g, ';').replace(/\\\\/g, '\\');
  }

  // Split ICS text into an array of {name, params, value} property lines (unfolded).
  function tokenize(text) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    const out = [];
    for (const raw of lines) {
      if (!raw) continue;
      if ((raw[0] === ' ' || raw[0] === '\t') && out.length) { out[out.length - 1] += raw.slice(1); continue; }
      out.push(raw);
    }
    return out.map((line) => {
      const idx = line.indexOf(':');
      let head = line, value = '';
      // Find the first ':' that isn't inside a quoted param value.
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (c === '"') inQ = !inQ;
        else if (c === ':' && !inQ) { head = line.slice(0, i); value = line.slice(i + 1); break; }
      }
      if (head === line && idx >= 0) { head = line.slice(0, idx); value = line.slice(idx + 1); }
      const [name, ...ps] = head.split(';');
      const params = {};
      for (const p of ps) {
        const eq = p.indexOf('=');
        if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
      }
      return { name: name.toUpperCase(), params, value };
    });
  }

  function parseRRule(v) {
    const r = {};
    for (const part of v.split(';')) {
      const [k, val] = part.split('=');
      if (!k || val === undefined) continue;
      r[k.toUpperCase()] = val;
    }
    const out = {
      freq: (r.FREQ || '').toUpperCase(),
      interval: Math.max(1, parseInt(r.INTERVAL || '1', 10)),
      count: r.COUNT ? parseInt(r.COUNT, 10) : null,
      until: r.UNTIL ? parseDateValue(r.UNTIL, {}) : null,
      byDay: r.BYDAY ? r.BYDAY.split(',').map((s) => {
        const m = /^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)$/.exec(s.trim());
        return m ? { ord: m[1] ? parseInt(m[1], 10) : 0, day: ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'].indexOf(m[2]) } : null;
      }).filter(Boolean) : null,
      byMonthDay: r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number) : null,
      byMonth: r.BYMONTH ? r.BYMONTH.split(',').map(Number) : null,
      wkst: r.WKST ? ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'].indexOf(r.WKST.toUpperCase()) : 1,
    };
    return out;
  }

  // Expand a recurring event's start instants within [winStart, winEnd] (ms).
  // Works in the event's wall-clock frame (zone of DTSTART) so DST is handled.
  function expandRRule(start, rule, winStart, winEnd, exdates, overridden) {
    const results = [];
    const untilMs = rule.until ? (rule.until.allDay ? rule.until.ms + 86399999 : rule.until.ms) : Infinity;
    const hardStop = Math.min(winEnd, untilMs);
    const MAX_ITER = 20000;
    let n = 0, emitted = 0;

    const toMs = (y, m, d) => start.allDay
      ? new Date(y, m - 1, d).getTime()
      : zonedToUtc(y, m, d, start.h, start.mi, start.s, start.zone);
    const push = (y, m, d) => {
      const ms = toMs(y, m, d);
      if (ms < start.ms - 1000) return true;              // before series start
      emitted++;
      if (rule.count && emitted > rule.count) return false;
      if (ms > hardStop) return false;
      if (ms + 86400000 * 366 >= winStart && !exdates.has(ms) && !overridden.has(ms)) {
        if (ms >= winStart - 86400000 * 366) results.push(ms);
      }
      return true;
    };

    const dim = (y, m) => new Date(y, m, 0).getDate();       // days in month (m 1-based)
    let y = start.y, m = start.m, d = start.d;

    if (rule.freq === 'DAILY') {
      let cur = new Date(y, m - 1, d);
      while (n++ < MAX_ITER) {
        const yy = cur.getFullYear(), mm = cur.getMonth() + 1, dd = cur.getDate();
        if (rule.byMonth && !rule.byMonth.includes(mm)) { /* skip */ }
        else if (rule.byDay && !rule.byDay.some((b) => b.day === cur.getDay())) { /* skip */ }
        else if (!push(yy, mm, dd)) break;
        cur.setDate(cur.getDate() + rule.interval);
        if (cur.getTime() > hardStop + 86400000) break;
      }
    } else if (rule.freq === 'WEEKLY') {
      const days = rule.byDay && rule.byDay.length ? rule.byDay.map((b) => b.day) : [new Date(y, m - 1, d).getDay()];
      let weekStart = new Date(y, m - 1, d);
      const diff = (weekStart.getDay() - rule.wkst + 7) % 7;
      weekStart.setDate(weekStart.getDate() - diff);
      outer: while (n++ < MAX_ITER) {
        for (let i = 0; i < 7; i++) {
          const cur = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + i);
          if (!days.includes(cur.getDay())) continue;
          if (rule.byMonth && !rule.byMonth.includes(cur.getMonth() + 1)) continue;
          if (!push(cur.getFullYear(), cur.getMonth() + 1, cur.getDate())) break outer;
        }
        weekStart.setDate(weekStart.getDate() + 7 * rule.interval);
        if (weekStart.getTime() > hardStop + 86400000) break;
      }
    } else if (rule.freq === 'MONTHLY') {
      let cy = y, cm = m;
      outer2: while (n++ < MAX_ITER) {
        if (!rule.byMonth || rule.byMonth.includes(cm)) {
          const candidates = [];
          if (rule.byDay && rule.byDay.length) {
            for (const b of rule.byDay) {
              const last = dim(cy, cm);
              const matching = [];
              for (let dd = 1; dd <= last; dd++) if (new Date(cy, cm - 1, dd).getDay() === b.day) matching.push(dd);
              if (b.ord > 0 && matching[b.ord - 1]) candidates.push(matching[b.ord - 1]);
              else if (b.ord < 0 && matching[matching.length + b.ord]) candidates.push(matching[matching.length + b.ord]);
              else if (b.ord === 0) candidates.push(...matching);
            }
          } else if (rule.byMonthDay) {
            for (const md of rule.byMonthDay) {
              const last = dim(cy, cm);
              const dd = md > 0 ? md : last + md + 1;
              if (dd >= 1 && dd <= last) candidates.push(dd);
            }
          } else if (d <= dim(cy, cm)) candidates.push(d);
          candidates.sort((a, b) => a - b);
          for (const dd of candidates) if (!push(cy, cm, dd)) break outer2;
        }
        cm += rule.interval;
        while (cm > 12) { cm -= 12; cy++; }
        if (new Date(cy, cm - 1, 1).getTime() > hardStop + 86400000) break;
      }
    } else if (rule.freq === 'YEARLY') {
      let cy = y;
      outer3: while (n++ < MAX_ITER) {
        const months = rule.byMonth || [m];
        for (const mm of months) {
          const candidates = [];
          if (rule.byDay && rule.byDay.length) {
            for (const b of rule.byDay) {
              const last = dim(cy, mm), matching = [];
              for (let dd = 1; dd <= last; dd++) if (new Date(cy, mm - 1, dd).getDay() === b.day) matching.push(dd);
              if (b.ord > 0 && matching[b.ord - 1]) candidates.push(matching[b.ord - 1]);
              else if (b.ord < 0 && matching[matching.length + b.ord]) candidates.push(matching[matching.length + b.ord]);
              else if (b.ord === 0) candidates.push(...matching);
            }
          } else if (rule.byMonthDay) candidates.push(...rule.byMonthDay.filter((x) => x >= 1 && x <= dim(cy, mm)));
          else if (d <= dim(cy, mm)) candidates.push(d);
          for (const dd of candidates.sort((a, b) => a - b)) if (!push(cy, mm, dd)) break outer3;
        }
        cy += rule.interval;
        if (new Date(cy, 0, 1).getTime() > hardStop + 86400000) break;
      }
    } else {
      results.push(start.ms);
    }
    return results.filter((ms) => ms >= winStart - 86400000 * 2 && ms <= winEnd);
  }

  // Parse ICS text and return event instances overlapping [winStart, winEnd].
  // Each: {uid, title, start, end, allDay, location, description, url}
  function parseICS(text, winStart, winEnd) {
    const props = tokenize(text);
    const events = [];
    let cur = null, depth = 0, inEvent = false;
    for (const p of props) {
      if (p.name === 'BEGIN') {
        depth++;
        if (p.value.toUpperCase() === 'VEVENT') { inEvent = true; cur = { props: [] }; }
        continue;
      }
      if (p.name === 'END') {
        depth--;
        if (p.value.toUpperCase() === 'VEVENT' && cur) { events.push(cur); cur = null; inEvent = false; }
        continue;
      }
      if (inEvent && cur) cur.props.push(p);
    }

    const parsed = events.map((ev) => {
      const o = { exdates: new Set(), rrule: null, recurrenceId: null };
      for (const p of ev.props) {
        switch (p.name) {
          case 'UID': o.uid = p.value; break;
          case 'SUMMARY': o.title = unescapeText(p.value); break;
          case 'LOCATION': o.location = unescapeText(p.value); break;
          case 'DESCRIPTION': o.description = unescapeText(p.value); break;
          case 'URL': o.url = p.value; break;
          case 'STATUS': o.status = p.value.toUpperCase(); break;
          case 'TRANSP': o.transp = p.value.toUpperCase(); break;
          case 'DTSTART': o.start = parseDateValue(p.value, p.params); break;
          case 'DTEND': o.end = parseDateValue(p.value, p.params); break;
          case 'DURATION': o.duration = parseDuration(p.value); break;
          case 'RRULE': o.rrule = parseRRule(p.value); break;
          case 'RECURRENCE-ID': o.recurrenceId = parseDateValue(p.value, p.params); break;
          case 'EXDATE':
            for (const v of p.value.split(',')) { const dv = parseDateValue(v, p.params); if (dv) o.exdates.add(dv.ms); }
            break;
          case 'X-MICROSOFT-CDO-ALLDAYEVENT': if (p.value.toUpperCase() === 'TRUE') o.msAllDay = true; break;
        }
      }
      return o;
    }).filter((o) => o.start);

    // Overrides (RECURRENCE-ID) replace the master's instance at that instant.
    const overridesByUid = new Map();
    for (const o of parsed) {
      if (o.recurrenceId && o.uid) {
        if (!overridesByUid.has(o.uid)) overridesByUid.set(o.uid, new Set());
        overridesByUid.get(o.uid).add(o.recurrenceId.ms);
      }
    }

    const out = [];
    for (const o of parsed) {
      if (o.status === 'CANCELLED') continue;
      const allDay = o.start.allDay;
      let durMs;
      if (o.end) durMs = o.end.ms - o.start.ms;
      else if (o.duration != null) durMs = o.duration;
      else durMs = allDay ? 86400000 : 0;
      if (allDay && durMs <= 0) durMs = 86400000;

      const starts = (o.rrule && !o.recurrenceId)
        ? expandRRule(o.start, o.rrule, winStart, winEnd, o.exdates, overridesByUid.get(o.uid) || new Set())
        : [o.start.ms];

      for (const s of starts) {
        const e = s + durMs;
        if (e <= winStart || s >= winEnd) continue;
        out.push({
          uid: o.uid || '', title: o.title || '(No title)', start: s, end: e, allDay,
          location: o.location || '', description: (o.description || '').slice(0, 500), url: o.url || '',
          busy: o.transp !== 'TRANSPARENT',
        });
      }
    }
    return out;
  }

  root.ICS = { parseICS, parseDateValue, parseRRule, expandRRule, tokenize, ianaZone };
})(typeof self !== 'undefined' ? self : this);
