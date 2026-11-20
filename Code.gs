// Soudia Coach Service dashboard (web app). Reads the monthly "Sale yyyy-mm" files in the "Sales Data" folder and the "Sales Target Plan" sheet.
const FOLDER_ID = '1_TNkmJ4jz0MEh3lgWSn4hZloP01M5zjX';
const MASTER_ID = '1P58XdF1BuzFxYEThATePL1AAkjTjfbWf6YWTSfbb2rQ';   // original "Sale" file: its Masterdata tab gives each counter's zone
const PLAN_ID = '1qqO2fc3y63yAlltMeNyQGZZw25bepnuQbaT81TIvXqA';     // "Sales Target Plan" Google Sheet
const MGMT_ID = '163SPGr-5GdgP6ZG04fmv6bRS-_86B3Tbwir9TxB2OX4';     // "Management report": tabs BP (bank) and PP Schedule (payments)
// Bank data (deposits and payments) is read from the "Bank Ledger (new)" tab. Last confirmed bank position (opening balance at the start of that day). The dashboard restarts its running balance from these figures.
// To confirm a newer position later, change the date and the four amounts (or type the balances into BP's "Op Balance" column on that day).
const CONFIRMED = { date: '2026-11-05', open: { 'BA 68': 256345, 'BA 125': 877, 'IBBL 2951': 29724, 'BA 145': 2371251 } };
const INCIDENT_ID = '1ZJJGuNTTGmtGRjg3u0eFxuMPbr5mcf_04w_p59MLNUk';   // "Incident Reporting Portal": tabs Reported and Followup
const MON_ID = '1tGU-ACen5D16BcwJTLczaVq0KDrL4JiPpofCr62T5lk';      // "Monitoring team": tab Overspeeding
const RT = ['AC', 'CDC', 'Aero', 'LR', 'MB-Lux', 'Teknaf'];
const MON = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Soudia Coach Service')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function file_(name) {   // if two files share a name, use the most recently updated one
  const it = DriveApp.getFolderById(FOLDER_ID).getFilesByName(name); let best = null;
  while (it.hasNext()) { const f = it.next(); if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f; }
  return best;
}
// One file per detail_YYYY-MM.json name (newest, non-empty), so a duplicate copy can never double-count.
function details_() {
  const it = DriveApp.getFolderById(FOLDER_ID).getFiles(), m = {};
  while (it.hasNext()) {
    const f = it.next(), n = f.getName(); if (!/^detail_\d{4}-\d{2}\.json$/.test(n) || f.getSize() === 0) continue;
    if (!m[n] || f.getLastUpdated() > m[n].getLastUpdated()) m[n] = f;
  }
  return Object.keys(m).sort().map(n => m[n]);
}
function read_(name) { const f = file_(name); return f ? f.getBlob().getDataAsString() : null; }
function save_(name, text) {
  const f = file_(name);
  if (f) f.setContent(text); else DriveApp.getFolderById(FOLDER_ID).createFile(name, text, MimeType.PLAIN_TEXT);
}
function day_(v, tz) {                        // Journey Date: real date or text like "02 Oct 2025"
  if (v instanceof Date) return Date.parse(Utilities.formatDate(v, tz, 'yyyy-MM-dd')) / 864e5;
  if (typeof v === 'number') return Math.floor(v) - 25569;
  const m = String(v).trim().match(/^(\d{1,2})[ \-\/](\w{3})\w*[ \-\/,]+(\d{4})$/);   // 02 Oct 2025, 2-Oct-2025, 2/Oct/2025
  return m && MON[m[2]] !== undefined ? Date.UTC(+m[3], MON[m[2]], +m[1]) / 864e5 : null;
}
const txt_ = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
const bus_ = v => String(v == null ? '' : v).trim().replace(/\s+/g, '').replace(/\.0+$/, '').toUpperCase();   // "11- 786" -> "11-786", "1171.0" -> "1171"

// Re-reads only the month files that changed, then rebuilds the lists and the unfiltered summary.
function refreshAll() { try { return refreshAll_w_(); } finally { bump_(); } }
function refreshAll_w_() { guard_('dc');
  const t0 = Date.now(), D = JSON.parse(read_('dict.json') || '{"c":[],"route":[],"coach":[],"bus":[],"z":[],"zn":[]}');
  const idx = JSON.parse(read_('detail_index.json') || '{}'), mp = {};
  ['c', 'route', 'coach', 'bus'].forEach(k => mp[k] = new Map(D[k].map((v, i) => [v, i])));
  const id_ = (k, v) => { let i = mp[k].get(v); if (i === undefined) { i = D[k].length; D[k].push(v); mp[k].set(v, i); } return i; };
  const it = DriveApp.getFolderById(FOLDER_ID).getFiles(), files = [];
  while (it.hasNext()) { const f = it.next(); if (/^Sale \d{4}-\d{2}$/.test(f.getName())) files.push(f); }
  files.sort((a, b) => a.getName() < b.getName() ? -1 : 1);
  let pending = 0;
  for (const f of files) {
    const key = f.getName().slice(5), upd = f.getLastUpdated().getTime();
    if (idx[key] === upd) continue;
    if (Date.now() - t0 > 230000) { pending++; continue; }          // stay under the 6-minute limit
    const ss = SpreadsheetApp.openById(f.getId()), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Sale Data');
    const n = sh.getLastRow() - 1, agg = {};
    for (let st = 0; st < n; st += 30000) {
      const v = sh.getRange(2 + st, 2, Math.min(30000, n - st), 11).getValues();   // columns B..L
      for (const r of v) {
        const d = day_(r[0], tz); if (d === null || !r[1]) continue;
        const c = id_('c', txt_(r[1])), rt = Math.max(0, RT.indexOf(txt_(r[10]))), ro = id_('route', txt_(r[7])), co = id_('coach', txt_(r[6])), bu = id_('bus', bus_(r[8]));
        const k = [d, c, rt, ro, co, bu].join('|'), o = agg[k] || (agg[k] = [d, c, rt, ro, co, bu, 0, 0]);
        o[6] += +r[3] || 0; o[7] += +r[5] || 0;
      }
    }
    save_('detail_' + key + '.json', JSON.stringify(Object.keys(agg).map(k => { const o = agg[k]; o[7] = Math.round(o[7]); return o; })));
    save_('dict.json', JSON.stringify(D));
    idx[key] = upd; save_('detail_index.json', JSON.stringify(idx));
  }
  const miss = zones_(D, false);
  save_('dict.json', JSON.stringify(D));
  save_('summary_all.json', query_({}));
  cdel_('meta'); cdel_('plan');
  return 'Refreshed ' + files.length + ' monthly files' + (pending ? '; ' + pending + ' still to read - run again' : '') + '.' +
    (miss.length ? ' ' + miss.length + ' counters have no zone in Masterdata (e.g. ' + miss.slice(0, 5).join(', ') + ').' : ' Every counter has a zone.');
}
// ZONES. Masterdata is the reference: each counter's zone comes from the Masterdata tab of the original "Sale" file
// (column A "Original Counter" -> column B "Zone"). It is read live whenever the dashboard loads (kept for 1 minute), so a change in
// Masterdata shows up on the next page load or Refresh data, with no need to rebuild the monthly files.
// Names are matched ignoring capitals, spaces and punctuation ("kamarpara Counter_TP" = "Kamarpara Counter_TP", "_TP" = "-TP", "02" = "2").
// If there is no exact match, names are matched again without the word "Counter" ("Nizampur College_TP" = "Nizampur College Counter_TP").
// A row with no zone never replaces a zone given on another row. Counters not found, or found with an empty Zone, stay "Unmapped".
const zkey_ = s => txt_(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\b0+(\d)/g, '$1').trim();
const zloose_ = k => k.replace(/\bcounter\b/g, ' ').replace(/\s+/g, ' ').trim();
function zoneMap_(fresh) {
  const cache = CacheService.getScriptCache(), hit = fresh ? null : cache.get('zonemap');
  if (hit) return JSON.parse(hit);
  const md = SpreadsheetApp.openById(MASTER_ID).getSheetByName('Masterdata').getDataRange().getValues(), zm = {}, zl = {};
  for (let i = 1; i < md.length; i++) { const n = txt_(md[i][0]), z = txt_(md[i][1]); if (n && z) { zm[zkey_(n)] = z; zl[zloose_(zkey_(n))] = z; } }
  const o = { zm: zm, zl: zl }; try { cache.put('zonemap', JSON.stringify(o), 60); } catch (e) {}
  return o;
}
// Sets D.z (zone index per counter) and D.zn (zone names); returns the counters that have no zone
function zones_(D, useCache) {
  const M = zoneMap_(!useCache), miss = [];
  D.zn = []; D.z = D.c.map(name => {
    const k = zkey_(name), zz = M.zm[k] || M.zl[zloose_(k)] || 'Unmapped'; if (zz === 'Unmapped') miss.push(name);
    let j = D.zn.indexOf(zz); if (j < 0) { j = D.zn.length; D.zn.push(zz); } return j;
  });
  return miss;
}
function dict_() { const D = JSON.parse(read_('dict.json')); try { zones_(D, true); } catch (e) {} return D; }   // if Masterdata cannot be read, the saved zones are used

// Short server-side cache (chunked: one cache entry holds at most 100 KB). Makes repeat loads fast; refreshAll() clears it.
function cget_(key) {
  try {
    const c = CacheService.getScriptCache(), n = +c.get(key + '#n'); if (!n) return null;
    const ks = []; for (let i = 0; i < n; i++) ks.push(key + '#' + i);
    const m = c.getAll(ks), out = []; for (let i = 0; i < n; i++) { if (m[ks[i]] == null) return null; out.push(m[ks[i]]); }
    return out.join('');
  } catch (e) { return null; }
}
function cput_(key, text, ttl) {
  try {
    const c = CacheService.getScriptCache(), SZ = 30000, o = {}; let n = 0;
    for (let i = 0; i < text.length; i += SZ) o[key + '#' + (n++)] = text.slice(i, i + SZ);
    if (n > 90) return; o[key + '#n'] = String(n); c.putAll(o, ttl);
  } catch (e) {}
}
function cdel_(key) {
  try { const c = CacheService.getScriptCache(), n = +c.get(key + '#n') || 0, ks = [key + '#n']; for (let i = 0; i < n; i++) ks.push(key + '#' + i); c.removeAll(ks); } catch (e) {}
}
function cached_(key, ttl, fn) { const h = cget_(key); if (h) return h; const v = fn(); cput_(key, v, ttl); return v; }
// Version-stamped result cache for the read functions (120 s). Every write function below bumps the version, so a save shows at once;
// edits made straight in the sheets show within 2 minutes. Cached copies are shared, so only functions with no per-user data are cached
// (the page permission check runs first, every time).
function dver_() { try { return CacheService.getScriptCache().get('dver') || '0'; } catch (e) { return '0'; } }
function bump_() { try { CacheService.getScriptCache().put('dver', String(Date.now()), 21600); } catch (e) {} }
function hash_(s) { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }
function memo_(name, args, ttl, fn) {
  const key = 'm_' + name + '_' + dver_() + '_' + (args.length ? hash_(JSON.stringify(args)) : '');
  const h = cget_(key); if (h !== null) { try { return JSON.parse(h); } catch (e) {} }
  const v = fn(); try { cput_(key, JSON.stringify(v === undefined ? null : v), ttl); } catch (e) {}
  return v;
}
function getMeta() { guard_(); return cached_("meta", 300, getMeta_); }
function getMeta_() {
  if (!file_('summary_all.json')) refreshAll();
  const D = dict_();
  return JSON.stringify({ c: D.c, z: D.z, zn: D.zn, rt: RT, route: D.route, coach: D.coach, bus: D.bus });
}
// sel = { counter:[], zone:[], rt:[], route:[], coach:[], bus:[] } (names); empty list = no filter on that field.
function query_(sel) {
  const D = dict_(), S = k => (sel[k] || []);
  const set = (k, arr) => S(k).length ? new Set(S(k).map(v => arr.indexOf(v)).filter(i => i >= 0)) : null;
  let cs = set('counter', D.c);
  if (S('zone').length) { const zi = new Set(S('zone').map(v => D.zn.indexOf(v))), zc = new Set(); D.z.forEach((z, i) => { if (zi.has(z)) zc.add(i); }); cs = cs ? new Set([...cs].filter(i => zc.has(i))) : zc; }
  const rs = set('rt', RT), os = set('route', D.route), ks = set('coach', D.coach), bs = set('bus', D.bus), agg = new Map();
  for (const f of details_()) {
    let rows; try { rows = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { Logger.log('Skipped unreadable ' + f.getName()); continue; }   // an empty or half-written month file is skipped, not fatal
    for (const x of rows) {
      if ((cs && !cs.has(x[1])) || (rs && !rs.has(x[2])) || (os && !os.has(x[3])) || (ks && !ks.has(x[4])) || (bs && !bs.has(x[5]))) continue;
      const k = (x[0] * 2048 + x[1]) * 8 + x[2], o = agg.get(k) || [x[0], x[1], x[2], 0, 0];
      o[3] += x[6]; o[4] += x[7]; agg.set(k, o);
    }
  }
  const r = [...agg.values()]; return JSON.stringify({ r: r, n: r.length });
}
function getFiltered(sel) { guard_(); return memo_('getFiltered', [sel], 120, () => getFiltered_c_(sel)); }
function getFiltered_c_(sel) {
  const any = Object.keys(sel || {}).some(k => (sel[k] || []).length);
  return (!any && read_('summary_all.json')) || query_(sel || {});
}
function setupDailyRefresh() {                 // run once: refreshes the data every morning
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshAll').timeBased().everyDays(1).atHour(6).create();
}

// ---------- Targets: rules in the "Sales Target Plan" sheet (tab Rules) + your own sales ----------
// Every day's target = base (average of recent normal days, per weekday) x stretch x month/weekday factor, or an event factor on event days.
// It is built in seats and in sales (fare = sales / seats). The day is then split by route type, corridor and counter using shares
// measured from your own sales, one set for normal days, one for event days and one for Ramadan.
function ymd_(v, tz) {
  if (v instanceof Date) return Date.parse(Utilities.formatDate(v, tz, 'yyyy-MM-dd')) / 864e5;
  if (typeof v === 'number') return Math.floor(v) - 25569;
  const m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 : day_(v, tz);
}
// GROWTH: year-on-year growth built into every target from the date given. To change it, add a line with the new date and
// percentage; earlier days keep the earlier percentage. The first date is also the day the counter-by-counter targets start.
const GROWTH_ = [['2026-11-01', 0.10]];
function getPlan() { guard_(); return cached_("plan", 300, getPlan_); }
function getPlan_() {
  const ss = SpreadsheetApp.openById(PLAN_ID), tz = ss.getSpreadsheetTimeZone(), R = ss.getSheetByName('Rules');
  const g = R.getRange(5, 2, 6, 1).getValues().map(r => r[0]);
  const stretch = +g[0] || 0, floorPct = +g[1] || .85, weeks = Math.max(2, +g[2] || 8), from = ymd_(g[3], tz), pEnd = ymd_(g[5], tz);
  const wdf = R.getRange(14, 2, 7, 1).getValues().map(r => +r[0] || 1), mf = R.getRange(24, 2, 12, 1).getValues().map(r => +r[0] || 1);
  const et = {};
  R.getRange(39, 1, 6, 5).getValues().forEach(r => { if (txt_(r[0])) et[txt_(r[0])] = { s: +r[1] || 1, v: +r[2] || 1, c: +r[3] || 0, l: txt_(r[4]) || 'Regular' }; });
  const evs = R.getRange(48, 1, 60, 4).getValues().filter(r => txt_(r[0]) && et[txt_(r[3])]).map(r => ({ n: txt_(r[0]), a: ymd_(r[1], tz), b: ymd_(r[2], tz), t: txt_(r[3]) }));
  const evOf = d => evs.find(e => d >= e.a && d <= e.b) || null;
  const inp = ss.getSheetByName('Inputs').getRange(5, 2, 5, 1).getValues().map(r => +r[0] || 0);
  const cor = ss.getSheetByName('Corridor').getRange(5, 1, 15, 3).getValues().filter(r => txt_(r[0]));
  const sraw = read_('summary_all.json'), dict = JSON.parse(read_('dict.json') || '{"c":[]}');
  if (!sraw) throw new Error('Sales summary is missing: run Refresh data first.');
  const rows = JSON.parse(sraw).r, dS = new Map(), dV = new Map();
  let last = 0;
  for (const x of rows) { dS.set(x[0], (dS.get(x[0]) || 0) + x[3]); dV.set(x[0], (dV.get(x[0]) || 0) + x[4]); if (x[0] > last) last = x[0]; }
  const wdi = d => (d + 3) % 7, mon = d => new Date(d * 864e5).getUTCMonth();
  const monday = d => d - wdi(d);
  const baseCache = {};
  const baseAt = end => {
    if (baseCache[end] !== undefined) return baseCache[end];
    let s = 0, v = 0, n = 0;
    for (let d = end - 7 * weeks; d < end; d++) {
      if (!dS.has(d) || d < from - 21 || evOf(d)) continue;
      const f = wdf[wdi(d)] * mf[mon(d)]; s += dS.get(d) / f; v += dV.get(d) / f; n++;
    }
    return baseCache[end] = n >= 10 ? { s: s / n, v: v / n, n: n, end: end } : null;
  };
  // shares by class: 0 normal (the base window), 1 event days, 2 Ramadan days
  const cls = [{ rt: [[], []], c: new Map(), tv: 0, ts: 0 }, { rt: [[], []], c: new Map(), tv: 0, ts: 0 }, { rt: [[], []], c: new Map(), tv: 0, ts: 0 }];
  const clsOf = d => { const e = evOf(d); if (!e) return (d >= last + 1 - 7 * weeks && d <= last) ? 0 : -1; return et[e.t].c === 0 ? -1 : et[e.t].c; };
  for (let k = 0; k < 3; k++) { for (let i = 0; i < 6; i++) { cls[k].rt[0][i] = 0; cls[k].rt[1][i] = 0; } }
  for (const x of rows) {
    const k = clsOf(x[0]); if (k < 0 || x[0] < from - 60) continue;
    const o = cls[k]; o.rt[0][x[2]] += x[4]; o.rt[1][x[2]] += x[3]; o.tv += x[4]; o.ts += x[3];
    const c = o.c.get(x[1]) || [0, 0]; c[0] += x[4]; c[1] += x[3]; o.c.set(x[1], c);
  }
  [1, 2].forEach(k => { if (!cls[k].tv) cls[k] = cls[0]; });
  const rtSh = (k, j) => cls[k].rt[j].map(v => cls[k].rt[j].reduce((a, b) => a + b, 0) ? v / cls[k].rt[j].reduce((a, b) => a + b, 0) : 0);
  const RS = [0, 1, 2].map(k => [rtSh(k, 0), rtSh(k, 1)]);
  // counters: top 100 by normal-day sales get their own share; the rest are "Other sales"
  const rank = [...cls[0].c.entries()].sort((a, b) => b[1][0] - a[1][0]).slice(0, 100);
  const cn = [], used = new Set(), r5 = v => Math.round(v * 1e5) / 1e5;
  rank.forEach(([c, v], i) => {
    used.add(c); const s = [], q = [];
    for (let k = 0; k < 3; k++) { const o = cls[k], own = o.c.get(c), z = cls[0].c.get(c) || [0, 0]; s.push(own ? r5(own[0] / (o.tv || 1)) : r5(z[0] / (cls[0].tv || 1))); q.push(own ? r5(own[1] / (o.ts || 1)) : r5(z[1] / (cls[0].ts || 1))); }
    cn.push({ n: txt_(dict.c[c]), t: i < 25 ? 'A' : 'B', s: s, q: q });
  });
  for (let k = 0; k < 3; k++) { const ss = cn.reduce((a, c) => a + c.s[k], 0), sq = cn.reduce((a, c) => a + c.q[k], 0); if (ss > .97) cn.forEach(c => c.s[k] = r5(c.s[k] * .97 / ss)); if (sq > .97) cn.forEach(c => c.q[k] = r5(c.q[k] * .97 / sq)); }
  const os = [], oq = [];
  for (let k = 0; k < 3; k++) { os.push(r5(Math.max(0, 1 - cn.reduce((a, c) => a + c.s[k], 0)))); oq.push(r5(Math.max(0, 1 - cn.reduce((a, c) => a + c.q[k], 0)))); }
  cn.push({ n: 'Other sales', t: 'C', s: os, q: oq });
  const P = { in: { fes: inp[0], reg: inp[1], str: inp[2], floor: inp[3], low: inp[4] }, fp: floorPct, crn: cor.map(r => txt_(r[0])), d: [], ty: [], ev: [], cl: [], tg: [], ts: [], fl: [], rt: [], rs: [], cr: [], cn: cn };
  const lastBase = baseAt(last + 1);
  const GR = GROWTH_.map(g => [Date.parse(g[0]) / 864e5, +g[1]]).sort((a, b) => a[0] - b[0]), NS = GR[0][0];
  const grow = d => { let v = GR[0][1]; for (const g of GR) if (g[0] <= d) v = g[1]; return v; };
  for (let d = from; d <= Math.min(pEnd, NS - 1); d++) {
    const b = baseAt(Math.min(monday(d), last + 1)); if (!b) continue;
    const e = evOf(d), t = e ? et[e.t] : null, k = t ? t.c : 0, kk = Math.max(0, Math.min(2, k));
    const fs = t ? t.s : wdf[wdi(d)] * mf[mon(d)], fv = t ? t.v : fs;
    const ts = Math.round(b.s * (1 + stretch) * fs), tg = Math.round(b.v * (1 + stretch) * fv);
    P.d.push(d); P.ty.push(t ? t.l : 'Regular'); P.ev.push(e ? e.n : ''); P.cl.push(kk); P.tg.push(tg); P.ts.push(ts); P.fl.push(Math.round(tg * floorPct));
    P.rt.push(RS[kk][0].map(s => Math.round(tg * s))); P.rs.push(RS[kk][1].map(s => Math.round(ts * s)));
    P.cr.push(cor.map(r => Math.round(tg * (+r[kk === 1 ? 2 : 1] || 0))));
  }
  // ----- From NS: targets built counter by counter and route type by route type (bottom-up) -----
  // Month base of a counter x route type = the same month last year, when that month is in the data for it;
  // otherwise its normal sales per weekday in the 8 weeks before the month starts (frozen on the 1st), with the month and event factors.
  // A counter x route type with no sale in the 14 days before the month starts gets no target. Growth is added per day.
  // The month is spread over its days by the route type's own weekday pattern (and the event factors on event days).
  let first = 1e9; for (const x of rows) if (x[0] < first) first = x[0];
  const NK = 8, ymOf = d => { const t = new Date(d * 864e5); return t.getUTCFullYear() * 12 + t.getUTCMonth(); }, ymStart = k => Date.UTC(Math.floor(k / 12), k % 12, 1) / 864e5;
  const ymKey = k => Math.floor(k / 12) + '-' + ('0' + (k % 12 + 1)).slice(-2);
  const mSum = new Map(), fs = new Map();          // (counter*8+rt) -> Map(month -> [sales, seats]); first sale day
  for (const x of rows) {
    const k = x[1] * NK + x[2]; let m = mSum.get(k); if (!m) mSum.set(k, m = new Map());
    const mk = ymOf(x[0]), o = m.get(mk) || [0, 0]; o[0] += x[4]; o[1] += x[3]; m.set(mk, o);
    if (!fs.has(k) || x[0] < fs.get(k)) fs.set(k, x[0]);
  }
  const winCache = {};
  const win = cut => {                              // normal-day sales by weekday in the 8 weeks before `cut`
    if (winCache[cut]) return winCache[cut];
    const a = Math.max(first, cut - 7 * weeks), W = { a: a, k: new Map(), rt: [], nd: [], mfw: 1 };
    for (let r = 0; r < 6; r++) W.rt.push([[0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0]]);
    let mfs = 0, mfn = 0;
    for (let d = a; d < Math.min(cut, last + 1); d++) if (!evOf(d)) { W.nd.push(d); mfs += mf[mon(d)]; mfn++; }
    W.mfw = mfn ? mfs / mfn : 1;
    for (const x of rows) {
      if (x[0] < a || x[0] >= cut) continue;
      const k = x[1] * NK + x[2]; let o = W.k.get(k); if (!o) W.k.set(k, o = { v: [0, 0, 0, 0, 0, 0, 0], s: [0, 0, 0, 0, 0, 0, 0], l: 0 });
      if (x[0] > o.l) o.l = x[0];
      if (evOf(x[0])) continue;
      const w = wdi(x[0]); o.v[w] += x[4]; o.s[w] += x[3]; W.rt[x[2]][0][w] += x[4]; W.rt[x[2]][1][w] += x[3];
    }
    return winCache[cut] = W;
  };
  const avg7 = a => a.reduce((x, y) => x + y, 0) / 7;
  const CMm = [], CMb = {}, CMy = {}, CMt = [];
  for (let mk = ymOf(NS); mk <= ymOf(pEnd); mk++) {
    const ms = ymStart(mk), me = ymStart(mk + 1) - 1, cut = Math.min(ms, last + 1), W = win(cut), ly = mk - 12, lyS = ymStart(ly), mi = CMm.length;
    CMm.push(ymKey(mk));
    const days = []; for (let d = Math.max(ms, NS); d <= Math.min(me, pEnd); d++) days.push(d);
    const whole = days.length === me - ms + 1;
    const tot = [[0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0]];
    mSum.forEach((m, k) => {
      const c = Math.floor(k / NK), rt = k % NK, o = W.k.get(k);
      if (!o || o.l < cut - 14) return;                                   // stopped selling: no target
      const L = m.get(ly), f0 = fs.get(k); let bv = 0, bs = 0, yoy = 0;
      if (whole && L && L[0] > 0 && lyS >= first && f0 <= lyS + 3) { bv = L[0]; bs = L[1]; yoy = 1; }
      else {
        const nd = [0, 0, 0, 0, 0, 0, 0]; for (const d of W.nd) if (d >= f0) nd[wdi(d)]++;
        const fill = arr => { const have = arr.filter((v, w) => nd[w]), m = have.length ? have.reduce((x, y) => x + y, 0) / have.length : 0; return arr.map((v, w) => nd[w] ? v : m); };   // a weekday with no normal day yet takes the average of the others
        const av = fill(o.v.map((v, w) => nd[w] ? v / nd[w] : 0)), as = fill(o.s.map((v, w) => nd[w] ? v / nd[w] : 0)), mv = avg7(av), msq = avg7(as), mfa = mf[mk % 12] / (W.mfw || 1);
        for (const d of days) { const e = evOf(d), t = e ? et[e.t] : null; if (t) { bv += mv * t.v; bs += msq * t.s; } else { bv += av[wdi(d)] * mfa; bs += as[wdi(d)] * mfa; } }
      }
      if (!(bv > 0)) return;
      const row = CMb[c] || (CMb[c] = []); while (row.length <= mi) row.push(0);
      if (!row[mi]) row[mi] = [0, 0, 0, 0, 0, 0]; row[mi][rt] = Math.round(bv);
      if (yoy) { const y = CMy[c] || (CMy[c] = []); while (y.length <= mi) y.push(0); y[mi] |= (1 << rt); }
      tot[0][rt] += Math.round(bv); tot[1][rt] += bs;
    });
    CMt.push(tot[0].slice());
    // spread each route type's month over the days
    const wt = [0, 1].map(j => W.rt.map(r => { const n = [0, 0, 0, 0, 0, 0, 0]; for (const d of W.nd) n[wdi(d)]++; const av = r[j].map((v, w) => n[w] ? v / n[w] : 0), m7 = avg7(av); return { av: av, m: m7 }; }));
    const wOf = (j, r, d) => { const e = evOf(d), t = e ? et[e.t] : null, o = wt[j][r]; if (!(o.m > 0)) return 1; return t ? o.m * (j ? t.s : t.v) : (o.av[wdi(d)] || o.m * 0.5); };
    const ws = [0, 1].map(j => [0, 1, 2, 3, 4, 5].map(r => days.reduce((a, d) => a + wOf(j, r, d), 0)));
    for (const d of days) {
      const e = evOf(d), t = e ? et[e.t] : null, k = t ? t.c : 0, kk = Math.max(0, Math.min(2, k)), g = 1 + grow(d);
      const rv = [0, 1, 2, 3, 4, 5].map(r => ws[0][r] ? Math.round(tot[0][r] * wOf(0, r, d) / ws[0][r] * g) : 0), rq = [0, 1, 2, 3, 4, 5].map(r => ws[1][r] ? Math.round(tot[1][r] * wOf(1, r, d) / ws[1][r] * g) : 0);
      const tg = rv.reduce((a, b) => a + b, 0), ts = rq.reduce((a, b) => a + b, 0);
      P.d.push(d); P.ty.push(t ? t.l : 'Regular'); P.ev.push(e ? e.n : ''); P.cl.push(kk); P.tg.push(tg); P.ts.push(ts); P.fl.push(Math.round(tg * floorPct));
      P.rt.push(rv); P.rs.push(rq); P.cr.push(cor.map(r => Math.round(tg * (+r[kk === 1 ? 2 : 1] || 0))));
    }
  }
  P.ns = NS; P.gr = GR; P.cm = { m: CMm, b: CMb, y: CMy, t: CMt };
  P.meta = { stretch: grow(Math.max(NS, last + 1)), weeks: weeks, last: last, baseSeats: lastBase ? Math.round(lastBase.s) : 0, baseSales: lastBase ? Math.round(lastBase.v) : 0 };
  return JSON.stringify(P);
}

// ---------- Bank, payments and monitoring sheets ----------
const num_ = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v == null ? '' : v).replace(/[,\s]/g, '')); return isNaN(n) ? 0 : n; };

// BP tab (raw bank data): one row per banking date and bank account.
// Returns { banks:[names], src:[deposit source names], r:[[day, bankIndex, opBalance, ...deposits by source, totalDep, partyPayment, loanClearance, totalPayment, closing]] }
// Columns are found by their header names (Banking Date, Account, Op Balance, <deposit sources...>, Total Dep, Party Payment, Loan Clearance, Total Payment),
// so inserting a column such as "Fund Transfer" or "LR Income" cannot shift the figures. Every column between Op Balance and Total Dep is a deposit source.
function bpMap_(sh) {
  const raw = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 19)).getValues()[0], h = raw.map(x => txt_(x).toLowerCase()), f = re => h.findIndex(x => re.test(x));
  const m = { date: f(/date/), acc: f(/^account/), op: f(/^op/), dep: f(/^total dep/), pp: f(/^party/), ln: f(/^loan/), tp: f(/^total pay/) };
  if (Object.keys(m).some(k => m[k] < 0) || m.dep <= m.op) throw new Error('The BP tab headers were not found. Needed: Banking Date, Account, Op Balance, Total Dep, Party Payment, Loan Clearance, Total Payment.');
  m.src = []; m.names = []; for (let i = m.op + 1; i < m.dep; i++) if (h[i]) { m.src.push(i); m.names.push(txt_(raw[i])); }
  m.width = Math.max(m.dep, m.pp, m.ln, m.tp) + 1;
  return m;
}
// ---------- Bank data now comes from the "Bank Ledger (new)" tab: one row per taka in or out (Date, Bank, Type, Category, Party / Source, Details, Amount) ----------
// Rows whose Party / Source is "Reconciliation" only keep the ledger tied to the bank-confirmed balance; the page ignores them. getBank and getPP rebuild exactly the same shapes the page has always received, so the Bank balance page, the print and the copy-as-photo layouts do not change.
const LEDGER = 'Bank Ledger (new)', LEDGER_TOP = 5;
const BANK_SRC_ = ['Dhk N AC', 'Non AC Bkash', 'Sylhet', 'Cox', 'Dhk MB', 'Cox Aero', 'Dhk AC', 'AC Bkash', 'Online', 'Fund Transfer', 'Reserve', 'LR Income'];
const DEP_CAT_ = ['Counter collection', 'bKash collection', 'Counter collection', 'Counter collection', 'Counter collection', 'Counter collection', 'Counter collection', 'bKash collection', 'Online sales', 'Fund transfer', 'Other deposit', 'LR income'];
const DEP_PARTY_ = ['Dhaka Non-AC counter', 'Non-AC bKash', 'Sylhet counter', 'Cox counter', 'Dhaka MB counter', 'Cox Aero counter', 'Dhaka AC counter', 'AC bKash', 'Online', 'Fund transfer', 'Reserve / other', 'LR income'];
function depIdx_(cat, party) {                  // which deposit column a ledger deposit belongs to; anything unknown is shown under Reserve (as the BP tab did)
  const c = cat.toLowerCase(), p = party.toLowerCase();
  for (let j = 0; j < 8; j++) if (DEP_CAT_[j].toLowerCase() === c && DEP_PARTY_[j].toLowerCase() === p) return j;
  if (c === 'online sales') return 8; if (c === 'fund transfer') return 9; if (c === 'lr income') return 11; return 10;
}
function depExact_(j, cat, party) { return cat.toLowerCase() === DEP_CAT_[j].toLowerCase() && (j === 8 || j === 9 || j === 11 || party.toLowerCase() === DEP_PARTY_[j].toLowerCase()); }
function payType_(cat) { const c = cat.toLowerCase(); return c === 'loan adjustment' ? 'Loan Adjustment' : c === 'fund transfer' ? 'Fund Transfer' : 'Payment'; }
function ledgerRows_(ss, tz) {                  // [{row, day, bank, type, cat, party, det, amt}]
  const sh = ss.getSheetByName(LEDGER); if (!sh) throw new Error('The tab "' + LEDGER + '" was not found.');
  const n = sh.getLastRow() - LEDGER_TOP, v = n > 0 ? sh.getRange(LEDGER_TOP + 1, 1, n, 7).getValues() : [], out = [];
  for (let i = 0; i < v.length; i++) {
    const r = v[i], d = day_(r[0], tz), a = num_(r[6]), b = txt_(r[1]), t = txt_(r[2]); if (d === null || !a || !b || (t !== 'Deposit' && t !== 'Payment')) continue;
    out.push({ row: i + LEDGER_TOP + 1, day: d, bank: b, type: t, cat: txt_(r[3]), party: txt_(r[4]), det: txt_(r[5]), amt: a });
  }
  return out;
}
function openingAnchor_(ss) {                   // balances before the first ledger date, from the Daily Summary (new) tab (row whose first cell starts with "Opening balance before")
  const sh = ss.getSheetByName('Daily Summary (new)'); if (!sh) return {};
  const v = sh.getRange(1, 1, sh.getLastRow(), 5).getValues();
  for (const r of v) if (/^opening balance before/i.test(txt_(r[0]))) { const o = {}; BANKS_.forEach((b, i) => { o[b] = num_(r[i + 1]); }); return o; }
  return {};
}
// Returns { banks, src:[deposit source names], ix:{dep,pp,ln,tp}, r:[[day, bank, opBalance, ...one amount per source, totalDep, partyPayment, loanClearance, totalPayment]] }
function getBank() { guard_('bk', 'en', 'rp', 'ov'); return memo_('getBank', [], 120, () => getBank_c_()); }
function getBank_c_() {
  const ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone(), L = ledgerRows_(ss, tz).filter(x => x.party.toLowerCase() !== 'reconciliation'), K = BANK_SRC_.length, banks = BANKS_.slice(), anchor = openingAnchor_(ss), conf = Date.parse(CONFIRMED.date) / 864e5;
  L.forEach(x => { if (banks.indexOf(x.bank) < 0) banks.push(x.bank); });
  const days = {}; L.forEach(x => { (days[x.day] = days[x.day] || []).push(x); });
  const ds = Object.keys(days).map(Number).sort((a, b) => a - b), out = [];
  ds.forEach((d, di) => {
    banks.forEach((b, bi) => {
      const dep = new Array(K).fill(0); let p = 0, q = 0;
      for (const x of days[d]) if (x.bank === b) { if (x.type === 'Deposit') dep[depIdx_(x.cat, x.party)] += x.amt; else if (x.cat.toLowerCase() === 'loan adjustment') q += x.amt; else p += x.amt; }
      const op = (d === conf && CONFIRMED.open[b] !== undefined) ? CONFIRMED.open[b] : (di === 0 ? (anchor[b] || 0) : 0);
      out.push([d, bi, op].concat(dep, [dep.reduce((s, v) => s + v, 0), p, q, p + q]));
    });
  });
  return JSON.stringify({ banks: banks, src: BANK_SRC_, ix: { dep: 3 + K, pp: 4 + K, ln: 5 + K, tp: 6 + K }, r: out });
}

// Payments: every ledger payment. Returns { r:[[day, to, details, amount, bank, type, rowNumber]] }
function getPP() { guard_('bk', 'en'); return memo_('getPP', [], 120, () => getPP_c_()); }
function getPP_c_() {
  const ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone();
  return JSON.stringify({ r: ledgerRows_(ss, tz).filter(x => x.type === 'Payment' && x.party.toLowerCase() !== 'reconciliation').map(x => [x.day, x.party, x.det, x.amt, x.bank, payType_(x.cat), x.row]) });
}

// Ledger rows for the Bank page: { r:[[day, bank, type, category, party, details, amount, rowNumber]] } (reconciliation rows left out, like getBank)
function getLedger() { guard_('bk', 'en'); return memo_('getLedger', [], 120, () => getLedger_c_()); }
function getLedger_c_() {
  const ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone();
  return JSON.stringify({ r: ledgerRows_(ss, tz).filter(x => x.party.toLowerCase() !== 'reconciliation').map(x => [x.day, x.bank, x.type, x.cat, x.party, x.det, x.amt, x.row]) });
}

// Everything the Bank page needs in one call: the bank sheet is opened and the ledger read once, instead of three separate calls
// (getBank, getPP, getLedger) that each opened the sheet and read the whole ledger again. Same shapes as the three functions above.
function getBankPack() { guard_('bk', 'en'); return memo_('getBankPack', [], 120, () => getBankPack_c_()); }
function getBankPack_c_() {
  const ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone(), all = ledgerRows_(ss, tz), L = all.filter(x => x.party.toLowerCase() !== 'reconciliation');
  const K = BANK_SRC_.length, banks = BANKS_.slice(), anchor = openingAnchor_(ss), conf = Date.parse(CONFIRMED.date) / 864e5;
  L.forEach(x => { if (banks.indexOf(x.bank) < 0) banks.push(x.bank); });
  const days = {}; L.forEach(x => { (days[x.day] = days[x.day] || []).push(x); });
  const ds = Object.keys(days).map(Number).sort((a, b) => a - b), r = [];
  ds.forEach((d, di) => {
    banks.forEach((b, bi) => {
      const dep = new Array(K).fill(0); let p = 0, q = 0;
      for (const x of days[d]) if (x.bank === b) { if (x.type === 'Deposit') dep[depIdx_(x.cat, x.party)] += x.amt; else if (x.cat.toLowerCase() === 'loan adjustment') q += x.amt; else p += x.amt; }
      const op = (d === conf && CONFIRMED.open[b] !== undefined) ? CONFIRMED.open[b] : (di === 0 ? (anchor[b] || 0) : 0);
      r.push([d, bi, op].concat(dep, [dep.reduce((s, v) => s + v, 0), p, q, p + q]));
    });
  });
  return JSON.stringify({
    b: { banks: banks, src: BANK_SRC_, ix: { dep: 3 + K, pp: 4 + K, ln: 5 + K, tp: 6 + K }, r: r },
    p: { r: L.filter(x => x.type === 'Payment').map(x => [x.day, x.party, x.det, x.amt, x.bank, payType_(x.cat), x.row]) },
    l: { r: L.map(x => [x.day, x.bank, x.type, x.cat, x.party, x.det, x.amt, x.row]) }
  });
}

// Overspeeding tab of the Monitoring team sheet. Removes repeated rows (the same event listed twice).
// Returns { r:[[day, trip, bus, startMinute, seconds, topSpeed, avgSpeed, lat, lon, place, driver, supervisor, alerts, monitor, note]], raw, buses, matched, unmatched, dateMismatch, roster }
// Columns are found by their header names (Date, Bus No, Start timing, End Timing, ... Duration, Top Speed, Average, Location, Driver, Supervisor, Alerts, Monitor, Note), so adding or moving a column does not break the page.
// The day of an event comes from its Start timing (the date and time the tracker recorded); the Date column is only used when the timing has no date.
// Driver and supervisor come from the event row, then from the Rotation tab (Date, Bus No, Driver Name, ..., Supervisor Name), then from the Driver/Supervisor list beside the data.
function busKey_(v) {                         // one number for a bus however it is written: "482", "0482", "11-0482", "11- 482"; Sheets-mangled dates are repaired
  const s = String(v == null ? '' : v).trim().replace(/\s+/g, '').toUpperCase();
  const m = s.match(/^[A-Z]{3}[A-Z]{3}\d{2}(\d{4})/); if (m) return String(+m[1]);
  const g = s.split(/-+/).filter(Boolean); if (!g.length) return '';
  let d = g[g.length - 1].replace(/\D/g, ''); if (g.length === 1 && d.length === 6 && d.slice(0, 2) === '11') d = d.slice(2);
  return d ? String(+d) : '';
}
function cols_(hdr, names) {                  // first header that contains any of the words
  const h = hdr.map(x => txt_(x).toLowerCase());
  const o = {}; for (const k in names) { o[k] = -1; for (const w of names[k]) { const i = h.findIndex(x => x.indexOf(w) >= 0); if (i >= 0) { o[k] = i; break; } } }
  return o;
}
// The Overspeeding tab has about 65,000 rows, and building the result from it takes minutes. So the finished result is kept as a saved copy (a file next to dict.json),
// rebuilt in the background every 10 minutes by refreshOverspeed (run overspeedCopySetup once to create that schedule). The page reads the saved copy; if it is missing or
// older than 60 minutes (the schedule stopped), it is built live as before.
const OS_FILE_ = 'overspeed_cache.json';
function osCopy_() { try { const f = file_(OS_FILE); return f && Date.now() - f.getLastUpdated().getTime() < 36e5 ? f.getBlob().getDataAsString() : null; } catch (e) { return null; } }
function getOverspeed() { guard_('mn', 'ov', 'ic'); return memo_('getOverspeed', [], 120, () => osCopy_() || getOverspeed_c_()); }
function refreshOverspeed() { const t0 = Date.now(), s = getOverspeed_c_(); save_(OS_FILE, s); console.log('Overspeed copy saved: ' + Math.round(s.length / 1024) + ' KB, built in ' + (Date.now() - t0) + ' ms'); return Math.round(s.length / 1024) + ' KB, ' + (Date.now() - t0) + ' ms'; }
function overspeedCopySetup() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'refreshOverspeed') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshOverspeed').timeBased().everyMinutes(10).create();
  return refreshOverspeed();
}
function getOverspeed_c_() {
  const ss = SpreadsheetApp.openById(MON_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Overspeeding');
  const hdr = sh.getRange(1, 1, 1, 14).getValues()[0], c = cols_(hdr, { place: ['place'], date: ['date'], coach: ['coach', 'trip'], bus: ['bus'], st: ['start'], en: ['end'], du: ['duration'], top: ['top'], avg: ['average', 'avg'], loc: ['location'], dr: ['driver'], su: ['supervisor'], al: ['alert'], mo: ['monitor'], no: ['note'] });
  const n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 14).getValues() : [], seen = {}, out = [], g = (r, k) => c[k] >= 0 ? r[c[k]] : '';
  const nm = x => { const t = txt_(x); return /^n\/?a$/i.test(t) ? '' : t; };
  let raw = 0, mism = 0;
  // roster: Rotation tab and the list beside the data (Date, Coach, Bus No, Driver, Supervisor)
  const roster = {};
  const addR = (bus, driver, sup, d) => { const k = busKey_(bus); if (!k || (!nm(driver) && !nm(sup))) return; (roster[k] = roster[k] || []).push([d || 0, nm(driver), nm(sup)]); };
  const rt = ss.getSheetByName('Rotation');
  if (rt && rt.getLastRow() > 1) { const rv = rt.getRange(1, 1, rt.getLastRow(), 7).getValues(), rc = cols_(rv[0], { date: ['date'], bus: ['bus'], dr: ['driver name', 'driver'], su: ['supervisor name', 'supervisor'] }); for (let i = 1; i < rv.length; i++) addR(rv[i][rc.bus], rv[i][rc.dr], rv[i][rc.su], day_(rv[i][rc.date], tz)); }
  if (sh.getLastColumn() >= 21 && n > 0) { const sv = sh.getRange(2, 17, n, 5).getValues(); for (const r of sv) addR(r[2], r[3], r[4], day_(r[0], tz)); }
  // account of each bus: Masterdata tab (Bus No, Account) in the same sheet
  const acc = {}, fleet = {}, mt = ss.getSheetByName('Masterdata');
  if (mt && mt.getLastRow() > 1) mt.getRange(2, 1, mt.getLastRow() - 1, 2).getValues().forEach(r => { const k = busKey_(r[0]), a = txt_(r[1]); if (k && a && !acc[k]) { acc[k] = a; fleet[a] = (fleet[a] || 0) + 1; } });
  const who = (bus, d) => { const a = roster[bus]; if (!a) return ['', '']; let b = null; for (const x of a) if (!x[0] || x[0] <= d) if (!b || x[0] >= b[0]) b = x; b = b || a[a.length - 1]; return [b[1], b[2]]; };
  for (const r of v) {
    const stv = g(r, 'st'); if (!stv && !g(r, 'bus')) continue;
    const st = stv instanceof Date ? Utilities.formatDate(stv, tz, 'yyyy-MM-dd hh:mm:ss a') : String(stv);
    const dm = st.match(/(\d{4})-(\d{2})-(\d{2})/), dc = day_(g(r, 'date'), tz);
    const d = dm ? Date.UTC(+dm[1], +dm[2] - 1, +dm[3]) / 864e5 : dc; if (d === null) continue;
    if (dm && dc !== null && dc !== d) mism++;
    const bus = busKey_(g(r, 'bus')); if (!bus) continue;
    raw++;
    const en = g(r, 'en') instanceof Date ? Utilities.formatDate(g(r, 'en'), tz, 'yyyy-MM-dd hh:mm:ss a') : String(g(r, 'en')), key = [d, bus, st, en, txt_(g(r, 'top'))].join('|');
    if (seen[key]) continue; seen[key] = 1;
    const m = st.match(/(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?/i);
    let mins = null; if (m) { const h = m[4] ? (+m[1]) % 12 + (/PM/i.test(m[4]) ? 12 : 0) : +m[1]; mins = h * 60 + (+m[2]); }
    const du = String(g(r, 'du')), dmn = du.match(/(\d+)\s*min/i), ds_ = du.match(/(\d+)\s*s\b/i), sec = (dmn ? +dmn[1] * 60 : 0) + (ds_ ? +ds_[1] : 0);
    const loc = txt_(g(r, 'loc')), lm = loc.match(/^(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*(.*)$/), w = who(bus, d);
    out.push([d, txt_(g(r, 'coach')), bus, mins, sec, parseFloat(String(g(r, 'top'))) || 0, parseFloat(String(g(r, 'avg'))) || 0,
      lm ? +lm[1] : null, lm ? +lm[2] : null, txt_(g(r, 'place')) || (lm ? lm[3].split(',').slice(0, 2).join(',').trim() : loc), nm(g(r, 'dr')) || w[0], nm(g(r, 'su')) || w[1], txt_(g(r, 'al')), txt_(g(r, 'mo')), txt_(g(r, 'no')), acc[bus] || '']);
  }
  // which of these buses also appear in the sales data (matched on the bus number, whatever its format)
  const D = JSON.parse(read_('dict.json') || '{"bus":[]}'), sales = {}; D.bus.forEach(b => { const k = busKey_(b); if (k) sales[k] = 1; });
  const buses = {}; out.forEach(x => buses[x[2]] = 1);
  const all = Object.keys(buses), un = all.filter(b => !sales[b]);
  return JSON.stringify({ r: out, raw: raw, buses: all.length, matched: all.length - un.length, unmatched: un.slice(0, 20), dateMismatch: mism, roster: Object.keys(roster).length, fleet: fleet });
}

// ===== Place names for overspeed locations =====
// The tracker website now gives only "lat, lon". fillPlaces() looks each pair up with Google Maps and writes the place name into a separate "Place" column
// (column N of the Overspeeding tab, header "Place"). Rows that already have a name after the coordinates just get that name copied over. Nothing in Location is changed.
// Run fillPlaces once from the editor (it asks for permission the first time), or run placesSetup once to repeat it every 15 minutes.
function placeShort_(a) {                     // keep the road / area / upazila; drop district, division, postcode and country
  return String(a || '').split(',').map(x => x.trim()).filter(x => x && !/(District|Division)$/i.test(x) && !/^\d{4}$/.test(x) && !/^Bangladesh$/i.test(x)).join(', ');
}
function fillPlaces() {
  const t0 = Date.now(), sh = SpreadsheetApp.openById(MON_ID).getSheetByName('Overspeeding'), last = sh.getLastRow();
  if (last < 2) return 'No rows.';
  const hdr = sh.getRange(1, 1, 1, 14).getValues()[0].map(x => txt_(x).toLowerCase());
  const lc = hdr.findIndex(x => x.indexOf('location') >= 0); if (lc < 0) throw new Error('No Location column found.');
  let pc = hdr.findIndex(x => x === 'place'); if (pc < 0) { pc = 13; if (hdr[13]) throw new Error('Column N is used by "' + hdr[13] + '". Rename its header to Place or clear it.'); sh.getRange(1, 14).setValue('Place'); }
  const n = last - 1, loc = sh.getRange(2, lc + 1, n, 1).getValues(), cur = sh.getRange(2, pc + 1, n, 1).getValues(), out = cur.map(r => [r[0]]);
  const cache = CacheService.getScriptCache(), geo = Maps.newGeocoder().setLanguage('en'), mem = {};
  let done = 0, copied = 0, failed = 0, left = 0;
  for (let i = 0; i < n; i++) {
    if (txt_(cur[i][0])) continue;
    const m = txt_(loc[i][0]).match(/^(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*(.*)$/); if (!m) continue;
    if (m[3]) { out[i][0] = placeShort_(m[3]).split(',').slice(0, 3).join(',').trim(); copied++; continue; }
    const lat = +m[1], lon = +m[2], key = 'pl' + lat.toFixed(3) + ',' + lon.toFixed(3);
    let name = mem[key] || cache.get(key);
    if (!name) {
      if (Date.now() - t0 > 240000 || done >= 450) { left++; continue; }
      try { const r = geo.reverseGeocode(lat, lon); name = placeShort_(r && r.results && r.results[0] ? r.results[0].formatted_address : ''); done++; Utilities.sleep(120); } catch (e) { failed++; continue; }
      if (name) cache.put(key, name, 21600);
    }
    mem[key] = name; if (name) out[i][0] = name; else failed++;
  }
  sh.getRange(2, pc + 1, n, 1).setValues(out); SpreadsheetApp.flush();
  return 'Looked up ' + done + ', copied ' + copied + ', failed ' + failed + (left ? ', ' + left + ' left for the next run.' : '.');
}
function placesSetup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fillPlaces').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('fillPlaces').timeBased().everyMinutes(15).create();
  return fillPlaces();
}

// ===== Access control ("Access" page, owner only) =====
// The two owner accounts below always have full access and are the only ones who can open the Access page.
// Everyone else is controlled from that page: a role, the pages they may open, the accounts they may see, and whether they may enter data.
// The list is kept in this script's properties (ACCESS_V1), so it is private to the owners. Every data function below checks it on the server.
const EDITORS = ['rh.saudiabus@gmail.com', 'network.rejaulsayem@gmail.com'];
const ALL_PAGES_ = ['ov', 'cl', 'sl', 'sr', 'bi', 'ro', 'bk', 'psc', 'dy', 'rt', 'pr', 'mn', 'ic', 'gp', 'rp', 'en', 'dc'];
const SCOPABLE_ = ['ro', 'dy'];                                   // pages that can be limited to some accounts; every other page is company-wide
const ROLES_ = {
  'Manager': { pages: ['ov', 'cl', 'sl', 'sr', 'bi', 'ro', 'bk', 'psc', 'dy', 'rt', 'pr', 'mn', 'ic', 'gp', 'rp', 'dc'], edit: false },
  'Accounts': { pages: ['ov', 'bk', 'psc', 'dy', 'rt', 'rp', 'en', 'dc'], edit: true },
  'Sales': { pages: ['sl', 'sr', 'bi', 'ro'], edit: false },
  'Operations': { pages: ['mn', 'ic', 'gp', 'cl'], edit: false },
  'Account head': { pages: ['ro', 'dy'], edit: false },
  'Viewer': { pages: ['ov'], edit: false }
};
let ACL_ = null, PERM_ = null;
function me_() { if (typeof FWD_EMAIL_ === 'string' && FWD_EMAIL_) return FWD_EMAIL_; try { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (x) { return ''; } }
function acl_() {
  if (ACL_) return ACL_;
  let c = {}; try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty('ACCESS_V1') || '{}'); } catch (x) { c = {}; }
  c.users = Array.isArray(c.users) ? c.users : []; c.restrict = !!c.restrict; c.log = Array.isArray(c.log) ? c.log : [];
  return (ACL_ = c);
}
// Who is this? { email, owner, allowed, pages, accounts (null = all), edit, restricted, name, role }
function perm_() {
  if (PERM_) return PERM_;
  const e = me_(), c = acl_(), none = { email: e, owner: false, allowed: false, pages: [], accounts: null, edit: false, restricted: c.restrict };
  if (e && EDITORS.indexOf(e) >= 0) return (PERM_ = { email: e, owner: true, allowed: true, pages: ALL_PAGES_.concat(['ax']), accounts: null, edit: true, restricted: c.restrict });
  const u = e ? c.users.filter(x => x.email === e)[0] : null;
  if (u) {
    if (u.off) return (PERM_ = none);
    const acc = Array.isArray(u.accounts) && u.accounts.length ? u.accounts.filter(a => RT.indexOf(a) >= 0) : null;
    let pg = (u.pages || []).filter(p => ALL_PAGES_.indexOf(p) >= 0); if (acc) pg = pg.filter(p => SCOPABLE_.indexOf(p) >= 0);
    return (PERM_ = { email: e, owner: false, allowed: true, pages: pg, accounts: acc, edit: !!u.edit && !acc, restricted: c.restrict, name: u.name || '', role: u.role || '' });
  }
  if (!c.restrict) return (PERM_ = { email: e, owner: false, allowed: true, pages: ALL_PAGES_.filter(p => p !== 'en'), accounts: null, edit: false, restricted: false });   // before the owner switches "Only listed emails" on, everyone who could open the link still sees what they saw before (read-only)
  return (PERM_ = none);
}
// Stops the call unless the person may use at least one of the pages named. With no pages named it only checks that the person is allowed in at all.
function guard_() {
  const p = perm_(); if (!p.allowed) throw new Error('This Google account (' + (p.email || 'unknown') + ') does not have access to the dashboard. Ask the owner to add it on the Access page.');
  if (p.owner || !arguments.length) return p;
  for (let i = 0; i < arguments.length; i++) if (p.pages.indexOf(arguments[i]) >= 0) return p;
  throw new Error('You do not have access to this page.');
}
function whoami() { const p = perm_(); return JSON.stringify({ email: p.email, can: p.edit, owner: p.owner, allowed: p.allowed, pages: p.pages, accounts: p.accounts, restricted: p.restricted, name: p.name || '', role: p.role || '' }); }
function isEditor_() { return !!perm_().edit; }
function calOk_() { const p = perm_(); return p.allowed && (p.owner || p.pages.indexOf('cl') >= 0); }
function ownerOnly_() { const p = perm_(); if (!p.owner) throw new Error('Only the owners can change access.'); return p; }
// Empties a value but keeps its shape, so a page written for the full data still draws (numbers 0, lists empty, text kept).
function blank_(v) {
  if (typeof v === 'number') return 0;
  if (Array.isArray(v)) return [];
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).forEach(k => { o[k] = blank_(v[k]); }); return o; }
  return v;
}
// Cash flow for a person limited to some accounts: only those accounts' rows and totals; company cash, notes, advances and expenses are emptied.
function scopeCash_(o, accs) {
  const ok = a => accs.indexOf(a) >= 0, by = (o.byAccount || []).filter(x => ok(x.account)), sum = f => by.reduce((t, x) => t + (+x[f] || 0), 0);
  const r = blank_(o);
  r.date = o.date; r.cashDate = o.cashDate; r.byAccount = by;
  r.approvedExpenses = (o.approvedExpenses || []).filter(x => ok(x.account));
  r.advRetDetails = (o.advRetDetails || []).filter(x => ok(x.account));
  r.paidToBank = (o.paidToBank || []).filter(x => ok(x.account));
  r.approvedTotal = r.approvedExpenses.reduce((t, x) => t + (+x.amount || 0), 0);
  r.collection = Object.assign(r.collection || {}, { sales: sum('sales'), trips: sum('trips'), income: sum('income'), approved: sum('approved'), advances: sum('advances'), bank: sum('bank'), online: sum('online'), payable: sum('payable'), received: sum('received'), toBank: sum('toBank'), diff: sum('diff') });
  if (o.today) r.today = blank_(o.today);
  r.scoped = accs;
  return r;
}
// Access page: read and save the list. Owners only.
function getAccess() {
  const p = ownerOnly_(), c = acl_();
  return JSON.stringify({ restrict: c.restrict, users: c.users, owners: EDITORS, roles: ROLES_, pages: ALL_PAGES_, scopable: SCOPABLE_, accounts: RT, log: c.log.slice(-15).reverse(), me: p.email });
}
function saveAccess(json) { try { return saveAccess_w_(json); } finally { bump_(); } }
function saveAccess_w_(json) {
  const p = ownerOnly_(), q = JSON.parse(json), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!q || !Array.isArray(q.users) || q.users.length > 150) throw new Error('Check the list of people.');
    const seen = {}, users = [];
    q.users.forEach(u => {
      const email = String(u.email || '').trim().toLowerCase();
      if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) throw new Error('"' + email + '" is not a valid email address.');
      if (EDITORS.indexOf(email) >= 0) throw new Error(email + ' is an owner and always has full access; remove it from the list.');
      if (seen[email]) throw new Error(email + ' is listed twice.'); seen[email] = 1;
      let accs = (Array.isArray(u.accounts) ? u.accounts : []).filter(a => RT.indexOf(a) >= 0);
      if (accs.length >= RT.length) accs = [];
      let pg = (Array.isArray(u.pages) ? u.pages : []).filter(x => ALL_PAGES_.indexOf(x) >= 0); if (accs.length) pg = pg.filter(x => SCOPABLE_.indexOf(x) >= 0);
      const role = (u.role === 'Custom' || ROLES_[u.role]) ? u.role : 'Custom';
      users.push({ email: email, name: String(u.name || '').trim().slice(0, 60), role: role, pages: pg, accounts: accs, edit: !!u.edit && !accs.length, off: !!u.off });
    });
    const c = acl_(), before = {}; c.users.forEach(u => { before[u.email] = u; });
    const log = c.log.slice(-60), now = Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyy-MM-dd HH:mm');
    users.forEach(u => { const b = before[u.email]; if (!b) log.push({ t: now, by: p.email, m: 'Added ' + u.email + ' (' + u.role + ')' }); else if (JSON.stringify([b.role, b.pages, b.accounts, !!b.edit, !!b.off]) !== JSON.stringify([u.role, u.pages, u.accounts, u.edit, u.off])) log.push({ t: now, by: p.email, m: 'Changed ' + u.email + (u.off && !b.off ? ' (switched off)' : '') }); });
    c.users.forEach(b => { if (!seen[b.email]) log.push({ t: now, by: p.email, m: 'Removed ' + b.email }); });
    if (!!q.restrict !== c.restrict) log.push({ t: now, by: p.email, m: q.restrict ? 'Switched ON: only listed emails can open the dashboard' : 'Switched OFF: anyone who can open the link can view it' });
    const out = { restrict: !!q.restrict, users: users, log: log.slice(-60) };
    PropertiesService.getScriptProperties().setProperty('ACCESS_V1', JSON.stringify(out));
    ACL_ = null; PERM_ = null;
    return 'ok';
  } finally { lock.releaseLock(); }
}
// ---------- Payment Schedule page ("SBS Payment Schedule" Google Sheet) ----------
// Tabs: Cash Payment and Cheque Payment (one row per scheduled payment to a party: Date, Party Name, Amount, Paid, Due, Status, Payment Plan, Payment Date)
//       Bill Payment (daily bills: Date, Amount, Details, Bus No, Party, Approved By, Submitted By, GRN, PO, Entered At, Status)
// The daily cash sheet (Cash Payment tab in the Bank sheet) is a different thing and is not used here.
const PSCHED_ID = '1xDXR1tym-dYnEG0vJM6J0K9UaOzi7ngppLFTJSR10Ss';
const PS_TABS = { Cash: 'Cash Payment', Cheque: 'Cheque Payment' };
const BILL_APPROVERS = ['CEO Sir', 'Shoybal Sir', 'Maruf Sir', 'Tanjid Sir'];
// A row counts as paid in full when the Paid cell holds a mark instead of an amount (a dot, a tick, any text).
// Returns { can, r:[[day, party, amount, paid, due, status, plan, paymentDay (0 if none), sheetRow, category]] }
function getPS() { guard_('psc', 'en', 'rp', 'ov', 'cl'); return getPS_(); }
function getPS_() { return memo_('getPS_', [], 120, () => getPS__c_()); }
function getPS__c_() {
  const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), out = [];
  Object.keys(PS_TABS).forEach(cat => {
    const sh = ss.getSheetByName(PS_TABS[cat]); if (!sh) return;
    const n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 8).getValues() : [];
    for (let i = 0; i < v.length; i++) {
      const r = v[i], d = day_(r[0], tz); if (d === null || !txt_(r[1])) continue;
      const amt = num_(r[2]), pr = r[3], mark = pr !== '' && pr != null && isNaN(parseFloat(String(pr).replace(/[,\s]/g, '')));
      const paid = mark ? amt : num_(pr), due = mark ? 0 : (r[4] === '' || r[4] == null ? Math.max(0, amt - paid) : num_(r[4])), pd = r[7] === '' ? null : day_(r[7], tz);
      out.push([d, txt_(r[1]), amt, paid, due, txt_(r[5]) || (mark ? 'Paid' : ''), txt_(r[6]), pd === null ? 0 : pd, i + 2, cat]);
    }
  });
  out.sort((a, b) => a[0] - b[0]);
  return JSON.stringify({ can: isEditor_(), r: out });
}
// Mark a scheduled payment paid (mode full), part paid (mode part, amount = paid now) or clear it (mode undo). Editors only.
// q.cat is the tab (Cash or Cheque). The row is checked against the party and amount the page saw, so a changed sheet never gets the wrong row updated.
function setPaid(json) { try { return setPaid_w_(json); } finally { bump_(); } }
function setPaid_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can change payments.');
  const q = JSON.parse(json), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName(PS_TABS[q.cat || 'Cash']);
    if (!sh) throw new Error('Choose Cash or Cheque.');
    if (!(q.row >= 2 && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
    const r = sh.getRange(q.row, 1, 1, 8).getValues()[0], amt = num_(r[2]);
    if (txt_(r[1]) !== q.party || amt !== q.amt) throw new Error('The sheet has changed: reload the page and try again.');
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    if (q.mode === 'undo') { sh.getRange(q.row, 4).clearContent(); sh.getRange(q.row, 6).clearContent(); sh.getRange(q.row, 8).clearContent(); }
    else {
      const was = isNaN(parseFloat(String(r[3]).replace(/[,\s]/g, ''))) && r[3] !== '' ? amt : num_(r[3]);
      let paid = q.mode === 'part' ? was + (+q.amount || 0) : amt;
      if (!(paid > 0)) throw new Error('Enter an amount above zero.');
      if (paid >= amt) paid = amt;
      sh.getRange(q.row, 4).setValue(paid);
      sh.getRange(q.row, 6).setValue(paid >= amt ? 'Paid' : 'Part paid');
      sh.getRange(q.row, 8).setValue(today).setNumberFormat('dd MMM yyyy');
    }
    if (sh.getRange(q.row, 5).getFormula() === '') sh.getRange(q.row, 5).setFormula('=C' + q.row + '-D' + q.row);
    SpreadsheetApp.flush();
    const x = sh.getRange(q.row, 1, 1, 8).getValues()[0], p = num_(x[3]), pd = x[7] === '' ? null : day_(x[7], tz);
    return JSON.stringify([day_(x[0], tz), txt_(x[1]), num_(x[2]), p, Math.max(0, num_(x[2]) - p), txt_(x[5]), txt_(x[6]), pd === null ? 0 : pd, q.row, q.cat || 'Cash']);
  } finally { lock.releaseLock(); }
}
// Bill Payment tab. Columns: Date, Amount, Details, Bus No, Party, Approved By, Submitted By, GRN, PO, Entered At, Status (Void), Paid (৳), Pay Status, Paid On.
// Returns { can, ap:[approvers], r:[[day, amount, details, bus, party, approvedBy, submittedBy, grn, po, sheetRow, paid, payStatus, paidOnDay]] } (voided rows left out). A blank Paid cell means unpaid.
function getBills() { guard_('psc', 'en', 'rp'); return memo_('getBills', [], 120, () => getBills_c_()); }
function getBills_c_() {
  const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Bill Payment'), out = [];
  const n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 14).getValues() : [];
  for (let i = 0; i < v.length; i++) {
    const r = v[i], d = day_(r[0], tz); if (d === null || txt_(r[10]) === 'Void') continue;
    const amt = num_(r[1]), paid = Math.min(amt, num_(r[11])), po = r[13] === '' ? null : day_(r[13], tz);
    out.push([d, amt, txt_(r[2]), txt_(r[3]), txt_(r[4]), txt_(r[5]), txt_(r[6]), txt_(r[7]), txt_(r[8]), i + 2, paid, paid <= 0 ? '' : (paid >= amt ? 'Paid' : (txt_(r[12]) || 'Part paid')), po === null ? 0 : po]);
  }
  out.sort((a, b) => b[0] - a[0] || b[9] - a[9]);
  return JSON.stringify({ can: isEditor_(), ap: BILL_APPROVERS, r: out });
}
// Writes the Paid, Pay Status and Paid On cells of one bill row. mode: '' unpaid, 'paid', 'half' (half of the bill), 'part' (the amount in partAmt), 'keep' (leave as is, but follow a changed bill amount).
function billPay_(sh, row, amt, mode, partAmt, tz) {
  if (sh.getRange(1, 12).getValue() === '') sh.getRange(1, 12, 1, 3).setValues([['Paid (৳)', 'Pay Status', 'Paid On']]);
  const cur = num_(sh.getRange(row, 12).getValue()), st = txt_(sh.getRange(row, 13).getValue());
  let paid, label;
  if (mode === 'keep') { if (cur <= 0) return; paid = st === 'Paid' ? amt : (st === 'Half paid' ? Math.round(amt / 2) : Math.min(cur, amt)); label = st; if (paid >= amt) { paid = amt; label = 'Paid'; } }
  else if (mode === 'paid') { paid = amt; label = 'Paid'; }
  else if (mode === 'half') { paid = Math.round(amt / 2); label = 'Half paid'; }
  else if (mode === 'part') { paid = +partAmt; if (!(paid > 0)) throw new Error('Enter the part amount paid.'); if (paid >= amt) { paid = amt; label = 'Paid'; } else label = 'Part paid'; }
  else { sh.getRange(row, 12, 1, 3).clearContent(); return; }
  sh.getRange(row, 12, 1, 2).setValues([[paid, label]]); sh.getRange(row, 12).setNumberFormat('#,##0');
  if (mode !== 'keep' || !sh.getRange(row, 14).getValue()) sh.getRange(row, 14).setValue(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')).setNumberFormat('d MMM yyyy');
}
const billOut_ = (sh, r, tz) => { const v = sh.getRange(r, 1, 1, 14).getValues()[0], amt = num_(v[1]), paid = Math.min(amt, num_(v[11])), po = v[13] === '' ? null : day_(v[13], tz);
  return [day_(v[0], tz), amt, txt_(v[2]), txt_(v[3]), txt_(v[4]), txt_(v[5]), txt_(v[6]), txt_(v[7]), txt_(v[8]), r, paid, paid <= 0 ? '' : (paid >= amt ? 'Paid' : (txt_(v[12]) || 'Part paid')), po === null ? 0 : po]; };
// Form on the Payment Schedule page: q = { day, amount, details, bus, party, approved (optional), grn, po, pay: ''|'paid'|'half'|'part', part }. Editors only.
function addBill(json) { try { return addBill_w_(json); } finally { bump_(); } }
function addBill_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can add bills.');
  const q = JSON.parse(json), day = dayOk_(q.day), amt = pos_(q.amount, 'Amount'), party = txt_(q.party), det = txt_(q.details), ap = txt_(q.approved);
  if (!party) throw new Error('Enter the party.');
  if (!det) throw new Error('Enter the details.');
  const who = String(Session.getActiveUser().getEmail() || '').toLowerCase(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Bill Payment'), r = Math.max(2, sh.getLastRow() + 1);
    putDate_(sh.getRange(r, 1), day);
    sh.getRange(r, 2, 1, 9).setValues([[amt, det, txt_(q.bus), party, ap, who, txt_(q.grn), txt_(q.po), Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm')]]);
    sh.getRange(r, 2).setNumberFormat('#,##0');
    billPay_(sh, r, amt, q.pay || '', q.part, tz);
    SpreadsheetApp.flush();
    return JSON.stringify(billOut_(sh, r, tz));
  } finally { lock.releaseLock(); }
}
// Edit button on the Bills table: q = { row, oldParty, oldAmt (what the page saw, so a changed sheet never gets the wrong row), day, amount, details, bus, party, approved, grn, po, pay: 'keep'|''|'paid'|'half'|'part', part }.
function editBill(json) { try { return editBill_w_(json); } finally { bump_(); } }
function editBill_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can change bills.');
  const q = JSON.parse(json), day = dayOk_(q.day), amt = pos_(q.amount, 'Amount'), party = txt_(q.party), det = txt_(q.details);
  if (!party) throw new Error('Enter the party.');
  if (!det) throw new Error('Enter the details.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Bill Payment');
    if (!(q.row >= 2 && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
    const o = sh.getRange(q.row, 1, 1, 11).getValues()[0];
    if (txt_(o[4]) !== q.oldParty || num_(o[1]) !== q.oldAmt || txt_(o[10]) === 'Void') throw new Error('The sheet has changed: reload the page and try again.');
    putDate_(sh.getRange(q.row, 1), day);
    sh.getRange(q.row, 2, 1, 5).setValues([[amt, det, txt_(q.bus), party, txt_(q.approved)]]);
    sh.getRange(q.row, 8, 1, 2).setValues([[txt_(q.grn), txt_(q.po)]]);
    sh.getRange(q.row, 2).setNumberFormat('#,##0');
    billPay_(sh, q.row, amt, q.pay == null ? 'keep' : q.pay, q.part, tz);
    SpreadsheetApp.flush();
    return JSON.stringify(billOut_(sh, q.row, tz));
  } finally { lock.releaseLock(); }
}
// Takes a wrongly entered bill off the page (the row stays in the sheet marked Void). The row is checked against the party and amount the page saw.
function voidBill(json) { try { return voidBill_w_(json); } finally { bump_(); } }
function voidBill_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can change bills.');
  const q = JSON.parse(json), sh = SpreadsheetApp.openById(PSCHED_ID).getSheetByName('Bill Payment');
  if (!(q.row >= 2 && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
  const r = sh.getRange(q.row, 1, 1, 5).getValues()[0];
  if (txt_(r[4]) !== q.party || num_(r[1]) !== q.amt) throw new Error('The sheet has changed: reload the page and try again.');
  sh.getRange(q.row, 11).setValue('Void'); return 'ok';
}
// ===== Party Ledger: one record per payment on a bill (Payments tab of the SBS Payment Schedule sheet) =====
// Payments columns: Date, Bill Row, Party, Amount, Method (Cash/Cheque/Bank), Collected By, Cheque No, Bank, Note, Entered By, Entered At, Status (Void).
// Party Names tab: Alias, Name (to merge different spellings of one party). The Paid (৳) cell of the bill stays the running total.
const PAYS_HEAD = ['Date', 'Bill Row', 'Party', 'Amount', 'Method', 'Collected By', 'Cheque No', 'Bank', 'Note', 'Entered By', 'Entered At', 'Status'];
function paysSheet_(ss) {
  let sh = ss.getSheetByName('Payments');
  if (!sh) { sh = ss.insertSheet('Payments'); sh.getRange(1, 1, 1, 12).setValues([PAYS_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
const payOut_ = (v, r, tz) => [day_(v[0], tz), num_(v[1]), txt_(v[2]), num_(v[3]), txt_(v[4]), txt_(v[5]), txt_(v[6]), txt_(v[7]), txt_(v[8]), r];
function getPayments() { guard_('psc', 'en', 'rp'); return memo_('getPayments', [], 120, () => getPayments_c_()); }
function getPayments_c_() {
  const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), sh = paysSheet_(ss), out = [], n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 12).getValues() : [];
  for (let i = 0; i < v.length; i++) { const d = day_(v[i][0], tz); if (d === null || txt_(v[i][11]) === 'Void') continue; out.push(payOut_(v[i], i + 2, tz)); }
  const an = ss.getSheetByName('Party Names'), alias = an && an.getLastRow() > 1 ? an.getRange(2, 1, an.getLastRow() - 1, 2).getValues().map(r => [txt_(r[0]), txt_(r[1])]).filter(r => r[0] && r[1]) : [];
  return JSON.stringify({ can: isEditor_(), pays: out, alias: alias });
}
// q = { row (bill row), oldParty, oldAmt, day, amount, method, by, chq, bank, note }. Editors only.
function addPayment(json) { try { return addPayment_w_(json); } finally { bump_(); } }
function addPayment_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can record payments.');
  const q = JSON.parse(json), day = dayOk_(q.day), amt = pos_(q.amount, 'Amount'), m = txt_(q.method), by = txt_(q.by), chq = txt_(q.chq), bank = txt_(q.bank);
  if (['Cash', 'Cheque', 'Bank'].indexOf(m) < 0) throw new Error('Choose Cash, Cheque or Bank.');
  if (m === 'Cash' && !by) throw new Error('Enter who collected the cash.');
  if (m === 'Cheque' && !chq) throw new Error('Enter the cheque no.');
  if (m === 'Bank' && !bank) throw new Error('Enter the bank.');
  const who = String(Session.getActiveUser().getEmail() || '').toLowerCase(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), bs = ss.getSheetByName('Bill Payment');
    if (!(q.row >= 2 && q.row <= bs.getLastRow())) throw new Error('Bill not found: reload the page.');
    const b = bs.getRange(q.row, 1, 1, 12).getValues()[0], bamt = num_(b[1]), cur = Math.min(bamt, num_(b[11]));
    if (txt_(b[4]) !== q.oldParty || bamt !== q.oldAmt || txt_(b[10]) === 'Void') throw new Error('The sheet has changed: reload the page and try again.');
    if (amt > bamt - cur + 0.5) throw new Error('This bill has only ' + (bamt - cur) + ' left to pay.');
    const ps = paysSheet_(ss), r = Math.max(2, ps.getLastRow() + 1);
    putDate_(ps.getRange(r, 1), day);
    ps.getRange(r, 2, 1, 10).setValues([[q.row, txt_(b[4]), amt, m, by, chq, bank, txt_(q.note), who, Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm')]]);
    ps.getRange(r, 4).setNumberFormat('#,##0');
    billPay_(bs, q.row, bamt, 'part', Math.min(bamt, cur + amt), tz);
    SpreadsheetApp.flush();
    return JSON.stringify({ pay: payOut_(ps.getRange(r, 1, 1, 12).getValues()[0], r, tz), bill: billOut_(bs, q.row, tz) });
  } finally { lock.releaseLock(); }
}
// Undo a wrongly recorded payment (the row stays in the sheet marked Void) and take it off the bill.
function voidPayment(json) { try { return voidPayment_w_(json); } finally { bump_(); } }
function voidPayment_w_(json) { guard_('psc', 'en');
  if (!isEditor_()) throw new Error('Only the owner can change payments.');
  const q = JSON.parse(json), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(PSCHED_ID), tz = ss.getSpreadsheetTimeZone(), ps = paysSheet_(ss), bs = ss.getSheetByName('Bill Payment');
    if (!(q.row >= 2 && q.row <= ps.getLastRow())) throw new Error('Payment not found: reload the page.');
    const p = ps.getRange(q.row, 1, 1, 12).getValues()[0];
    if (txt_(p[11]) === 'Void' || num_(p[3]) !== q.amt || num_(p[1]) !== q.bill) throw new Error('The sheet has changed: reload the page and try again.');
    const br = num_(p[1]), b = bs.getRange(br, 1, 1, 12).getValues()[0], bamt = num_(b[1]), left = Math.max(0, Math.min(bamt, num_(b[11])) - num_(p[3]));
    ps.getRange(q.row, 12).setValue('Void');
    billPay_(bs, br, bamt, left > 0 ? 'part' : '', left, tz);
    SpreadsheetApp.flush();
    return JSON.stringify({ bill: billOut_(bs, br, tz) });
  } finally { lock.releaseLock(); }
}


// Cash Payment tab: cash paid out. Returns { r:[[day, party, amount, to, remark, remark2, tag]] } (tag = the extra column, for example LR, Teknaf or IBBL)
function getCash() { guard_('en'); return memo_('getCash', [], 120, () => getCash_c_()); }
function getCash_c_() {
  const ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName('Cash Payment');
  const n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 7).getValues() : [], out = [];
  for (let i = 0; i < v.length; i++) {
    const r = v[i], d = day_(r[0], tz), a = num_(r[2]); if (d === null || !a) continue;
    out.push([d, txt_(r[1]), a, txt_(r[3]), txt_(r[4]), txt_(r[5]), txt_(r[6]), i + 2]);
  }
  return JSON.stringify({ r: out });
}


// ---------- Entry (owner only): payments, cash payments and daily bank deposits ----------
// New rows go to the first empty row of the tab. Dates are written as real dates. The BP tab's formula columns (Total Dep, Total Payment, Closing) are never written.
const BANKS_ = ['BA 68', 'BA 125', 'IBBL 2951', 'BA 145'];
function needEditor_() { if (!isEditor_()) throw new Error('Only the owner can change data.'); }
function dstr_(d) { return Utilities.formatDate(new Date(d * 864e5), 'UTC', 'yyyy-MM-dd'); }
function putDate_(cell, d) { cell.setValue(dstr_(d)).setNumberFormat('d MMM yyyy'); }
function lastRowA_(sh) {                       // last row that has something in column A (the sheet's own last row can be pushed down by formulas)
  const v = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues(); let i = v.length - 1;
  while (i >= 0 && (v[i][0] === '' || v[i][0] == null)) i--; return i + 1;
}
function nextRow_(sh) { const r = lastRowA_(sh) + 1; if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100); return r; }
function pos_(v, name) { const a = +v; if (!(a > 0)) throw new Error(name + ' must be above zero.'); return a; }
function dayOk_(d) { if (!(d > 19000 && d < 25000)) throw new Error('Check the date.'); return +d; }
// BP row for a banking date and account. Creates it (and the day's other accounts) when missing.
function bpRow_(bp, tz, day, bank, create) {
  const n = lastRowA_(bp), v = n > 1 ? bp.getRange(2, 1, n - 1, 2).getValues() : []; let hasDay = false;
  for (let i = 0; i < v.length; i++) if (day_(v[i][0], tz) === day) { hasDay = true; if (txt_(v[i][1]) === bank) return i + 2; }
  if (!create) return 0;
  const add = hasDay ? [bank] : BANKS_.concat(BANKS_.indexOf(bank) < 0 ? [bank] : []);
  let row = 0;
  for (const b of add) {
    const r = nextRow_(bp); putDate_(bp.getRange(r, 1), day); bp.getRange(r, 2).setValue(b);
    try { bp.getRange(2, 2).copyTo(bp.getRange(r, 2), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false); } catch (e) {}
    if (b === bank) row = r;
  }
  return row;
}
// BP's Party Payment and Loan Clearance for a day and account are the PP Schedule totals, so they are set from it.
function syncBP_(ss, tz, day, bank) {
  const pp = ss.getSheetByName('PP Schedule'), bp = ss.getSheetByName('BP'), n = lastRowA_(pp);
  const v = n > 1 ? pp.getRange(2, 1, n - 1, 6).getValues() : []; let p = 0, q = 0;
  for (const r of v) if (day_(r[0], tz) === day && txt_(r[4]) === bank) { if (txt_(r[5]) === 'Loan Adjustment') q += num_(r[3]); else p += num_(r[3]); }
  const row = bpRow_(bp, tz, day, bank, p > 0 || q > 0); if (!row) return;
  const m = bpMap_(bp); bp.getRange(row, m.pp + 1).setValue(p || ''); bp.getRange(row, m.ln + 1).setValue(q || '');
}
function withLock_(fn) { needEditor_(); const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

function addPay(json) { try { return addPay_w_(json); } finally { bump_(); } }
function addPay_w_(json) { guard_('en');
  return withLock_(() => {
    const q = JSON.parse(json), day = dayOk_(q.day), amt = pos_(q.amount, 'Amount'), to = txt_(q.to), bank = txt_(q.bank), type = txt_(q.type) || 'Payment';
    if (!to) throw new Error('Enter who the payment is to.');
    if (BANKS_.indexOf(bank) < 0) throw new Error('Choose the bank.');
    if (['Payment', 'Loan Adjustment'].indexOf(type) < 0) throw new Error('Choose the type.');
    const ss = SpreadsheetApp.openById(MGMT_ID), sh = ss.getSheetByName(LEDGER), r = nextRow_(sh);
    putDate_(sh.getRange(r, 1), day); sh.getRange(r, 2, 1, 6).setValues([[bank, 'Payment', type === 'Loan Adjustment' ? 'Loan adjustment' : 'Party payment', to, txt_(q.details), amt]]);
    SpreadsheetApp.flush(); return 'ok';
  });
}
const ENTRY_DEP_ = ['Counter collection', 'bKash collection', 'Online sales', 'LR income', 'Fund transfer', 'Other deposit'];
const ENTRY_PAY_ = ['Party payment', 'Group payment', 'Loan adjustment', 'Salary', 'Rent', 'Document expense', 'Fuel', 'VAT', 'Bank charge', 'Fund transfer', 'Other payment'];
function growTo_(sh, lastRow) { if (lastRow > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), lastRow - sh.getMaxRows() + 50); }
// Data Entry page, Bank position: q = { rows:[{day, bank, type, cat, party, details, amount}] }. Every row is checked first; nothing is written unless all rows are valid.
function addLedgerRows(json) { try { return addLedgerRows_w_(json); } finally { bump_(); } }
function addLedgerRows_w_(json) { guard_('en');
  return withLock_(() => {
    const q = JSON.parse(json), rows = q.rows || []; if (!rows.length) throw new Error('Nothing to save.');
    const out = rows.map((x, i) => {
      const n = 'Row ' + (i + 1) + ': ', day = dayOk_(x.day), bank = txt_(x.bank), type = txt_(x.type), cat = txt_(x.cat), party = txt_(x.party), amt = +x.amount;
      if (BANKS_.indexOf(bank) < 0) throw new Error(n + 'choose the bank.');
      if (type !== 'Deposit' && type !== 'Payment') throw new Error(n + 'choose Deposit or Payment.');
      if ((type === 'Deposit' ? ENTRY_DEP_ : ENTRY_PAY_).indexOf(cat) < 0) throw new Error(n + 'choose the category.');
      if (!(amt > 0)) throw new Error(n + 'enter the amount (above zero).');
      if (type === 'Payment' && !party && cat !== 'Bank charge') throw new Error(n + 'enter who the payment is to.');
      return [dstr_(day), bank, type, cat, party, txt_(x.details), amt];
    });
    const sh = SpreadsheetApp.openById(MGMT_ID).getSheetByName(LEDGER), r = nextRow_(sh); growTo_(sh, r + out.length - 1);
    sh.getRange(r, 1, out.length, 7).setValues(out); sh.getRange(r, 1, out.length, 1).setNumberFormat('d MMM yyyy');
    SpreadsheetApp.flush(); return String(out.length);
  });
}
// Data Entry page, Cash position (expenses): q = { rows:[{day, party, amount, to, r1, r2, tag}] } saved to the Cash Payment tab.
function addCashRows(json) { try { return addCashRows_w_(json); } finally { bump_(); } }
function addCashRows_w_(json) { guard_('en');
  return withLock_(() => {
    const q = JSON.parse(json), rows = q.rows || []; if (!rows.length) throw new Error('Nothing to save.');
    const out = rows.map((x, i) => {
      const n = 'Row ' + (i + 1) + ': ', day = dayOk_(x.day), party = txt_(x.party), amt = +x.amount;
      if (!party) throw new Error(n + 'enter the party.'); if (!(amt > 0)) throw new Error(n + 'enter the amount (above zero).');
      return [dstr_(day), party, amt, txt_(x.to), txt_(x.r1), txt_(x.r2), txt_(x.tag)];
    });
    const sh = SpreadsheetApp.openById(MGMT_ID).getSheetByName('Cash Payment'), r = nextRow_(sh); growTo_(sh, r + out.length - 1);
    sh.getRange(r, 1, out.length, 7).setValues(out); sh.getRange(r, 1, out.length, 1).setNumberFormat('d MMM yyyy'); sh.getRange(r, 3, out.length, 1).setNumberFormat('#,##0');
    SpreadsheetApp.flush(); return String(out.length);
  });
}
// ---------- Data Entry page, Cash position: Daily Cashflow sheet (Expenses and Advances tabs) ----------
// Lists tab: approvers C2:C20, expense heads I2:I40, "advance given from" K2:K25. Data rows start at row 3 (rows 1-2 are headings).
function cfList_(sh, a1) { return sh.getRange(a1).getValues().map(r => txt_(r[0])).filter(String); }
// Returns the drop-down lists plus the latest entries of both tabs (to check what was saved).
function getCashForm() { guard_('en', 'dy');
  const ss = SpreadsheetApp.openById(CASHFLOW_ID), L = ss.getSheetByName('Lists');
  const last = (name, cols, n) => { const r = cfRows_(ss, name, cols).filter(x => cfKey_(x[0])); return r.slice(-n).reverse(); };
  return JSON.stringify({
    can: isEditor_(), heads: cfList_(L, 'I2:I40'), approvers: cfList_(L, 'C2:C20'), from: cfList_(L, 'K2:K25'),
    accounts: cfList_(L, 'A2:A20'), banks: cfList_(L, 'G2:G20'),
    // Cash Receiving rows (latest 90): [date, account, sales, trips, income, approved, advances, bank, online, payable, notes x10, received, paid to bank, bank account, slip, difference, remarks]
    rec: cfRows_(ss, 'Cash Receiving', 32).filter(x => cfKey_(x[0])).slice(-90).map(r => [cfKey_(r[0]), txt_(r[1])].concat(r.slice(2, 22).map(cfNum_), [txt_(r[22]), txt_(r[23]), cfNum_(r[24]), txt_(r[25]), cfNum_(r[26]), cfNum_(r[28]), cfNum_(r[29]), cfNum_(r[30]), cfNum_(r[31])])),
    exp: last('Expenses', 18, 12).map(r => [cfKey_(r[0]), txt_(r[2]), txt_(r[3]), txt_(r[4]), txt_(r[5]), cfNum_(r[16]) || r.slice(6, 16).reduce((t, v, i) => t + cfNum_(v) * CF_NOTES[i], 0)]),
    adv: last('Advances', 32, 12).map(r => [cfKey_(r[0]), txt_(r[2]), txt_(r[3]), txt_(r[4]), txt_(r[5]), cfNum_(r[17]) || r.slice(7, 17).reduce((t, v, i) => t + cfNum_(v) * CF_NOTES[i], 0), txt_(r[31])])
  });
}
// Data Entry page, Cash receiving: one row per date and account in the Cash Receiving tab.
// q = { day, account, sales, trips, income, bank, online, notes:[10 counts], toBank, bankAcc, ref, remarks }
// If the row for that date and account exists it is updated, otherwise a new row is added.
// Only the typed columns are written (C:E, H:I, K:T, V:X, Z). Approved Expenses, Advances, Cash Payable, Cash Received and Difference are the sheet's own formulas.
function saveCashReceiving(json) { try { return saveCashReceiving_w_(json); } finally { bump_(); } }
function saveCashReceiving_w_(json) { guard_('en', 'dy');
  return withLock_(() => {
    const q = JSON.parse(json), day = dayOk_(q.day), key = dstr_(day), ss = SpreadsheetApp.openById(CASHFLOW_ID), L = ss.getSheetByName('Lists');
    const acc = txt_(q.account); if (cfList_(L, 'A2:A20').indexOf(acc) < 0) throw new Error('Choose the account.');
    const num = (v, name) => { if (v === '' || v === null || v === undefined) return ''; const n = +String(v).replace(/,/g, ''); if (isNaN(n) || n < 0) throw new Error(name + ' must be a number, zero or above.'); return n; };
    const notes = CF_NOTES.map((_, i) => { const n = num((q.notes || [])[i], 'Note count'); if (n !== '' && n !== Math.round(n)) throw new Error('Note counts must be whole numbers.'); return n || ''; });
    const bankAcc = txt_(q.bankAcc), toBank = num(q.toBank, 'Paid to bank');
    if (bankAcc && cfList_(L, 'G2:G20').indexOf(bankAcc) < 0) throw new Error('Choose the bank account from the list.');
    if (toBank > 0 && !bankAcc) throw new Error('Choose the bank account for the amount paid to bank.');
    const sh = ss.getSheetByName('Cash Receiving'), mx = sh.getMaxRows(), v = sh.getRange(3, 1, Math.max(1, mx - 2), 2).getValues();
    let row = 0, last = 2;
    for (let i = 0; i < v.length; i++) { if (v[i][0] === '' || v[i][0] === null) continue; last = i + 3; if (!row && cfKey_(v[i][0]) === key && txt_(v[i][1]).toLowerCase() === acc.toLowerCase()) row = i + 3; }
    const isNew = !row; if (isNew) { row = last + 1; growTo_(sh, row); sh.getRange(row, 1, 1, 2).setValues([[key, acc]]); sh.getRange(row, 1).setNumberFormat('dd-MMM-yyyy'); }
    sh.getRange(row, 3, 1, 3).setValues([[num(q.sales, 'Sales'), num(q.trips, 'Trips'), num(q.income, 'Income')]]);
    sh.getRange(row, 8, 1, 2).setValues([[num(q.bank, 'Bank deposit'), num(q.online, 'Online')]]);
    sh.getRange(row, 11, 1, 10).setValues([notes]);
    sh.getRange(row, 22, 1, 3).setValues([[toBank, bankAcc, txt_(q.ref)]]);
    sh.getRange(row, 26).setValue(txt_(q.remarks));
    sh.getRange(row, 27).setValue(num(q.buses, 'Buses'));   // AA Buses
    // AC:AF online split (BDtickets, Shohoz, Soudia Web, Jatri). Online (column I) is the total: if the split adds up to more than the typed Online, Online becomes the split total.
    const sp = ['bd', 'sh', 'sw', 'ja'].map(k => num(q[k], 'Online split')); sh.getRange(row, 29, 1, 4).setValues([sp]);
    const spT = sp.reduce((t, v) => t + (+v || 0), 0); if (spT > (+num(q.online, 'Online') || 0)) sh.getRange(row, 9).setValue(spT);
    SpreadsheetApp.flush(); return isNew ? 'added' : 'updated';
  });
}
function cfNextRow_(sh) {
  const mx = sh.getMaxRows(), v = sh.getRange(3, 1, Math.max(1, mx - 2), 1).getValues(); let last = 2;
  for (let i = 0; i < v.length; i++) if (v[i][0] !== '' && v[i][0] !== null) last = i + 3;
  return last + 1;
}
// q = { kind:'exp'|'adv', day, voucher, head, details, to, by, remarks, from, purpose, category, notes:[10 counts], amount, byMe }
function addCashflowEntry(json) { try { return addCashflowEntry_w_(json); } finally { bump_(); } }
function addCashflowEntry_w_(json) { guard_('en', 'dy');
  return withLock_(() => {
    const q = JSON.parse(json), day = dayOk_(q.day), ss = SpreadsheetApp.openById(CASHFLOW_ID), L = ss.getSheetByName('Lists');
    const notes = CF_NOTES.map((_, i) => { const n = Math.round(+((q.notes || [])[i]) || 0); if (n < 0) throw new Error('Note counts cannot be negative.'); return n; });
    const nsum = notes.reduce((t, n, i) => t + n * CF_NOTES[i], 0), typed = +q.amount || 0;
    if (nsum > 0 && typed > 0 && Math.round(typed) !== nsum) throw new Error('The notes add up to ' + nsum.toLocaleString('en-US') + ' but the amount is ' + Math.round(typed).toLocaleString('en-US') + '. Fix one of them.');
    const amount = nsum > 0 ? nsum : Math.round(typed * 100) / 100; if (!(amount > 0)) throw new Error('Enter the amount or the note counts.');
    const heads = cfList_(L, 'I2:I40'), appr = cfList_(L, 'C2:C20'), from = cfList_(L, 'K2:K25');
    const by = txt_(q.by); if (by && appr.indexOf(by) < 0) throw new Error('Choose the approver from the list.');
    const nv = notes.map(n => n || '');
    if (q.kind === 'exp') {
      const head = txt_(q.head); if (heads.indexOf(head) < 0) throw new Error('Choose the expense head.');
      if (!txt_(q.details) && !txt_(q.to)) throw new Error('Enter the details or who it was paid to.');
      const sh = ss.getSheetByName('Expenses'), r = cfNextRow_(sh); growTo_(sh, r);
      const f = nsum > 0 ? '=G' + r + '*1000+H' + r + '*500+I' + r + '*200+J' + r + '*100+K' + r + '*50+L' + r + '*20+M' + r + '*10+N' + r + '*5+O' + r + '*2+P' + r : amount;
      sh.getRange(r, 1, 1, 18).setValues([[dstr_(day), txt_(q.voucher), head, txt_(q.details), txt_(q.to), by].concat(nv, [f, txt_(q.remarks)])]);
      sh.getRange(r, 1).setNumberFormat('dd-MMM-yyyy'); sh.getRange(r, 17).setNumberFormat('#,##0');
    } else if (q.kind === 'adv') {
      const to = txt_(q.to); if (!to) throw new Error('Enter who the advance is given to.');
      const fr = txt_(q.from) || 'Me'; if (from.indexOf(fr) < 0) throw new Error('Choose where the advance is given from.');
      const cat = txt_(q.category); if (cat && heads.indexOf(cat) < 0) throw new Error('Choose the category from the list.');
      const sh = ss.getSheetByName('Advances'), r = cfNextRow_(sh); growTo_(sh, r);
      sh.getRange(r, 1, 1, 18).setValues([[dstr_(day), txt_(q.voucher), to, fr, txt_(q.purpose), cat, by].concat(nv, [amount])]);
      sh.getRange(r, 32).setValue(fr === 'Me' ? 'Yes' : 'No'); sh.getRange(r, 1).setNumberFormat('dd-MMM-yyyy'); sh.getRange(r, 18).setNumberFormat('#,##0');
    } else throw new Error('Choose Expense or Advance.');
    SpreadsheetApp.flush(); return String(amount);
  });
}
function addCash(json) { try { return addCash_w_(json); } finally { bump_(); } }
function addCash_w_(json) { guard_('en', 'dy');
  return withLock_(() => {
    const q = JSON.parse(json), day = dayOk_(q.day), amt = pos_(q.amount, 'Amount'), party = txt_(q.party);
    if (!party) throw new Error('Enter the party.');
    const ss = SpreadsheetApp.openById(MGMT_ID), sh = ss.getSheetByName('Cash Payment'), r = nextRow_(sh);
    putDate_(sh.getRange(r, 1), day); sh.getRange(r, 2, 1, 6).setValues([[party, amt, txt_(q.to), txt_(q.r1), txt_(q.r2), txt_(q.tag)]]);
    sh.getRange(r, 3).setNumberFormat('#,##0'); SpreadsheetApp.flush(); return 'ok';
  });
}
// Remove a row the owner entered by mistake (the page checks it is still the same row). Payments are deleted from the Bank Ledger, cash payments are cleared.
function delRow(json) { try { return delRow_w_(json); } finally { bump_(); } }
function delRow_w_(json) { guard_('en', 'dy');
  return withLock_(() => {
    const q = JSON.parse(json), ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone(), pp = q.kind === 'pp';
    const sh = ss.getSheetByName(pp ? LEDGER : 'Cash Payment');
    if (!(q.row > (pp ? LEDGER_TOP : 1) && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
    const r = sh.getRange(q.row, 1, 1, 7).getValues()[0], a = num_(r[pp ? 6 : 2]);
    if (day_(r[0], tz) !== q.day || txt_(r[pp ? 4 : 1]) !== q.party || a !== q.amount) throw new Error('The sheet has changed: reload the page and try again.');
    if (pp) sh.deleteRow(q.row); else sh.getRange(q.row, 1, 1, 7).clearContent();
    SpreadsheetApp.flush(); return 'ok';
  });
}
// Daily deposits: q = { day, v: { "BA 68": { "0": 123, "3": null, ... } } } where the key is the deposit source's position in the list getBank returns (src); null clears it.
// Each box is one ledger row (date, bank, Deposit, the source's category and name); saving replaces that row, an empty box removes it.
function saveDep(json) { try { return saveDep_w_(json); } finally { bump_(); } }
function saveDep_w_(json) { guard_('en');
  return withLock_(() => {
    const q = JSON.parse(json), day = dayOk_(q.day), ss = SpreadsheetApp.openById(MGMT_ID), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName(LEDGER);
    let n = 0;
    for (const bank in q.v) {
      if (BANKS_.indexOf(bank) < 0) throw new Error('Unknown bank ' + bank);
      const cells = q.v[bank];
      for (const k of Object.keys(cells)) {
        const j = +k; if (!(j >= 0 && j < BANK_SRC_.length)) throw new Error('Unknown deposit column.');
        const a = cells[k], hit = ledgerRows_(ss, tz).filter(x => x.day === day && x.bank === bank && x.type === 'Deposit' && depExact_(j, x.cat, x.party)).map(x => x.row);
        if (a === null || a === '') { hit.sort((x, y) => y - x).forEach(r => sh.deleteRow(r)); }
        else {
          const amt = pos_(a, 'Deposit');
          if (hit.length) { sh.getRange(hit[0], 7).setValue(amt); hit.slice(1).sort((x, y) => y - x).forEach(r => sh.deleteRow(r)); }
          else { const r = nextRow_(sh); putDate_(sh.getRange(r, 1), day); sh.getRange(r, 2, 1, 6).setValues([[bank, 'Deposit', DEP_CAT_[j], DEP_PARTY_[j], '', amt]]); }
        }
        n++;
      }
    }
    SpreadsheetApp.flush(); return String(n);
  });
}

// Sales report by coach number: month totals for the coaches asked for (an empty list = every coach), in one pass over the monthly detail files.
// Returns { coachName: { "yyyy-mm": [seats, sales] } }
function getCoachMonthly(json) { guard_('sl', 'sr', 'bi', 'ov'); return memo_('getCoachMonthly', [json], 120, () => getCoachMonthly_c_(json)); }
function getCoachMonthly_c_(json) {
  const want = new Set(JSON.parse(json)), D = JSON.parse(read_('dict.json')), ix = new Set(), out = {};
  D.coach.forEach((c, i) => { if (want.has(c)) ix.add(i); });
  for (const f of details_()) {
    for (const x of JSON.parse(f.getBlob().getDataAsString())) {
      if (want.size && !ix.has(x[4])) continue;
      const c = D.coach[x[4]], m = new Date(x[0] * 864e5).toISOString().slice(0, 7), o = out[c] || (out[c] = {}), t = o[m] || (o[m] = [0, 0]);
      t[0] += x[6]; t[1] += x[7];
    }
  }
  return JSON.stringify(out);
}

// Sales report, coach profile: everything one coach number sold, in one pass over the monthly detail files.
// Returns { rows:[[day, routeType, seats, sales]], routes:[[name, seats, sales]], buses:[[bus, seats, sales, days]], counters:[[name, seats, sales]] }
function getCoachProfile(name) { guard_('sl', 'sr', 'bi', 'ov'); return memo_('getCoachProfile', [name], 120, () => getCoachProfile_c_(name)); }
function getCoachProfile_c_(name) {
  const D = JSON.parse(read_('dict.json')), ci = D.coach.indexOf(name), rows = {}, ro = {}, bu = {}, co = {};
  if (ci < 0) return JSON.stringify({ rows: [], routes: [], buses: [], counters: [] });
  for (const f of details_()) {
    for (const x of JSON.parse(f.getBlob().getDataAsString())) {
      if (x[4] !== ci) continue;
      const k = x[0] + '|' + x[2], r = rows[k] || (rows[k] = [x[0], x[2], 0, 0]); r[2] += x[6]; r[3] += x[7];
      const a = ro[x[3]] || (ro[x[3]] = [0, 0]); a[0] += x[6]; a[1] += x[7];
      const b = bu[x[5]] || (bu[x[5]] = [0, 0, {}]); b[0] += x[6]; b[1] += x[7]; b[2][x[0]] = 1;
      const c = co[x[1]] || (co[x[1]] = [0, 0]); c[0] += x[6]; c[1] += x[7];
    }
  }
  const top = (o, nm, n) => Object.keys(o).map(k => [nm(k), o[k][0], o[k][1], o[k][2] ? Object.keys(o[k][2]).length : 0]).sort((p, q) => q[2] - p[2]).slice(0, n);
  return JSON.stringify({
    rows: Object.keys(rows).map(k => rows[k]),
    routes: top(ro, k => D.route[k], 15).map(r => r.slice(0, 3)),
    buses: top(bu, k => D.bus[k], 15),
    counters: top(co, k => D.c[k], 15).map(r => r.slice(0, 3))
  });
}

// ---------- Operations records: expenses, maintenance, documents, fleet ----------
// Kept in their own spreadsheet, "MM Operations Data", created in the Sales Data folder the first time the owner saves a record.
const OPSMOD = {
  expenses:    { tab: 'Expenses',    cols: ['Date', 'Category', 'Description', 'Amount', 'Bus / Coach', 'Paid via', 'Notes'], dates: [0], nums: [3], need: [0, 1, 3] },
  maintenance: { tab: 'Maintenance', cols: ['Date', 'Bus', 'Type', 'Description', 'Cost', 'Workshop', 'Next due'], dates: [0, 6], nums: [4], need: [0, 1, 2] },
  documents:   { tab: 'Documents',   cols: ['Bus', 'Document', 'Number', 'Issue date', 'Expiry date', 'Cost', 'Notes'], dates: [3, 4], nums: [5], need: [0, 1, 4] },
  fleet:       { tab: 'Fleet',       cols: ['Bus', 'Type', 'Model', 'Registration', 'Seats', 'Status', 'Notes'], dates: [], nums: [4], need: [0] },
  fuel:        { tab: 'Fuel',        cols: ['Date', 'Bus', 'Litres', 'Amount', 'Odometer km', 'Station', 'Notes'], dates: [0], nums: [2, 3, 4], need: [0, 1, 2] },
  incidents:   { tab: 'Incidents',   cols: ['Date', 'Bus', 'Type', 'Description', 'Cost', 'Status', 'Follow-up date', 'Responsible'], dates: [0, 6], nums: [4], need: [0, 2, 5] },
  staff:       { tab: 'Staff',       cols: ['Name', 'Role', 'Phone', 'Licence no', 'Licence expiry', 'Joined', 'Status', 'Monthly pay'], dates: [4, 5], nums: [7], need: [0, 1] },
  rent:        { tab: 'Counter rent', cols: ['Counter', 'Monthly rent', 'Landlord', 'Advance', 'Lease end', 'Status', 'Notes'], dates: [4], nums: [1, 3], need: [0, 1] },
  salary:      { tab: 'Salary',      cols: ['Salary month', 'Staff', 'Amount', 'Paid on', 'Paid via', 'Notes'], dates: [0, 3], nums: [2], need: [0, 1, 2] },
  loans:       { tab: 'Loans',       cols: ['Lender', 'Loan ref', 'Purpose', 'Principal', 'Outstanding', 'Monthly instalment', 'Next due', 'End date', 'Status'], dates: [6, 7], nums: [3, 4, 5], need: [0, 3] },
  dues:        { tab: 'Payments',    cols: ['Due date', 'Type', 'Pay to', 'Reference', 'Amount', 'Status', 'Paid on', 'Notes'], dates: [0, 6], nums: [4], need: [0, 1, 2, 4] },
  purchases:   { tab: 'Purchases',   cols: ['Date', 'Supplier', 'Item', 'Category', 'Amount', 'Paid', 'Bus'], dates: [0], nums: [4, 5], need: [0, 1, 4] }
};
function opsTabs_(ss) {
  for (const k in OPSMOD) {
    const m = OPSMOD[k]; if (ss.getSheetByName(m.tab)) continue;
    const sh = ss.insertSheet(m.tab);
    sh.getRange(1, 1, 1, m.cols.length).setValues([m.cols]).setFontWeight('bold').setBackground('#0b3a5c').setFontColor('#ffffff'); sh.setFrozenRows(1);
  }
}
function ops_(create) {
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('OPS_ID');
  if (id) { try { const o = SpreadsheetApp.openById(id); if (create) opsTabs_(o); return o; } catch (e) {} }
  if (!create) return null;
  const ss = SpreadsheetApp.create('MM Operations Data'); let i = 0;
  for (const k in OPSMOD) {
    const m = OPSMOD[k], sh = i === 0 ? ss.getSheets()[0].setName(m.tab) : ss.insertSheet(m.tab); i++;
    sh.getRange(1, 1, 1, m.cols.length).setValues([m.cols]).setFontWeight('bold').setBackground('#0b3a5c').setFontColor('#ffffff'); sh.setFrozenRows(1);
  }
  try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(FOLDER_ID)); } catch (e) {}
  opsTabs_(ss); P.setProperty('OPS_ID', ss.getId()); return ss;
}
// Returns { ready, can, mods:{ expenses:[[sheetRow, v1, v2, ...]], ... } }; dates are day numbers, amounts are numbers
function getRecords() { guard_('ov', 'en', 'ro', 'rp'); return memo_('getRecords', [], 120, () => getRecords_c_()); }
function getRecords_c_() {
  const ss = ops_(false), out = { ready: !!ss, can: isEditor_(), mods: {} };
  for (const k in OPSMOD) {
    const m = OPSMOD[k], sh = ss ? ss.getSheetByName(m.tab) : null, rows = [];
    if (sh && sh.getLastRow() > 1) {
      const tz = ss.getSpreadsheetTimeZone(), v = sh.getRange(2, 1, sh.getLastRow() - 1, m.cols.length).getValues();
      v.forEach((r, i) => {
        if (r.every(x => x === '')) return;
        const o = [i + 2]; r.forEach((x, j) => o.push(m.dates.indexOf(j) >= 0 ? (x === '' ? 0 : day_(x, tz) || 0) : m.nums.indexOf(j) >= 0 ? num_(x) : txt_(x)));
        rows.push(o);
      });
    }
    out.mods[k] = rows;
  }
  return JSON.stringify(out);
}
// q = { mod, v:[value per column; dates as day numbers, numbers as numbers, null when empty] }. Owner only.
function recAdd(json) { try { return recAdd_w_(json); } finally { bump_(); } }
function recAdd_w_(json) { guard_('en', 'ro', 'ov');
  return withLock_(() => {
    const q = JSON.parse(json), m = OPSMOD[q.mod]; if (!m) throw new Error('Unknown record type.');
    if (!Array.isArray(q.v) || q.v.length !== m.cols.length) throw new Error('Check the form and try again.');
    m.need.forEach(j => { if (q.v[j] === null || q.v[j] === '' || q.v[j] === undefined) throw new Error(m.cols[j] + ' is needed.'); });
    const ss = ops_(true), sh = ss.getSheetByName(m.tab), r = nextRow_(sh);
    m.cols.forEach((c, j) => {
      const x = q.v[j], cell = sh.getRange(r, j + 1);
      if (m.dates.indexOf(j) >= 0) { if (x) putDate_(cell, dayOk_(x)); }
      else if (m.nums.indexOf(j) >= 0) { if (x !== null && x !== '') { cell.setValue(+x); cell.setNumberFormat('#,##0'); } }
      else cell.setValue(txt_(x));
    });
    SpreadsheetApp.flush(); return 'ok';
  });
}
// Clears one record the owner entered by mistake; q = { mod, row, sig } where sig is the text of the record's key column
function recDel(json) { try { return recDel_w_(json); } finally { bump_(); } }
function recDel_w_(json) { guard_('en', 'ro', 'ov');
  return withLock_(() => {
    const q = JSON.parse(json), m = OPSMOD[q.mod]; if (!m) throw new Error('Unknown record type.');
    const ss = ops_(false); if (!ss) throw new Error('Nothing to delete.');
    const sh = ss.getSheetByName(m.tab); if (!(q.row >= 2 && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
    sh.getRange(q.row, 1, 1, m.cols.length).clearContent(); SpreadsheetApp.flush(); return 'ok';
  });
}

// Changes some columns of one record (owner only). q = { mod, row, set: { "<column index>": value } }; dates as day numbers, numbers as numbers
function recUpd(json) { try { return recUpd_w_(json); } finally { bump_(); } }
function recUpd_w_(json) { guard_('en', 'ro', 'ov');
  return withLock_(() => {
    const q = JSON.parse(json), m = OPSMOD[q.mod]; if (!m) throw new Error('Unknown record type.');
    const ss = ops_(false); if (!ss) throw new Error('Nothing to change.');
    const sh = ss.getSheetByName(m.tab); if (!sh || !(q.row >= 2 && q.row <= sh.getLastRow())) throw new Error('Row not found: reload the page.');
    for (const k in (q.set || {})) {
      const j = +k, x = q.set[k]; if (!(j >= 0 && j < m.cols.length)) continue;
      const cell = sh.getRange(q.row, j + 1);
      if (m.dates.indexOf(j) >= 0) { if (x) putDate_(cell, dayOk_(x)); else cell.clearContent(); }
      else if (m.nums.indexOf(j) >= 0) { if (x === null || x === '') cell.clearContent(); else { cell.setValue(+x); cell.setNumberFormat('#,##0'); } }
      else cell.setValue(txt_(x));
    }
    SpreadsheetApp.flush(); return 'ok';
  });
}
// Route profitability: seats and sales per route, and per route and bus, for a day range. q = { from, to } as day numbers
// Returns { routes:[[routeIdx, seats, sales]], rb:[[routeIdx, busIdx, seats, sales]] } (indexes into dict.json route / bus)
function getRouteEcon(json) { guard_('ro', 'ov'); if (perm_().accounts) return JSON.stringify({ routes: [], rb: [] }); return memo_('getRouteEcon', [json], 120, () => getRouteEcon_c_(json)); }
function getRouteEcon_c_(json) {
  const q = JSON.parse(json), a = +q.from, b = +q.to; if (!(a > 19000 && b >= a && b < 25000)) throw new Error('Check the dates.');
  const ym = d => new Date(d * 864e5).toISOString().slice(0, 7), ro = {}, rb = {}, A = ym(a), B = ym(b);
  for (const f of details_()) {
    const mm = /^detail_(\d{4}-\d{2})\.json$/.exec(f.getName()); if (!mm || mm[1] < A || mm[1] > B) continue;
    for (const x of JSON.parse(f.getBlob().getDataAsString())) {
      if (x[0] < a || x[0] > b) continue;
      const r = ro[x[3]] || (ro[x[3]] = [0, 0]); r[0] += x[6]; r[1] += x[7];
      const k = x[3] + '|' + x[5], p = rb[k] || (rb[k] = [0, 0]); p[0] += x[6]; p[1] += x[7];
    }
  }
  return JSON.stringify({
    routes: Object.keys(ro).map(k => [+k, ro[k][0], ro[k][1]]),
    rb: Object.keys(rb).map(k => { const s = k.split('|'); return [+s[0], +s[1], rb[k][0], rb[k][1]]; })
  });
}

// ===== Route-type books (Route profitability > By route type) =====
// One Google Sheet per route type, named "<type> Combined" (MB-Lux -> "MB Lux Combined", LR -> "LR Combined", CDC, Aero, AC, Teknaf).
// Put the sheet anywhere in Drive that this account can open; it is found by name the first time and remembered.
// Columns (any order, found by heading): Journey Date, Bus, Route, Seat Capacity, Seat Sold, Total Sale, Salary, Fuel Lt., Fuel, Toll, Others, Chalan, Total Exp
// Every tab of the sheet that has such a heading row is read (a "Notes" tab is ignored). To pin a sheet to a type, add its id below.
const RT_BOOKS = {};   // optional: pin a type to its own sheet id (used only when the ALL tab has no rows for that account)
function rtNames_(t) { return [t.replace(/-/g, ' ') + ' Combined', t + ' Combined']; }
function rtBook_(t) {
  const P = PropertiesService.getScriptProperties(), k = 'RTB_' + t; let id = RT_BOOKS[t] || P.getProperty(k);
  if (id) return id;
  for (const n of rtNames_(t)) {
    const it = DriveApp.getFilesByName(n);
    while (it.hasNext()) { const f = it.next(); if (f.getMimeType() === MimeType.GOOGLE_SHEETS) { P.setProperty(k, f.getId()); return f.getId(); } }
  }
  return null;
}
function rtNum_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).replace(/[,\s৳]/g, ''); if (!s || s === '-') return 0;
  const neg = /^\(.*\)$/.test(s); s = s.replace(/[()]/g, ''); const n = parseFloat(s); return isNaN(n) ? 0 : (neg ? -n : n);
}
function rtDay_(v, tz) {
  if (v instanceof Date || typeof v === 'number') return day_(v, tz);
  const s = String(v == null ? '' : v).trim(); if (!s) return null;
  let m = s.match(/^(\d{1,2})[ \-\/\.](\w{3})\w*[ \-\/,\.]+(\d{2,4})$/);
  if (m && MON[m[2][0].toUpperCase() + m[2].slice(1).toLowerCase()] !== undefined) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return Date.UTC(y, MON[m[2][0].toUpperCase() + m[2].slice(1).toLowerCase()], +m[1]) / 864e5; }
  m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);           // dd/mm/yyyy
  if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return Date.UTC(y, +m[2] - 1, +m[1]) / 864e5; }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 : null;
}
const RT_COLS = { date: ['journeydate', 'date'], bus: ['bus', 'vehicleno', 'vehicle', 'busno'], route: ['route'], cap: ['seatcapacity', 'capacity'], sold: ['seatsold', 'sold'],
  sale: ['totalsale', 'sale', 'sales'], sal: ['salary'], fl: ['fuellt', 'fuelltr', 'fuellitre'], fuel: ['fuel', 'fuelcost'], toll: ['toll'], oth: ['others', 'other'], chalan: ['chalan', 'challan'], texp: ['totalexp', 'totalexpense'], buf: ['buffet'] };
// ONE book for every account: the Income Statement sheet, tab "ALL". Column "Account" (column P) says which account each row belongs to
// (MB-Lux, Aero, CDC, Teknaf, AC, LR ... or any new name). Every distinct account name shows up in Route profitability automatically.
const ROUTES_ID = '1ET4pkDNRz2ZpaCdj39UQ_GlWD6Ntve7rYCjeN5Imm9I'; // SBS Routes: Route Map tab
const RT_ALL = { id: '1K_PPyIWBsb-pbcnUuGA_jW4-ebrPR08q3db8fJkH0xI', tab: 'ALL' };
function rtAcct_(v) {
  const s = String(v == null ? '' : v).trim().replace(/\s+/g, ' '); if (!s) return '';
  const k = s.toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const t of RT) if (t.toLowerCase().replace(/[^a-z0-9]/g, '') === k) return t;
  return s;
}
function rtParse_(sh, tz, put) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2 || lc < 3) return false;
  const V = sh.getRange(1, 1, lr, lc).getValues(); let hr = -1, C = null, ca = -1;
  for (let r = 0; r < Math.min(8, V.length) && hr < 0; r++) {
    const H = V[r].map(h => String(h).toLowerCase().replace(/[^a-z]/g, '')), c = {};
    for (const k in RT_COLS) { const j = H.findIndex(h => RT_COLS[k].indexOf(h) >= 0); if (j >= 0) c[k] = j; }
    if (c.date !== undefined && c.sale !== undefined && c.route !== undefined) { hr = r; C = c; ca = H.findIndex(h => h === 'account' || h === 'accounttype' || h === 'routetype'); }
  }
  if (hr < 0) return false;
  const g = (row, k) => C[k] === undefined ? 0 : rtNum_(row[C[k]]);
  for (let r = hr + 1; r < V.length; r++) {
    const row = V[r], d = rtDay_(row[C.date], tz), route = txt_(row[C.route]), acct = ca < 0 ? '' : rtAcct_(row[ca]);
    if (d == null || !route) { put(acct, null); continue; }
    const bus = C.bus === undefined ? '' : bus_(row[C.bus]), sale = g(row, 'sale'), sal = g(row, 'sal'), fuel = g(row, 'fuel'), toll = g(row, 'toll'), oth = g(row, 'oth'), ch = g(row, 'chalan');
    const buf = g(row, 'buf'), te = C.texp !== undefined && row[C.texp] !== '' ? g(row, 'texp') : sal + fuel + toll + oth + ch + buf;
    put(acct, { d, bus, route, vals: [g(row, 'cap'), g(row, 'sold'), sale, sal, g(row, 'fl'), fuel, toll, oth, ch, te, buf], blank: row.every(x => x === '' || x == null) });
  }
  return true;
}
// Returns JSON { types:[{ n, ok, err, rows:[[day,busIdx,routeIdx,cap,sold,sale,salary,fuelLt,fuel,toll,others,chalan,totalExp]], buses:[], routes:[], skipped, tabs }] }
function getRouteTypes() { guard_('ro', 'dy', 'ov', 'rp'); const a = perm_().accounts, j = getRouteTypes_(); if (!a) return j; const o = JSON.parse(j); o.types = o.types.filter(t => a.indexOf(t.n) >= 0); return JSON.stringify(o); }
function getRouteTypes_() { return memo_('getRouteTypes_', [], 120, () => getRouteTypes__c_()); }
function getRouteTypes__c_() {
  const T = {}, order = [];
  const mk = t => T[t] || (T[t] = (order.push(t), { n: t, ok: false, rows: [], buses: [], routes: [], skipped: 0, tabs: 0, bi: {}, ri: {} }));
  RT.forEach(mk);
  const add = (o, x) => {
    if (!x) return;
    const b = x.bus in o.bi ? o.bi[x.bus] : (o.bi[x.bus] = o.buses.push(x.bus || '(no bus)') - 1), q = x.route in o.ri ? o.ri[x.route] : (o.ri[x.route] = o.routes.push(x.route) - 1);
    o.rows.push([x.d, b, q].concat(x.vals));
  };
  let allErr = '';
  try {                                              // 1) the ALL tab, split by the Account column
    const ss = SpreadsheetApp.openById(RT_ALL.id), tz = ss.getSpreadsheetTimeZone(), sh = ss.getSheetByName(RT_ALL.tab);
    if (!sh) allErr = 'Tab "' + RT_ALL.tab + '" not found in the Income Statement sheet';
    else {
      let noAcct = 0;
      const ok = rtParse_(sh, tz, (acct, x) => {
        if (!x) return;
        if (!acct) { noAcct++; acct = 'MB-Lux'; }    // a row with no account name is treated as MB-Lux (the original sheet)
        const o = mk(acct); add(o, x);
      });
      if (!ok) allErr = 'Tab "' + RT_ALL.tab + '" has no heading row (Journey Date, Route, Total Sale ...)';
      order.forEach(t => { if (T[t].rows.length) T[t].tabs = 1; });
      if (noAcct) T['MB-Lux'].err2 = noAcct + ' rows had no account name and were counted as MB-Lux';
    }
  } catch (e) { allErr = String(e.message || e); }
  order.slice().forEach(t => {                       // 2) an account with nothing in ALL: its own "<type> Combined" book, if one exists
    const o = T[t]; if (o.rows.length) return;
    try {
      const id = rtBook_(t);
      if (!id || id === RT_ALL.id) return;
      const ss = SpreadsheetApp.openById(id), tz = ss.getSpreadsheetTimeZone();
      ss.getSheets().forEach(sh => { if (rtParse_(sh, tz, (a, x) => { if (x) add(o, x); })) o.tabs++; });
    } catch (e) { o.err = String(e.message || e); }
  });
  const types = order.map(t => {
    const o = T[t]; delete o.bi; delete o.ri; o.ok = o.rows.length > 0;
    if (!o.ok && !o.err) o.err = allErr || 'No rows with Account = "' + t + '" in the ALL tab yet';
    return o;
  });
  return JSON.stringify({ types: types, rmap: rtRouteMap_() });
}
// "Route Map" tab of the Income Statement sheet: column "Route in income statement" -> column "Route group".
// Every raw route name listed there is shown under its group (both directions together) in Route profitability and Route statements.
// Names not listed are grouped automatically. Edit the Route group column to rename or merge routes; add a row for any new raw name.
function rtRouteMap_() {
  const m = {};
  try {
    const sh = SpreadsheetApp.openById(ROUTES_ID).getSheetByName('Route Map'); if (!sh || sh.getLastRow() < 2) return m;
    const V = sh.getRange(1, 1, sh.getLastRow(), Math.max(3, sh.getLastColumn())).getValues(), H = V[0].map(h => String(h).toLowerCase());
    let cr = H.findIndex(h => h.indexOf('route in') === 0 || h.indexOf('raw') >= 0); if (cr < 0) cr = 1;
    let cg = H.findIndex(h => h.indexOf('group') >= 0); if (cg < 0) cg = 2;
    for (let r = 1; r < V.length; r++) { const k = String(V[r][cr] || '').toLowerCase().replace(/\s+/g, ''), g = String(V[r][cg] || '').trim(); if (k && g) m[k] = g; }
  } catch (e) {}
  return m;
}

// Bank deposit, web & online and previous-day fuel per route type per day, for the Route profitability Summary
// and the day-by-day table. Returns JSON { rows: [[dayNumber, type, bankDeposit, webAndOnline, prevDayFuel]] }
// Source 1: a tab named "Deposits" in the type's own book (headings: Date, Bank Deposit, Web & Online, Less: Prev Day Fuel).
// Source 2 (a type with no Deposits tab): the Daily Cashflow "Cash Receiving" tab, account name = route type.
function getTypeDeposits() { guard_('ro', 'dy', 'ov', 'rp'); const a = perm_().accounts, j = getTypeDeposits_(); if (!a) return j; const o = JSON.parse(j); o.rows = o.rows.filter(r => a.indexOf(r[1]) >= 0); return JSON.stringify(o); }
function getTypeDeposits_() { return memo_('getTypeDeposits_', [], 120, () => getTypeDeposits__c_()); }
function getTypeDeposits__c_() {
  const rows = [], have = {};
  ['MB-Lux'].forEach(t => {                          // the Deposits tab of the Income Statement sheet belongs to MB-Lux; other accounts use Cash Receiving below
    try {
      const id = RT_ALL.id;
      const ss = SpreadsheetApp.openById(id), sh = ss.getSheetByName('Deposits'); if (!sh) return;
      const tz = ss.getSpreadsheetTimeZone(), lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 3) return;
      const V = sh.getRange(1, 1, lr, lc).getValues(), norm = h => String(h).toLowerCase().replace(/[^a-z]/g, '');
      let hr = -1; for (let r = 0; r < Math.min(6, V.length); r++) if (V[r].map(norm).indexOf('date') >= 0 && V[r].map(norm).indexOf('totaldeposit') >= 0) { hr = r; break; }
      if (hr < 0) return;
      const H = V[hr].map(norm), ix = k => H.indexOf(k), cd = ix('date'), cb = ix('bankdeposit'), cw = ix('webonline'), cf = ix('lessprevdayfuel');
      for (let r = hr + 1; r < V.length; r++) {
        const d = rtDay_(V[r][cd], tz); if (d == null) continue;
        rows.push([d, t, cb < 0 ? 0 : rtNum_(V[r][cb]), cw < 0 ? 0 : rtNum_(V[r][cw]), cf < 0 ? 0 : rtNum_(V[r][cf])]);
      }
      have[t] = true;
    } catch (e) {}
  });
  const ss = SpreadsheetApp.openById(CASHFLOW_ID);
  cfRows_(ss, 'Cash Receiving', 9).forEach(r => {
    const t = String(r[1] || '').trim(), k = cfKey_(r[0]); if (!t || have[t] || !k) return;
    rows.push([Date.parse(k) / 864e5, t, cfNum_(r[7]), cfNum_(r[8]), 0]);
  });
  return JSON.stringify({ rows: rows });
}

// ===== Rent (page "Rent") =====
// Reads the "Rent" Google Sheet: one tab per month named like "Sep 2026" with the headings
// Dept, Zone, Counter Name, M. Rent, Adjustment, Pay. Rent, Remarks. A final row whose Dept is "Total" is used only to check the sums.
// Add a new tab each month (copy last month's) and the page picks it up. Edit the rent in the sheet, not on the page.
const RENT_ID = '11frxmTBy6naLmslZm2cxwqk1bavzHKVDnbl9P6iTMSk';
const RENT_COLS = { dept: ['dept', 'department'], zone: ['zone'], name: ['countername', 'counter'], rent: ['mrent', 'monthlyrent', 'rent'], adj: ['adjustment', 'adj'], pay: ['payrent', 'payablerent', 'payable'], rem: ['remarks', 'remark', 'note', 'notes'] };
// Returns JSON { months:[{ m:'2026-09', tab, rows:[[dept, zone, counter, rent, adjustment, payable, remarks]], total:[rent, adjustment, payable] | null }] } oldest first
function getRent() { guard_('rt'); return memo_('getRent', [], 120, () => getRent_c_()); }
function getRent_c_() {
  const ss = SpreadsheetApp.openById(RENT_ID), months = [];
  ss.getSheets().forEach(sh => {
    const nm = sh.getName().trim(), m = /^([A-Za-z]{3})[A-Za-z]*[ \-'’]*(\d{2,4})$/.exec(nm); if (!m) return;
    const mo = MON[m[1][0].toUpperCase() + m[1].slice(1).toLowerCase()]; if (mo === undefined) return;
    const y = +m[2] < 100 ? 2000 + +m[2] : +m[2], lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2 || lc < 3) return;
    const V = sh.getRange(1, 1, lr, lc).getValues(); let hr = -1, C = null;
    for (let r = 0; r < Math.min(8, V.length) && hr < 0; r++) {
      const H = V[r].map(h => String(h).toLowerCase().replace(/[^a-z]/g, '')), c = {};
      for (const k in RENT_COLS) { const j = H.findIndex(h => RENT_COLS[k].indexOf(h) >= 0); if (j >= 0) c[k] = j; }
      if (c.name !== undefined && c.rent !== undefined && c.pay !== undefined) { hr = r; C = c; }
    }
    if (hr < 0) return;
    const g = (row, k) => C[k] === undefined ? '' : row[C[k]], rows = []; let total = null;
    for (let r = hr + 1; r < V.length; r++) {
      const row = V[r], dept = txt_(g(row, 'dept')), name = txt_(g(row, 'name'));
      if (/^total$/i.test(dept) || /^total$/i.test(name)) { total = [rtNum_(g(row, 'rent')), rtNum_(g(row, 'adj')), rtNum_(g(row, 'pay'))]; continue; }
      if (!name) continue;
      rows.push([dept, txt_(g(row, 'zone')), name, rtNum_(g(row, 'rent')), rtNum_(g(row, 'adj')), rtNum_(g(row, 'pay')), txt_(g(row, 'rem'))]);
    }
    months.push({ m: y + '-' + ('0' + (mo + 1)).slice(-2), tab: nm, rows: rows, total: total });
  });
  months.sort((a, b) => a.m < b.m ? -1 : 1);
  return JSON.stringify({ months: months });
}

// ===== Daily cash (page "Daily cash") =====
// Source: the "Daily Cashflow" Google Sheet (tabs Cash Receiving, Expenses, Advances, Lists). Add rows there; the page works out the figures.
/**
 * Daily Cashflow data for the management dashboard.
 * Reads the "Daily Cashflow" Google Sheet directly (Cash Receiving, Expenses,
 * Advances, Lists) and works out the figures itself, so the dashboard never
 * changes the report date cells in the sheet.
 *
 * Call from the page:
 *   google.script.run.withSuccessHandler(render).getCashflowData('2026-11-09');
 * Leave the date out (or pass '') for today.
 */

const CASHFLOW_ID = '1V8oUmKKI8SnfU6B7X2oBLM-UtkMx9Kit-Ydeg0ks0k8';
const CF_TZ = 'Asia/Dhaka';
const CF_NOTES = [1000, 500, 200, 100, 50, 20, 10, 5, 2, 1];

function cfKey_(d) {
  if (d === '' || d === null || d === undefined) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, CF_TZ, 'yyyy-MM-dd');
  }
  const s = String(d).trim(), m = /^(\d{1,2})[ \-\/]([A-Za-z]{3})[A-Za-z]*[ \-\/,]+(\d{4})$/.exec(s);   // text such as "08-Nov-2026"
  if (m && MON[m[2][0].toUpperCase() + m[2].slice(1).toLowerCase()] !== undefined) return m[3] + '-' + ('0' + (MON[m[2][0].toUpperCase() + m[2].slice(1).toLowerCase()] + 1)).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s.slice(0, 10);
}

function cfNum_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function cfRows_(ss, name, lastCol) {
  const sh = ss.getSheetByName(name);
  const last = sh.getLastRow();
  if (last < 3) return [];
  return sh.getRange(3, 1, last - 2, lastCol).getValues()
    .filter(r => r[0] !== '' && r[0] !== null);
}

function cfSum_(arr, f) {
  return arr.reduce((s, x) => s + f(x), 0);
}

function cfGroup_(arr, key, val) {
  const m = {};
  arr.forEach(x => {
    const k = x[key] || '(blank)';
    m[k] = (m[k] || 0) + val(x);
  });
  return Object.keys(m)
    .map(k => ({ name: k, amount: m[k] }))
    .sort((a, b) => b.amount - a.amount);
}

// Expenses by head, split into approved (paid by the accounts) and from my hand
function cfSplit_(apxList, expList) {
  const m = {};
  const add = (h, k, v) => {
    const n = h || '(no head)';
    if (!m[n]) m[n] = { name: n, approved: 0, hand: 0, total: 0 };
    m[n][k] += v; m[n].total += v;
  };
  apxList.forEach(a => add(a.head, 'approved', a.amount));
  expList.forEach(e => add(e.head, 'hand', e.amount));
  return Object.keys(m).map(k => m[k]).sort((a, b) => b.total - a.total);
}

// Sales and running income split by brand for each account and day, from the Income Statement ALL tab (column Q "Brand").
// Mercedes and Royalty are shown together as "Aircon", Silky as "Silky"; any other brand name is shown as it is. Days without a brand entered give no split.
function cfBrands_() {
  const cache = CacheService.getScriptCache(), c = cache.get('cfbrands'); if (c) { try { return JSON.parse(c); } catch (x) {} }
  const out = {};
  try {
    const ss = SpreadsheetApp.openById(RT_ALL.id), sh = ss.getSheetByName(RT_ALL.tab), tz = ss.getSpreadsheetTimeZone(), lr = sh.getLastRow();
    if (lr > 1) {
      const n = Math.min(lr - 1, 2500), r0 = lr - n + 1, V = sh.getRange(r0, 1, n, 17).getValues();   // A Date ... G Total Sale ... O Income, P Account, Q Brand
      V.forEach(r => {
        const d = rtDay_(r[0], tz), a = rtAcct_(r[15]), b0 = String(r[16] == null ? '' : r[16]).trim(); if (d == null || !a || !b0) return;
        const bl = b0.toLowerCase(), g = (bl === 'mercedes' || bl === 'royalty' || bl === 'aircon' || bl === 'ac') ? 'Aircon' : (bl === 'silky' ? 'Silky' : b0);
        const k = calIso_(d) + '|' + a.toLowerCase() + '|' + g, o = out[k] || (out[k] = { s: 0, i: 0, n: 0 });
        o.s += rtNum_(r[6]); o.i += rtNum_(r[14]); o.n += 1;
      });
    }
  } catch (x) {}
  try { cache.put('cfbrands', JSON.stringify(out), 120); } catch (x) {}
  return out;
}
function cfBrandsFor_(bm, day, acc) {
  const pre = day + '|' + String(acc).trim().toLowerCase() + '|', r = [];
  Object.keys(bm).forEach(k => { if (k.indexOf(pre) === 0) r.push({ b: k.slice(pre.length), sales: bm[k].s, income: bm[k].i, n: bm[k].n }); });
  const ord = { Aircon: 0, Silky: 1 }; r.sort((x, y) => (ord[x.b] != null ? ord[x.b] : 9) - (ord[y.b] != null ? ord[y.b] : 9) || x.b.localeCompare(y.b));
  return r;
}

function getCashflowData(dateStr) { guard_('dy', 'ov', 'rp'); const o = getCashflowData_(dateStr), a = perm_().accounts; return a ? scopeCash_(o, a) : o; }
function getCashflowData_(dateStr) { return memo_('getCashflowData_', [dateStr], 120, () => getCashflowData__c_(dateStr)); }
function getCashflowData__c_(dateStr) {
  const ss = SpreadsheetApp.openById(CASHFLOW_ID);
  const day = dateStr || Utilities.formatDate(new Date(), CF_TZ, 'yyyy-MM-dd');
  const month = day.slice(0, 7);

  // Lists: accounts in A, approvers in C, opening cash in E2
  const lists = ss.getSheetByName('Lists').getRange('A2:E20').getValues();
  // My cash lag: Expenses / Advances typed on or after Lists!M2 belong to the previous
  // accounts day (accounts report day D, I handle that cash on D+1). Blank M2 = no lag.
  const lagFrom = cfKey_(ss.getSheetByName('Lists').getRange('M2').getValue());
  const lagDay_next = k => {
    if (!lagFrom) return k;
    const t = new Date(k + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() + 1);
    const n = t.toISOString().slice(0, 10);
    return n >= lagFrom ? n : k;
  };
  const lagDay = k => {
    if (!k || !lagFrom || k < lagFrom) return k;
    const t = new Date(k + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() - 1);
    return t.toISOString().slice(0, 10);
  };
  const accounts = lists.map(r => r[0]).filter(String);
  const openingBase = cfNum_(lists[0][4]);

  // Cash Receiving: A Date, B Account, C Sales, D Trips (shown only, never added
  // or deducted), E Income (after running expense), F Approved Expenses (sum of the
  // Approved Expenses tab), G Bank Deposit, H Online, I Cash Payable, J:S notes,
  // T Cash Received, U Paid to Bank (instead of cash), V Bank Account,
  // W Deposit Slip / Ref, X Difference, Y Remarks.
  // Paid to Bank settles the payable but never enters cash in hand.
  const rec = cfRows_(ss, 'Cash Receiving', 32).map(r => ({
    d: cfKey_(r[0]), acc: r[1],
    sales: cfNum_(r[2]), trips: cfNum_(r[3]), income: cfNum_(r[4]), appr: cfNum_(r[5]),
    adv: cfNum_(r[6]),
    bank: cfNum_(r[7]), online: cfNum_(r[8]), payable: cfNum_(r[9]),
    onl: [cfNum_(r[28]), cfNum_(r[29]), cfNum_(r[30]), cfNum_(r[31])],   // AC:AF BDtickets, Shohoz, Soudia Web, Jatri (part of Online)
    notes: r.slice(10, 20).map(cfNum_), received: cfNum_(r[20]),
    toBank: cfNum_(r[21]), toBankAcc: String(r[22] || ''),
    toBankRef: String(r[23] || ''), diff: cfNum_(r[24]),
    remarks: String(r[25] || ''),
    buses: cfNum_(r[26])   // AA Buses (typed): used for the income per bus figures
  }));

  // Approved Expenses: A Date, B Account, C Head, D Bus / Counter, E Details,
  // F Amount, G Approved By, H Remarks (the breakdown behind column E above)
  const apx = ss.getSheetByName('Approved Expenses') ? cfRows_(ss, 'Approved Expenses', 8).map(r => ({
    d: cfKey_(r[0]), acc: String(r[1]), head: String(r[2]), bus: String(r[3]),
    details: String(r[4]), amount: cfNum_(r[5]), approvedBy: String(r[6]),
    remarks: String(r[7])
  })) : [];

  // Expenses: A Date, B Voucher, C Head, D Details, E Paid To,
  // F Approved By, G:P notes, Q Amount, R Remarks
  const exp = cfRows_(ss, 'Expenses', 18).map(r => ({
    d: cfKey_(r[0]), voucher: String(r[1]), head: String(r[2]),
    details: String(r[3]), paidTo: String(r[4]), approvedBy: String(r[5]),
    notes: r.slice(6, 16).map(cfNum_), amount: cfNum_(r[16]) || r.slice(6, 16).reduce(function (s, v, i) { return s + cfNum_(v) * CF_NOTES[i]; }, 0),
    remarks: String(r[17])
  }));

  // Advances: A Date Given, B Voucher, C Given To, D Given From (an account, or
  // "Me" / blank = from my cash in hand), E Purpose, F Category, G Approved By,
  // H:Q notes given, R Amount Given, S Settlement Date, T Bills,
  // T:AC notes returned, AD Cash Returned
  // (a Category column sits after Purpose, so everything from Approved By moves one right)
  const advAll = cfRows_(ss, 'Advances', 34).map(r => ({
    d: cfKey_(r[0]), voucher: String(r[1]), to: String(r[2]),
    from: String(r[3] || '').trim(),
    purpose: String(r[4]), category: String(r[5] || ''), approvedBy: String(r[6]),
    notesOut: r.slice(7, 17).map(cfNum_),
    given: cfNum_(r[17]) || r.slice(7, 17).reduce(function (s, v, i) { return s + cfNum_(v) * CF_NOTES[i]; }, 0),
    sd: cfKey_(r[18]), bills: cfNum_(r[19]),
    notesIn: r.slice(20, 30).map(cfNum_),
    returned: cfNum_(r[30]) || r.slice(20, 30).reduce(function (s, v, i) { return s + cfNum_(v) * CF_NOTES[i]; }, 0),
    byMe: String(r[31] || '').trim().toLowerCase()   // AF "By Me?" Yes / No
  }));
  exp.forEach(e => { e.d = lagDay(e.d); });
  // only advances that go through me follow my cash day (D+1); advances given by an account keep the accounts date
  advAll.forEach(a => { if (a.byMe !== 'no' && a.byMe !== 'n') { a.d = lagDay(a.d); a.sd = lagDay(a.sd); } });
  // Returned notes / cash with no settlement date are counted on the date given
  advAll.forEach(a => { if (!a.sd && (a.returned !== 0 || a.bills > 0)) a.sd = a.d; });
  // Advance by me comes out of my cash in hand; an advance given from an
  // account is paid out of that account's collection (it lowers its cash payable)
  // Default = by me (always deducted from cash in hand); only an explicit "No" in AF excludes it
  const isMine = a => a.byMe !== 'no' && a.byMe !== 'n';
  const adv = advAll.filter(isMine);
  const advAcc = advAll.filter(a => !isMine(a));

  // Cash in hand up to a date (inclusive or before)
  const cashUpTo = (k, inclusive) => {
    const ok = d => d && (inclusive ? d <= k : d < k);
    return openingBase
      + cfSum_(rec.filter(r => ok(r.d)), r => r.received)
      - cfSum_(exp.filter(e => ok(e.d)), e => e.amount)
      - cfSum_(adv.filter(a => ok(a.d)), a => a.given)
      + cfSum_(adv.filter(a => ok(a.sd)), a => a.returned);
  };

  // Selected day
  const recDay = rec.filter(r => r.d === day);
  const expDay = exp.filter(e => e.d === day);
  const today = {
    opening: cashUpTo(day, false),
    received: cfSum_(recDay, r => r.received),
    advReturned: cfSum_(adv.filter(a => a.sd === day), a => Math.max(a.returned, 0)),
    advExtra: cfSum_(adv.filter(a => a.sd === day), a => Math.max(-a.returned, 0)),   // a negative Cash Returned = extra cash paid to settle the advance
    advGiven: cfSum_(adv.filter(a => a.d === day), a => a.given),
    advFromAccounts: cfSum_(advAcc.filter(a => a.d === day), a => a.given),
    expenses: cfSum_(expDay, e => e.amount),
    toBank: cfSum_(recDay, r => r.toBank)
  };
  today.closing = today.opening + today.received + today.advReturned
    - today.advExtra - today.advGiven - today.expenses;

  const brandMap = cfBrands_();
  const byAccount = accounts.map(acc => {
    const rs = recDay.filter(r => r.acc === acc);
    return {
      account: acc, brands: cfBrandsFor_(brandMap, day, acc),
      sales: cfSum_(rs, r => r.sales), trips: cfSum_(rs, r => r.trips), income: cfSum_(rs, r => r.income),
      approved: cfSum_(rs, r => r.appr), advances: cfSum_(rs, r => r.adv), bank: cfSum_(rs, r => r.bank),
      online: cfSum_(rs, r => r.online), onl: [0, 1, 2, 3].map(i => cfSum_(rs, r => r.onl[i])), payable: cfSum_(rs, r => r.payable),
      received: cfSum_(rs, r => r.received), toBank: cfSum_(rs, r => r.toBank),
      diff: cfSum_(rs, r => r.diff), buses: cfSum_(rs, r => r.buses)
    };
  });

  // Notes in hand as of the selected day
  const upto = d => d && d <= day;
  // Where a row has an amount but no note counts (typed amount only), the notes are
  // worked out automatically: opening cash and typed returns are split into notes
  // (largest first); typed expenses / advances are paid from the notes in hand
  // (largest first), breaking a bigger note for change when needed.
  const nsum = a => a.reduce((s, v) => s + v, 0);
  const pile = CF_NOTES.map((n, i) =>
    cfSum_(rec.filter(r => upto(r.d)), r => r.notes[i])
    - cfSum_(exp.filter(e => upto(e.d)), e => e.notes[i])
    - cfSum_(adv.filter(a => upto(a.d)), a => a.notesOut[i])
    + cfSum_(adv.filter(a => upto(a.sd)), a => a.notesIn[i]));
  const addAmt = amt => { CF_NOTES.forEach((n, i) => { const k = Math.floor(amt / n); pile[i] += k; amt -= k * n; }); };
  const payAmt = amt => {
    CF_NOTES.forEach((n, i) => { const k = Math.min(Math.floor(amt / n), Math.max(pile[i], 0)); pile[i] -= k; amt -= k * n; });
    if (amt > 0) {
      for (let i = CF_NOTES.length - 1; i >= 0; i--) {
        if (CF_NOTES[i] > amt && pile[i] > 0) { pile[i] -= 1; addAmt(CF_NOTES[i] - amt); amt = 0; break; }
      }
    }
    if (amt > 0) { pile[CF_NOTES.length - 1] -= amt; }   // nothing left to break: shows as a shortfall
  };
  addAmt(openingBase);
  exp.filter(e => upto(e.d) && nsum(e.notes) === 0 && e.amount > 0).forEach(e => payAmt(e.amount));
  adv.filter(a => upto(a.d) && nsum(a.notesOut) === 0 && a.given > 0).forEach(a => payAmt(a.given));
  adv.filter(a => upto(a.sd) && nsum(a.notesIn) === 0 && a.returned !== 0).forEach(a => { if (a.returned > 0) addAmt(a.returned); else payAmt(-a.returned); });
  const notes = CF_NOTES.map((n, i) => ({ note: n, count: pile[i], amount: pile[i] * n }));

  // Advances still open as of the selected day
  const outstanding = advAll.filter(a => upto(a.d)).map(a => {
    const settled = upto(a.sd);
    return {
      date: a.d, voucher: a.voucher, to: a.to, from: a.from || 'Me', purpose: a.purpose, category: a.category,
      given: a.given,
      outstanding: a.given - (settled ? a.bills + a.returned : 0)
    };
  }).filter(a => a.outstanding > 0);

  // Month to date (1st of the month up to the selected day)
  const inMonth = d => d && d.slice(0, 7) === month && d <= day;
  const mRec = rec.filter(r => inMonth(r.d));
  const mExp = exp.filter(e => inMonth(e.d));
  const mApx = apx.filter(a => inMonth(a.d));
  const apxDay = apx.filter(a => a.d === day);
  const daily = [];
  for (let i = 1; i <= Number(day.slice(8, 10)); i++) {
    const k = month + '-' + String(i).padStart(2, '0');
    daily.push({
      date: k,
      received: cfSum_(rec.filter(r => r.d === k), r => r.received),
      expenses: cfSum_(exp.filter(e => e.d === k), e => e.amount),
      approved: cfSum_(apx.filter(a => a.d === k), a => a.amount),
      sales: cfSum_(rec.filter(r => r.d === k), r => r.sales),
      closing: cashUpTo(k, true)
    });
  }

  // Date my cash is actually handled: the day after the accounts day once the lag is on
  const cashDate = (() => { const n = lagDay_next(day); return n; })();
  return {
    date: day,
    cashDate: cashDate,
    today: today,
    byAccount: byAccount,
    approvedExpenses: apx.filter(a => a.d === day).map(a => ({
      account: a.acc, head: a.head, bus: a.bus, details: a.details,
      amount: a.amount, approvedBy: a.approvedBy, remarks: a.remarks
    })),
    advDetails: advAll.filter(a => a.d === day).map(a => ({ account: isMine(a) ? 'Me' : a.from, to: a.to, purpose: a.purpose, category: a.category, amount: a.given })),
    // advances closed / returned on this day for an account: they give cash back (lower the advance expense, raise the cash payable)
    advRetDetails: advAll.filter(a => a.sd === day && a.returned !== 0 && !isMine(a)).map(a => ({ account: a.from, to: a.to, purpose: a.purpose, category: a.category, amount: a.returned, given: a.given, dateGiven: a.d })),
    paidToBank: recDay.filter(r => r.toBank).map(r => ({
      account: r.acc, amount: r.toBank, bank: r.toBankAcc, ref: r.toBankRef
    })),
    expenses: expDay.map(e => ({
      voucher: e.voucher, head: e.head, details: e.details,
      paidTo: e.paidTo, approvedBy: e.approvedBy,
      amount: e.amount, remarks: e.remarks
    })),
    expensesTotal: today.expenses,
    approvedTotal: cfSum_(apxDay, a => a.amount),
    expensesByHead: cfSplit_(apxDay, expDay),
    collection: {
      sales: cfSum_(recDay, r => r.sales), trips: cfSum_(recDay, r => r.trips), income: cfSum_(recDay, r => r.income),
      approved: cfSum_(recDay, r => r.appr), advances: cfSum_(recDay, r => r.adv), bank: cfSum_(recDay, r => r.bank),
      online: cfSum_(recDay, r => r.online), onl: [0, 1, 2, 3].map(i => cfSum_(recDay, r => r.onl[i])), payable: cfSum_(recDay, r => r.payable),
      received: cfSum_(recDay, r => r.received), toBank: cfSum_(recDay, r => r.toBank),
      diff: cfSum_(recDay, r => r.diff)
    },
    notes: notes,
    notesTotal: cfSum_(notes, n => n.amount),
    outstanding: outstanding,
    outstandingTotal: cfSum_(outstanding, a => a.outstanding),
    month: {
      label: month,
      received: cfSum_(mRec, r => r.received),
      payable: cfSum_(mRec, r => r.payable),
      toBank: cfSum_(mRec, r => r.toBank),
      sales: cfSum_(mRec, r => r.sales), trips: cfSum_(mRec, r => r.trips), income: cfSum_(mRec, r => r.income),
      online: cfSum_(mRec, r => r.online), onl: [0, 1, 2, 3].map(i => cfSum_(mRec, r => r.onl[i])),
      approved: cfSum_(mRec, r => r.appr),
      advances: cfSum_(mRec, r => r.adv),
      advancesByMe: cfSum_(adv.filter(a => inMonth(a.d)), a => a.given),
      approvedByHead: cfGroup_(apx.filter(a => inMonth(a.d)), 'head', a => a.amount),
      shortExcess: cfSum_(mRec, r => r.diff),
      expenses: cfSum_(mExp, e => e.amount),
      byHead: cfGroup_(mExp, 'head', e => e.amount),
      byApprover: cfGroup_(mExp.map(e => ({ who: e.approvedBy, v: e.amount }))
        .concat(mApx.map(a => ({ who: a.approvedBy, v: a.amount }))), 'who', x => x.v),
      approvedExp: cfSum_(mApx, a => a.amount),
      expenseHeads: cfSplit_(mApx, mExp),
      apxRows: mApx.map(a => ({ date: a.d, account: a.acc, head: a.head, bus: a.bus, details: a.details, amount: a.amount, approvedBy: a.approvedBy })),
      expRows: mExp.map(e => ({ date: e.d, voucher: e.voucher, head: e.head, details: e.details, paidTo: e.paidTo, amount: e.amount, approvedBy: e.approvedBy })),
      bank: cfSum_(mRec, r => r.bank),
      daily: daily
    }
  };
}

/** Run this once from the editor to check the numbers and grant access. */
function testCashflowData() {
  const d = getCashflowData_('2026-11-09');
  Logger.log(JSON.stringify(d.today));      // closing should be 87650
  Logger.log('Notes total: ' + d.notesTotal); // should match closing
}

// ===== Salary (page "Payroll") =====
// Source: the "MM Motors Salarysheet" Google Sheet, tabs Salary Register, Garage Register and Ticket Sales. HR adds the new month's rows to those tabs; the page reads them as they are.
// Only what HR entered is shown (sales included). Nothing is compared with the dashboard's own sales.
const SALARY_ID = '1l5UOWptKXHqqzPe1WxPgPbOziuKhZzTdGp5bp-NvNkg';
const sHd_ = v => String(v == null ? '' : v).toLowerCase().replace(/[^a-z]/g, '');
const SAL_COLS = {
  rec: 'rec', month: 'month', unit: 'unit', zone: 'departmentzone', sec: 'sectioncounterasinsource', eid: 'employeeid', name: 'nameofemployee', desig: 'designation',
  days: 'daysbasismonthdays', gross: 'monthlygrosssalary', abs: 'totalabsent', dabs: 'deductionforabsent', dadv: 'deductionforadvancetax', dfood: 'deductionforfooding',
  ded: 'totaldeduction', pay: 'payablesalary', arr: 'arrear', net: 'netpayablesalary', idn: 'idnumbersource', tk: 'totalsoldtickets', tkd: 'soldticketsperday',
  disc: 'totaldiscountamount', comm: 'totalcommissionamount', rem: 'remarkssource'
};
const SAL_GCOLS = {
  month: 'month', garage: 'garagesoudiashop', grp: 'sectiongroup', eid: 'empid', name: 'name', desig: 'designationasprinted', gross: 'grosspay', days: 'daysworked',
  tot: 'totalsalary', oth: 'overtimepay', pay: 'totalpayable', dabs: 'deductionabsent', dadv: 'deductionadv', doth: 'deductionother', ded: 'totaldeduction', net: 'netpayable', bank: 'bank', cash: 'cash'
};
function salMonth_(v) {                       // "Sep 2026" -> "2026-09"
  const m = /^([A-Za-z]{3})[A-Za-z]*[ \-']*(\d{2,4})$/.exec(txt_(v)); if (!m) return '';
  const mo = MON[m[1][0].toUpperCase() + m[1].slice(1).toLowerCase()]; if (mo === undefined) return '';
  const y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; return y + '-' + ('0' + (mo + 1)).slice(-2);
}
function salSec_(s) {                         // tidy the counter name typed by HR: drop phone numbers and trailing colons, fix common typos
  let t = txt_(s).replace(/\(?\b0?1\d{9,11}(\s*\/\s*\d+)?\)?/g, ' ').replace(/[;:\-,\s]+$/, '').replace(/\bcountre\b/ig, 'Counter').replace(/^ECURITY/, 'Security').replace(/^Machanical/, 'Mechanical');
  return txt_(t.replace(/\s+([:;])/g, '$1'));
}
function salKind_(zone, sec) {                // cc = Call Center, ctr = counter, chk = checker, off = office and support staff
  const z = txt_(zone).toLowerCase(), s = txt_(sec).toLowerCase();
  if (z === 'call center') return 'cc';
  if (z === 'checker') return 'chk';
  if (/office|account|store|driver|ecurity|mechanical|machanical|coaster|staff/.test(s) || /^abdullahpur$/.test(s)) return 'off';
  return 'ctr';
}
function salGrid_(sh) {
  if (!sh) return null;
  const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return null;
  const V = sh.getRange(1, 1, lr, lc).getValues(); return { H: V[0].map(sHd_), V: V.slice(1) };
}
function salCol_(H, map) { const c = {}; for (const k in map) { const j = H.indexOf(map[k]); if (j >= 0) c[k] = j; } return c; }
const salN_ = v => { const n = Number(v); return isNaN(n) || v === '' || v === null ? 0 : n; };

function getSalaryData() { guard_('pr'); return memo_('getSalaryData', [], 120, () => getSalaryData_c_()); }
function getSalaryData_c_() {
  const ss = SpreadsheetApp.openById(SALARY_ID), reg = salGrid_(ss.getSheetByName('Salary Register')), gar = salGrid_(ss.getSheetByName('Garage Register')), tks = salGrid_(ss.getSheetByName('Ticket Sales'));
  if (!reg) throw new Error('The Salary Register tab was not found in the Salarysheet.');
  const C = salCol_(reg.H, SAL_COLS);
  ['rec', 'month', 'unit', 'zone', 'name', 'net'].forEach(k => { if (C[k] === undefined) throw new Error('Salary Register is missing the heading for "' + k + '".'); });
  // HR sales: the sales column that belongs to the salary month itself (for a Sep 2026 sheet, the "Sep 26" column)
  const sales = {};
  if (tks) {
    const T = salCol_(tks.H, { rec: 'recsalaryregister', month: 'salarymonthfile', per: 'salesperiodcolumnheadinginsource', amt: 'totalsaleamount' });
    if (T.rec !== undefined && T.per !== undefined && T.amt !== undefined) for (const r of tks.V) {
      const mk = salMonth_(r[T.month]), pk = salMonth_(String(r[T.per]).replace(/^(\w{3})\s*(\d{2})$/, '$1 $2'));
      if (mk && pk === mk && r[T.rec] !== '') sales[r[T.rec]] = (sales[r[T.rec]] || 0) + salN_(r[T.amt]);
    }
  }
  const g = (r, k) => C[k] === undefined ? '' : r[C[k]], M = {};
  const mo = (m, lab) => M[m] || (M[m] = { m: m, lab: lab, r: [], g: [] });
  for (const r of reg.V) {
    const mk = salMonth_(g(r, 'month')); if (!mk || !txt_(g(r, 'name'))) continue;
    const zone = txt_(g(r, 'zone')), secRaw = txt_(g(r, 'sec')), kind = salKind_(zone, secRaw), rec = g(r, 'rec');
    const sec = salSec_(secRaw) || (kind === 'cc' ? 'Call Center' : '(no counter name)');
    const sl = sales[rec];
    mo(mk, txt_(g(r, 'month'))).r.push([txt_(g(r, 'unit')), zone, sec, txt_(g(r, 'eid')), txt_(g(r, 'name')), txt_(g(r, 'desig')), salN_(g(r, 'days')), salN_(g(r, 'gross')), salN_(g(r, 'abs')),
      salN_(g(r, 'dabs')), salN_(g(r, 'dadv')), salN_(g(r, 'dfood')), salN_(g(r, 'ded')), salN_(g(r, 'pay')), salN_(g(r, 'arr')), salN_(g(r, 'net')), txt_(g(r, 'idn')),
      kind === 'ctr' || kind === 'cc' ? (sl === undefined ? null : Math.round(sl * 100) / 100) : null, g(r, 'tk') === '' ? null : salN_(g(r, 'tk')), kind, txt_(g(r, 'rem'))]);
  }
  if (gar) {
    const G = salCol_(gar.H, SAL_GCOLS);
    if (G.month !== undefined && G.garage !== undefined && G.name !== undefined && G.net !== undefined) for (const r of gar.V) {
      const mk = salMonth_(r[G.month]); if (!mk || !txt_(r[G.name])) continue;
      const v = k => G[k] === undefined ? '' : r[G[k]];
      mo(mk, txt_(r[G.month])).g.push([txt_(v('garage')), txt_(v('grp')), txt_(v('eid')), txt_(v('name')), txt_(v('desig')), salN_(v('days')), salN_(v('gross')), salN_(v('tot')), salN_(v('oth')), salN_(v('pay')),
        salN_(v('dabs')), salN_(v('dadv')), salN_(v('doth')), salN_(v('ded')), salN_(v('net')), salN_(v('bank')), salN_(v('cash'))]);
    }
  }
  const months = Object.keys(M).sort().map(k => M[k]);
  return JSON.stringify({ months: months });
}


// Incident Reporting Portal: tab Reported (one row per incident) and tab Followup (one row per Incident ID).
// Returns { r:[{id, d, time, type, reason, sev, bus, loc, route, driver, sup, inj, injd, desc, drs, sus, act, auth, est, rem, inc, sub, fu}] }
// d is the day number (days since 1970), bus is the same bus key the Monitoring data uses, fu is the follow-up row or null.
// Columns are found by their header names, so adding or moving a column does not break the page.
function getIncidents() { guard_('ic', 'ov'); return memo_('getIncidents', [], 120, () => getIncidents_c_()); }
function getIncidents_c_() {
  const ss = SpreadsheetApp.openById(INCIDENT_ID), tz = ss.getSpreadsheetTimeZone();
  const grid = n => { const sh = ss.getSheetByName(n); if (!sh || sh.getLastRow() < 2) return { h: [], v: [] }; const lc = sh.getLastColumn(); const a = sh.getRange(1, 1, sh.getLastRow(), lc).getValues(); return { h: a[0].map(x => txt_(x).toLowerCase()), v: a.slice(1) }; };
  const ix = (h, ...res) => { for (const re of res) { const i = h.findIndex(x => re.test(x)); if (i >= 0) return i; } return -1; };
  const num = v => { const n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, '')); return isNaN(n) ? 0 : n; };
  const clock = v => v instanceof Date ? Utilities.formatDate(v, tz, 'h:mm a') : txt_(v);
  const stamp = v => v instanceof Date ? Utilities.formatDate(v, tz, 'd MMM yyyy h:mm a') : txt_(v);
  const R = grid('Reported'), F = grid('Followup');
  const c = { id: ix(R.h, /^id$/, /^incident id/), d: ix(R.h, /^date$/, /^date/), t: ix(R.h, /^time/), ty: ix(R.h, /^type/), re: ix(R.h, /^reason/), sv: ix(R.h, /^severity/), bus: ix(R.h, /^bus/), lo: ix(R.h, /^location/), ro: ix(R.h, /^route/), dr: ix(R.h, /^driver$/, /^driver/), su: ix(R.h, /^supervisor$/, /^supervisor/), inj: ix(R.h, /^any injury/, /^injur/), ind: ix(R.h, /^injury details/), de: ix(R.h, /^incident description/, /^description/), ds: ix(R.h, /^driver.s statement/), ss: ix(R.h, /^supervisor.s statement/), ac: ix(R.h, /^immediate action/), au: ix(R.h, /^action authorised/, /^action authorized/), es: ix(R.h, /^estimated repair/), rm: ix(R.h, /^in-charge remarks/, /remarks/), ic: ix(R.h, /^in-charge name/), sb: ix(R.h, /^report submission/) };
  const g = (r, k) => c[k] >= 0 ? r[c[k]] : '';
  const f = { id: ix(F.h, /^incident id/, /^id$/), fi: ix(F.h, /^investigation/), rc: ix(F.h, /^confirmed root/), ac: ix(F.h, /^actual repair/), co: ix(F.h, /^repair \/ corrective/, /^corrective/), rp: ix(F.h, /^action responsible/), ap: ix(F.h, /^management approval/), rd: ix(F.h, /^resolution date/), st: ix(F.h, /^case status/), ev: ix(F.h, /^evidence/), rep: ix(F.h, /^repeat/) };
  const fu = {};
  for (const r of F.v) { const id = txt_(f.id >= 0 ? r[f.id] : ''); if (!id) continue; const gf = k => f[k] >= 0 ? r[f[k]] : ''; fu[id] = { findings: txt_(gf('fi')), cause: txt_(gf('rc')), actual: num(gf('ac')), action: txt_(gf('co')), who: txt_(gf('rp')), approval: txt_(gf('ap')), resolved: gf('rd') ? day_(gf('rd'), tz) : null, status: txt_(gf('st')), evidence: txt_(gf('ev')), repeat: txt_(gf('rep')) }; }
  const out = [];
  for (const r of R.v) {
    const id = txt_(g(r, 'id')); if (!id && !txt_(g(r, 'bus')) && !txt_(g(r, 'de'))) continue;
    const d = day_(g(r, 'd'), tz);
    out.push({ id: id, d: d, time: clock(g(r, 't')), type: txt_(g(r, 'ty')), reason: txt_(g(r, 're')), sev: txt_(g(r, 'sv')), bus: busKey_(g(r, 'bus')), busRaw: txt_(g(r, 'bus')), loc: txt_(g(r, 'lo')), route: txt_(g(r, 'ro')), driver: txt_(g(r, 'dr')), sup: txt_(g(r, 'su')), inj: txt_(g(r, 'inj')), injd: txt_(g(r, 'ind')), desc: txt_(g(r, 'de')), drs: txt_(g(r, 'ds')), sus: txt_(g(r, 'ss')), act: txt_(g(r, 'ac')), auth: txt_(g(r, 'au')), est: num(g(r, 'es')), rem: txt_(g(r, 'rm')), inc: txt_(g(r, 'ic')), sub: stamp(g(r, 'sb')), fu: fu[id] || null });
  }
  return JSON.stringify({ r: out });
}
// ---------- Data Entry page, Fixed costs: "Counter Fixed Costs" Google Sheet (tab Entry, Lists, Counters) ----------
// Entry: headings in row 4, row 5 is a grey example line (ignored), real lines start at row 6. A:L are typed, M:N are the sheet's own formulas (never written here).
const FIXED_ID = '1P58XdF1BuzFxYEThATePL1AAkjTjfbWf6YWTSfbb2rQ';
// Counter list for the Fixed costs form: our own counters first (tab Counters), then every other counter name in the Masterdata tab (all counters of the booking system), no repeats.
function fxCounters_(ss) {
  const seen = {}, out = [], add = v => { const t = txt_(v), k = t.toLowerCase(); if (t && !seen[k]) { seen[k] = 1; out.push(t); } };
  const C = ss.getSheetByName('Counters'), M = ss.getSheetByName('Masterdata');
  if (C && C.getLastRow() > 1) C.getRange(2, 1, C.getLastRow() - 1, 1).getValues().forEach(r => add(r[0]));
  const own = out.length, rest = [];
  if (M && M.getLastRow() > 1) M.getRange(2, 1, M.getLastRow() - 1, 2).getValues().forEach(r => { const t = txt_(r[0]), k = t.toLowerCase(); if (t && txt_(r[1]) !== 'System' && !seen[k]) { seen[k] = 1; rest.push(t); } });   // 'System' rows (Shohoz, Jatri, Call Center ...) are sales channels, not counters
  rest.sort((a, b) => a.localeCompare(b));
  return out.concat(rest);
}
function getFixedForm() { guard_('en');
  const ss = SpreadsheetApp.openById(FIXED_ID), L = ss.getSheetByName('Lists'), C = ss.getSheetByName('Counters'), E = ss.getSheetByName('Entry');
  const cn = fxCounters_(ss);
  const n = Math.max(0, E.getLastRow() - 5), rows = n ? E.getRange(6, 1, n, 14).getDisplayValues() : [];
  const out = []; rows.forEach((r, i) => { if (txt_(r[1]) || txt_(r[2])) out.push([i + 6].concat(r.map(txt_))); });
  return JSON.stringify({ can: isEditor_(), counters: cn, heads: cfList_(L, 'A2:A40'), freqs: cfList_(L, 'C2:C12'), from: cfList_(L, 'F2:F20'), status: cfList_(L, 'H2:H12'), rows: out.slice(-40).reverse() });
}
// q = { day, counter, head, details, amount, freq, to, from, since, status, told, notes }  ->  'added'
function addFixedCost(json) { try { return addFixedCost_w_(json); } finally { bump_(); } }
function addFixedCost_w_(json) { guard_('en');
  return withLock_(() => {
    const q = JSON.parse(json), ss = SpreadsheetApp.openById(FIXED_ID), L = ss.getSheetByName('Lists'), C = ss.getSheetByName('Counters'), E = ss.getSheetByName('Entry');
    const day = dayOk_(q.day), counter = txt_(q.counter), head = txt_(q.head), freq = txt_(q.freq), from = txt_(q.from), status = txt_(q.status);
    const counters = fxCounters_(ss);
    if (counters.indexOf(counter) < 0) throw new Error('Choose the counter from the list.');
    if (cfList_(L, 'A2:A40').indexOf(head) < 0) throw new Error('Choose the cost head.');
    if (freq && cfList_(L, 'C2:C12').indexOf(freq) < 0) throw new Error('Choose how often it is paid from the list.');
    if (from && cfList_(L, 'F2:F20').indexOf(from) < 0) throw new Error('Choose the paid-from option from the list.');
    if (status && cfList_(L, 'H2:H12').indexOf(status) < 0) throw new Error('Choose the status from the list.');
    const raw = String(q.amount === undefined ? '' : q.amount).replace(/,/g, '').trim(), amount = raw === '' ? '' : +raw;
    if (amount !== '' && (isNaN(amount) || amount < 0)) throw new Error('Amount must be a number, zero or above.');
    if (amount !== '' && !freq) throw new Error('Choose how often it is paid, so the monthly and daily figures can be worked out.');
    let since = ''; if (txt_(q.since)) { const t = Date.parse(txt_(q.since) + ' UTC'); if (isNaN(t)) throw new Error('Check the Since date.'); since = Utilities.formatDate(new Date(t), 'UTC', 'yyyy-MM-dd'); }
    const mx = E.getMaxRows(), v = E.getRange(6, 2, Math.max(1, mx - 5), 2).getValues(); let last = 5;
    for (let i = 0; i < v.length; i++) if (txt_(v[i][0]) || txt_(v[i][1])) last = i + 6;
    const r = last + 1; growTo_(E, r);
    E.getRange(r, 1, 1, 12).setValues([[dstr_(day), counter, head, txt_(q.details), amount, freq, txt_(q.to), from, since, status, txt_(q.told), txt_(q.notes)]]);
    E.getRange(r, 1).setNumberFormat('dd MMM yyyy'); E.getRange(r, 9).setNumberFormat('dd MMM yyyy'); E.getRange(r, 5).setNumberFormat('#,##0');
    SpreadsheetApp.flush(); return 'added';
  });
}

// ---------- Daily cash: Cash Statement (bank style: Out | In | Balance), the whole history in one call ----------
// Same day logic as getCashflowData: accounts day for receipts, my cash day (lagged) for expenses and the advances I give.
// Returns { opening, min, rows:[ [day 'yyyy-MM-dd', particulars, out, in, approvedBy, balanceAfter] ] }
function getCashStatement() { guard_('dy'); if (perm_().accounts) throw new Error('The company cash ledger is not available for account-limited access.'); return memo_('getCashStatement', [], 120, () => getCashStatement_c_()); }
function getCashStatement_c_() {
  const ss = SpreadsheetApp.openById(CASHFLOW_ID), L = ss.getSheetByName('Lists');
  const lagFrom = cfKey_(L.getRange('M2').getValue()), opening = cfNum_(L.getRange('E2').getValue());
  const lagDay = k => { if (!k || !lagFrom || k < lagFrom) return k; const t = new Date(k + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() - 1); return t.toISOString().slice(0, 10); };
  const sh = d0 => { const p = d0.split('-'); return p[2] + '-' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+p[1] - 1]; };
  const join = (a, sep) => a.filter(s => s && String(s).trim()).join(sep);
  const ev = [];   // [day, order, text, out, in, by]
  cfRows_(ss, 'Cash Receiving', 32).forEach(r => { const v = cfNum_(r[20]); if (v) ev.push([cfKey_(r[0]), 0, 'Cash received from ' + txt_(r[1]) + ' account', 0, v, '']); });
  cfRows_(ss, 'Expenses', 18).forEach(r => {
    const d0 = cfKey_(r[0]), d = lagDay(d0), amt = cfNum_(r[16]) || r.slice(6, 16).reduce((t, v, i) => t + cfNum_(v) * CF_NOTES[i], 0); if (!amt) return;
    ev.push([d, 1, 'Expense: ' + join([join([txt_(r[2]), txt_(r[3])], ', '), txt_(r[4])], ' – ') + (d0 !== d ? ' (entered ' + sh(d0) + ')' : ''), amt, 0, txt_(r[5])]);
  });
  cfRows_(ss, 'Advances', 34).forEach(r => {
    const by = String(r[31] || '').trim().toLowerCase(); if (by === 'no' || by === 'n') return;   // advances given by an account never touch my cash
    const d0 = cfKey_(r[0]), d = lagDay(d0), s0 = cfKey_(r[18]), amt = cfNum_(r[17]) || r.slice(7, 17).reduce((t, v, i) => t + cfNum_(v) * CF_NOTES[i], 0), ret = cfNum_(r[30]) || r.slice(20, 30).reduce((t, v, i) => t + cfNum_(v) * CF_NOTES[i], 0), who = join([txt_(r[2]), txt_(r[4])], ' – ');
    if (amt) ev.push([d, 2, 'Advance given: ' + who + (d0 !== d ? ' (entered ' + sh(d0) + ')' : ''), amt, 0, txt_(r[6])]);
    if (ret) ev.push([lagDay(s0 || d0), 3, (ret > 0 ? 'Advance cash returned: ' : 'Extra paid on advance settlement: ') + who, ret < 0 ? -ret : 0, ret > 0 ? ret : 0, '']);
  });
  ev.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]);
  let bal = opening; const rows = ev.filter(e => e[0]).map(e => { bal += e[4] - e[3]; return [e[0], e[2], e[3], e[4], e[5], bal]; });
  return JSON.stringify({ opening: opening, min: rows.length ? rows[0][0] : '', rows: rows });
}

// ---------- Calendar page: reminders, schedule and daily routine ("Soudia Calendar" Google Sheet) ----------
// Tabs: Schedule (ID, Title, Date, Time, Type, Repeat, Notes, Status, Created by) · Routine (ID, Task, Time, Sun..Sat = Y, Active) · Done (Date, Item ID, Done at, Done by)
// Only the two editor accounts can read or change it.
const CAL_ID = '1Fi-sFfDbTU8PtDsXCQJz_WUzufxkyveZMkmrOMBWPjc';
function calSh_(n) { const s = SpreadsheetApp.openById(CAL_ID).getSheetByName(n); if (!s) throw new Error('Tab "' + n + '" not found in the Soudia Calendar sheet'); return s; }
function calRows_(n, w) { const s = calSh_(n), r = s.getLastRow(); return r > 1 ? s.getRange(2, 1, r - 1, w).getDisplayValues().map((x, i) => [i + 2].concat(x.map(txt_))) : []; }
const CAL_CHECK_FROM = '2026-11-15', CAL_ALERT_TO = ['rh.saudiabus@gmail.com'], CAL_ACCTS = ['MB-Lux', 'LR', 'CDC', 'AC', 'Aero', 'Teknaf'];
const calIso_ = n => new Date(n * 864e5).toISOString().slice(0, 10);
const calToday_ = () => Math.floor((Date.now() + 216e5) / 864e5);      // today in Bangladesh time, as a day number
// What the other sheets already hold, kept for 90 seconds so the Calendar and the Overview strip stay quick.
function calHave_() {
  const cache = CacheService.getScriptCache(), c = cache.get('calhave'); if (c) { try { return JSON.parse(c); } catch (x) {} }
  const today = calToday_(), from = today - 60, have = { stmt: {}, cash: {}, apx: {}, bank: {}, exp: {}, ps: [], err: {} };
  try {
    const ss = SpreadsheetApp.openById(RT_ALL.id), sh = ss.getSheetByName(RT_ALL.tab), tz = ss.getSpreadsheetTimeZone(), lr = sh.getLastRow();
    if (lr > 1) { const n = Math.min(lr - 1, 3500), r0 = lr - n + 1, A = sh.getRange(r0, 1, n, 1).getValues(), F = sh.getRange(r0, 7, n, 1).getValues(), O = sh.getRange(r0, 16, n, 1).getValues();   // G = Total Sale, P = Account (Trip is column D)
      for (let i = 0; i < n; i++) { const d = rtDay_(A[i][0], tz), a = rtAcct_(O[i][0]); if (d == null || d < from || !a) continue; const k = calIso_(d) + '|' + a.toLowerCase(); have.stmt[k] = (have.stmt[k] || 0) + rtNum_(F[i][0]); } }
  } catch (x) { have.err.stmt = String(x.message || x); }
  try { cfRows_(SpreadsheetApp.openById(CASHFLOW_ID), 'Cash Receiving', 9).forEach(r => { const k = cfKey_(r[0]), a = String(r[1] || '').trim().toLowerCase(); if (k && a) have.cash[k + '|' + a] = (have.cash[k + '|' + a] || 0) + cfNum_(r[2]); }); } catch (x) { have.err.cash = String(x.message || x); }
  try { cfRows_(SpreadsheetApp.openById(CASHFLOW_ID), 'Approved Expenses', 6).forEach(r => { const k = cfKey_(r[0]), a = String(r[1] || '').trim().toLowerCase(); if (k && a) have.apx[k + '|' + a] = (have.apx[k + '|' + a] || 0) + cfNum_(r[5]); }); } catch (x) { have.err.apx = String(x.message || x); }
  try { const ss = SpreadsheetApp.openById(MGMT_ID); ledgerRows_(ss, ss.getSpreadsheetTimeZone()).forEach(r => { if (r.day >= from) have.bank[calIso_(r.day)] = 1; }); } catch (x) { have.err.bank = String(x.message || x); }
  try { const ss = SpreadsheetApp.openById(CASHFLOW_ID); cfRows_(ss, 'Expenses', 18).forEach(r => { const k = cfKey_(r[0]); if (k) have.exp[k] = 1; }); cfRows_(ss, 'Advances', 34).forEach(r => { const k = cfKey_(r[0]); if (k) have.exp[k] = 1; }); } catch (x) { have.err.exp = String(x.message || x); }
  try { JSON.parse(getPS_()).r.forEach(r => { if (r[0] >= today - 45 && r[0] <= today + 30) have.ps.push([calIso_(r[0]), r[1], r[2], r[4], r[5], r[8]]); }); } catch (x) { have.err.ps = String(x.message || x); }
  try { cache.put('calhave', JSON.stringify(have), 90); } catch (x) {}
  return have;
}
// A routine task can complete itself. Auto = "stmt|AC|1" (income statement rows for AC exist in the Income Statement ALL tab), "cash|LR|1" (Cash Receiving row for LR),
// "bank|Any|0" (any Bank Ledger entry), "exp|Any|0" (any expense or advance in Daily Cashflow). The last number is how many days earlier the data is dated.
function calAutoFor_(ro, have) {
  const out = {}, today = calToday_();
  ro.filter(r => /^(stmt|cash|bank|exp)\|/.test(r[4])).forEach(r => { const p = r[4].split('|'), acc = (p[1] || '').toLowerCase(), lag = +p[2] || 0, ok = [];
    for (let d = today - 45; d <= today + 1; d++) { const src = calIso_(d - lag); if ((p[0] === 'bank' || p[0] === 'exp') ? have[p[0]][src] : have[p[0]][src + '|' + acc] != null) ok.push(calIso_(d)); }
    out[r[0]] = ok; });
  return out;
}
// Differences between the income statement and Cash Receiving, for the last 7 days up to yesterday (from CAL_CHECK_FROM on).
function calChecks_(have) {
  const out = [], today = calToday_(), f0 = Date.parse(CAL_CHECK_FROM) / 864e5;
  for (let d = Math.max(today - 7, f0); d <= today - 1; d++) { const k = calIso_(d);
    CAL_ACCTS.forEach(a => { const key = k + '|' + a.toLowerCase(), s = have.stmt[key], c = have.cash[key];
      if (s != null && c == null) out.push([k, 'Cash receiving is not entered for ' + a]);
      else if (c != null && s == null) out.push([k, 'No income statement rows for ' + a]);
      else if (s != null && Math.round(s) !== Math.round(c)) out.push([k, a + ': income statement sales ' + Math.round(s).toLocaleString('en-US') + ' but Cash Receiving shows ' + Math.round(c).toLocaleString('en-US')]); }); }
  return out;
}
// Table of accounts x checks for each of the last days (from CAL_CHECK_FROM): income statement rows, approved expense details and cash receiving are read from the sheets;
// "Approved expense details" can also be ticked by hand when an account had none, and "Operating check" is ticked by hand in the Calendar (saved in the Done tab as AX:<account> / OC:<account>).
function calMatrix_(have) {
  const today = calToday_(), f0 = Date.parse(CAL_CHECK_FROM) / 864e5, days = [], M = { accts: CAL_ACCTS, days: days, stmt: {}, cash: {}, apx: {} };
  for (let d = Math.max(today - 14, f0); d <= today - 1; d++) { const k = calIso_(d); days.push(k);
    CAL_ACCTS.forEach(a => { const key = k + '|' + a.toLowerCase(); if (have.stmt[key] != null) M.stmt[key] = 1; if ((have.cash[key] || 0) > 0) M.cash[key] = 1; if (have.apx[key] != null) M.apx[key] = 1; }); }
  return M;
}
function calData_() {
  const ev = calRows_('Schedule', 9).filter(r => r[8] !== 'Deleted' && r[1]).map(r => [r[1], r[2], r[3], r[4], r[5], r[6], r[7]]);
  const ro = calRows_('Routine', 12).filter(r => r[1] && String(r[11]).toUpperCase() !== 'N').map(r => [r[1], r[2], r[3], [4, 5, 6, 7, 8, 9, 10].map(i => String(r[i]).toUpperCase() === 'Y' ? 1 : 0).join(''), r[12] || '']);
  const lim = new Date(); lim.setDate(lim.getDate() - 400); const lk = Utilities.formatDate(lim, 'UTC', 'yyyy-MM-dd');
  const dn = calRows_('Done', 5).filter(r => r[1] >= lk).map(r => [r[1], r[2], r[5] || '']);
  const have = calHave_();
  return { ev: ev, ro: ro, dn: dn, au: calAutoFor_(ro, have), ck: calChecks_(have), mx: calMatrix_(have), ps: have.ps, err: have.err };
}
function getCalendar() { guard_('cl'); return memo_('getCalendar', [], 120, () => getCalendar_c_()); }
function getCalendar_c_() {
  if (!calOk_()) return JSON.stringify({ can: false });
  const d = calData_(); d.can = true; return JSON.stringify(d);
}
const calOn_ = (e, k) => { if (k < e[2]) return false; if (e[4] === 'Every day') return true; const a = new Date(e[2] + 'T00:00:00Z'), b = new Date(k + 'T00:00:00Z'); if (e[4] === 'Every week') return a.getUTCDay() === b.getUTCDay(); if (e[4] === 'Every month') return a.getUTCDate() === b.getUTCDate(); return e[2] === k; };
// Emails reminders that came due and are not done, plus one morning note about unpaid payments and data checks. Run calSetupAlerts once to start it (every 15 minutes).
function calAlerts() {
  const tz = 'Asia/Dhaka', now = new Date(), k = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), nowM = +Utilities.formatDate(now, tz, 'H') * 60 + +Utilities.formatDate(now, tz, 'm'), dow = (+Utilities.formatDate(now, tz, 'u')) % 7;
  const d = calData_(), done = {}; d.dn.forEach(x => { done[x[0] + '|' + x[1]] = 1; });
  const P = PropertiesService.getScriptProperties(); let sent = {}; try { sent = JSON.parse(P.getProperty('calsent') || '{}'); } catch (x) {}
  Object.keys(sent).forEach(s => { if (s.indexOf(k) !== 0) delete sent[s]; });
  const mins = h => (+h.slice(0, 2)) * 60 + (+h.slice(3, 5)), due = [];
  d.ro.forEach(r => { if (r[3][dow] === '1' && r[2] && mins(r[2]) <= nowM && !done[k + '|R:' + r[0]] && (d.au[r[0]] || []).indexOf(k) < 0) due.push({ key: k + '|R:' + r[0], t: r[1], h: r[2] }); });
  d.ev.forEach(e => { if (e[3] && calOn_(e, k) && mins(e[3]) <= nowM && !done[k + '|E:' + e[0]]) due.push({ key: k + '|E:' + e[0], t: e[1], h: e[3] }); });
  const fresh = due.filter(x => !sent[x.key]);
  if (fresh.length) { MailApp.sendEmail(CAL_ALERT_TO.join(','), 'Reminder: ' + fresh.map(x => x.t).slice(0, 2).join('; ') + (fresh.length > 2 ? ' +' + (fresh.length - 2) + ' more' : ''), fresh.map(x => x.h + '  ' + x.t).join('\n') + '\n\nTick them off in the dashboard Calendar.'); fresh.forEach(x => { sent[x.key] = 1; }); }
  if (nowM >= 9 * 60 && !sent[k + '|digest']) {
    const unpaid = d.ps.filter(p => p[0] <= k && p[3] > 0).map(p => p[1] + ' ' + Math.round(p[3]).toLocaleString('en-US') + ' (' + p[0] + ')'), ck = d.ck.map(c => c[0] + ' ' + c[1]);
    if (unpaid.length || ck.length) MailApp.sendEmail(CAL_ALERT_TO.join(','), 'Morning check: ' + unpaid.length + ' unpaid, ' + ck.length + ' data checks', (unpaid.length ? 'Payments not yet marked paid:\n' + unpaid.join('\n') + '\n\n' : '') + (ck.length ? 'Data checks:\n' + ck.join('\n') : ''));
    sent[k + '|digest'] = 1;
  }
  P.setProperty('calsent', JSON.stringify(sent));
}
function calSetupAlerts() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'calAlerts') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('calAlerts').timeBased().everyMinutes(15).create();
}
// q = { op: 'addEv'|'delEv'|'addRo'|'setRo'|'delRo'|'tick', ... }  ->  new id or 'ok'
function calSave(json) { try { return calSave_w_(json); } finally { bump_(); } }
function calSave_w_(json) { guard_('cl');
  return withLock_(() => {
    const q = JSON.parse(json), me = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    const okD = d => /^\d{4}-\d{2}-\d{2}$/.test(String(d)), okT = t => t === '' || /^([01]\d|2[0-3]):[0-5]\d$/.test(String(t));
    const ids = (sh, c) => { const r = sh.getLastRow(); return r > 1 ? sh.getRange(2, c, r - 1, 1).getDisplayValues().map(x => txt_(x[0])) : []; };
    const newId = p => p + Date.now().toString(36).toUpperCase();
    const days = s => { s = String(s || '1111111'); if (!/^[01]{7}$/.test(s)) throw new Error('Choose the days.'); return s.split('').map(c => c === '1' ? 'Y' : ''); };
    if (q.op === 'addEv') {
      const t = txt_(q.title); if (!t) throw new Error('Write what to schedule.'); if (!okD(q.date)) throw new Error('Check the date.'); if (!okT(q.time || '')) throw new Error('Check the time.');
      if (['Meeting', 'Payment', 'Deadline', 'Other'].indexOf(q.type) < 0) throw new Error('Choose the type.');
      if (['Does not repeat', 'Every day', 'Every week', 'Every month'].indexOf(q.repeat) < 0) throw new Error('Choose how it repeats.');
      const s = calSh_('Schedule'), id = newId('E'), r = Math.max(2, s.getLastRow() + 1);
      s.getRange(r, 1, 1, 9).setValues([[id, t, q.date, q.time || '', q.type, q.repeat, txt_(q.notes), 'Open', me]]); return id;
    }
    if (q.op === 'delEv') { const s = calSh_('Schedule'), i = ids(s, 1).indexOf(txt_(q.id)); if (i < 0) throw new Error('That item was not found.'); s.getRange(i + 2, 8).setValue('Deleted'); return 'ok'; }
    if (q.op === 'addRo' || q.op === 'setRo') {
      const t = txt_(q.task); if (!t) throw new Error('Write the task.'); if (!okT(q.time || '')) throw new Error('Check the time.'); const dd = days(q.days), s = calSh_('Routine'), au = txt_(q.auto);
      if (au && !/^(stmt|cash|bank|exp)\|[A-Za-z0-9 \-]{1,20}\|[01]$/.test(au)) throw new Error('Check the automatic completion setting.');
      if (q.op === 'addRo') { const id = newId('R'), r = Math.max(2, s.getLastRow() + 1); s.getRange(r, 1, 1, 12).setValues([[id, t, q.time || '', dd[0], dd[1], dd[2], dd[3], dd[4], dd[5], dd[6], 'Y', au]]); return id; }
      const i = ids(s, 1).indexOf(txt_(q.id)); if (i < 0) throw new Error('That task was not found.'); s.getRange(i + 2, 2, 1, 9).setValues([[t, q.time || ''].concat(dd)]); s.getRange(i + 2, 12).setValue(au); return 'ok';
    }
    if (q.op === 'delRo') { const s = calSh_('Routine'), i = ids(s, 1).indexOf(txt_(q.id)); if (i < 0) throw new Error('That task was not found.'); s.getRange(i + 2, 11).setValue('N'); return 'ok'; }
    if (q.op === 'tick') {
      if (!okD(q.date)) throw new Error('Check the date.'); const id = txt_(q.id); if (!id) throw new Error('Missing item.');
      const s = calSh_('Done'), r = s.getLastRow(), V = r > 1 ? s.getRange(2, 1, r - 1, 2).getDisplayValues() : [];
      let at = -1; for (let i = V.length - 1; i >= 0; i--) if (V[i][0] === q.date && V[i][1] === id) { at = i + 2; break; }
      if (q.on) { if (at < 0) s.getRange(Math.max(2, r + 1), 1, 1, 4).setValues([[q.date, id, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), me]]); }
      else if (at > 0) s.deleteRow(at);
      return 'ok';
    }
    if (q.op === 'note') {
      if (!okD(q.date)) throw new Error('Check the date.'); const id = txt_(q.id), note = txt_(q.note).slice(0, 300); if (!id) throw new Error('Missing item.');
      const s = calSh_('Done'), r = s.getLastRow(), V = r > 1 ? s.getRange(2, 1, r - 1, 2).getDisplayValues() : [];
      let at = -1; for (let i = V.length - 1; i >= 0; i--) if (V[i][0] === q.date && V[i][1] === id) { at = i + 2; break; }
      if (at > 0) s.getRange(at, 5).setValue(note); else s.getRange(Math.max(2, r + 1), 1, 1, 5).setValues([[q.date, id, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), me, note]]);
      return 'ok';
    }
    throw new Error('Unknown action.');
  });
}

// ---------- Garage profiling: the six workbooks in SBS / Garage ----------
// Garages (Baizid, Vatiyari, 1kilo Body), their staff, buses, job cards, stock, attendance and daily log.
// getGarage() is the light first call; getGarageInv(), getGarageHR() and getGarageBus() are read only when that view is opened.
const GARAGE_ = {
  MASTER: '19uLlpLuKXN5lcRd7fnzLf-Tz5W1kiJA7YSfrfNv7nE4', JOBS: '1InSgz1Tg2N9TU72jh3VG8MqL1COOYTk6iP4e81C534U',
  INV: '1Wn6A50xZ4FnR8ODdNN83uX14IspRl8KlxYV04GE7Gyk', HR: '1q83U_Kaixe1L7eFh1fGbAKojzR4x1o8kHA0AxO8evB0',
  DAILY: '1i5N0WO4sSf0oqLniHWb2cDI5z_C6SgUOIDz0CluikHQ', REPORTS: '11oejgOwjZqpc0y1vpx3WAdYogNF-avREAM76RywhbYM'
};
function gpGrid_(key, tab, cols) {
  const sh = SpreadsheetApp.openById(GARAGE_[key]).getSheetByName(tab);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols || sh.getLastColumn()).getValues();
}
const gpN_ = v => typeof v === 'number' ? v : (parseFloat(String(v == null ? '' : v).replace(/,/g, '')) || 0);
const gpDay_ = v => v instanceof Date ? Math.floor(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) / 864e5) : null;
const gpMon_ = v => v instanceof Date ? v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) : txt_(v).slice(0, 7);
function gpParts_() { const pm = {}; gpGrid_('MASTER', 'Parts', 7).forEach(r => { pm[txt_(r[0])] = { n: txt_(r[1]), c: txt_(r[2]), u: txt_(r[3]), min: gpN_(r[4]), cost: gpN_(r[5]), v: txt_(r[6]) }; }); return pm; }

function getGarage() { guard_('gp', 'ov'); return memo_('getGarage', [], 120, () => getGarage_c_()); }
function getGarage_c_() {
  const gar = gpGrid_('MASTER', 'Garages', 2).map(r => [txt_(r[0]), txt_(r[1])]);
  const sec = gpGrid_('MASTER', 'Sections', 3).map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2])]);
  const staff = gpGrid_('MASTER', 'Staff', 6).map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2]), txt_(r[3]), txt_(r[4])]);
  const buses = gpGrid_('MASTER', 'Buses', 7).map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2]), txt_(r[3]), txt_(r[6])]);
  const rep = gpGrid_('REPORTS', 'By Garage', 11).map(r => [gpMon_(r[0]), txt_(r[1])].concat([2, 3, 4, 5, 6, 7, 8, 9, 10].map(i => gpN_(r[i]))));
  const inv = gpGrid_('REPORTS', 'By Function', 5).filter(r => txt_(r[2]) === 'Inventory').map(r => [gpMon_(r[0]), txt_(r[1]), txt_(r[2]), txt_(r[3]), gpN_(r[4])]);
  const busRep = gpGrid_('REPORTS', 'By Bus', 11).map(r => [txt_(r[0]), txt_(r[1]), gpN_(r[2]), gpN_(r[3]), gpN_(r[4]), gpN_(r[5]), gpN_(r[6]), gpDay_(r[7]), txt_(r[8]), txt_(r[9]), gpN_(r[10])]);
  const open = gpGrid_('JOBS', 'Job Cards', 8).filter(r => txt_(r[7]) !== 'Closed').map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2]), gpDay_(r[3]), txt_(r[5]), txt_(r[6]), txt_(r[7])]);
  const dl = gpGrid_('DAILY', 'Daily Log', 8); let asOf = 0; dl.forEach(r => { const d = gpDay_(r[0]); if (d > asOf) asOf = d; });
  const daily = dl.map(r => [gpDay_(r[0]), txt_(r[1]), gpN_(r[2]), gpN_(r[3]), gpN_(r[4]), gpN_(r[5]), gpN_(r[6]), txt_(r[7])]).filter(r => r[0] > asOf - 120);
  const stamp = SpreadsheetApp.openById(GARAGE_.REPORTS).getSheetByName('By Garage').getRange(1, 13).getValue();
  return JSON.stringify({ gar: gar, sec: sec, staff: staff, buses: buses, rep: rep, inv: inv, busRep: busRep, open: open, daily: daily, asOf: asOf, stamp: txt_(stamp) });
}

function getGarageInv() { guard_('gp', 'ov'); return memo_('getGarageInv', [], 120, () => getGarageInv_c_()); }
function getGarageInv_c_() {
  const pm = gpParts_(), st = {}, buy = {}, iss = {};
  const add = (g, c, q) => { const k = g + '|' + c; st[k] = (st[k] || 0) + q; };
  gpGrid_('INV', 'Stock In', 9).forEach(r => { const g = txt_(r[1]), c = txt_(r[2]); add(g, c, gpN_(r[3])); const k = gpMon_(r[0]) + '|' + txt_(r[6]); buy[k] = (buy[k] || 0) + gpN_(r[5]); });
  gpGrid_('INV', 'Stock Out', 7).forEach(r => { const g = txt_(r[1]), c = txt_(r[2]), q = gpN_(r[3]); add(g, c, -q); const p = pm[c], k = gpMon_(r[0]) + '|' + g; iss[k] = (iss[k] || 0) + q * (p ? p.cost : 0); });
  gpGrid_('INV', 'Transfers', 6).forEach(r => { const c = txt_(r[3]), q = gpN_(r[4]); add(txt_(r[1]), c, -q); add(txt_(r[2]), c, q); });
  return JSON.stringify({
    parts: Object.keys(pm).map(c => [c, pm[c].n, pm[c].c, pm[c].u, pm[c].min, pm[c].cost, pm[c].v]),
    stock: Object.keys(st).map(k => { const a = k.split('|'); return [a[0], a[1], st[k]]; }),
    buy: Object.keys(buy).map(k => { const a = k.split('|'); return [a[0], a[1], Math.round(buy[k])]; }),
    iss: Object.keys(iss).map(k => { const a = k.split('|'); return [a[0], a[1], Math.round(iss[k])]; }),
    vendors: gpGrid_('MASTER', 'Vendors', 2).map(r => [txt_(r[0]), txt_(r[1])])
  });
}

function getGarageHR() { guard_('gp', 'ov'); return memo_('getGarageHR', [], 120, () => getGarageHR_c_()); }
function getGarageHR_c_() {
  const A = gpGrid_('HR', 'Attendance', 5); let max = 0;
  const D = A.map(r => gpDay_(r[0])); for (let i = 0; i < D.length; i++) if (D[i] > max) max = D[i];
  const mon = {}, day = {}, ab = {};
  for (let i = 0; i < A.length; i++) {
    const r = A[i], g = txt_(r[2]), s = txt_(r[3]), ot = gpN_(r[4]), d = D[i];
    const km = gpMon_(r[0]) + '|' + g, o = mon[km] || (mon[km] = [0, 0, 0, 0, 0]);
    const j = s === 'Present' ? 0 : s === 'Absent' ? 1 : s === 'Leave' ? 2 : 3; o[j]++; o[4] += ot;
    if (d > max - 30) { const kd = d + '|' + g, p = day[kd] || (day[kd] = [0, 0, 0, 0]); p[j]++; }
    if (d > max - 90 && s === 'Absent') { const e = txt_(r[1]); ab[e] = (ab[e] || 0) + 1; }
  }
  return JSON.stringify({
    asOf: max,
    mon: Object.keys(mon).map(k => { const a = k.split('|'); return [a[0], a[1]].concat(mon[k]); }),
    day: Object.keys(day).map(k => { const a = k.split('|'); return [+a[0], a[1]].concat(day[k]); }),
    abs: Object.keys(ab).map(e => [e, ab[e]]).sort((x, y) => y[1] - x[1]).slice(0, 20)
  });
}

// One bus: every job card with its labour and parts cost, and the parts it used most.
function getGarageBus(bus) { guard_('gp', 'ov'); return memo_('getGarageBus', [bus], 120, () => getGarageBus_c_(bus)); }
function getGarageBus_c_(bus) {
  const key = txt_(bus), ids = {}, pm = gpParts_(), pu = {};
  const J = gpGrid_('JOBS', 'Job Cards', 11).filter(r => txt_(r[1]) === key);
  J.forEach(r => { ids[txt_(r[0])] = { lab: 0, par: 0 }; });
  gpGrid_('JOBS', 'Job Lines', 8).forEach(r => { const o = ids[txt_(r[0])]; if (o) o.lab += gpN_(r[6]); });
  gpGrid_('JOBS', 'Parts Used', 7).forEach(r => { const o = ids[txt_(r[0])]; if (!o) return; const a = gpN_(r[4]), c = txt_(r[1]), p = pu[c] || (pu[c] = { q: 0, a: 0 }); o.par += a; p.q += gpN_(r[2]); p.a += a; });
  const jobs = J.map(r => { const o = ids[txt_(r[0])]; return [txt_(r[0]), txt_(r[2]), gpDay_(r[3]), gpDay_(r[4]), txt_(r[5]), txt_(r[6]), txt_(r[7]), gpN_(r[8]), o.lab, o.par]; }).sort((a, b) => b[2] - a[2]);
  const parts = Object.keys(pu).map(c => [c, pm[c] ? pm[c].n : c, pu[c].q, pu[c].a]).sort((a, b) => b[3] - a[3]).slice(0, 12);
  return JSON.stringify({ bus: key, jobs: jobs, parts: parts });
}

// ---------- Garage v2: Ask the garage, Service due, Fleet health ----------
// Reads the report tabs written by GarageSeed.gs (Cost Summary, Service Due, Service Intervals, Fleet Health) so every question is a filter and a sum.
GARAGE_.ODO = '10vLpSmxPlTXlafT4L0pjA8uDhhRSPo-yTmdxXTE72t0';
function gaSum_() { return gpGrid_('REPORTS', 'Cost Summary', 12); }
function getGarageAsk(fj) { guard_('gp', 'ov'); return memo_('getGarageAsk', [fj], 120, () => getGarageAsk_c_(fj)); }
function getGarageAsk_c_(fj) {
  const f = fj ? JSON.parse(fj) : {}, R = gaSum_(), uniq = i => { const o = {}; R.forEach(r => { o[txt_(r[i])] = 1; }); return Object.keys(o).sort(); };
  const months = {}; R.forEach(r => { months[gpMon_(r[0])] = 1; });
  const meta = { months: Object.keys(months).sort(), causes: uniq(4), systems: uniq(5), comps: uniq(6), gars: gpGrid_('MASTER', 'Garages', 2).map(r => [txt_(r[0]), txt_(r[1])]), accts: uniq(2), buses: uniq(1), stamp: txt_(SpreadsheetApp.openById(GARAGE_.REPORTS).getSheetByName('Cost Summary').getRange(1, 14).getValue()) };
  if (!f.by) return JSON.stringify({ meta: meta });
  const has = (a, v) => !a || !a.length || a.indexOf(v) >= 0, KEY = { m: 0, bus: 1, acct: 2, gid: 3, ca: 4, sy: 5, co: 6 };
  const g = {}; let n = 0;
  R.forEach(r => {
    const m = gpMon_(r[0]);
    if (f.from && m < f.from) return; if (f.to && m > f.to) return;
    if (!has(f.cause, txt_(r[4])) || !has(f.system, txt_(r[5])) || !has(f.comp, txt_(r[6])) || !has(f.gid, txt_(r[3])) || !has(f.acct, txt_(r[2]))) return;
    if (f.bus && txt_(r[1]) !== f.bus) return;
    const dims = f.by.map(k => k === 'm' ? m : txt_(r[KEY[k]])), key = dims.join('|'), o = g[key] || (g[key] = { d: dims, lab: 0, par: 0, jobs: 0, last: 0 });
    o.lab += gpN_(r[7]); o.par += gpN_(r[8]); o.jobs += gpN_(r[10]); const ld = gpDay_(r[11]); if (ld > o.last) o.last = ld; n++;
  });
  const rows = Object.keys(g).map(k => { const o = g[k]; return o.d.concat([Math.round(o.lab), Math.round(o.par), Math.round(o.lab + o.par), o.jobs, o.last]); });
  return JSON.stringify({ meta: meta, rows: rows.slice(0, 3000), more: rows.length > 3000, n: n });
}
function getGarageService() { guard_('gp', 'ov'); return memo_('getGarageService', [], 120, () => getGarageService_c_()); }
function getGarageService_c_() {
  const due = gpGrid_('REPORTS', 'Service Due', 14).map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2]), gpDay_(r[3]), gpN_(r[4]), gpN_(r[5]), gpN_(r[6]), gpN_(r[7]), gpN_(r[8]), txt_(r[9]), gpN_(r[10]), gpN_(r[11]), txt_(r[12]), gpN_(r[13])]);
  let max = 0; const all = gpGrid_('REPORTS', 'Service Intervals', 9).map(r => [txt_(r[0]), txt_(r[1]), txt_(r[2]), gpDay_(r[3]), gpN_(r[4]), gpN_(r[5]), gpN_(r[6]), gpN_(r[7]), txt_(r[8])]);
  all.forEach(r => { if (r[3] > max) max = r[3]; });
  const ints = all.filter(r => r[3] > max - 400);
  const plan = gpGrid_('MASTER', 'Service Plan', 6).map(r => [txt_(r[0]), txt_(r[1]), gpN_(r[2]), gpN_(r[3]), txt_(r[4]), gpN_(r[5])]);
  const stamp = SpreadsheetApp.openById(GARAGE_.REPORTS).getSheetByName('Service Due').getRange(1, 16).getValue();
  return JSON.stringify({ due: due, ints: ints, plan: plan, asOf: max, stamp: txt_(stamp) });
}
function getGarageFleet() { guard_('gp', 'ov'); return memo_('getGarageFleet', [], 120, () => getGarageFleet_c_()); }
function getGarageFleet_c_() {
  const fh = gpGrid_('REPORTS', 'Fleet Health', 10).map(r => [txt_(r[0]), txt_(r[1]), gpN_(r[2]), gpN_(r[3]), gpN_(r[4]), gpN_(r[5]), gpN_(r[6]), gpN_(r[7]), gpN_(r[8]), r[9] === '' ? null : gpN_(r[9])]);
  const jc = gpGrid_('JOBS', 'Job Cards', 17); let max = 0; jc.forEach(r => { const d = gpDay_(r[3]); if (d > max) max = d; });
  const rep = jc.filter(r => txt_(r[16]) && gpDay_(r[3]) > max - 180).map(r => [txt_(r[0]), txt_(r[1]), gpDay_(r[3]), txt_(r[2]), txt_(r[5]), txt_(r[13]), txt_(r[16])]).sort((a, b) => b[2] - a[2]).slice(0, 300);
  return JSON.stringify({ fh: fh, rep: rep, asOf: max });
}
