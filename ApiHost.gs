/** Add this file to the same Apps Script project as Code.gs (use the Code.gs from this package: me_() has one extra line).
 *  Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone. Copy its /exec address into Cloudflare as APPS_URL.
 *  Script properties: RPC_SECRET (long random text, same as in Cloudflare), CLOUD_URL (https://xxxx.pages.dev), PUSH_TOKEN (same as in Cloudflare).
 *  Run setupCloud() once. */
let FWD_EMAIL_ = '';
const NOCALL_ = ['doGet', 'doPost', 'setupCloud', 'cloudPrewarm', 'cloudPurge', 'onSheetChange', 'refreshOverspeed', 'overspeedCopySetup', 'watchSheets', 'setupWatch', 'watchSales_', 'watchMonitoring_', 'refreshOverspeedInc'];

function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  try {
    const b = JSON.parse(e.postData.contents), sec = PropertiesService.getScriptProperties().getProperty('RPC_SECRET');
    if (!sec || b.secret !== sec) return out({ ok: false, error: 'Not allowed.' });
    const fn = String(b.fn || '');
    if (!/^[A-Za-z][A-Za-z0-9]*$/.test(fn) || fn.slice(-1) === '_' || NOCALL_.indexOf(fn) >= 0 || typeof this[fn] !== 'function') return out({ ok: false, error: 'Unknown function: ' + fn });
    FWD_EMAIL_ = String(b.email || '').toLowerCase();                 // me_() in Code.gs uses this instead of the Google session
    const r = this[fn].apply(null, Array.isArray(b.args) ? b.args : []);
    return out({ ok: true, data: r === undefined ? null : r });
  } catch (x) { return out({ ok: false, error: String(x && x.message || x) }); }
}
function cloudCall_(path) {
  const P = PropertiesService.getScriptProperties(), u = P.getProperty('CLOUD_URL'), t = P.getProperty('PUSH_TOKEN'); if (!u || !t) return;
  try { UrlFetchApp.fetch(u.replace(/\/$/, '') + '/api/' + path, { method: 'post', headers: { Authorization: 'Bearer ' + t }, muteHttpExceptions: true }); } catch (x) {}
}
function cloudPrewarm() { cloudCall_('prewarm'); }          // every 5 minutes: refresh what people use
function onSheetChange() { cloudCall_('purge'); }           // when any dashboard sheet changes: next view is fresh
function setupCloud() {
  ScriptApp.getProjectTriggers().forEach(t => { const h = t.getHandlerFunction(); if (h === 'cloudPrewarm' || h === 'onSheetChange') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('cloudPrewarm').timeBased().everyMinutes(5).create();
  const ids = [MASTER_ID, PLAN_ID, MGMT_ID, INCIDENT_ID, MON_ID, PSCHED_ID, ROUTES_ID, RENT_ID, CASHFLOW_ID, SALARY_ID, CAL_ID, '1K_PPyIWBsb-pbcnUuGA_jW4-ebrPR08q3db8fJkH0xI'];
  [...new Set(ids)].forEach(id => { try { ScriptApp.newTrigger('onSheetChange').forSpreadsheet(id).onChange().create(); } catch (x) { console.warn('No trigger for ' + id + ': ' + x.message); } });
  return 'Triggers set. Dashboard sheets now clear the cache when they change.';
}

// Every minute this watcher checks two things:
// 1) Sales files ("Sale yyyy-mm" in the Sales Data folder; not covered by the onChange triggers): when this month's or last month's file was edited, refreshAll re-reads only
//    the changed file, the page cache is cleared and Cloudflare is told to refresh.
// 2) The Monitoring sheet: when it changed, only the new Overspeeding rows are added to the saved copy (refreshOverspeedInc), then Cloudflare is told to refresh.
// It costs under a second a minute when nothing changed. Run setupWatch once to create the 1-minute schedule.
function watchSheets() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return;     // a previous run (or the full overspeed rebuild) is still going
  const P = PropertiesService.getScriptProperties();
  try { watchSales_(P); } catch (e) { console.error('Sales watch: ' + (e && e.message || e)); }
  try { watchMonitoring_(P); } catch (e) { console.error('Monitoring watch: ' + (e && e.message || e)); }
  finally { lock.releaseLock(); }
}
function watchSales_(P) {
  const bad = +P.getProperty('SALES_ERR_AT') || 0; if (Date.now() - bad < 3e5) return;     // after an error wait 5 minutes before trying again
  const tz = Session.getScriptTimeZone(), d = new Date(), pv = new Date(d.getFullYear(), d.getMonth() - 1, 15), sig = {};
  [Utilities.formatDate(d, tz, 'yyyy-MM'), Utilities.formatDate(pv, tz, 'yyyy-MM')].forEach(m => { const f = file_('Sale ' + m); if (f) sig['Sale ' + m] = f.getLastUpdated().getTime(); });
  const now = JSON.stringify(sig); if (now === P.getProperty('SALES_SIG')) return;
  FWD_EMAIL_ = EDITORS[0];                                       // refreshAll checks the caller: act as the first owner
  try { refreshAll(); P.setProperty('SALES_SIG', now); P.deleteProperty('SALES_ERR_AT'); cloudCall_('purge'); console.log('Sales refreshed after a change: ' + now); }
  catch (e) { P.setProperty('SALES_ERR_AT', String(Date.now())); console.error('Sales refresh failed: ' + (e && e.message || e)); }
  finally { FWD_EMAIL_ = ''; }
}
function watchMonitoring_(P) {
  const bad = +P.getProperty('MON_ERR_AT') || 0; if (Date.now() - bad < 3e5) return;
  const sig = String(DriveApp.getFileById(MON_ID).getLastUpdated().getTime()); if (sig === P.getProperty('MON_SIG')) return;
  try { const m = refreshOverspeedInc(); P.setProperty('MON_SIG', sig); P.deleteProperty('MON_ERR_AT'); if (/added/.test(m)) cloudCall_('purge'); }
  catch (e) { P.setProperty('MON_ERR_AT', String(Date.now())); console.error('Monitoring refresh failed: ' + (e && e.message || e)); }
}
function setupWatch() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'watchSheets') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('watchSheets').timeBased().everyMinutes(1).create();
  const P = PropertiesService.getScriptProperties(); P.deleteProperty('SALES_SIG'); P.deleteProperty('MON_SIG');
  watchSheets();
}
