/** Add this file to the same Apps Script project as Code.gs (use the Code.gs from this package: me_() has one extra line).
 *  Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone. Copy its /exec address into Cloudflare as APPS_URL.
 *  Script properties: RPC_SECRET (long random text, same as in Cloudflare), CLOUD_URL (https://xxxx.pages.dev), PUSH_TOKEN (same as in Cloudflare).
 *  Run setupCloud() once. */
let FWD_EMAIL_ = '';
const NOCALL_ = ['doGet', 'doPost', 'setupCloud', 'cloudPrewarm', 'cloudPurge', 'onSheetChange', 'refreshOverspeed', 'overspeedCopySetup'];

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
