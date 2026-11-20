// Cloudflare Pages Function. Every call the dashboard makes goes through POST /api/rpc.
//  - Sign-in: Cloudflare Access. The Access token is verified here (signature, audience, expiry), the email header alone is never trusted.
//  - Reads (names starting get...) are kept in D1 and answered at once; a stale answer is returned and refreshed in the background.
//  - Everything else (saves, edits, refresh) goes straight to Apps Script as the signed-in person, then the cache is cleared.
// Env: DB (D1), CF_TEAM, CF_AUD, APPS_URL, RPC_SECRET, PUSH_TOKEN
const FRESH = 120e3, STALE = 7 * 864e5, MAXKEYS = 60, WARM = 7 * 864e5;
// pages people open first or wait for most: kept ready for everyone who used the dashboard in the last week
const HEAVY = [['getMeta', []], ['getFiltered', [{}]], ['getPlan', []], ['getBank', []], ['getPayments', []], ['getOverspeed', []], ['getRouteTypes', []], ['getTypeDeposits', []]];
const J = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { 'content-type': 'application/json', 'cache-control': 'no-store' } });
const b64u = s => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/')), c => c.charCodeAt(0));
let JW = null, JWT_AT = 0;

async function email(request, env) {
  const tok = request.headers.get('Cf-Access-Jwt-Assertion');
  if (!tok || !env.CF_TEAM || !env.CF_AUD) return null;
  const [h, p, s] = tok.split('.'); if (!s) return null;
  const head = JSON.parse(new TextDecoder().decode(b64u(h))), pay = JSON.parse(new TextDecoder().decode(b64u(p)));
  if (!JW || Date.now() - JWT_AT > 36e5 || !JW.keys.some(k => k.kid === head.kid)) {
    const r = await fetch(`https://${env.CF_TEAM}.cloudflareaccess.com/cdn-cgi/access/certs`); if (!r.ok) return null; JW = await r.json(); JWT_AT = Date.now();
  }
  const jwk = JW.keys.find(k => k.kid === head.kid); if (!jwk) return null;
  const key = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
  if (!await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, b64u(s), new TextEncoder().encode(h + '.' + p))) return null;
  const aud = [].concat(pay.aud || []);
  if (!aud.includes(env.CF_AUD) || pay.exp * 1000 < Date.now() || pay.iss !== `https://${env.CF_TEAM}.cloudflareaccess.com`) return null;
  return String(pay.email || '').toLowerCase() || null;
}
async function gz(text) { const s = new Blob([text]).stream().pipeThrough(new CompressionStream('gzip')); const u = new Uint8Array(await new Response(s).arrayBuffer()); let b = ''; for (let i = 0; i < u.length; i += 8192) b += String.fromCharCode(...u.subarray(i, i + 8192)); return btoa(b); }
async function ungz(z) { const u = Uint8Array.from(atob(z), c => c.charCodeAt(0)); return new Response(new Blob([u]).stream().pipeThrough(new DecompressionStream('gzip'))).text(); }

async function apps(env, em, fn, args) {              // one call to the Apps Script web app, as that person
  const r = await fetch(env.APPS_URL, { method: 'POST', headers: { 'content-type': 'text/plain' }, redirect: 'follow', body: JSON.stringify({ secret: env.RPC_SECRET, email: em, fn, args }) });
  const t = await r.text(); try { return JSON.parse(t); } catch (e) { return { ok: false, error: 'Apps Script answered with an unexpected page (status ' + r.status + ').' }; }
}
async function fetchAndStore(env, em, fn, args, k) {
  const o = await apps(env, em, fn, args); if (!o.ok) return o;
  const body = JSON.stringify(o.data === undefined ? null : o.data), now = Date.now();
  try { if (body.length < 6e6) await env.DB.prepare('INSERT INTO rpc (k,fn,em,ar,z,t,la) VALUES (?,?,?,?,?,?,?) ON CONFLICT(k) DO UPDATE SET z=excluded.z,t=excluded.t').bind(k, fn, em, JSON.stringify(args), await gz(body), now, now).run(); } catch (e) {}
  return { ok: true, body };
}
const bearer = (r, env) => env.PUSH_TOKEN && r.headers.get('authorization') === 'Bearer ' + env.PUSH_TOKEN;

export async function onRequest({ request, env, params, waitUntil }) {
  const path = (params.path || []).join('/'), wu = waitUntil || (() => {});
  if (request.method !== 'POST') return J({ error: 'POST only' }, 405);
  if (path === 'purge') { if (!bearer(request, env)) return J({ error: 'no' }, 401); await env.DB.prepare('UPDATE rpc SET t=MIN(t,?)').bind(Date.now() - FRESH - 1e3).run(); return J({ ok: true }); }   // old copy still shown at once, rebuilt in the background
  if (path === 'prewarm') {                         // called every 5 minutes by an Apps Script trigger: rebuild what people use, oldest first
    if (!bearer(request, env)) return J({ error: 'no' }, 401);
    const now = Date.now();
    const { results } = await env.DB.prepare('SELECT k,fn,em,ar FROM rpc WHERE t<? AND la>? ORDER BY t ASC LIMIT ?').bind(now - 9e4, now - WARM, MAXKEYS).all();
    const jobs = results.map(r => [r.em, r.fn, JSON.parse(r.ar), r.k]);
    const { results: us } = await env.DB.prepare('SELECT em, MAX(la) m FROM rpc WHERE la>? GROUP BY em ORDER BY m DESC LIMIT 12').bind(now - WARM).all();
    for (const u of us) for (const [fn, ar] of HEAVY) {   // first view of a main page must never wait: make the copy before they ask
      const k = u.em + '|' + fn + '|' + JSON.stringify(ar);
      if (!(await env.DB.prepare('SELECT 1 x FROM rpc WHERE k=?').bind(k).first()) && jobs.length < MAXKEYS + 12) jobs.push([u.em, fn, ar, k]);
    }
    wu((async () => { const t0 = Date.now(); for (let i = 0; i < jobs.length && Date.now() - t0 < 20e3; i += 4) await Promise.all(jobs.slice(i, i + 4).map(j => fetchAndStore(env, ...j).catch(() => {}))); })());
    return J({ refreshing: jobs.length });
  }
  if (path !== 'rpc') return J({ error: 'not found' }, 404);
  const em = await email(request, env); if (!em) return J({ ok: false, error: 'Sign in required.' }, 401);
  let b; try { b = await request.json(); } catch (e) { return J({ ok: false, error: 'Bad request.' }, 400); }
  const fn = String(b.fn || ''), args = Array.isArray(b.args) ? b.args : [];
  if (!/^[A-Za-z][A-Za-z0-9]{0,63}$/.test(fn)) return J({ ok: false, error: 'Bad name.' }, 400);
  if (!/^(get|whoami$)/.test(fn)) {                // a save, edit or refresh: straight through, then forget the cache
    const o = await apps(env, em, fn, args);
    if (o.ok) {                                      // the saver sees their change at once; everyone else keeps seeing the old page for a moment, then it refreshes
      const { results } = await env.DB.prepare('SELECT fn,ar,k FROM rpc WHERE em=? ORDER BY la DESC LIMIT 8').bind(em).all();
      await env.DB.prepare('DELETE FROM rpc WHERE em=?').bind(em).run();
      await env.DB.prepare('UPDATE rpc SET t=MIN(t,?) WHERE em<>?').bind(Date.now() - FRESH - 1e3, em).run();
      wu((async () => { for (let i = 0; i < results.length; i += 4) await Promise.all(results.slice(i, i + 4).map(r => fetchAndStore(env, em, r.fn, JSON.parse(r.ar), r.k).catch(() => {}))); })());
    }
    return o.ok ? J({ ok: true, data: o.data === undefined ? null : o.data }) : J(o, 400);
  }
  const k = em + '|' + fn + '|' + JSON.stringify(args), now = Date.now();
  const row = await env.DB.prepare('SELECT z,t,la FROM rpc WHERE k=?').bind(k).first();
  if (row && now - row.t < STALE) {
    if (now - row.la > 36e5) wu(env.DB.prepare('UPDATE rpc SET la=? WHERE k=?').bind(now, k).run());
    if (now - row.t >= FRESH) wu(fetchAndStore(env, em, fn, args, k));
    return new Response('{"ok":true,"data":' + await ungz(row.z) + '}', { headers: { 'content-type': 'application/json', 'cache-control': 'no-store', 'x-cache': now - row.t < FRESH ? 'hit' : 'stale' } });
  }
  const o = await fetchAndStore(env, em, fn, args, k);
  return o.ok ? new Response('{"ok":true,"data":' + o.body + '}', { headers: { 'content-type': 'application/json', 'cache-control': 'no-store', 'x-cache': 'miss' } }) : J(o, 400);
}
