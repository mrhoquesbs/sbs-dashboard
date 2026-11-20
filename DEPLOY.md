# Dashboard on Cloudflare (free) — all pages, all data

How it works
- Your Google Sheets and Drive files stay the place where data is typed and kept. Nothing is dropped.
- The page itself is served from Cloudflare, so it opens at once.
- Every call the page makes (all 35 kinds: every page, every save, every edit) goes through Cloudflare:
  - Reads are stored in a free Cloudflare database. Repeat views come back in milliseconds. A result older than 2 minutes is shown immediately and refreshed in the background.
  - Saves and edits go straight to Apps Script as the signed-in person, with your existing page permissions and account limits, then the stored results are cleared so the next view is fresh.
  - Every 5 minutes a trigger refreshes the views people actually use, and any change in a dashboard sheet clears the stored results.
- First-ever view of a page (or a view nobody has opened for a day) still takes as long as it does today, because Apps Script builds it. Everything after that is fast.
- Sign-in is Cloudflare Access: only the emails you list can open the site, and the server checks the Access token itself.

Folder: index.html (the page), functions/api/[[path]].js (the Cloudflare code), schema.sql (database), Code.gs (your Code.gs with one extra line in me_()), ApiHost.gs (new Apps Script file).

## Setup (about 30 minutes)
A. **Apps Script**
 1. Replace the Code.gs in your project with this Code.gs. Add ApiHost.gs as a new file.
 2. Project settings > Script properties: add `RPC_SECRET` (a long random text) and `PUSH_TOKEN` (another long random text). Keep both.
 3. Deploy > New deployment > Web app > Execute as: **Me**, Who has access: **Anyone**. Copy the web app address ending in /exec. (Keep your old deployment as it is; it keeps working as a backup.)
B. **Cloudflare**
 1. Sign up free at cloudflare.com. Put this folder in a free GitHub repository.
 2. Storage & Databases > D1 > create `sbsdb`, open its Console, paste schema.sql, Run.
 3. Workers & Pages > Create > Pages > Connect to Git > your repository. No build command, output directory `/`.
 4. Pages project > Settings > Bindings: D1 database, variable name `DB` = sbsdb. Settings > Variables and secrets: `APPS_URL` (the /exec address), `RPC_SECRET`, `PUSH_TOKEN` (same texts as in Apps Script), `CF_TEAM` (your Zero Trust team name, the part before .cloudflareaccess.com), `CF_AUD` (from step 5). Redeploy once.
 5. Zero Trust (free) > Access > Applications > Add > Self-hosted. Domain: your `xxxx.pages.dev` (also add `*.xxxx.pages.dev` so preview links are covered). Policy: Allow > Emails > the approved people. Copy the Application Audience (AUD) tag into `CF_AUD`.
C. **Finish**
 1. In Apps Script add script property `CLOUD_URL` = `https://xxxx.pages.dev`, then run `setupCloud` once and approve.
 2. Open the site, sign in, check Overview and one page that saves something (use a test entry).

## Notes
- Access control still lives in your Access page (ACL in Apps Script). Per-page and per-account limits are enforced by Apps Script exactly as now, because each call is made as the signed-in person.
- Cached results are kept per person, so one person's limits never leak to another.
- A free Cloudflare plan allows far more than this use (D1 100,000 writes a day; the dashboard uses a few thousand).
- Tested here: the Cloudflare code with a mock database and signed test tokens (wrong token, wrong audience, expired token, cache hit, stale refresh, errors, saves clear the cache, prewarm, purge); syntax of all files. Not tested against a real Cloudflare or Apps Script deployment, so do the check in C.2 and send me anything that errors.
