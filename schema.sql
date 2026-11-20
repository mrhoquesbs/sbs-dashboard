-- Cloudflare D1: edge cache of every dashboard call
CREATE TABLE IF NOT EXISTS rpc (
  k  TEXT PRIMARY KEY,     -- email | function | arguments
  fn TEXT NOT NULL,
  em TEXT NOT NULL,
  ar TEXT NOT NULL,
  z  TEXT NOT NULL,        -- gzip + base64 of the answer
  t  INTEGER NOT NULL,     -- when it was fetched from Apps Script
  la INTEGER NOT NULL      -- when it was last used
);
CREATE INDEX IF NOT EXISTS rpc_t ON rpc (t);
