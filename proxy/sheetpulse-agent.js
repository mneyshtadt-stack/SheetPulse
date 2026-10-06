// Read-only access to the SheetPulse spreadsheet for Claude (the assistant that maintains SheetPulse),
// so checks against the live sheet no longer need an .xlsx download.
//
// Mounted in index.js BEFORE the Google sign-in check (it has its own secret instead), with:
//   const { handleAgentRoutes } = require('./sheetpulse-agent');
//   if (await handleAgentRoutes(req, res, parsed.pathname, send)) return;
//
//   GET /agent/tabs                  ->  {tabs: [...]}           (the readable tab names)
//   GET /agent/sheet?name=<tab>      ->  CSV of that tab         (only tabs in AGENT_TABS)
//   Header:  X-Agent-Token: <SHEETPULSE_AGENT_TOKEN>
//
// Safety: read-only (values.get only -- there is no write call in this file), a fixed list of
// tabs, constant-time token check, and nothing at all unless SHEETPULSE_AGENT_TOKEN is set in .env.
// Revoke: delete the SHEETPULSE_AGENT_TOKEN line from ~/sheet-proxy/.env and restart sheet-proxy.
// The token lives on the VM (.env) and on the user's PC outside OneDrive -- never in this repo.
//
// .env:
//   SHEETPULSE_AGENT_TOKEN   the shared secret (64 hex characters; generated on the VM by the
//                            install command in Diagnostics -> Troubleshooting)
//   SHEETPULSE_SHEET_ID      optional, the Tracker spreadsheet (default: SheetPulse's own)
// Uses the same GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN as index.js.

const crypto = require('crypto');

const SHEET_ID = process.env.SHEETPULSE_SHEET_ID || '1eKHhkNP9F39jwqlwokLoJ0zC3i4jA5u91zg5WyNPb0g';
const AGENT_TABS = ['Tracker', 'Transactions', 'Dividends', 'IntradayLog', 'PriceHistory', 'Investment Universe', 'Pulse Check'];
const MAX_PER_MINUTE = 30;
let window_ = { start: 0, count: 0 };

function tokenOk(req) {
  const want = process.env.SHEETPULSE_AGENT_TOKEN || '';
  const got = String(req.headers['x-agent-token'] || '');
  if (want.length < 32 || got.length !== want.length) return false;
  return crypto.timingSafeEqual(Buffer.from(got), Buffer.from(want));
}

function csvCell(v) {
  const s = v === undefined || v === null ? '' : String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

async function readTabCsv(name) {
  const { google } = require('googleapis');
  const auth = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  const sheets = google.sheets({ version: 'v4', auth });
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${name}'`, valueRenderOption: 'FORMATTED_VALUE' });
  return (r.data.values || []).map((row) => row.map(csvCell).join(',')).join('\n');
}

async function handleAgentRoutes(req, res, pathname, send) {
  if (!pathname || !pathname.startsWith('/agent/')) return false;
  const reply = (status, body, type) => send(res, req, status, body, type || 'application/json');
  if (!process.env.SHEETPULSE_AGENT_TOKEN) { reply(404, JSON.stringify({ error: 'Not found' })); return true; }
  if (req.method !== 'GET') { reply(405, JSON.stringify({ error: 'Read-only: use GET' })); return true; }
  if (!tokenOk(req)) { reply(401, JSON.stringify({ error: 'Bad or missing X-Agent-Token' })); return true; }
  const now = Date.now();
  if (now - window_.start > 60000) window_ = { start: now, count: 0 };
  if (++window_.count > MAX_PER_MINUTE) { reply(429, JSON.stringify({ error: 'Too many requests, wait a minute' })); return true; }

  if (pathname === '/agent/tabs') { reply(200, JSON.stringify({ tabs: AGENT_TABS })); return true; }
  if (pathname === '/agent/sheet') {
    const name = new URL(req.url || pathname, 'http://local').searchParams.get('name') || '';
    if (AGENT_TABS.indexOf(name) === -1) { reply(400, JSON.stringify({ error: 'Tab not allowed', allowed: AGENT_TABS })); return true; }
    try {
      reply(200, await readTabCsv(name), 'text/csv; charset=utf-8');
    } catch (e) {
      console.error('agent read failed:', e.message);
      reply(502, JSON.stringify({ error: 'Could not read the sheet' }));
    }
    return true;
  }
  reply(404, JSON.stringify({ error: 'Not found' }));
  return true;
}

module.exports = { handleAgentRoutes };
