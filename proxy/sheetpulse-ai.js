// AI analysis for SheetPulse -- the dashboard's "AI analysis" and "AI daily brief" panels, answered
// by Google Gemini.
//
// Mounted in index.js after the sign-in check, like the MedTracker routes, so only the one allowed
// account can use it:
//   POST /ai/analyze        {holdings, summary, question?}  ->  {text, model, usedToday, limit}
//   GET  /ai/status                                         ->  {configured, model, usedToday, limit}
//   GET  /ai/brief[?date=YYYY-MM-DD]                        ->  {brief | null, dates}
//   POST /ai/brief/refresh                                  ->  {brief, usedToday, limit}
//
// Daily brief: written here on the server, automatically at 23:50 Israel time Monday-Friday (after
// TASE's 17:30 / Friday 14:00 close and the USA's 23:00 close), from the Tracker sheet itself, and
// saved in ai-briefs.json next to this file (the last 7 trading days). The dashboard just shows the
// saved one; Refresh writes a new one on demand -- "so far today" while a market trades, otherwise
// the last finished session again, never a "0% because the markets are closed" non-brief.
//
// Uses Gemini's free tier, where Google may use what it's sent to improve its products -- so only
// percentages ever leave this server. sanitizeHoldings / sanitizeSummary keep a fixed whitelist of
// fields and drop anything else, so no ILS / USD amount, total value or share count can reach
// Google -- for the dashboard's requests and for the brief this server builds itself alike.
//
// .env:
//   GEMINI_API_KEY          required (from aistudio.google.com -> API keys; no billing = free tier)
//   GEMINI_MODEL            optional, default gemini-3.8-flash
//   GEMINI_FALLBACK_MODELS  optional, comma-separated, default gemini-3.7-flash,gemini-3.6-flash,gemini-3.5-flash,
//                           gemini-3.5-flash-lite (the other free-tier Flash models; older ones are often
//                           less crowded when the newest is "high demand")
//                           -- tried in order when the main model is busy ("high demand"), over its
//                           free quota, or unavailable
//   AI_DAILY_LIMIT          optional, default 40 answered requests per day (Israel time); the automatic
//                           23:50 brief doesn't count
//   SHEETPULSE_SHEET_ID     optional, the Tracker spreadsheet (default: SheetPulse's own)
//   SHEETPULSE_TASE_TICKERS optional, comma-separated TASE holdings, default IBI.F35,IBI.FK4
// The brief reads the sheet with the same GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET /
// GOOGLE_REFRESH_TOKEN index.js already uses.

const fs = require('fs');
const path = require('path');

const MODEL = process.env.GEMINI_MODEL || 'gemini-3.8-flash';
const FALLBACK_MODELS = (process.env.GEMINI_FALLBACK_MODELS || process.env.GEMINI_FALLBACK_MODEL || 'gemini-3.7-flash,gemini-3.6-flash,gemini-3.5-flash,gemini-3.5-flash-lite')
  .split(',').map((m) => m.trim()).filter((m) => m && m !== MODEL);
// Free-tier "high demand" (503) and per-minute quota (429) spikes usually pass within seconds.
const RETRYABLE = [429, 500, 503, 504];
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const DAILY_LIMIT = Number(process.env.AI_DAILY_LIMIT) || 40;
const TIMEOUT_MS = 90000;
const MAX_HOLDINGS = 200;
const MAX_QUESTION_CHARS = 500;

const SHEET_ID = process.env.SHEETPULSE_SHEET_ID || '1eKHhkNP9F39jwqlwokLoJ0zC3i4jA5u91zg5WyNPb0g';
const TASE_TICKERS = (process.env.SHEETPULSE_TASE_TICKERS || 'IBI.F35,IBI.FK4').split(',').map((t) => t.trim()).filter(Boolean);
const BRIEF_FILE = path.join(__dirname, 'ai-briefs.json');
const BRIEF_KEEP = 7;

const SYSTEM_INSTRUCTION = [
  "You analyze a private investment portfolio for its owner, who lives in Israel and holds Israeli (TASE) funds and USA-listed ETFs and stocks.",
  "You receive percentages only: each holding's weight in the portfolio, its total return since purchase, today's change, sector and role, and sometimes its pre-market or after-hours change. There are no amounts -- never estimate or mention money values.",
  "USA holdings are valued in ILS at the live USD/ILS rate, so a USD/ILS move changes their ILS value even when their USD prices don't move; call that out when it matters.",
  "Explain, don't advise: describe what drives the returns, concentration (single holdings, sectors, roles, markets, currency), outliers, and what moved today. Do not recommend buying, selling or resizing specific securities and give no price targets; if the owner asks, lay out the considerations and say plainly that this isn't financial advice.",
  "Use only the data given -- you have no news or live market access. Name tickers and cite the numbers. Say so when the data can't answer something.",
  "Write in English, about 250-400 words, as Markdown with '## ' section headings, '- ' bullets and **bold** -- nothing else (no tables, no links, no code). Without a question, use the sections: Overview, What drives the return, Concentration and exposure, Today, Worth watching. With a question, answer it directly first.",
].join('\n');

// Israel-time day for the daily limit; kept in memory -- a proxy restart resets it, which is fine.
const usage = { day: '', count: 0 };
function israelDay() {
  return new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Jerusalem' });
}
function usedToday() {
  if (usage.day !== israelDay()) { usage.day = israelDay(); usage.count = 0; }
  return usage.count;
}

const num = (v, dp = 2) => (typeof v === 'number' && Number.isFinite(v) ? Math.round(v * 10 ** dp) / 10 ** dp : null);
const text = (v, max) => (typeof v === 'string' ? v.replace(/[\r\n\t]+/g, ' ').trim().slice(0, max) : '');

// The only fields that may reach Gemini, per holding: all percentages or labels.
function sanitizeHoldings(list) {
  if (!Array.isArray(list)) return [];
  return list.slice(0, MAX_HOLDINGS).map((h) => {
    const ticker = text(h && h.ticker, 20);
    if (!/^[A-Za-z0-9.:\-]+$/.test(ticker)) return null;
    return {
      ticker,
      market: h.market === 'TASE' ? 'TASE' : 'USA',
      weightPct: num(h.weightPct),
      returnPct: num(h.returnPct),
      dayPct: num(h.dayPct),
      extPct: num(h.extPct),
      sector: text(h.sector, 40),
      role: text(h.role, 40),
    };
  }).filter((h) => h && h.weightPct !== null);
}

function sanitizeSummary(s) {
  s = s || {};
  return {
    asOf: text(s.asOf, 40),
    dayPct: num(s.dayPct),
    totalReturnPct: num(s.totalReturnPct),
    usaWeightPct: num(s.usaWeightPct),
    taseWeightPct: num(s.taseWeightPct),
    usdIls: num(s.usdIls, 4),
    usdIlsChangePct: num(s.usdIlsChangePct),
    taseOpen: s.taseOpen === true,
    usaOpen: s.usaOpen === true,
    extSession: s.extSession === 'Pre-market' || s.extSession === 'After hours' ? s.extSession : '',
  };
}

function buildPrompt(holdings, summary, question) {
  const f = (v) => (v === null ? '' : String(v));
  const lines = [
    `Portfolio snapshot${summary.asOf ? ' as of ' + summary.asOf + ' (Israel time)' : ''} -- percentages only.`,
    `Today: ${f(summary.dayPct)}% | total return since purchase: ${f(summary.totalReturnPct)}%`,
    `Split: USA ${f(summary.usaWeightPct)}% / TASE ${f(summary.taseWeightPct)}% | USD/ILS ${f(summary.usdIls)} (${f(summary.usdIlsChangePct)}% vs. last close)`,
    `Markets now: TASE ${summary.taseOpen ? 'open' : 'closed'}, USA ${summary.usaOpen ? 'open' : 'closed'}`
      + (summary.extSession ? ` | extPct = ${summary.extSession.toLowerCase()} change vs. last close` : ''),
    '',
    'ticker,market,weightPct,returnPct,dayPct,extPct,sector,role',
    ...holdings.map((h) => [h.ticker, h.market, f(h.weightPct), f(h.returnPct), f(h.dayPct), f(h.extPct),
      h.sector.replace(/,/g, ';'), h.role.replace(/,/g, ';')].join(',')),
    '',
    question ? `Owner's question: ${question}` : 'Write the standard analysis.',
  ];
  return lines.join('\n');
}

async function callGemini(model, prompt) {
  const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': process.env.GEMINI_API_KEY },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: SYSTEM_INSTRUCTION }] },
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.4 },
    }),
    signal: AbortSignal.timeout(TIMEOUT_MS),
  });
  let data = null;
  try { data = await res.json(); } catch (e) { /* non-JSON error page */ }
  return { status: res.status, data };
}

function extractText(data) {
  const cand = data && data.candidates && data.candidates[0];
  const parts = (cand && cand.content && cand.content.parts) || [];
  // Thinking models can return thought parts first -- keep only the answer.
  return parts.filter((p) => typeof p.text === 'string' && !p.thought).map((p) => p.text).join('').trim();
}

// One answer from Gemini: the main model, again after a short pause if it was only busy, then each
// fallback model. A 404 (model closed to new keys) moves straight on; any other error (e.g. a bad
// request) stops. Resolves {ok: true, text, model} or {ok: false, httpStatus, error}.
async function generate(prompt) {
  const attempts = [
    { model: MODEL, wait: 0 }, { model: MODEL, wait: 2500 },
    ...FALLBACK_MODELS.map((m) => ({ model: m, wait: 0 })),
  ];
  let r = null, model = MODEL, busy = true;
  try {
    for (const a of attempts) {
      if (r && r.status === 404 && a.model === model) continue;  // don't retry a missing model
      if (a.wait) await sleep(a.wait);
      model = a.model;
      r = await callGemini(model, prompt);
      if (r.status === 200) break;
      const msg = (r.data && r.data.error && r.data.error.message) || `HTTP ${r.status}`;
      console.error(new Date().toISOString(), 'gemini error:', model, r.status, msg);
      if (!RETRYABLE.includes(r.status) && r.status !== 404) { busy = false; break; }
    }
  } catch (err) {
    console.error(new Date().toISOString(), 'gemini request failed:', err.message);
    return { ok: false, httpStatus: 504, error: err.name === 'TimeoutError' ? 'Gemini took too long to answer -- try again.' : 'Could not reach Gemini: ' + err.message };
  }
  if (r.status !== 200) {
    const msg = (r.data && r.data.error && r.data.error.message) || `HTTP ${r.status}`;
    return busy
      ? { ok: false, httpStatus: 503, error: `Gemini's free tier is busy right now (tried ${[MODEL, ...FALLBACK_MODELS].join(', ')}). Spikes usually pass within a few minutes -- try again shortly.` }
      : { ok: false, httpStatus: 502, error: 'Gemini error: ' + msg };
  }
  const answer = extractText(r.data);
  if (!answer) {
    const reason = (r.data && r.data.promptFeedback && r.data.promptFeedback.blockReason)
      || (r.data && r.data.candidates && r.data.candidates[0] && r.data.candidates[0].finishReason) || 'empty answer';
    return { ok: false, httpStatus: 502, error: 'Gemini returned no text (' + reason + ')' };
  }
  return { ok: true, text: answer, model };
}

// ---------------------------------------------------------------------------------------------
// Daily brief
// ---------------------------------------------------------------------------------------------

// Date and time parts in a given zone: {date: 'YYYY-MM-DD', hm: 2350, wd: 'Mon'}.
function zoneParts(timeZone, d = new Date()) {
  const f = new Intl.DateTimeFormat('en-US', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'short', hourCycle: 'h23' });
  const p = Object.fromEntries(f.formatToParts(d).map((x) => [x.type, x.value]));
  return { date: `${p.year}-${p.month}-${p.day}`, hm: Number(p.hour) * 100 + Number(p.minute), wd: p.weekday };
}
const WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
function isWeekdayDate(date) {
  const d = new Date(date + 'T12:00:00Z').getUTCDay();
  return d >= 1 && d <= 5;
}
function shiftDate(date, days) {
  const d = new Date(date + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}
function prevWeekday(date) {
  let d = shiftDate(date, -1);
  while (!isWeekdayDate(d)) d = shiftDate(d, -1);
  return d;
}
function dayLabel(date) {
  return new Date(date + 'T12:00:00Z').toLocaleDateString('en-US', { timeZone: 'UTC', weekday: 'long', month: 'short', day: 'numeric' });
}
function taseOpenNow(now = new Date()) {
  const p = zoneParts('Asia/Jerusalem', now);
  if (p.wd === 'Fri') return p.hm >= 1000 && p.hm <= 1400;
  return ['Mon', 'Tue', 'Wed', 'Thu'].includes(p.wd) && p.hm >= 1000 && p.hm <= 1730;
}
function usaOpenNow(now = new Date()) {
  const p = zoneParts('America/New_York', now);
  return WEEKDAYS.includes(p.wd) && p.hm >= 930 && p.hm <= 1600;
}
// Which trading day a brief written now is about: today once TASE has opened (10:00) on a weekday --
// "so far" while a market still trades -- otherwise the previous weekday (early morning, weekends).
function currentSession(now = new Date()) {
  const p = zoneParts('Asia/Jerusalem', now);
  const live = taseOpenNow(now) || usaOpenNow(now);
  if (WEEKDAYS.includes(p.wd) && p.hm >= 1000) return { date: p.date, live };
  return { date: prevWeekday(p.date), live: false };
}

const parseNum = (s) => {
  if (typeof s === 'number') return s;
  if (s === undefined || s === null) return NaN;
  return parseFloat(String(s).replace(/[₪$,%\s]/g, '').replace(/^\((.*)\)$/, '-$1'));
};

async function sheetsClient() {
  const { google } = require('googleapis');
  const auth = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return google.sheets({ version: 'v4', auth });
}
async function readTab(sheets, name) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `'${name}'`, valueRenderOption: 'FORMATTED_VALUE' });
  return r.data.values || [];
}

// The brief's data, read from the Tracker exactly the way the dashboard reads it (header row with
// "Ticker" and "Shares"; Current Value (ILS), Day Change (%), Total Return (%), Sector, Role Group,
// Total Cost (ILS); USD/ILS in Q1, its change in Q2), plus extended-hours moves from Investment
// Universe's Ext columns. Returns the percentages-only {holdings, summary} the prompt takes.
// The portfolio's day change is the size-weighted change of the holdings themselves, so it always
// agrees with them (the dashboard's own figure resets to 0 overnight, until the next open).
function snapshotFromTables(tracker, universe, now = new Date()) {
  const headerIdx = tracker.findIndex((r) => r.includes('Ticker') && r.includes('Shares'));
  if (headerIdx === -1) throw new Error('Could not find the Tracker holdings header row');
  const h = tracker[headerIdx].map((x) => String(x).trim());
  const c = (name) => h.indexOf(name);
  const cTicker = c('Ticker'), cValue = c('Current Value (ILS)'), cDay = c('Day Change (%)'), cRet = c('Total Return (%)'),
    cSector = c('Sector'), cRole = c('Role Group'), cCost = c('Total Cost (ILS)'), cPrice = c('Current Price');
  if ([cTicker, cValue, cDay, cRet, cSector].some((i) => i === -1)) throw new Error('Tracker columns have changed -- expected headers not found');

  const rows = [];
  for (let i = headerIdx + 1; i < tracker.length; i++) {
    const r = tracker[i];
    const ticker = String(r[cTicker] || '').trim();
    if (!ticker) break;
    const value = parseNum(r[cValue]);
    if (!(value > 0) || (cPrice >= 0 && !(parseNum(r[cPrice]) > 0))) continue;  // no live price (#N/A)
    rows.push({
      ticker, value, cost: cCost >= 0 ? parseNum(r[cCost]) : NaN,
      day: parseNum(r[cDay]), ret: parseNum(r[cRet]),
      sector: (String(r[cSector] || 'Other').trim() || 'Other').split('/')[0].trim(),
      role: cRole >= 0 ? (String(r[cRole] || 'Other').trim() || 'Other') : 'Other',
      tase: TASE_TICKERS.includes(ticker),
    });
  }
  if (!rows.length) throw new Error('No holdings with a live price in the Tracker');
  const total = rows.reduce((s, r) => s + r.value, 0);
  const cost = rows.reduce((s, r) => s + (r.cost > 0 ? r.cost : r.value), 0);
  const usa = rows.filter((r) => !r.tase).reduce((s, r) => s + r.value, 0);
  const dayPct = rows.reduce((s, r) => s + (Number.isFinite(r.day) ? r.value / total * r.day : 0), 0);

  // USD/ILS: Q1 = the rate; Q2 = "Sep 21, 18:00  USD / ILS 3.014 ▼-0.0074 (-0.0226)" -- the change is
  // a percent when it carries "%", else a fraction.
  const q1 = tracker[0] ? tracker[0][16] : undefined, q2 = tracker[1] ? tracker[1][16] : undefined;
  const fx = parseNum(q1);
  const m = typeof q2 === 'string' ? /USD\s*\/\s*ILS\s+[\d.]+\s*[▲▼]?\s*([+-][\d.]+)(%)?/.exec(q2) : null;
  const fxChangePct = m ? parseFloat(m[1]) * (m[2] ? 1 : 100) : NaN;

  // Extended-hours moves (ExtendedHours.gs): a pre-market one only until that day's 09:30 New York.
  const ext = {};
  let extSession = '';
  if (universe && universe.length) {
    const uh = universe[0].map((x) => String(x).trim());
    const uT = uh.indexOf('Ticker'), uPct = uh.indexOf('Ext Change %'), uS = uh.indexOf('Ext Session'), uU = uh.indexOf('Ext Updated');
    const ny = zoneParts('America/New_York', now);
    if (uT !== -1 && uPct !== -1 && uS !== -1) {
      for (let i = 1; i < universe.length; i++) {
        const r = universe[i];
        const pct = parseNum(r[uPct]), session = String(r[uS] || '').trim();
        if (!Number.isFinite(pct) || !session) continue;
        if (session === 'Pre-market') {
          const upd = uU !== -1 ? new Date(String(r[uU] || '')) : null;
          if (!upd || isNaN(upd) || zoneParts('America/New_York', upd).date !== ny.date || ny.hm >= 930) continue;
        }
        ext[String(r[uT] || '').trim().split(':').pop().toUpperCase()] = pct;
        extSession = extSession || session;
      }
    }
  }

  const holdings = rows.map((r) => ({
    ticker: r.ticker, market: r.tase ? 'TASE' : 'USA', weightPct: r.value / total * 100,
    returnPct: r.ret, dayPct: r.day, extPct: r.tase ? null : ext[r.ticker.split(':').pop().toUpperCase()],
    sector: r.sector, role: r.role,
  }));
  const summary = {
    asOf: now.toLocaleString('en-GB', { timeZone: 'Asia/Jerusalem', day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }),
    dayPct, totalReturnPct: cost ? (total - cost) / cost * 100 : NaN,
    usaWeightPct: usa / total * 100, taseWeightPct: (total - usa) / total * 100,
    usdIls: fx, usdIlsChangePct: fxChangePct,
    taseOpen: taseOpenNow(now), usaOpen: usaOpenNow(now), extSession,
  };
  return { holdings: sanitizeHoldings(holdings), summary: sanitizeSummary(summary) };
}

function briefQuestion(session) {
  return `Write the daily brief for ${dayLabel(session.date)}`
    + (session.live ? ' -- so far today; at least one market is still trading' : ' -- the finished trading day')
    + ', about 150-200 words. The portfolio day change here is the size-weighted day change of the holdings in their own currencies; the USA holdings\' ILS value also moves with USD/ILS.'
    + ' Use the sections: Today (the portfolio\'s day change, and TASE vs. USA), Biggest movers (the holdings with the largest effect on the portfolio, weight x day change -- up and down),'
    + ' Currency (how the USD/ILS move affects the USA holdings), and Worth a look (anything unusual, including after-hours or pre-market moves if the data has them).';
}

function loadBriefs() {
  try { return JSON.parse(fs.readFileSync(BRIEF_FILE, 'utf8')); } catch (e) { return []; }
}
function saveBrief(brief) {
  const all = loadBriefs().filter((b) => b.date !== brief.date);
  all.push(brief);
  all.sort((a, b) => (a.date < b.date ? 1 : -1));
  fs.writeFileSync(BRIEF_FILE, JSON.stringify(all.slice(0, BRIEF_KEEP), null, 2));
}

// Reads the sheet, asks Gemini, saves the brief. Resolves {ok: true, brief} or {ok: false, httpStatus, error}.
async function writeBrief(kind, now = new Date()) {
  const session = currentSession(now);
  let snap;
  try {
    const sheets = await sheetsClient();
    const [tracker, universe] = await Promise.all([readTab(sheets, 'Tracker'), readTab(sheets, 'Investment Universe').catch(() => [])]);
    snap = snapshotFromTables(tracker, universe, now);
  } catch (err) {
    console.error(new Date().toISOString(), 'daily brief: sheet read failed:', err.message);
    return { ok: false, httpStatus: 502, error: 'Could not read the Tracker sheet: ' + err.message };
  }
  const g = await generate(buildPrompt(snap.holdings, snap.summary, briefQuestion(session)));
  if (!g.ok) return g;
  const brief = {
    date: session.date, label: dayLabel(session.date), live: session.live, final: !session.live,
    kind, text: g.text, model: g.model, generatedAt: new Date().toISOString(),
  };
  saveBrief(brief);
  console.log(new Date().toISOString(), `daily brief written for ${brief.date} (${kind}, ${g.model})`);
  return { ok: true, brief };
}

// 23:50 Monday-Friday (Israel): write that day's final brief. Checked every minute; if Gemini is
// busy it tries again every 5 minutes until 03:00 -- after midnight still for the day just ended.
let lastScheduledTry = 0;
function startBriefScheduler() {
  if (!process.env.GEMINI_API_KEY) return;
  setInterval(async () => {
    const p = zoneParts('Asia/Jerusalem');
    const target = p.hm >= 2350 ? p.date : p.hm < 300 ? shiftDate(p.date, -1) : null;
    if (!target || !isWeekdayDate(target)) return;
    if (loadBriefs().some((b) => b.date === target && b.final)) return;
    if (Date.now() - lastScheduledTry < 5 * 60000) return;
    lastScheduledTry = Date.now();
    const r = await writeBrief('scheduled');
    if (!r.ok) console.error(new Date().toISOString(), 'daily brief: scheduled run failed:', r.error);
  }, 60000).unref();
}
startBriefScheduler();

// ---------------------------------------------------------------------------------------------
// Routes
// ---------------------------------------------------------------------------------------------

async function handleAiRoutes(req, res, pathname, { send, readJsonBody }) {
  const json = (status, obj) => send(res, req, status, JSON.stringify(obj), 'application/json');
  const url = new URL(req.url || pathname, 'http://local');

  if (pathname === '/ai/status' && req.method === 'GET') {
    json(200, { configured: !!process.env.GEMINI_API_KEY, model: MODEL, usedToday: usedToday(), limit: DAILY_LIMIT });
    return true;
  }

  if (pathname === '/ai/brief' && req.method === 'GET') {
    const all = loadBriefs();
    const want = url.searchParams.get('date');
    json(200, { brief: (want ? all.find((b) => b.date === want) : all[0]) || null, dates: all.map((b) => b.date) });
    return true;
  }

  if (pathname !== '/ai/analyze' && pathname !== '/ai/brief/refresh') return false;
  if (req.method !== 'POST') { json(405, { error: 'Use POST' }); return true; }
  if (!process.env.GEMINI_API_KEY) {
    json(503, { error: 'AI is not set up on the server: add GEMINI_API_KEY to ~/sheet-proxy/.env and restart sheet-proxy.' });
    return true;
  }
  if (usedToday() >= DAILY_LIMIT) {
    json(429, { error: `Daily AI limit reached (${DAILY_LIMIT}). It resets at midnight Israel time.` });
    return true;
  }

  // Only answered requests count toward the daily limit -- retrying while Google's free tier is
  // "busy" mustn't use it up.
  if (pathname === '/ai/brief/refresh') {
    const r = await writeBrief('manual');
    if (!r.ok) { json(r.httpStatus, { error: r.error }); return true; }
    usage.count++;
    json(200, { brief: r.brief, dates: loadBriefs().map((b) => b.date), usedToday: usage.count, limit: DAILY_LIMIT });
    return true;
  }

  let body;
  try { body = await readJsonBody(req); } catch (e) { json(400, { error: 'Invalid JSON' }); return true; }
  const holdings = sanitizeHoldings(body.holdings);
  if (!holdings.length) { json(400, { error: 'No holdings sent' }); return true; }
  const summary = sanitizeSummary(body.summary);
  const question = text(body.question, MAX_QUESTION_CHARS);
  const g = await generate(buildPrompt(holdings, summary, question));
  if (!g.ok) { json(g.httpStatus, { error: g.error }); return true; }
  usage.count++;
  json(200, { text: g.text, model: g.model, usedToday: usage.count, limit: DAILY_LIMIT });
  return true;
}

module.exports = { handleAiRoutes, sanitizeHoldings, sanitizeSummary, buildPrompt, snapshotFromTables, currentSession, briefQuestion };
