// AI analysis for SheetPulse -- the dashboard's "AI analysis" panel, answered by Google Gemini.
//
// Mounted in index.js after the sign-in check, like the MedTracker routes, so only the one allowed
// account can use it:
//   POST /ai/analyze  {holdings: [...], summary: {...}, question?}  ->  {text, model, usedToday, limit}
//   GET  /ai/status                                                 ->  {configured, model, usedToday, limit}
//
// Uses Gemini's free tier, where Google may use what it's sent to improve its products -- so only
// percentages ever leave this server. The dashboard sends just weights, returns and changes, and
// sanitizeHoldings / sanitizeSummary below keep a fixed whitelist of fields and drop anything else,
// so no ILS / USD amount, total value or share count can reach Google even if one were sent.
//
// .env:
//   GEMINI_API_KEY          required (from aistudio.google.com -> API keys; no billing = free tier)
//   GEMINI_MODEL            optional, default gemini-3.8-flash
//   GEMINI_FALLBACK_MODEL   optional, default gemini-3.5-flash-lite (tried once when the main model
//                           is busy, over its free quota, or unavailable)
//   AI_DAILY_LIMIT          optional, default 40 analyses per day (Israel time)

const MODEL = process.env.GEMINI_MODEL || 'gemini-3.8-flash';
const FALLBACK_MODEL = process.env.GEMINI_FALLBACK_MODEL || 'gemini-3.5-flash-lite';
const DAILY_LIMIT = Number(process.env.AI_DAILY_LIMIT) || 40;
const TIMEOUT_MS = 90000;
const MAX_HOLDINGS = 200;
const MAX_QUESTION_CHARS = 500;

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

async function handleAiRoutes(req, res, pathname, { send, readJsonBody }) {
  const json = (status, obj) => send(res, req, status, JSON.stringify(obj), 'application/json');

  if (pathname === '/ai/status' && req.method === 'GET') {
    json(200, { configured: !!process.env.GEMINI_API_KEY, model: MODEL, usedToday: usedToday(), limit: DAILY_LIMIT });
    return true;
  }
  if (pathname !== '/ai/analyze') return false;
  if (req.method !== 'POST') { json(405, { error: 'Use POST' }); return true; }
  if (!process.env.GEMINI_API_KEY) {
    json(503, { error: 'AI is not set up on the server: add GEMINI_API_KEY to ~/sheet-proxy/.env and restart sheet-proxy.' });
    return true;
  }
  if (usedToday() >= DAILY_LIMIT) {
    json(429, { error: `Daily AI limit reached (${DAILY_LIMIT}). It resets at midnight Israel time.` });
    return true;
  }

  let body;
  try { body = await readJsonBody(req); } catch (e) { json(400, { error: 'Invalid JSON' }); return true; }
  const holdings = sanitizeHoldings(body.holdings);
  if (!holdings.length) { json(400, { error: 'No holdings sent' }); return true; }
  const summary = sanitizeSummary(body.summary);
  const question = text(body.question, MAX_QUESTION_CHARS);
  const prompt = buildPrompt(holdings, summary, question);

  usage.count++;
  try {
    let model = MODEL;
    let r = await callGemini(model, prompt);
    // Busy, over the free quota, or the model name retired -> one try on the lighter model.
    if ([404, 429, 500, 503].includes(r.status) && FALLBACK_MODEL && FALLBACK_MODEL !== MODEL) {
      model = FALLBACK_MODEL;
      r = await callGemini(model, prompt);
    }
    if (r.status !== 200) {
      const msg = (r.data && r.data.error && r.data.error.message) || `HTTP ${r.status}`;
      console.error(new Date().toISOString(), 'gemini error:', r.status, msg);
      json(r.status === 429 ? 429 : 502, { error: r.status === 429 ? 'Gemini free-tier limit reached -- try again in a minute.' : 'Gemini error: ' + msg });
      return true;
    }
    const answer = extractText(r.data);
    if (!answer) {
      const reason = (r.data && r.data.promptFeedback && r.data.promptFeedback.blockReason)
        || (r.data && r.data.candidates && r.data.candidates[0] && r.data.candidates[0].finishReason) || 'empty answer';
      json(502, { error: 'Gemini returned no text (' + reason + ')' });
      return true;
    }
    json(200, { text: answer, model, usedToday: usage.count, limit: DAILY_LIMIT });
  } catch (err) {
    console.error(new Date().toISOString(), 'gemini request failed:', err.message);
    json(504, { error: err.name === 'TimeoutError' ? 'Gemini took too long to answer -- try again.' : 'Could not reach Gemini: ' + err.message });
  }
  return true;
}

module.exports = { handleAiRoutes, sanitizeHoldings, sanitizeSummary, buildPrompt };
