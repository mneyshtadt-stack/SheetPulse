// Pre-market and after-hours prices for Market Universe's price card, read from Google Finance's
// own quote page (google.com/finance/quote/SYM:EXCHANGE). GOOGLEFINANCE() has no extended-hours
// attribute -- its "price" stays at the last regular close until the next open -- but the quote
// page shows the extended-hours trade right under the regular price, and that part is plain HTML.
//
// Google Finance is the primary source; for any ticker it gives nothing for (page failed, no
// extended-hours block, or a layout change broke the parsing) Yahoo Finance's chart data is tried
// as a backup. Yahoo is unofficial and can throttle, so it's never the first choice.
//
// Writes six columns into the Investment Universe tab (created at its right edge if missing), for
// each current US holding in the Tracker:
//   Ext Price | Ext Change | Ext Change % | Ext Session | Ext Updated | Ext Source
//   (vs. the last regular close; "Pre-market" / "After hours"; ISO time of the reading, UTC;
//    "Google" / "Yahoo")
// Pre-market values are cleared once the regular session opens; after-hours values stay up
// overnight (and over the weekend) until the next pre-market replaces them -- as Google shows them.
//
// Setup (once, in the Apps Script project bound to the Tracker spreadsheet):
//   1. Paste this file as ExtendedHours.gs.
//   2. Run updateExtendedHours() once by hand (approve the "connect to an external service"
//      permission). Outside pre-market / after-hours it just returns.
//   3. Run installExtendedHoursTrigger() -- repeats it every 5 minutes.

const EXT_SHEET = 'Investment Universe';
const EXT_HEADERS = ['Ext Price', 'Ext Change', 'Ext Change %', 'Ext Session', 'Ext Updated', 'Ext Source'];
// Tried in this order to find a bare ticker's listing; the answer is remembered per ticker.
const EXT_EXCHANGES = ['NASDAQ', 'NYSEARCA', 'NYSE', 'BATS', 'NYSEAMERICAN'];
const EXT_BATCH = 6;             // pages fetched in parallel (each is ~1.3 MB)
const EXT_MAX_DISCOVER = 8;      // new tickers whose exchange is looked up per run

// 'pre' 04:00-09:30, 'regular' 09:30-16:00, 'post' 16:00-20:05 New York time, Mon-Fri; else 'closed'.
// The extra 5 minutes after 20:00 let a drifting trigger still catch the final after-hours trade.
function extSessionNow_() {
  const now = new Date();
  const wd = Utilities.formatDate(now, 'America/New_York', 'EEE');
  const hm = Utilities.formatDate(now, 'America/New_York', 'HHmm');
  if (['Mon', 'Tue', 'Wed', 'Thu', 'Fri'].indexOf(wd) === -1) return 'closed';
  if (hm >= '0400' && hm < '0930') return 'pre';
  if (hm >= '0930' && hm < '1600') return 'regular';
  if (hm >= '1600' && hm <= '2005') return 'post';
  return 'closed';
}

function updateExtendedHours() {
  const session = extSessionNow_();
  if (session === 'closed') return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXT_SHEET);
  if (!sheet) throw new Error('No "' + EXT_SHEET + '" tab');
  const props = PropertiesService.getScriptProperties();

  if (session === 'regular') {
    // A pre-market figure means nothing once the real session trades -- clear it, once a day.
    const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
    if (props.getProperty('ext_cleared_date') !== today) {
      extWrite_(sheet, {});
      props.setProperty('ext_cleared_date', today);
    }
    return;
  }

  const symbols = extResolveSymbols_(extUsHoldings_(ss), props);
  const results = {};
  const nowIso = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const put = (ticker, q, source) => {
    results[extKey_(ticker)] = [q.price, q.change, q.pct, q.session, nowIso, source];
  };
  for (let i = 0; i < symbols.length; i += EXT_BATCH) {
    const batch = symbols.slice(i, i + EXT_BATCH);
    const responses = UrlFetchApp.fetchAll(batch.map(s => extRequest_(s.url)));
    responses.forEach((res, j) => {
      if (res.getResponseCode() !== 200) return;
      const q = extParse_(res.getContentText());
      if (q) put(batch[j].ticker, q, 'Google');
    });
  }

  // Backup: Yahoo for whatever Google gave nothing for. Tickers Google knows no US listing for
  // (the TASE funds) aren't in symbols at all, so they're never sent here.
  const missing = symbols.filter(s => !results[extKey_(s.ticker)]);
  let fromYahoo = 0;
  for (let i = 0; i < missing.length; i += EXT_BATCH) {
    const batch = missing.slice(i, i + EXT_BATCH);
    const responses = UrlFetchApp.fetchAll(batch.map(s => ({url: extYahooUrl_(s.ticker), muteHttpExceptions: true})));
    responses.forEach((res, j) => {
      if (res.getResponseCode() !== 200) return;
      let json;
      try { json = JSON.parse(res.getContentText()); } catch (e) { return; }
      const q = extParseYahoo_(json, session);
      if (q) { put(batch[j].ticker, q, 'Yahoo'); fromYahoo++; }
    });
  }

  extWrite_(sheet, results);
  console.log('Extended hours (' + session + '): ' + Object.keys(results).length + ' of '
    + symbols.length + ' tickers have a price (' + fromYahoo + ' from Yahoo)');
}

// Every 5 minutes; the function itself decides whether it's pre-market / after-hours.
// Safe to re-run -- replaces any existing one.
function installExtendedHoursTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateExtendedHours')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateExtendedHours').timeBased().everyMinutes(5).create();
}

// Forgets every remembered exchange (e.g. after a ticker moved listing, or to retry one that
// wasn't found). The next runs look them up again.
function resetExtendedHoursCache() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties())
    .filter(k => k.indexOf('ext_exch_') === 0)
    .forEach(k => props.deleteProperty(k));
}

// "BATS:COWZ" and "COWZ" name the same row -- Tracker and Investment Universe may write it either way.
function extKey_(ticker) {
  const i = ticker.indexOf(':');
  return (i === -1 ? ticker : ticker.slice(i + 1)).toUpperCase();
}

function extRequest_(url) {
  return {url: url, muteHttpExceptions: true, followRedirects: true, headers: {'Accept-Language': 'en-US'}};
}

function extQuoteUrl_(sym, exchange) {
  return 'https://www.google.com/finance/quote/' + encodeURIComponent(sym) + ':' + exchange;
}

// The Tracker's current holdings, by ticker as written there (e.g. "SMH", "BATS:COWZ").
function extUsHoldings_(ss) {
  const data = ss.getSheetByName('Tracker').getDataRange().getValues();
  let hdr = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].indexOf('Ticker') !== -1 && data[i].indexOf('Shares') !== -1) { hdr = i; break; }
  }
  if (hdr === -1) throw new Error('Could not find the Tracker holdings header row');
  const cTicker = data[hdr].indexOf('Ticker');
  const out = [];
  for (let r = hdr + 1; r < data.length; r++) {
    const t = String(data[r][cTicker] || '').trim();
    if (!t) break;
    if (out.indexOf(t) === -1) out.push(t);
  }
  return out;
}

// Ticker -> quote-page URL. "EXCH:SYM" (Sheets style) is used as given; a bare ticker's exchange
// is found by trying EXT_EXCHANGES and keeping the page titled "... (SYM) Price ..." -- remembered in
// Script Properties, including "NONE" for tickers not listed on a US exchange (the TASE funds).
function extResolveSymbols_(tickers, props) {
  const out = [];
  let discovered = 0;
  tickers.forEach(ticker => {
    const i = ticker.indexOf(':');
    if (i !== -1) {
      out.push({ticker: ticker, url: extQuoteUrl_(ticker.slice(i + 1), ticker.slice(0, i).toUpperCase())});
      return;
    }
    const key = 'ext_exch_' + ticker;
    let exch = props.getProperty(key);
    if (!exch) {
      if (discovered >= EXT_MAX_DISCOVER) return;  // the rest get looked up on later runs
      discovered++;
      const responses = UrlFetchApp.fetchAll(EXT_EXCHANGES.map(e => extRequest_(extQuoteUrl_(ticker, e))));
      const title = '(' + ticker.toUpperCase() + ')';
      const hit = responses.findIndex(res => {
        if (res.getResponseCode() !== 200) return false;
        const m = /<title>([^<]*)<\/title>/.exec(res.getContentText());
        return !!m && m[1].indexOf(title) !== -1;
      });
      exch = hit === -1 ? 'NONE' : EXT_EXCHANGES[hit];
      props.setProperty(key, exch);
    }
    if (exch !== 'NONE') out.push({ticker: ticker, url: extQuoteUrl_(ticker, exch)});
  });
  return out;
}

// Reads the extended-hours block off a quote page: its "Pre-market" / "After hours" label, and the
// two prices right before it -- the last regular close, then the extended-hours trade. Found by
// that label rather than Google's generated class names, which change without notice. Returns
// null when the page has no such block (no trade yet, or a regular session).
function extParse_(html) {
  const m = /(?:>|<\/i>)(Pre-market|After[ -]?hours)(?=&nbsp;|<)/i.exec(html);
  if (!m) return null;
  const before = html.slice(Math.max(0, m.index - 2500), m.index);
  const prices = [];
  const re = /<span>\$([0-9,]+(?:\.[0-9]+)?)<\/span>/g;
  let p;
  while ((p = re.exec(before)) !== null) prices.push(Number(p[1].replace(/,/g, '')));
  if (prices.length < 2) return null;
  const regular = prices[prices.length - 2], ext = prices[prices.length - 1];
  if (!(regular > 0 && ext > 0) || Math.abs(ext / regular - 1) > 0.5) return null;
  const change = ext - regular;
  return {
    price: ext,
    change: Math.round(change * 100) / 100,
    pct: Math.round(change / regular * 10000) / 100,
    session: /^pre/i.test(m[1]) ? 'Pre-market' : 'After hours',
  };
}

function extYahooUrl_(ticker) {
  return 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(extKey_(ticker))
    + '?interval=5m&range=1d&includePrePost=true';
}

// Yahoo's 5-minute chart for today, extended hours included: the last bar inside the current
// session's pre-market / after-hours window, against regularMarketPrice -- which during pre-market
// is still yesterday's close and during after-hours today's close, the same reference Google uses.
function extParseYahoo_(json, session) {
  const res = json && json.chart && json.chart.result && json.chart.result[0];
  if (!res || !res.meta || !res.meta.currentTradingPeriod) return null;
  const period = res.meta.currentTradingPeriod[session === 'pre' ? 'pre' : 'post'];
  const ts = res.timestamp || [];
  const quote = res.indicators && res.indicators.quote && res.indicators.quote[0];
  const close = (quote && quote.close) || [];
  let ext = null;
  for (let i = 0; i < ts.length; i++) {
    if (period && ts[i] >= period.start && ts[i] < period.end && typeof close[i] === 'number') ext = close[i];
  }
  const regular = res.meta.regularMarketPrice;
  if (!(ext > 0 && regular > 0) || Math.abs(ext / regular - 1) > 0.5) return null;
  const change = ext - regular;
  return {
    price: Math.round(ext * 100) / 100,
    change: Math.round(change * 100) / 100,
    pct: Math.round(change / regular * 10000) / 100,
    session: session === 'pre' ? 'Pre-market' : 'After hours',
  };
}

// Writes results ({SYMBOL: [price, change, pct, session, updated, source]}, keyed by extKey_) into the Ext columns; every
// other row's Ext cells are blanked, so a ticker without a trade this run shows nothing.
function extWrite_(sheet, results) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const cTicker = header.indexOf('Ticker');
  if (cTicker === -1) throw new Error('No "Ticker" column in "' + EXT_SHEET + '"');
  const cols = EXT_HEADERS.map(h => {
    let c = header.indexOf(h);
    if (c === -1) {
      c = header.length;
      header.push(h);
      sheet.getRange(1, c + 1).setValue(h).setFontWeight('bold');
    }
    return c;
  });
  const rows = sheet.getLastRow() - 1;
  if (rows < 1) return;
  const tickers = sheet.getRange(2, cTicker + 1, rows, 1).getValues().map(r => String(r[0]).trim());
  EXT_HEADERS.forEach((h, k) => {
    const values = tickers.map(t => { const r = t && results[extKey_(t)]; return [r ? r[k] : '']; });
    const range = sheet.getRange(2, cols[k] + 1, rows, 1);
    if (h === 'Ext Updated' || h === 'Ext Session' || h === 'Ext Source') range.setNumberFormat('@');
    range.setValues(values);
  });
}
