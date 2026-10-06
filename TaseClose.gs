// Official TASE closing prices for the Israeli funds, from Yahoo Finance, into the TaseClose tab.
//
// Why: GOOGLEFINANCE gives these funds' last continuous trade, not TASE's closing-auction price,
// so after the close the Tracker kept e.g. IBI.F35 at 7465 while the official close (and Meitav,
// and Yahoo) was 7452 (broker comparison, 6 Oct 2026). Yahoo's chart API has the auction close
// (price, previous close, quote time), in agorot like GOOGLEFINANCE.
//
// While TASE is in session (Mon-Thu 10:00-17:30, Fri 10:00-14:00, Israel time) the price cells
// are left empty, so the Tracker formulas fall back to GOOGLEFINANCE's live price; outside the
// session they hold Yahoo's official close. Tracker formulas (row of IBI.F35 shown; same for FK4):
//   Price:          =IFERROR(VLOOKUP(A23, TaseClose!A:B, 2, FALSE)/1, GOOGLEFINANCE(A23, "price"))
//   Prev Day Close: =IFERROR(VLOOKUP(A23, TaseClose!A:C, 3, FALSE)/1, GOOGLEFINANCE(A23, "closeyest"))
//   Day Change %:   =IFERROR(VLOOKUP(A23, TaseClose!A:D, 4, FALSE)/1, GOOGLEFINANCE(A23, "changepct"))
// (An empty cell makes the "/1" an error, so IFERROR falls back to GOOGLEFINANCE.)
//
// Setup: run updateTaseCloses() once (creates the tab), then installTaseCloseTrigger() once.

const TASE_CLOSE_SHEET = 'TaseClose';
const TASE_CLOSE_TICKERS = ['IBI.F35', 'IBI.FK4'];   // Google symbols; Yahoo: IBI-F35.TA, IBI-FK4.TA
const TASE_CLOSE_HEADER = ['Ticker', 'Price', 'Prev Close', 'Change %', 'Quote time', 'Updated'];

function updateTaseCloses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TASE_CLOSE_SHEET) || ss.insertSheet(TASE_CLOSE_SHEET);
  const now = new Date();
  const inSession = taseInSession_(now);
  const rows = TASE_CLOSE_TICKERS.map(t => {
    const q = inSession ? null : taseYahooClose_(t);
    // In session, or Yahoo unreachable: empty cells, so the Tracker uses GOOGLEFINANCE.
    if (!q) return [t, '', '', '', '', now];
    return [t, q.price, q.prevClose, (q.price / q.prevClose - 1) * 100, q.time, now];
  });
  sh.getRange(1, 1, 1, TASE_CLOSE_HEADER.length).setValues([TASE_CLOSE_HEADER]);
  sh.getRange(2, 1, rows.length, TASE_CLOSE_HEADER.length).setValues(rows);
  Logger.log((inSession ? 'TASE in session: cells cleared (GOOGLEFINANCE in use). ' : 'Official closes written. ') + JSON.stringify(rows));
}

// Every 10 minutes. Safe to run again: it replaces this script's own trigger only.
function installTaseCloseTrigger() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'updateTaseCloses') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('updateTaseCloses').timeBased().everyMinutes(10).create();
  Logger.log('updateTaseCloses now runs every 10 minutes.');
}

// TASE's regular session in Israel time (the same hours as the dashboard's getTaseStatus).
function taseInSession_(now) {
  const tz = 'Asia/Jerusalem';
  const wd = Utilities.formatDate(now, tz, 'EEE');
  const mins = Number(Utilities.formatDate(now, tz, 'H')) * 60 + Number(Utilities.formatDate(now, tz, 'm'));
  if (wd === 'Fri') return mins >= 600 && mins < 840;
  return ['Mon', 'Tue', 'Wed', 'Thu'].indexOf(wd) !== -1 && mins >= 600 && mins < 1050;
}

// Yahoo's latest regular-session price (the closing-auction price once the session has ended),
// the previous close and the quote time, for a Google-style TASE symbol. Null on any failure.
function taseYahooClose_(googleSymbol) {
  const symbol = googleSymbol.replace(/\./g, '-') + '.TA';
  try {
    const resp = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol), {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json'
      }
    });
    if (resp.getResponseCode() !== 200) return null;
    const r = JSON.parse(resp.getContentText()).chart.result[0], m = r && r.meta;
    const price = m && m.regularMarketPrice, prev = m && (m.chartPreviousClose != null ? m.chartPreviousClose : m.previousClose);
    if (!(price > 0) || !(prev > 0)) return null;
    return {price: price, prevClose: prev, time: m.regularMarketTime ? new Date(m.regularMarketTime * 1000) : ''};
  } catch (e) {
    return null;
  }
}
