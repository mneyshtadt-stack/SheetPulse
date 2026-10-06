// Yahoo Finance prices for the Israeli funds (IBI.F35, IBI.FK4), all day, into the TasePrices tab.
//
// Why: GOOGLEFINANCE gives these funds' last continuous trade, not TASE's closing-auction price,
// so after the close the Tracker kept e.g. IBI.F35 at 7465 while the official close (and Meitav,
// and Yahoo) was 7452 (broker comparison, 6 Oct 2026). Yahoo's chart API has the live price during
// the session (delayed about 20 minutes, like Google's) and the auction close after it, in agorot
// like GOOGLEFINANCE.
//
// The Tracker uses Yahoo's figures only while they're fresh (updated in the last 30 minutes);
// otherwise, e.g. if Yahoo is unreachable or the trigger stopped, it falls back to GOOGLEFINANCE,
// so a price can never freeze. Tracker formulas (IBI.F35's row shown; same for IBI.FK4's row):
//   Price:          =IFERROR(IF(NOW()-VLOOKUP(A23,TasePrices!A:F,6,FALSE)<1/48, VLOOKUP(A23,TasePrices!A:B,2,FALSE)/1, GOOGLEFINANCE(A23,"price")), GOOGLEFINANCE(A23,"price"))
//   Prev Day Close: =IFERROR(IF(NOW()-VLOOKUP(A23,TasePrices!A:F,6,FALSE)<1/48, VLOOKUP(A23,TasePrices!A:C,3,FALSE)/1, GOOGLEFINANCE(A23,"closeyest")), GOOGLEFINANCE(A23,"closeyest"))
//   Day Change %:   =IFERROR(IF(NOW()-VLOOKUP(A23,TasePrices!A:F,6,FALSE)<1/48, VLOOKUP(A23,TasePrices!A:D,4,FALSE)/1, GOOGLEFINANCE(A23,"changepct")), GOOGLEFINANCE(A23,"changepct"))
//
// Setup: run updateTasePrices() once (creates the tab), then installTasePricesTrigger() once.

const TASE_PRICES_SHEET = 'TasePrices';
const TASE_PRICES_TICKERS = ['IBI.F35', 'IBI.FK4'];   // Google symbols; Yahoo: IBI-F35.TA, IBI-FK4.TA
const TASE_PRICES_HEADER = ['Ticker', 'Price', 'Prev Close', 'Change %', 'Quote time', 'Updated'];

function updateTasePrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TASE_PRICES_SHEET) || ss.insertSheet(TASE_PRICES_SHEET);
  sh.getRange(1, 1, 1, TASE_PRICES_HEADER.length).setValues([TASE_PRICES_HEADER]);
  const now = new Date();
  TASE_PRICES_TICKERS.forEach((t, i) => {
    const q = tasePricesYahoo_(t);
    // Yahoo unreachable: leave the last row as it is; its Updated time ages past 30 minutes and
    // the Tracker falls back to GOOGLEFINANCE on its own.
    if (!q) return;
    sh.getRange(2 + i, 1, 1, TASE_PRICES_HEADER.length)
      .setValues([[t, q.price, q.prevClose, (q.price / q.prevClose - 1) * 100, q.time, now]]);
  });
}

// Every 5 minutes, like the intraday log. Safe to run again: it replaces this script's own trigger only.
function installTasePricesTrigger() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'updateTasePrices') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('updateTasePrices').timeBased().everyMinutes(5).create();
  Logger.log('updateTasePrices now runs every 5 minutes.');
}

// Yahoo's latest regular-session price (the closing-auction price once the session has ended),
// the previous close and the quote time, for a Google-style TASE symbol. Null on any failure.
function tasePricesYahoo_(googleSymbol) {
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
