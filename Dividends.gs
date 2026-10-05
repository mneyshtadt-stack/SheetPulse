// Dividend history for the dashboard: dividend income per holding (both brokers), and Psagot's own
// average cost, which subtracts every dividend once paid, after 25% Israeli tax.
//
// One tab, written by this script and read by the dashboard:
//   Dividends -- Ticker | Ex-Date | Pay Date | Amount | Source | Added
//                One row per dividend per share, for every US ticker in the Transactions tab (both
//                brokers), from the first purchase on. Dates are "yyyy-mm-dd" text.
//
// Source: Yahoo Finance's chart API (ex-date and amount). Yahoo has no payment date, so Pay Date is
// estimated from the ex-date (divPayLagDays_): ETFs pay within days, stocks weeks later. Checked
// against Psagot's statements of 4 and 5 Oct 2026 (38 holdings): 35 matched to the cent; the rest were
// dividends Yahoo hadn't listed yet (DBMF, VO) or adjusted for a spin-off (SPGI). Rows whose Source says "estimated" get their
// Pay Date recalculated on every run (so a better rule fixes old rows too); change Source to
// "manual" after correcting a Pay Date or Amount by hand from the broker's statement and it's never
// touched again. TASE funds (IBI.*) are skipped: no USD dividends.
//
// Setup (once, in the Apps Script project bound to the Tracker spreadsheet):
//   1. Paste this file as Dividends.gs and run updateDividends() -- creates and fills the tab.
//   2. Run installDividendsTrigger() -- repeats it every morning at 08:00.

const DIV_SHEET = 'Dividends';
const DIV_TX_SHEET = 'Transactions';
const DIV_HEADER = ['Ticker', 'Ex-Date', 'Pay Date', 'Amount', 'Source', 'Added'];
const DIV_ESTIMATED = 'Yahoo, pay date estimated';
// Days from ex-date to payment. ETFs: 1-6 (JEPI and PFF 6+, SCHD 6, sector SPDRs 2, DBMF 1). Stocks: 14-32 (BAM 29,
// GS 28, MCO 26, SPGI 19, MS 15). Per-ticker exceptions where the type misleads.
const DIV_LAG_ETF = 6;
const DIV_LAG_STOCK = 21;
const DIV_LAG_TICKER = {ASGI: 8, POWR: 3, VCSH: 2};   // closed-end fund; an ETF Yahoo calls a stock; Vanguard bond ETF paying in 2 days

function divPayLagDays_(ticker, type) {
  if (DIV_LAG_TICKER[ticker] != null) return DIV_LAG_TICKER[ticker];
  return type === 'ETF' ? DIV_LAG_ETF : DIV_LAG_STOCK;
}

function updateDividends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const firstBuy = divFirstBuyDates_(ss);
  const sh = ss.getSheetByName(DIV_SHEET) || ss.insertSheet(DIV_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(DIV_HEADER);
  sh.getRange('A:C').setNumberFormat('@');   // dates stay text, never reformatted by the sheet

  const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, DIV_HEADER.length).getValues() : [];
  const have = {};
  existing.forEach(r => { have[String(r[0]).trim() + '|' + String(r[1]).trim()] = true; });

  const today = fmt(new Date());
  const added = [], types = {};
  Object.keys(firstBuy).sort().forEach(ticker => {
    const y = divFetchYahoo_(ticker);
    if (!y) { Logger.log('No Yahoo data for ' + ticker); return; }
    types[ticker] = y.type;
    const lag = divPayLagDays_(ticker, y.type);
    const from = firstBuy[ticker].getTime() - 86400000;
    y.divs.forEach(d => {
      if (d.date.getTime() < from) return;
      const ex = fmt(d.date);
      if (have[ticker + '|' + ex]) return;
      added.push([ticker, ex, fmt(new Date(d.date.getTime() + lag * 86400000)), d.amount, DIV_ESTIMATED, today]);
      have[ticker + '|' + ex] = true;
    });
  });

  // Estimated rows already in the tab: Pay Date recalculated with the current rule.
  let fixed = 0;
  existing.forEach(r => {
    if (String(r[4]).indexOf('estimated') === -1) return;
    const ticker = String(r[0]).trim(), ex = divParseDate_(r[1]);
    if (!ex || !(ticker in types)) return;
    const pay = fmt(new Date(ex.getTime() + divPayLagDays_(ticker, types[ticker]) * 86400000));
    if (String(r[2]).trim() !== pay) { r[2] = pay; fixed++; }
  });
  if (fixed) sh.getRange(2, 1, existing.length, DIV_HEADER.length).setValues(existing);

  if (added.length) sh.getRange(sh.getLastRow() + 1, 1, added.length, DIV_HEADER.length).setValues(added);
  // Newest first within each ticker, tickers A-Z.
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, DIV_HEADER.length).sort([{column: 1, ascending: true}, {column: 2, ascending: false}]);
  Logger.log('Dividends: ' + added.length + ' new row(s), ' + fixed + ' pay date(s) updated');
}

function installDividendsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateDividends')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateDividends').timeBased().everyDays(1).atHour(8).create();
}

// Earliest purchase per US ticker in the Transactions tab (Ticker | Shares | Buy Date | ...).
function divFirstBuyDates_(ss) {
  const sh = ss.getSheetByName(DIV_TX_SHEET);
  if (!sh) throw new Error('No "' + DIV_TX_SHEET + '" tab');
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim().toLowerCase());
  const cT = h.indexOf('ticker'), cS = h.indexOf('shares');
  const cD = h.findIndex(x => /^(buy )?date$/.test(x));
  if (cT === -1 || cS === -1 || cD === -1) throw new Error('Transactions needs Ticker, Shares and Buy Date columns');
  const out = {};
  for (let r = 1; r < v.length; r++) {
    const t = String(v[r][cT] || '').trim().toUpperCase();
    if (!t || /^IBI\./.test(t) || !(Number(v[r][cS]) > 0)) continue;
    const d = divParseDate_(v[r][cD]);
    if (d && (!out[t] || d < out[t])) out[t] = d;
  }
  return out;
}

// {type: 'ETF' | 'EQUITY' | ..., divs: [{date, amount}]} from Yahoo's chart API, or null.
// BRK.B -> BRK-B (Yahoo's share-class style).
function divFetchYahoo_(ticker) {
  const sym = ticker.replace(/\.([A-Za-z])$/, '-$1');
  const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(sym) + '?range=10y&interval=1mo&events=div';
  try {
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, headers: {'User-Agent': 'Mozilla/5.0'}});
    if (res.getResponseCode() !== 200) return null;
    const result = JSON.parse(res.getContentText()).chart.result;
    if (!result || !result[0]) return null;
    const ev = (result[0].events && result[0].events.dividends) || {};
    return {
      type: (result[0].meta && result[0].meta.instrumentType) || '',
      divs: Object.keys(ev).map(k => ({date: new Date(ev[k].date * 1000), amount: Number(ev[k].amount)})).filter(d => d.amount > 0)
    };
  } catch (e) { return null; }
}

// A date cell, "yyyy-mm-dd" text or "dd/mm/yyyy" text.
function divParseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}
