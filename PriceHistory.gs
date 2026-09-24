// Daily closing-price history for the dashboard's 1M / 6M / YTD / 1Y / 5Y / MAX chart ranges.
//
// Two tabs:
//   History       -- YOUR formula tab, edited by hand like the Tracker. B1 = start date; row 3 =
//                    one symbol per column pair (CURRENCY:USDILS, VOO, SMH, ...); row 4 =
//                    =QUERY(GOOGLEFINANCE(symbol,"close",$B$1,TODAY(),"DAILY"),"select * order by Col1 desc",1)
//                    under each symbol, spilling Date | Close down the pair, newest date first.
//                    Add a ticker = add a new column pair.
//   PriceHistory  -- plain values written by this script, read by the dashboard: row 1 =
//                    Date | USDILS | <ticker> ...; one row per date, newest first ("yyyy-mm-dd" text); prices in
//                    each ticker's own quote currency (USD, or agorot for the TASE funds). Blank =
//                    no close that day; the dashboard carries the last one forward.
//
// Why the copy: the Sheets API (which the dashboard's proxy uses) returns #N/A for GOOGLEFINANCE
// history of US-listed tickers, while Apps Script can read it -- so the values are copied out.
// A column prefix like "NYSEARCA:VOO" is stored as "VOO", and CURRENCY:USDILS as "USDILS".
//
// Setup (once, in the Apps Script project bound to the Tracker spreadsheet):
//   1. Optional: run createHistoryTab() -- builds History with every Market Universe ticker,
//      starting from your earliest Tracker "Buy Date". Skip it if you'd rather build it yourself.
//   2. Wait until the History formulas have loaded, then run copyHistoryToPriceHistory().
//   3. Run installPriceHistoryTrigger() -- repeats the copy every morning at 07:00.

const PH_SOURCE_SHEET = 'History';
const PH_SHEET = 'PriceHistory';
const PH_FX_COL = 'USDILS';

function copyHistoryToPriceHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(PH_SOURCE_SHEET);
  if (!src) throw new Error('No "' + PH_SOURCE_SHEET + '" tab -- create it first (see createHistoryTab)');
  const tz = ss.getSpreadsheetTimeZone();
  const values = src.getDataRange().getValues();
  // A GOOGLEFINANCE block errors out once its growing history reaches the sheet's last row --
  // keep a margin so that never happens.
  if (src.getMaxRows() - values.length < 60) src.insertRowsAfter(src.getMaxRows(), 250);

  // The formulas' own "Date" headers mark each block; the symbol sits directly above.
  let headerRow = -1;
  for (let r = 1; r < Math.min(values.length, 15); r++) {
    if (values[r].indexOf('Date') !== -1) { headerRow = r; break; }
  }
  if (headerRow === -1) throw new Error('No loaded GOOGLEFINANCE blocks found in "' + PH_SOURCE_SHEET + '" (no "Date" headers yet)');

  const fetched = {};
  const empty = [];
  for (let c = 0; c < values[headerRow].length; c++) {
    const symbol = String(values[headerRow - 1][c] || '').trim();
    if (!symbol) continue;
    const col = phColumnName_(symbol);
    if (values[headerRow][c] !== 'Date') { empty.push(col); continue; }
    const closes = {};
    for (let r = headerRow + 1; r < values.length; r++) {
      const dt = values[r][c], close = values[r][c + 1];
      if (!(dt instanceof Date)) break;
      if (typeof close === 'number') closes[Utilities.formatDate(dt, tz, 'yyyy-MM-dd')] = close;
    }
    if (Object.keys(closes).length) fetched[col] = closes; else empty.push(col);
  }

  const rowCount = phMergeAndWrite_(ss, fetched);
  console.log('PriceHistory: ' + Object.keys(fetched).length + ' series copied, ' + rowCount + ' dates in the tab'
    + (empty.length ? '. No data (yet) for: ' + empty.join(', ') : ''));
}

// 07:00 in this Apps Script project's own time zone (Project Settings; should be Asia/Jerusalem):
// after USA's close, before TASE's open. Safe to re-run -- replaces any existing one.
function installPriceHistoryTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'copyHistoryToPriceHistory')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('copyHistoryToPriceHistory').timeBased().everyDays(1).atHour(7).create();
}

// One-time starter layout for the History tab. Refuses to touch an existing one.
function createHistoryTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(PH_SOURCE_SHEET)) throw new Error('"' + PH_SOURCE_SHEET + '" already exists -- not overwriting it');
  const symbols = ['CURRENCY:USDILS'].concat(phTickers_(ss));
  const sh = ss.insertSheet(PH_SOURCE_SHEET);
  const start = phEarliestBuyDate_(ss);
  const rowsNeeded = Math.ceil((Date.now() - start.getTime()) / 86400000) + 400;
  phEnsureSize_(sh, rowsNeeded, symbols.length * 2);

  sh.getRange('A1').setValue('Start');
  sh.getRange('B1').setValue(start).setNumberFormat('yyyy-mm-dd');
  symbols.forEach((s, i) => {
    const c = i * 2 + 1;
    sh.getRange(3, c).setValue(s).setFontWeight('bold');
    const a1 = sh.getRange(3, c).getA1Notation();
    sh.getRange(4, c).setFormula(phDescendingFormula_('GOOGLEFINANCE(' + a1 + ',"close",$B$1,TODAY(),"DAILY")'));
    sh.getRange(5, c, rowsNeeded - 4, 1).setNumberFormat('yyyy-mm-dd');
  });
  sh.setFrozenRows(4);
}

// GOOGLEFINANCE always returns oldest-first; QUERY re-sorts newest-first and keeps the Date | Close
// header (the 1) on top.
function phDescendingFormula_(googleFinanceCall) {
  return '=QUERY(' + googleFinanceCall + ',"select * order by Col1 desc",1)';
}

// One-time: wraps every plain GOOGLEFINANCE history formula in the History tab in the QUERY above,
// so each block lists its newest date first. Formulas already wrapped are left alone; safe to re-run.
function convertHistoryToDescending() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PH_SOURCE_SHEET);
  if (!sh) throw new Error('No "' + PH_SOURCE_SHEET + '" tab');
  const rows = Math.min(sh.getLastRow(), 15), cols = sh.getLastColumn();
  const formulas = sh.getRange(1, 1, rows, cols).getFormulas();
  let converted = 0;
  formulas.forEach((row, r) => row.forEach((f, c) => {
    if (/^=\s*GOOGLEFINANCE\(/i.test(f) && !/QUERY\(/i.test(f)) {
      sh.getRange(r + 1, c + 1).setFormula(phDescendingFormula_(f.replace(/^=\s*/, '')));
      converted++;
    }
  }));
  console.log('History: ' + converted + ' formula(s) converted to newest-first');
}

// "NYSEARCA:VOO" -> "VOO", "CURRENCY:USDILS" -> "USDILS", "SMH" -> "SMH".
function phColumnName_(symbol) {
  const i = symbol.lastIndexOf(':');
  return i === -1 ? symbol : symbol.slice(i + 1);
}

// Merges the copied closes into whatever PriceHistory already holds -- a block that's temporarily
// #N/A or loading never erases values copied on an earlier run -- and rewrites the tab sorted
// newest date first. Dates where only USD/ILS has a value (weekend FX quotes) are dropped. Returns
// the date count.
function phMergeAndWrite_(ss, fetched) {
  let sh = ss.getSheetByName(PH_SHEET);
  if (!sh) sh = ss.insertSheet(PH_SHEET);
  const existing = sh.getLastRow() > 0 ? sh.getDataRange().getValues() : [];
  const header = existing.length ? existing[0].map(String) : ['Date', PH_FX_COL];
  const cols = header.slice(1);
  Object.keys(fetched).forEach(c => { if (cols.indexOf(c) === -1) cols.push(c); });

  const byDate = {};
  for (let r = 1; r < existing.length; r++) {
    const key = String(existing[r][0]).trim();
    if (!key) continue;
    const row = byDate[key] = {};
    for (let c = 1; c < header.length; c++) {
      if (typeof existing[r][c] === 'number') row[header[c]] = existing[r][c];
    }
  }
  Object.keys(fetched).forEach(col => {
    const closes = fetched[col];
    Object.keys(closes).forEach(key => { (byDate[key] || (byDate[key] = {}))[col] = closes[key]; });
  });

  const keys = Object.keys(byDate)
    .filter(k => cols.some(c => c !== PH_FX_COL && typeof byDate[k][c] === 'number'))
    .sort()
    .reverse();
  const out = [['Date'].concat(cols)].concat(
    keys.map(k => [k].concat(cols.map(c => (typeof byDate[k][c] === 'number' ? byDate[k][c] : '')))));

  phEnsureSize_(sh, out.length, out[0].length);
  sh.clearContents();
  sh.getRange(1, 1, out.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.setFrozenRows(1);
  return keys.length;
}

// Every Market Universe ticker, plus any Tracker holding not listed there (for createHistoryTab).
function phTickers_(ss) {
  const tickers = [];
  const add = t => { t = String(t || '').trim(); if (t && tickers.indexOf(t) === -1) tickers.push(t); };
  const uni = ss.getSheetByName('Investment Universe');
  if (uni && uni.getLastRow() > 1) {
    const v = uni.getDataRange().getValues();
    const c = v[0].map(h => String(h).trim()).indexOf('Ticker');
    if (c !== -1) for (let r = 1; r < v.length; r++) add(v[r][c]);
  }
  phHoldings_(ss).forEach(h => add(h.ticker));
  return tickers;
}

function phHoldings_(ss) {
  const data = ss.getSheetByName('Tracker').getDataRange().getValues();
  let hdr = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].indexOf('Ticker') !== -1 && data[i].indexOf('Shares') !== -1) { hdr = i; break; }
  }
  if (hdr === -1) throw new Error('Could not find the Tracker holdings header row');
  const h = data[hdr].map(x => String(x).trim());
  const cTicker = h.indexOf('Ticker'), cBuy = h.indexOf('Buy Date');
  const out = [];
  for (let r = hdr + 1; r < data.length; r++) {
    const ticker = String(data[r][cTicker] || '').trim();
    if (!ticker) break;
    out.push({ ticker: ticker, buyDate: cBuy === -1 ? null : phParseDate_(data[r][cBuy]) });
  }
  return out;
}

function phEarliestBuyDate_(ss) {
  const dates = phHoldings_(ss).map(h => h.buyDate).filter(d => d);
  if (dates.length) return new Date(Math.min.apply(null, dates.map(d => d.getTime())));
  const fallback = new Date();
  fallback.setFullYear(fallback.getFullYear() - 5);
  return fallback;
}

// Buy Date may be a real date cell or dd/mm/yyyy text.
function phParseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

function phEnsureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
