// Dividend history for the dashboard: dividend income per holding (both brokers), and Psagot's own
// average cost, which subtracts every dividend once paid, after 25% Israeli tax.
//
// One tab, written by this script and read by the dashboard:
//   Dividends -- Ticker | Ex-Date | Pay Date | Amount | Source | Added
//                One row per dividend per share, for every US ticker in the Transactions tab (both
//                brokers), from the first purchase on. Dates are "yyyy-mm-dd" text.
//
// Sources, in order:
//   1. Alpha Vantage (free key, 25 requests a day) -- the REAL payment date and the declared amount.
//      Each run checks up to DIV_AV_PER_RUN tickers: first those with a recent estimated row, then the
//      rest in rotation, so every ticker is refreshed every few days. Its rows say "Alpha Vantage" and
//      win over Yahoo's for the same dividend. Checked against Psagot's statements (Oct 2026): payment
//      dates exact or within a day.
//   2. Yahoo Finance's chart API -- backup: every ticker, every run, ex-date and amount but NO payment
//      date, so Pay Date is estimated (divPayLagDays_: ETFs within days, stocks weeks later) and Source
//      says "Yahoo, pay date estimated". It fills in dividends Alpha Vantage hasn't listed or checked yet.
//      Yahoo sometimes adjusts old amounts for spin-offs (SPGI May 2026) -- Alpha Vantage's row fixes that.
// Rows whose Source says "estimated" get their Pay Date recalculated on every run. Change Source to
// "manual" after correcting a row by hand from the broker's statement and it's never touched again.
// Without an Alpha Vantage key the script runs on Yahoo alone. TASE funds (IBI.*) are skipped.
//
// Setup (once, in the Apps Script project bound to the Tracker spreadsheet):
//   1. Paste this file as Dividends.gs and run updateDividends() -- creates and fills the tab.
//   2. Run installDividendsTrigger() -- repeats it every morning at 08:00.
//   3. Alpha Vantage (first source): get a free key at alphavantage.co, paste it into setAlphaVantageKey()
//      below, run that function once, then put the placeholder back and save (the key lives only in
//      Script Properties, never in this file). checkAlphaVantageKey() confirms it's stored.

const DIV_SHEET = 'Dividends';
const DIV_TX_SHEET = 'Transactions';
const DIV_HEADER = ['Ticker', 'Ex-Date', 'Pay Date', 'Amount', 'Source', 'Added'];
const DIV_ESTIMATED = 'Yahoo, pay date estimated';
// Days from ex-date to payment. ETFs: 1-6 (JEPI and PFF 6+, SCHD 6, sector SPDRs 2, DBMF 1). Stocks: 14-32 (BAM 29,
// GS 28, MCO 26, SPGI 19, MS 15). Per-ticker exceptions where the type misleads.
const DIV_LAG_ETF = 6;
const DIV_LAG_STOCK = 21;
const DIV_LAG_TICKER = {ASGI: 8, POWR: 3, VCSH: 2};   // closed-end fund; an ETF Yahoo calls a stock; Vanguard bond ETF paying in 2 days

const DIV_AV_PROP = 'ALPHAVANTAGE_KEY';
const DIV_AV_CURSOR = 'div_av_cursor';
const DIV_AV_PER_RUN = 15;          // of the free tier's 25 a day; ~13 s apart (its per-minute limit)
const DIV_AV_RECENT_DAYS = 45;      // estimated rows this recent are checked first

// One-time: paste the key between the quotes, run, then restore the placeholder and save.
function setAlphaVantageKey() {
  const key = 'PASTE_YOUR_KEY_HERE';
  if (!key || /PASTE_YOUR_KEY/.test(key)) throw new Error('Paste your Alpha Vantage key into setAlphaVantageKey() first');
  PropertiesService.getScriptProperties().setProperty(DIV_AV_PROP, key.trim());
  Logger.log('Saved. Now put PASTE_YOUR_KEY_HERE back in the code and save the file.');
}
function checkAlphaVantageKey() {
  const k = PropertiesService.getScriptProperties().getProperty(DIV_AV_PROP);
  Logger.log(k ? 'Alpha Vantage key is stored (ends with ...' + k.slice(-3) + ')' : 'No Alpha Vantage key stored');
}

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
  // Header row lost (it has happened): put it back above the data -- the dashboard finds its
  // columns by these names.
  if (String(sh.getRange(1, 1).getValue()).trim() !== DIV_HEADER[0]) {
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, DIV_HEADER.length).setValues([DIV_HEADER]);
  }
  sh.getRange('A:C').setNumberFormat('@');   // dates stay text, never reformatted by the sheet

  const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, DIV_HEADER.length).getValues() : [];
  const have = {};
  existing.forEach(r => { have[String(r[0]).trim() + '|' + String(r[1]).trim()] = true; });

  const today = fmt(new Date());
  const all = existing.slice();

  // Yahoo for every ticker: its instrument type (for the payment-date estimate) and its dividends.
  const types = {}, yahoo = {};
  Object.keys(firstBuy).sort().forEach(ticker => {
    const y = divFetchYahoo_(ticker);
    if (!y) { Logger.log('No Yahoo data for ' + ticker); return; }
    types[ticker] = y.type; yahoo[ticker] = y.divs;
  });

  // 1. Alpha Vantage first: real payment dates and amounts, and dividends missing from the tab.
  const av = divApplyAlphaVantage_(all, firstBuy, types, fmt, today);

  // 2. Yahoo fills in what's still missing (a dividend already in the tab within a day is the same one).
  const byKey = {};
  all.forEach(r => { byKey[String(r[0]).trim() + '|' + String(r[1]).trim()] = true; });
  const known = (t, d) => [0, -1, 1].some(off => byKey[t + '|' + fmt(new Date(d.getTime() + off * 86400000))]);
  let yAdded = 0;
  Object.keys(yahoo).forEach(ticker => {
    const lag = divPayLagDays_(ticker, types[ticker]);
    const from = firstBuy[ticker].getTime() - 86400000;
    yahoo[ticker].forEach(d => {
      if (d.date.getTime() < from || known(ticker, d.date)) return;
      const r = [ticker, fmt(d.date), fmt(new Date(d.date.getTime() + lag * 86400000)), d.amount, DIV_ESTIMATED, today];
      all.push(r); byKey[ticker + '|' + r[1]] = true; yAdded++;
    });
  });

  // Every estimated row: Pay Date recalculated with the current rule.
  let fixed = 0;
  all.forEach(r => {
    if (String(r[4]).indexOf('estimated') === -1) return;
    const ticker = String(r[0]).trim(), ex = divParseDate_(r[1]);
    if (!ex || !(ticker in types)) return;
    const pay = fmt(new Date(ex.getTime() + divPayLagDays_(ticker, types[ticker]) * 86400000));
    if (String(r[2]).trim() !== pay) { r[2] = pay; fixed++; }
  });
  if (all.length) sh.getRange(2, 1, all.length, DIV_HEADER.length).setValues(all);
  // Newest first within each ticker, tickers A-Z.
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, DIV_HEADER.length).sort([{column: 1, ascending: true}, {column: 2, ascending: false}]);
  Logger.log('Alpha Vantage: ' + av + ' | Yahoo (backup): ' + yAdded + ' new | ' + fixed + ' estimated pay date(s) updated');
}

function installDividendsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateDividends')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateDividends').timeBased().everyDays(1).atHour(8).create();
}

// Updates rows in place (and appends to `rows`) from Alpha Vantage. Returns a short summary.
function divApplyAlphaVantage_(rows, firstBuy, types, fmt, today) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty(DIV_AV_PROP);
  if (!key) return 'no key, skipped';
  const tickers = Object.keys(firstBuy).sort();
  if (!tickers.length) return 'nothing to check';
  // Tickers with a recent estimated row first, then the rest in rotation.
  const recent = Date.now() - DIV_AV_RECENT_DAYS * 86400000;
  const urgent = tickers.filter(t => rows.some(r => String(r[0]).trim() === t && String(r[4]).indexOf('estimated') !== -1
    && (divParseDate_(r[1]) || new Date(0)).getTime() > recent));
  let cursor = Number(props.getProperty(DIV_AV_CURSOR) || 0) % tickers.length;
  const pick = urgent.slice(0, DIV_AV_PER_RUN);
  while (pick.length < Math.min(DIV_AV_PER_RUN, tickers.length)) {
    const t = tickers[cursor]; cursor = (cursor + 1) % tickers.length;
    if (pick.indexOf(t) === -1) pick.push(t);
  }
  props.setProperty(DIV_AV_CURSOR, String(cursor));

  const byKey = {};
  rows.forEach(r => { byKey[String(r[0]).trim() + '|' + String(r[1]).trim()] = r; });
  const near = (t, ex) => {   // Yahoo's ex-date can be a day off (time zone)
    const d = divParseDate_(ex);
    for (const off of [0, -1, 1]) { const r = byKey[t + '|' + fmt(new Date(d.getTime() + off * 86400000))]; if (r) return r; }
    return null;
  };
  let checked = 0, updated = 0, addedN = 0;
  for (let i = 0; i < pick.length; i++) {
    if (i) Utilities.sleep(13000);
    const t = pick[i];
    let data;
    try {
      const res = UrlFetchApp.fetch('https://www.alphavantage.co/query?function=DIVIDENDS&symbol=' + encodeURIComponent(t) + '&apikey=' + encodeURIComponent(key), {muteHttpExceptions: true});
      data = JSON.parse(res.getContentText());
    } catch (e) { continue; }
    if (data.Information || data.Note) { Logger.log('Alpha Vantage limit: ' + (data.Information || data.Note)); break; }
    if (!Array.isArray(data.data)) continue;
    checked++;
    const from = firstBuy[t].getTime() - 86400000;
    data.data.forEach(d => {
      const ex = divParseDate_(d.ex_dividend_date), amount = Number(d.amount);
      if (!ex || !(amount > 0) || ex.getTime() < from) return;
      const pay = divParseDate_(d.payment_date);
      const row = near(t, fmt(ex));
      if (row) {
        if (String(row[4]).indexOf('manual') !== -1 || String(row[4]) === 'Alpha Vantage') return;
        if (!pay) return;   // no real date yet -- keep the estimate
        row[2] = fmt(pay); row[3] = amount; row[4] = 'Alpha Vantage'; updated++;
      } else {
        const p = pay || new Date(ex.getTime() + divPayLagDays_(t, types[t]) * 86400000);
        const r = [t, fmt(ex), fmt(p), amount, pay ? 'Alpha Vantage' : 'Alpha Vantage, pay date estimated', today];
        rows.push(r); byKey[t + '|' + r[1]] = r; addedN++;
      }
    });
  }
  return checked + ' ticker(s) checked, ' + updated + ' pay date(s) confirmed, ' + addedN + ' dividend(s) added';
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
