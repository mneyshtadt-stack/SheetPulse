// Transactions tab: every purchase (and sale) as one line -- the Tracker keeps one row per ticker
// and takes its Shares, price and Buy Date from here by formula.
//
//   Transactions:  Date | Ticker | Shares | Price | Note
//     - one line per purchase: the date, the Tracker's ticker exactly as written there, the number of
//       shares, and the price per share in the ticker's own currency (the same as the Tracker's
//       price column, the "D" in Total Cost = B*D)
//     - a sale: negative Shares, Price left empty (average-cost method: the average price stays,
//       Shares -- and so Total Cost -- go down)
//
//   Tracker, per holding row (formulas written by migrateToTransactions):
//     Shares    = total shares in Transactions for that ticker (buys minus sales)
//     Price     = share-weighted average price of the buys
//     Buy Date  = the first purchase date
//   Total Cost (=B*D), Total Cost (ILS), Current Value etc. keep working unchanged on top of them.
//
// One-time setup (in the Apps Script project bound to the Tracker spreadsheet):
//   0. File -> Make a copy of the spreadsheet first, as a backup.
//   1. Paste this file as Transactions (Apps Script adds .gs).
//   2. Run previewTransactionsMigration() -- changes nothing; the Execution log lists the columns it
//      found and the transactions it would create. Check them.
//   3. Run migrateToTransactions() -- creates the Transactions tab from the Tracker's current rows
//      and replaces Shares / Price / Buy Date with the formulas, then checks every row still shows
//      the same values as before.
// After that, a new purchase is one new line in Transactions -- nothing to change in the Tracker.

const TX_SHEET = 'Transactions';
const TX_HEADER = ['Date', 'Ticker', 'Shares', 'Price', 'Note'];

// Finds the Tracker's holdings table and the columns the migration needs. The Shares and price
// columns are taken from the first holding's own Total Cost formula (=B5*D5), so nothing is guessed.
function txTrackerLayout_(ss) {
  const sh = ss.getSheetByName('Tracker');
  if (!sh) throw new Error('No "Tracker" tab');
  const values = sh.getDataRange().getValues();
  let hdr = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i].indexOf('Ticker') !== -1 && values[i].indexOf('Shares') !== -1) { hdr = i; break; }
  }
  if (hdr === -1) throw new Error('Could not find the Tracker header row (with "Ticker" and "Shares")');
  const h = values[hdr].map((x) => String(x).trim());
  const cTicker = h.indexOf('Ticker'), cShares = h.indexOf('Shares'), cBuyDate = h.indexOf('Buy Date');
  const cTotalCost = h.findIndex((x) => /^total cost/i.test(x) && !/ils|₪/i.test(x));
  if (cBuyDate === -1) throw new Error('No "Buy Date" column in the Tracker header');
  if (cTotalCost === -1) throw new Error('No "Total Cost" column (the non-ILS one) in the Tracker header');

  const firstRow = hdr + 2;  // 1-based sheet row of the first holding
  const formula = sh.getRange(firstRow, cTotalCost + 1).getFormula();
  const m = /^=\s*\$?([A-Z]+)\$?(\d+)\s*\*\s*\$?([A-Z]+)\$?(\d+)\s*$/i.exec(formula);
  if (!m) throw new Error('Total Cost on row ' + firstRow + ' is "' + formula + '" -- expected a formula like =B' + firstRow + '*D' + firstRow);
  const colOf = (letters) => letters.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  let a = colOf(m[1]), b = colOf(m[3]);
  if (b === cShares) { const t = a; a = b; b = t; }            // written as =D*B
  if (a !== cShares) throw new Error('Total Cost formula ' + formula + ' does not use the Shares column (' + txLetter_(cShares) + ')');
  const cPrice = b;

  const rows = [];
  for (let r = hdr + 1; r < values.length; r++) {
    const ticker = String(values[r][cTicker] || '').trim();
    if (!ticker) break;
    rows.push({ sheetRow: r + 1, ticker, shares: values[r][cShares], price: values[r][cPrice], buyDate: values[r][cBuyDate] });
  }
  return { sh, hdr, cTicker, cShares, cPrice, cBuyDate, cTotalCost, header: h, rows };
}

function txLetter_(c) {
  let s = '', n = c + 1;
  while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Changes nothing: logs what migrateToTransactions would do.
function previewTransactionsMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const L = txTrackerLayout_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const fmtDate = (d) => (d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : String(d));
  console.log('Tracker header on row ' + (L.hdr + 1) + ': Ticker=' + txLetter_(L.cTicker) + ', Shares=' + txLetter_(L.cShares)
    + ', Price=' + txLetter_(L.cPrice) + ' ("' + L.header[L.cPrice] + '"), Buy Date=' + txLetter_(L.cBuyDate)
    + ', Total Cost=' + txLetter_(L.cTotalCost) + '; ' + L.rows.length + ' holdings.');
  if (ss.getSheetByName(TX_SHEET)) console.log('NOTE: a "' + TX_SHEET + '" tab already exists -- migrateToTransactions will refuse to run.');
  const dups = txDuplicates_(L.rows);
  if (dups.length) console.log('NOTE: these tickers have more than one Tracker row: ' + dups.join(', ') + ' -- migrateToTransactions will refuse; keep one row each (add the extra purchase as a Transactions line afterwards).');
  const bad = L.rows.filter((r) => !(Number(r.shares) > 0) || !(Number(r.price) > 0) || !(r.buyDate instanceof Date));
  if (bad.length) console.log('NOTE: rows missing Shares, Price or a real Buy Date (they are copied as they are): ' + bad.map((r) => r.ticker + ' (row ' + r.sheetRow + ')').join(', '));
  L.rows.slice(0, 8).forEach((r) => console.log('  would add: ' + fmtDate(r.buyDate) + ' | ' + r.ticker + ' | ' + r.shares + ' | ' + r.price));
  if (L.rows.length > 8) console.log('  ... and ' + (L.rows.length - 8) + ' more');
}

function txDuplicates_(rows) {
  const seen = {}, dups = [];
  rows.forEach((r) => { if (seen[r.ticker] && dups.indexOf(r.ticker) === -1) dups.push(r.ticker); seen[r.ticker] = true; });
  return dups;
}

// One-time: creates Transactions from the Tracker's rows and switches the Tracker to formulas.
function migrateToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(TX_SHEET)) throw new Error('"' + TX_SHEET + '" already exists -- not running twice');
  const L = txTrackerLayout_(ss);
  const dups = txDuplicates_(L.rows);
  if (dups.length) throw new Error('More than one Tracker row for: ' + dups.join(', ') + ' -- keep one row per ticker first');

  // 1. The Transactions tab: one line per current holding, oldest first.
  const tx = ss.insertSheet(TX_SHEET);
  const lines = L.rows.map((r) => [r.buyDate, r.ticker, r.shares, r.price, 'from Tracker'])
    .sort((x, y) => (x[0] instanceof Date ? x[0].getTime() : 0) - (y[0] instanceof Date ? y[0].getTime() : 0));
  tx.getRange(1, 1, 1, TX_HEADER.length).setValues([TX_HEADER]).setFontWeight('bold');
  if (lines.length) tx.getRange(2, 1, lines.length, TX_HEADER.length).setValues(lines);
  tx.getRange('A:A').setNumberFormat(L.sh.getRange(L.rows[0].sheetRow, L.cBuyDate + 1).getNumberFormat() || 'dd/mm/yyyy');
  tx.getRange('D:D').setNumberFormat(L.sh.getRange(L.rows[0].sheetRow, L.cPrice + 1).getNumberFormat() || '0.00');
  tx.setFrozenRows(1);
  tx.autoResizeColumns(1, TX_HEADER.length);

  // 2. The Tracker's Shares / Price / Buy Date become formulas over Transactions.
  const T = "'" + TX_SHEET + "'!";
  const col = (c) => '$' + txLetter_(c);
  L.rows.forEach((r) => {
    const t = col(L.cTicker) + r.sheetRow;
    L.sh.getRange(r.sheetRow, L.cShares + 1).setFormula(`=SUMIFS(${T}$C$2:$C, ${T}$B$2:$B, ${t})`);
    L.sh.getRange(r.sheetRow, L.cPrice + 1).setFormula(
      `=IFERROR(SUMPRODUCT((${T}$B$2:$B=${t})*(${T}$C$2:$C>0), ${T}$C$2:$C, ${T}$D$2:$D) / SUMIFS(${T}$C$2:$C, ${T}$B$2:$B, ${t}, ${T}$C$2:$C, ">0"), "")`);
    L.sh.getRange(r.sheetRow, L.cBuyDate + 1).setFormula(`=IFERROR(1/(1/MINIFS(${T}$A$2:$A, ${T}$B$2:$B, ${t}, ${T}$C$2:$C, ">0")), "")`);
  });
  SpreadsheetApp.flush();

  // 3. Every row must show the same Shares, Price and Buy Date as before.
  const after = txTrackerLayout_(ss).rows;
  const same = (x, y) => (x instanceof Date && y instanceof Date) ? x.getTime() === y.getTime() : Math.abs(Number(x) - Number(y)) < 1e-9;
  const diffs = L.rows.filter((r, i) => !(same(r.shares, after[i].shares) && same(r.price, after[i].price) && same(r.buyDate, after[i].buyDate)));
  if (diffs.length) {
    console.log('CHECK THESE ROWS -- their values changed: ' + diffs.map((r) => r.ticker + ' (row ' + r.sheetRow + ')').join(', '));
  } else {
    console.log('Done: ' + lines.length + ' transactions created; all ' + L.rows.length + ' Tracker rows show the same values as before.');
  }
}
