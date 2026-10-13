// SheetPulse write endpoints for the proxy (runs on the Oracle VM, next to index.js).
//
// POST /sheetpulse/append  {sheetId, tab, rows: [[cell, ...], ...]}
//   Adds rows at the bottom of one tab of the SheetPulse Payments spreadsheet: events that belong in rows of
//   their own (an order's status, a bill marked paid, a bill's note, a new contract, receipt or routine bill).
//
// POST /sheetpulse/update  {sheetId, tab, key: {Column: value, ...}, set: {Column: value, ...}}
//   Changes cells of ONE existing row in place, so an edit (a contract's note, what a receipt was for, a
//   routine bill's amount) doesn't add a second row. The row is the first one whose key columns all match
//   (trimmed, ignoring case): Name + Reference in Contracts, Message id in Receipts, Payee in Routine bills.
//   Only the columns listed below can be set; the header row is never written. No match: 404, and the
//   dashboard falls back to adding a row.
//
// POST /sheetpulse/clear-avg  {ticker, value}
//   The one write to the Tracker spreadsheet: clears the hand-typed Broker Avg Cost cell(s) of one ticker in its
//   Transactions tab, only where the cell holds that value (the Dividends tab's "Delete now"). Nothing else
//   in the Tracker can be written.
//
// Deliberately narrow: only the Payments spreadsheet, only the tabs below, never deletes a row, at most 20
// rows of 20 cells per append. Both routes sit after the shared Google ID-token check in index.js, so only the
// allowed account can call them.
//
// Cells: numbers are written as numbers, a yyyy-mm-dd string as a date, and any other string that Sheets
// would turn into a number, date or formula (an order no. like 405-1244688, a message id like 188e20, "2026-07",
// "=...") is written as text (a leading apostrophe), so nothing is reinterpreted.

const PAYMENTS_SHEET_ID = '1CYWfMLK8oJTz7LA85iaoVoYcDYia5uN3I6TEuz0unaA';
const ALLOWED_TABS = ['Payments', 'Shopping', 'Contracts', 'Receipts', 'Routine bills'];
const UPDATE_RULES = {
  'Contracts':     { key: ['Name', 'Reference'], set: ['Provider', 'Type', 'Start', 'End', 'Cost', 'Per', 'Covers', 'Notes', 'Link'] },
  'Receipts':      { key: ['Message id'],        set: ['Amount', 'Currency', 'Notes', 'Product', 'Kind'] },
  'Routine bills': { key: ['Payee'],             set: ['Category', 'Amount', 'Every', 'Day', 'Start', 'End', 'Card', 'Notes'] },
};
const MAX_ROWS = 20, MAX_CELLS = 20, MAX_LEN = 500;

class SheetPulseAppendError extends Error {}
class SheetPulseNotFound extends SheetPulseAppendError {}

function cellValue(v){
  if(v === null || v === undefined) return '';
  if(typeof v === 'number'){
    if(!isFinite(v)) throw new SheetPulseAppendError('Not a valid number');
    return v;
  }
  if(typeof v === 'boolean') return v;
  const s = String(v);
  if(s.length > MAX_LEN) throw new SheetPulseAppendError('A cell is too long');
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;                        // a date
  if(/^[=+\-@]/.test(s) || /^[\d\s.,:\/eE%$-]+$/.test(s) || /^(true|false)$/i.test(s)) return "'" + s;
  return s;
}

function checkSheetAndTab(body, tabs){
  if(!body || body.sheetId !== PAYMENTS_SHEET_ID) throw new SheetPulseAppendError('Unknown sheet');
  if(!tabs.includes(body.tab)) throw new SheetPulseAppendError('Tab not allowed');
}

function validateAppend(body){
  checkSheetAndTab(body, ALLOWED_TABS);
  const rows = body.rows;
  if(!Array.isArray(rows) || !rows.length || rows.length > MAX_ROWS) throw new SheetPulseAppendError('rows must hold 1 to ' + MAX_ROWS + ' rows');
  return rows.map(r => {
    if(!Array.isArray(r) || !r.length || r.length > MAX_CELLS) throw new SheetPulseAppendError('Each row must hold 1 to ' + MAX_CELLS + ' cells');
    return r.map(cellValue);
  });
}

async function tabExists(sheets, tab){
  const meta = await sheets.spreadsheets.get({ spreadsheetId: PAYMENTS_SHEET_ID, fields: 'sheets.properties.title' });
  if(!meta.data.sheets.some(s => s.properties.title === tab)) throw new SheetPulseAppendError(`No "${tab}" tab in the spreadsheet yet`);
}

async function handleSheetPulseAppend(sheets, body){
  const values = validateAppend(body);
  await tabExists(sheets, body.tab);
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId: PAYMENTS_SHEET_ID, range: `'${body.tab}'!A1`, valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
    requestBody: { values },
  });
  return { ok: true, tab: body.tab, rows: values.length, range: (res.data.updates || {}).updatedRange || '' };
}

function colLetter(i){ let s = ''; i += 1; while(i > 0){ const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
const norm = v => String(v === null || v === undefined ? '' : v).trim().toLowerCase();

async function handleSheetPulseUpdate(sheets, body){
  checkSheetAndTab(body, Object.keys(UPDATE_RULES));
  const rule = UPDATE_RULES[body.tab];
  const key = body.key || {}, set = body.set || {};
  if(Object.keys(key).length !== rule.key.length || !rule.key.every(k => k in key)) throw new SheetPulseAppendError('key must be ' + rule.key.join(' + '));
  const setCols = Object.keys(set);
  if(!setCols.length || setCols.some(c => !rule.set.includes(c))) throw new SheetPulseAppendError('Only these can be changed: ' + rule.set.join(', '));
  const values = setCols.map(c => cellValue(set[c]));
  await tabExists(sheets, body.tab);
  const grid = (await sheets.spreadsheets.values.get({ spreadsheetId: PAYMENTS_SHEET_ID, range: `'${body.tab}'!A1:Z` })).data.values || [];
  const head = (grid[0] || []).map(norm);
  const col = name => head.indexOf(norm(name));
  // A cleared header cell in column A still means the first key column (it happened to Contracts' "Name").
  const keyIdx = rule.key.map((k, j) => { const i = col(k); return i < 0 && j === 0 && head[0] === '' ? 0 : i; });
  if(keyIdx.some(i => i < 0)) throw new SheetPulseAppendError('The tab has no ' + rule.key.join(' / ') + ' column');
  const setIdx = setCols.map(c => col(c));
  if(setIdx.some(i => i < 0)) throw new SheetPulseAppendError('The tab has no ' + setCols.filter((c, j) => setIdx[j] < 0).join(', ') + ' column');
  let rowNo = -1;
  for(let r = 1; r < grid.length; r++){
    const row = grid[r] || [];
    if(rule.key.every((k, j) => norm(row[keyIdx[j]]) === norm(key[k]))){ rowNo = r + 1; break; }
  }
  if(rowNo < 0) throw new SheetPulseNotFound('Row not found');
  const data = setIdx.map((i, j) => ({ range: `'${body.tab}'!${colLetter(i)}${rowNo}`, values: [[values[j]]] }));
  await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: PAYMENTS_SHEET_ID, requestBody: { valueInputOption: 'USER_ENTERED', data } });
  return { ok: true, tab: body.tab, row: rowNo, cells: setCols };
}

const TRACKER_SHEET_ID = '1eKHhkNP9F39jwqlwokLoJ0zC3i4jA5u91zg5WyNPb0g';
async function handleSheetPulseClearAvg(sheets, body){
  const ticker = norm(body && body.ticker), value = Number(body && body.value);
  if(!ticker || ticker.length > 20 || !isFinite(value) || value <= 0) throw new SheetPulseAppendError('ticker and value are needed');
  const grid = (await sheets.spreadsheets.values.get({ spreadsheetId: TRACKER_SHEET_ID, range: "'Transactions'!A1:Z", valueRenderOption: 'UNFORMATTED_VALUE' })).data.values || [];
  const head = (grid[0] || []).map(norm);
  const cT = head.indexOf('ticker'), cA = head.findIndex(h => /^broker\s*avg\.?\s*cost$/.test(h));
  if(cT < 0 || cA < 0) throw new SheetPulseAppendError('The Transactions tab has no Ticker / Broker Avg Cost column');
  const data = [];
  for(let r = 1; r < grid.length; r++){
    const row = grid[r] || [], v = Number(String(row[cA] === undefined ? '' : row[cA]).replace(/[$,\s]/g, ''));
    if(norm(row[cT]) === ticker && isFinite(v) && Math.abs(v - value) < 0.005) data.push({ range: `'Transactions'!${colLetter(cA)}${r + 1}`, values: [['']] });
  }
  if(!data.length) throw new SheetPulseNotFound('No such Broker Avg Cost cell');
  await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: TRACKER_SHEET_ID, requestBody: { valueInputOption: 'RAW', data } });
  return { ok: true, cleared: data.map(d => d.range) };
}

module.exports = { handleSheetPulseAppend, handleSheetPulseUpdate, handleSheetPulseClearAvg, SheetPulseAppendError, SheetPulseNotFound, PAYMENTS_SHEET_ID };
