// SheetPulse write endpoint for the proxy (runs on the Oracle VM, next to index.js).
//
// POST /sheetpulse/append  {sheetId, tab, rows: [[cell, ...], ...]}
// Adds rows at the bottom of one tab of the SheetPulse Payments spreadsheet: what the dashboard's
// "saved on this device" changes used to ask you to paste by hand (a routine bill, a contract, a receipt
// or its description, an order's status or note, a bill marked paid or its note).
//
// Deliberately narrow: only the Payments spreadsheet, only the tabs below, only appends (never edits or
// deletes a row), at most 20 rows of 20 cells per request. The dashboard reads rows with the same message
// id or key as one record, so appending is enough. The route sits after the shared Google ID-token check
// in index.js, so only the allowed account can call it.
//
// Cells: numbers are written as numbers, a yyyy-mm-dd string as a date, and any other string that Sheets
// would turn into a number, date or formula (an order no. like 405-1244688, a message id like 188e20, "2026-07",
// "=...") is written as text (a leading apostrophe), so nothing is reinterpreted.

const PAYMENTS_SHEET_ID = '1CYWfMLK8oJTz7LA85iaoVoYcDYia5uN3I6TEuz0unaA';
const ALLOWED_TABS = ['Payments', 'Shopping', 'Contracts', 'Receipts', 'Routine bills'];
const MAX_ROWS = 20, MAX_CELLS = 20, MAX_LEN = 500;

class SheetPulseAppendError extends Error {}

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

function validate(body){
  if(!body || body.sheetId !== PAYMENTS_SHEET_ID) throw new SheetPulseAppendError('Unknown sheet');
  if(!ALLOWED_TABS.includes(body.tab)) throw new SheetPulseAppendError('Tab not allowed');
  const rows = body.rows;
  if(!Array.isArray(rows) || !rows.length || rows.length > MAX_ROWS) throw new SheetPulseAppendError('rows must hold 1 to ' + MAX_ROWS + ' rows');
  return rows.map(r => {
    if(!Array.isArray(r) || !r.length || r.length > MAX_CELLS) throw new SheetPulseAppendError('Each row must hold 1 to ' + MAX_CELLS + ' cells');
    return r.map(cellValue);
  });
}

async function handleSheetPulseAppend(sheets, body){
  const values = validate(body);
  const spreadsheetId = PAYMENTS_SHEET_ID;
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties.title' });
  if(!meta.data.sheets.some(s => s.properties.title === body.tab)) throw new SheetPulseAppendError(`No "${body.tab}" tab in the spreadsheet yet`);
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId, range: `'${body.tab}'!A1`, valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
    requestBody: { values },
  });
  return { ok: true, tab: body.tab, rows: values.length, range: (res.data.updates || {}).updatedRange || '' };
}

module.exports = { handleSheetPulseAppend, SheetPulseAppendError, PAYMENTS_SHEET_ID };
