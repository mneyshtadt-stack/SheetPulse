#!/bin/bash
# Installs SheetPulse's save route (POST /sheetpulse/append) into the proxy on the Oracle VM.
# Run with:  bash ~/install-append.sh
set -e
cd /home/ubuntu/sheet-proxy
cp index.js index.js.before-append

cat > sheetpulse-append.js <<'APPEND_EOF'
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
APPEND_EOF

cat > patch-append.js <<'APPEND_EOF'
// One-time patch for /home/ubuntu/sheet-proxy/index.js: adds SheetPulse's save route (POST /sheetpulse/append).
// Safe to re-run -- it does nothing if the route is already there. Keeps a backup as index.js.bak-append.
// Usage (in /home/ubuntu/sheet-proxy):  node patch-append.js

const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, 'index.js');
let src = fs.readFileSync(file, 'utf8');

if (src.includes('/sheetpulse/append')) {
  console.log('Already patched - nothing to do.');
  process.exit(0);
}

const requireAnchor = "} = require('@simplewebauthn/server');";
// The route goes after the shared token check, so it gets exactly the same protection as the reads
// (and as the MedTracker save route, which uses the same anchor).
const routeAnchor = '  const sheetId = parsed.query.sheetId;';
if (!src.includes(requireAnchor) || !src.includes(routeAnchor) || !src.includes('getSheetsClient') || !src.includes('readJsonBody')) {
  console.error('Could not find the expected lines in index.js - not changed.');
  process.exit(1);
}

const requireLine = "\nconst { handleSheetPulseAppend, SheetPulseAppendError } = require('./sheetpulse-append');";
const route = `  // --- SheetPulse: add rows to the Payments spreadsheet (same token check as the reads above) ---
  if (parsed.pathname === '/sheetpulse/append' && req.method === 'POST') {
    try {
      const body = await readJsonBody(req);
      const sheets = await getSheetsClient();
      const result = await handleSheetPulseAppend(sheets, body);
      return send(res, req, 200, JSON.stringify(result), 'application/json');
    } catch (err) {
      const status = err instanceof SheetPulseAppendError ? 400 : 500;
      console.error(new Date().toISOString(), 'sheetpulse append error:', err.message);
      return send(res, req, status, JSON.stringify({ error: err.message }), 'application/json');
    }
  }

`;

fs.writeFileSync(file + '.bak-append', src);
src = src.replace(requireAnchor, requireAnchor + requireLine);
src = src.replace(routeAnchor, route + routeAnchor);
fs.writeFileSync(file, src);
console.log('Patched index.js (backup: index.js.bak-append).');
APPEND_EOF

node --check sheetpulse-append.js
node patch-append.js
node --check index.js
sudo systemctl restart sheet-proxy
sleep 2
echo "sheet-proxy is: $(systemctl is-active sheet-proxy)"
sudo journalctl -u sheet-proxy -n 5 --no-pager
