// One-time patch for /home/ubuntu/sheet-proxy/index.js: adds SheetPulse's in-place edit route
// (POST /sheetpulse/update), next to /sheetpulse/append (which patch-append.js added).
// Safe to re-run -- it does nothing if the route is already there. Keeps a backup as index.js.bak-update.
// Usage (in /home/ubuntu/sheet-proxy):  node patch-update.js

const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, 'index.js');
let src = fs.readFileSync(file, 'utf8');

if (src.includes('/sheetpulse/update')) {
  console.log('Already patched - nothing to do.');
  process.exit(0);
}

const oldRequire = "const { handleSheetPulseAppend, SheetPulseAppendError } = require('./sheetpulse-append');";
const routeAnchor = '  const sheetId = parsed.query.sheetId;';
if (!src.includes(oldRequire) || !src.includes(routeAnchor)) {
  console.error('Could not find the expected lines in index.js (is the append route installed?) - not changed.');
  process.exit(1);
}

const newRequire = "const { handleSheetPulseAppend, handleSheetPulseUpdate, SheetPulseAppendError, SheetPulseNotFound } = require('./sheetpulse-append');";
const route = `  // --- SheetPulse: change cells of one existing row in place (same token check as the reads above) ---
  if (parsed.pathname === '/sheetpulse/update' && req.method === 'POST') {
    try {
      const body = await readJsonBody(req);
      const sheets = await getSheetsClient();
      const result = await handleSheetPulseUpdate(sheets, body);
      return send(res, req, 200, JSON.stringify(result), 'application/json');
    } catch (err) {
      const status = err instanceof SheetPulseNotFound ? 404 : err instanceof SheetPulseAppendError ? 400 : 500;
      if (status !== 404) console.error(new Date().toISOString(), 'sheetpulse update error:', err.message);
      return send(res, req, status, JSON.stringify({ error: err.message }), 'application/json');
    }
  }

`;

fs.writeFileSync(file + '.bak-update', src);
src = src.replace(oldRequire, newRequire);
src = src.replace(routeAnchor, route + routeAnchor);
fs.writeFileSync(file, src);
console.log('Patched index.js (backup: index.js.bak-update).');
