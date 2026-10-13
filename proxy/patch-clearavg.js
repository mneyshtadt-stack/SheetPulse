// One-time patch for /home/ubuntu/sheet-proxy/index.js: adds POST /sheetpulse/clear-avg (clears one ticker's
// hand-typed Broker Avg Cost in the Tracker's Transactions tab), next to /sheetpulse/update.
// Safe to re-run -- it does nothing if the route is already there. Keeps a backup as index.js.bak-clearavg.
// Usage (in /home/ubuntu/sheet-proxy):  node patch-clearavg.js

const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, 'index.js');
let src = fs.readFileSync(file, 'utf8');

if (src.includes('/sheetpulse/clear-avg')) {
  console.log('Already patched - nothing to do.');
  process.exit(0);
}

const oldRequire = "const { handleSheetPulseAppend, handleSheetPulseUpdate, SheetPulseAppendError, SheetPulseNotFound } = require('./sheetpulse-append');";
const routeAnchor = '  const sheetId = parsed.query.sheetId;';
if (!src.includes(oldRequire) || !src.includes(routeAnchor)) {
  console.error('Could not find the expected lines in index.js (is the update route installed?) - not changed.');
  process.exit(1);
}

const newRequire = "const { handleSheetPulseAppend, handleSheetPulseUpdate, handleSheetPulseClearAvg, SheetPulseAppendError, SheetPulseNotFound } = require('./sheetpulse-append');";
const route = `  // --- SheetPulse: clear a hand-typed Broker Avg Cost in the Tracker (same token check as the reads above) ---
  if (parsed.pathname === '/sheetpulse/clear-avg' && req.method === 'POST') {
    try {
      const body = await readJsonBody(req);
      const sheets = await getSheetsClient();
      const result = await handleSheetPulseClearAvg(sheets, body);
      return send(res, req, 200, JSON.stringify(result), 'application/json');
    } catch (err) {
      const status = err instanceof SheetPulseNotFound ? 404 : err instanceof SheetPulseAppendError ? 400 : 500;
      console.error(new Date().toISOString(), 'sheetpulse clear-avg error:', err.message);
      return send(res, req, status, JSON.stringify({ error: err.message }), 'application/json');
    }
  }

`;

fs.writeFileSync(file + '.bak-clearavg', src);
src = src.replace(oldRequire, newRequire);
src = src.replace(routeAnchor, route + routeAnchor);
fs.writeFileSync(file, src);
console.log('Patched index.js (backup: index.js.bak-clearavg).');
