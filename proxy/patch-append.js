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
