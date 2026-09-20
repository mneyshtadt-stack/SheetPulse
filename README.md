# SheetPulse

A single-file, client-side dashboard for tracking a personal investment portfolio and net worth over time. No framework, no build step — just one HTML file that pulls live data from Google Sheets through a small private proxy, behind a Google Sign-In wall.

## Features

- **Portfolio overview** — Israeli (TASE) and USA holdings, day gain/loss, a size-weighted treemap of positions, and a live USD/ILS rate pulled straight from the sheet.
- **Sector & role-group allocation** — visual breakdowns of what each holding is for (growth, income, hedge, etc.), not just what it is.
- **Net Worth tracker** — assets, liabilities, and pension tracked monthly, with:
  - A hero net worth figure with month-over-month and year-to-date comparisons.
  - Trend charts for Net Worth, Portfolio, Study Funds, and Pension.
  - A **year switcher** that automatically detects each year's tab in the underlying spreadsheet (2026, 2027, …) with no code changes needed as new years are added.
- **Market Universe** — a sortable research watchlist of tickers with sector, category, and role classifications, plus hover tooltips for the fuller "why I hold this" and "what this fund is" notes.
- **Live refresh** — open detail views refresh in place on a timer without losing your current search/sort/filter state.
- **Installable as an app** — a web app manifest and service worker let Android's Chrome (and other browsers) install this as a real full-screen app with its own icon, no browser chrome, via "Add to Home screen" / "Install app".
- **Biometric quick-unlock** — after signing in with Google once on a device, you can enable Face ID / Fingerprint (via WebAuthn) as a faster way back in on that same device, without repeating the full Google sign-in flow every time.

## How it works

The page fetches sheet data with `fetch()` and parses it client-side with [PapaParse](https://www.papaparse.com/); charts and the treemap are drawn with [D3.js](https://d3js.org/). Both libraries are loaded from a CDN — there's nothing to install and nothing to build.

Data doesn't come from Google Sheets' public "Publish to web" CSV export. Instead, the page calls a small self-hosted proxy (a Node.js process behind Caddy, on a free-tier Oracle Cloud VM) which authenticates to the Google Sheets API v4 with an OAuth refresh token and returns the same CSV-shaped data. The underlying spreadsheets are **not** published to the web at all, so the raw data isn't reachable by anyone who merely has a URL.

Access to the page itself is gated by **Google Sign-In**: the dashboard shows a "Sign in with Google" screen (via Google Identity Services) instead of any password. Once signed in, every request to the proxy carries the viewer's own Google ID token, and the proxy independently re-verifies that token on **every single request** — checking its signature against Google's own public keys and confirming its email matches the one allowed account — before it will read anything from the sheets. There is no static secret embedded in this page's source for a reader to find and reuse; the one Client ID that is embedded is meant to be public (that's how Google's browser sign-in flow is designed to work) and can't by itself be used to read any data.

Once signed in with Google, a device can additionally register a **WebAuthn** (Face ID / Fingerprint) credential as a per-device shortcut back in. Registering one still requires a fresh, real Google ID token; unlocking with the credential afterward has the proxy verify the biometric assertion itself and hand back a short-lived session token, used exactly like a Google ID token from then on. The biometric read never leaves your device — the proxy only ever stores a public key, and only for devices you've explicitly enabled.

## Setup

This dashboard is wired to one specific Google account's sheets and one allowed sign-in email via hardcoded constants in the HTML (`SHEET_ID`, `NET_WORTH_SHEET_ID`, `PROXY_BASE_URL`, `GOOGLE_WEB_CLIENT_ID`, `ALLOWED_GOOGLE_EMAIL`) — it isn't meant to be reconfigured per-viewer. To point it at your own sheets and account:

1. Stand up the proxy: a small Node.js server (using the `googleapis` package with an OAuth2 client + refresh token for Sheets access, plus `google-auth-library` to verify incoming sign-ins) behind a reverse proxy with HTTPS, reachable at some `PROXY_BASE_URL`. It exposes `?sheetId=...&sheetName=...` (returns that tab as CSV) and `?sheetId=...&action=list` (returns all tab names — used for the Net Worth year switcher).
2. Create two OAuth clients in Google Cloud: a **Desktop app** client (for the proxy's own read access to your Sheets, via a one-time refresh-token exchange) and a **Web application** client (for Google Sign-In on the page itself), with the page's own origin listed under its Authorized JavaScript origins.
3. Update the `SHEET_ID` / `NET_WORTH_SHEET_ID` constants in the HTML to your own spreadsheet IDs (from each sheet's normal `.../spreadsheets/d/<ID>/edit` URL — the file must be in native Google Sheets format, not an uploaded Excel file).
4. Set `GOOGLE_WEB_CLIENT_ID` in the HTML to the Web client's ID, and set `ALLOWED_GOOGLE_EMAIL` to the Google account that should be allowed to sign in. On the proxy, set matching `WEB_CLIENT_ID` and `ALLOWED_EMAIL` values so it can independently verify the same thing server-side, and restrict its CORS to the origin you're hosting this page on.

## Local testing

`SheetPulse.html` is always the source of truth — it's the one file that ever gets edited by hand and the one file the deploy script reads. To try a change before it goes live:

1. Edit `SheetPulse.html`.
2. Run `TestLocal.bat` — it copies `SheetPulse.html` to `TestLocal.html` fresh (so the test copy can never drift out of sync), starts a local server, and opens `http://localhost:8000/TestLocal.html`. The URL itself makes it obvious you're looking at the local test copy, not production. `TestLocal.html` is disposable and gets overwritten every run — never edit it directly, since those changes would be silently lost.
3. If it looks good, run `deploy-to-github.ps1` — it reads `SheetPulse.html` directly and pushes it to GitHub as `index.html`. `TestLocal.html`/`TestLocal.bat` are never touched by the deploy script and never get pushed anywhere.

Testing locally works for Google Sign-In and live sheet data (the proxy's CORS and the Google OAuth client both additionally allow `http://localhost:8000` for this reason), but not for biometric quick-unlock — a WebAuthn credential is tied to its exact domain by design, so that one can only be tested on the real deployed URL.

## Tech stack

- Plain HTML/CSS/JavaScript — no framework, no bundler.
- [D3.js](https://d3js.org/) for charts and the treemap.
- [PapaParse](https://www.papaparse.com/) for CSV parsing.
- [Google Identity Services](https://developers.google.com/identity/gsi/web) for Sign In With Google.
- A Node.js proxy (OAuth2 + Google Sheets API v4 + `google-auth-library` for ID token verification) behind Caddy for automatic HTTPS, self-hosted on a free-tier Oracle Cloud VM.
- Hosted as a static page (e.g. GitHub Pages).
- A web app manifest (`manifest.json`) and a minimal service worker (`sw.js`) for PWA installability.
- The native browser WebAuthn API (`navigator.credentials`) on the frontend, and `@simplewebauthn/server` + `jsonwebtoken` on the proxy, for biometric quick-unlock.

## Privacy

No data is sent anywhere except from your browser to the private proxy, and from the proxy to Google's Sheets API. There is no analytics and no third party involved in loading your data. The spreadsheets themselves are not published to the web; access requires signing in with the one allowed Google account, which the proxy independently re-verifies on every request alongside an origin restriction. If you enable biometric unlock, your fingerprint/face data itself never leaves your device — only a public key is stored on the proxy, standard for WebAuthn.
