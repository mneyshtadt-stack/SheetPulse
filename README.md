# Personal Finance & Investment Dashboard

A single-file, client-side dashboard for tracking a personal investment portfolio and net worth over time. No framework, no build step — just one HTML file that pulls live data from Google Sheets through a small private proxy.

## Features

- **Portfolio overview** — Israeli (TASE) and USA holdings, day gain/loss, a size-weighted treemap of positions, and a live USD/ILS rate pulled straight from the sheet.
- **Sector & role-group allocation** — visual breakdowns of what each holding is for (growth, income, hedge, etc.), not just what it is.
- **Net Worth tracker** — assets, liabilities, and pension tracked monthly, with:
  - A hero net worth figure with month-over-month and year-to-date comparisons.
  - Trend charts for Net Worth, Portfolio, Study Funds, and Pension.
  - A **year switcher** that automatically detects each year's tab in the underlying spreadsheet (2026, 2027, …) with no code changes needed as new years are added.
- **Market Universe** — a sortable research watchlist of tickers with sector, category, and role classifications, plus hover tooltips for the fuller "why I hold this" and "what this fund is" notes.
- **Live refresh** — open detail views refresh in place on a timer without losing your current search/sort/filter state.

## How it works

The page fetches sheet data with fetch() and parses it client-side with PapaParse (https://www.papaparse.com/); charts and the treemap are drawn with D3.js (https://d3js.org/). Both libraries are loaded from a CDN — there's nothing to install and nothing to build.

Data no longer comes from Google Sheets' public "Publish to web" CSV export. Instead, the page calls a small self-hosted proxy (a Node.js process behind Caddy, on a free-tier VM) which authenticates to the Google Sheets API v4 with an OAuth refresh token and returns the same CSV-shaped data. The proxy checks a bearer token and only accepts requests from this page's own origin (CORS-restricted), and the underlying spreadsheets are not published to the web at all — so the raw data is no longer reachable by anyone who merely has a URL.

That said, this is a static, client-only page: the bearer token has to ship inside this page's JavaScript for the browser to use it, so it's still visible to anyone who reads this file's source. This setup raises the bar against casual scraping and search-engine indexing of a bare public link — it does not provide true secrecy against a determined reader of the page's own code. A lightweight client-side password screen (SHA-256, checked in-browser) separately gates viewing the page itself.

## Setup

This dashboard is wired to one specific Google account's sheets via hardcoded constants in the HTML (SHEET_ID, NET_WORTH_SHEET_ID, PROXY_BASE_URL, PROXY_AUTH_TOKEN) — it isn't meant to be reconfigured per-viewer. To point it at different sheets:

1. Stand up the proxy: a small Node.js server (using the googleapis package with an OAuth2 client + refresh token) behind a reverse proxy with HTTPS, reachable at some PROXY_BASE_URL. It exposes ?sheetId=...&sheetName=... (returns that tab as CSV) and ?sheetId=...&action=list (returns all tab names — used for the Net Worth year switcher).
2. Update the SHEET_ID / NET_WORTH_SHEET_ID constants in the HTML to your own spreadsheet IDs (from each sheet's normal .../spreadsheets/d/<ID>/edit URL — the file must be in native Google Sheets format, not an uploaded Excel file).
3. Set PROXY_AUTH_TOKEN to match the token your proxy expects, and restrict the proxy's CORS to the origin you're hosting this page on.
4. Set your own password by hashing a passphrase with SHA-256 and swapping it into the lock screen constant in the HTML.

## Tech stack

- Plain HTML/CSS/JavaScript — no framework, no bundler.
- D3.js (https://d3js.org/) for charts and the treemap.
- PapaParse (https://www.papaparse.com/) for CSV parsing.
- A Node.js proxy (OAuth2 + Google Sheets API v4) behind Caddy for automatic HTTPS, self-hosted on a free-tier VM.
- Hosted as a static page (e.g. GitHub Pages).

## Privacy

No data is sent anywhere except from your browser to the private proxy, and from the proxy to Google's Sheets API. There is no analytics and no third party involved in loading your data. The spreadsheets themselves are not published to the web; access requires the bearer token this page sends, which the proxy checks alongside an origin restriction.
