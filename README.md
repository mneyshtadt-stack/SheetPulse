# Personal Finance & Investment Dashboard

A single-file, client-side dashboard for tracking a personal investment portfolio and net worth over time. No backend, no database, no build step — just one HTML file that pulls live data from Google Sheets.

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

Everything is driven by Google Sheets' **File → Share → Publish to web → CSV** feature. The page fetches that published CSV with `fetch()` and parses it client-side with [PapaParse](https://www.papaparse.com/); charts and the treemap are drawn with [D3.js](https://d3js.org/). Both libraries are loaded from a CDN — there's nothing to install and nothing to build.

Because the data source is a public "published" link rather than a private API, this only works for data you're comfortable being reachable by anyone with the link. A lightweight client-side password screen (SHA-256, checked in-browser) adds a layer of privacy for hosting the page itself, but it does **not** make the underlying spreadsheet link itself private — treat the CSV URL with the same care as the spreadsheet's sharing settings.

## Setup

1. Publish your portfolio sheet and (optionally) a separate net worth sheet to the web as CSV.
2. Open the dashboard, click **Sheet URL**, and paste in your published CSV URL(s).
3. For the Net Worth year switcher, publish the *entire* workbook (not a single sheet) so every year's tab is reachable from one URL — the dashboard finds each year's tab automatically.
4. Set your own password by hashing a passphrase with SHA-256 and swapping it into the lock screen constant in the HTML.

## Tech stack

- Plain HTML/CSS/JavaScript — no framework, no bundler.
- [D3.js](https://d3js.org/) for charts and the treemap.
- [PapaParse](https://www.papaparse.com/) for CSV parsing.
- Hosted as a static page (e.g. GitHub Pages).

## Privacy

No data is sent anywhere except directly from your browser to Google, to fetch the published CSV. There is no server, no analytics, and no third party involved in loading your data.
