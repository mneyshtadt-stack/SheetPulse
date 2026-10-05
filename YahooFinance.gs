// Fetches a ticker's live data from Yahoo Finance's own (unofficial, undocumented) chart API --
// for manual side-by-side comparison against GOOGLEFINANCE in Tracker_Yahoo. Not wired into the
// main Tracker sheet, IntradayLog, or any trigger; purely a formula you call directly in cells.
//
// Usage: =YAHOOFINANCE("AAPL")                 -> current price
//        =YAHOOFINANCE("AAPL", "changepct")    -> day change %
//        =YAHOOFINANCE("AAPL", "change")       -> day change, in the ticker's own currency
//        =YAHOOFINANCE("AAPL", "previousclose")
//        =YAHOOFINANCE("AAPL", "currency")     -> e.g. "USD"
//        =YAHOOFINANCE("AAPL", "name")
//
// Israeli/TASE tickers: Yahoo's own symbol isn't the same one Google uses. Confirmed working for
// IBI.F35 / IBI.FK4 specifically as IBI-F35.TA / IBI-FK4.TA (periods become hyphens, plus a ".TA"
// suffix) -- tried automatically as a second fallback below.
//
// IMPORTANT unit mismatch on TASE tickers: Yahoo quotes these in AGOROT (1/100 ILS) -- e.g.
// IBI-F35.TA's price comes back as ~7873, which is really ~78.73 ILS. "changepct" is a ratio so
// it's unaffected, but "price"/"previousclose"/"change" are 100x too large if used as-is in place
// of GOOGLEFINANCE's ILS-denominated figures. Divide by 100 (or use "changepct" only) for any TASE
// ticker before comparing against Google's numbers.
//
// Returns the string "#N/A" on any failure (ticker not found, Yahoo unreachable, bad field name)
// so it behaves the same way GOOGLEFINANCE does on error -- IFERROR(GOOGLEFINANCE(...),
// YAHOOFINANCE(...)) works as expected once/if this gets wired in as a real fallback.
function YAHOOFINANCE(ticker, field) {
  field = (field || 'price').toString().toLowerCase();
  if (!ticker) return '#N/A';

  let parsed = fetchYahooQuote_(ticker);
  // Google and Yahoo disagree on share-class tickers -- Google uses a period (BRK.B), Yahoo a
  // hyphen (BRK-B). Retry with the substitution only when the plain symbol didn't resolve and it
  // actually looks like a share-class suffix (a single letter after the last period), so this
  // doesn't misfire on tickers where a period means something else.
  if (!parsed && /\.[A-Za-z]$/.test(ticker)) {
    parsed = fetchYahooQuote_(ticker.replace(/\.([A-Za-z])$/, '-$1'));
  }
  // TASE tickers: Google's "IBI.F35" style becomes Yahoo's "IBI-F35.TA" -- every period turns into
  // a hyphen, plus a ".TA" suffix. Tried last since it's the most specific/unlikely transform.
  if (!parsed && ticker.indexOf('.') !== -1) {
    parsed = fetchYahooQuote_(ticker.replace(/\./g, '-') + '.TA');
  }
  if (!parsed) return '#N/A';

  const price = parsed.price;
  const prevClose = parsed.previousClose;

  switch (field) {
    case 'price': return price != null ? price : '#N/A';
    case 'previousclose': return prevClose != null ? prevClose : '#N/A';
    case 'change': return (price != null && prevClose != null) ? (price - prevClose) : '#N/A';
    case 'changepct': return (price != null && prevClose) ? ((price - prevClose) / prevClose * 100) : '#N/A';
    case 'currency': return parsed.currency || '#N/A';
    case 'name': return parsed.name || '#N/A';
    default: return '#N/A';
  }
}

// Returns a plain object ({price, previousClose, currency, name}) for one symbol, or null on any
// failure (not found, Yahoo unreachable, bad response shape) -- never throws, so callers can just
// check truthiness. Custom functions can recalculate far more often than a human would ever
// refresh by hand, so each symbol's result is cached for 60 seconds to avoid hitting Yahoo once
// per ticker per recalculation tick.
function fetchYahooQuote_(symbol) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'yf_' + symbol;
  const cached = cache.get(cacheKey);
  if (cached) return cached === 'null' ? null : JSON.parse(cached);

  let result = null;
  try {
    // Yahoo appears to reject Apps Script's own default request signature outright (a
    // connection-level failure, not even a proper HTTP error response) -- a realistic
    // browser-like User-Agent is what got a plain server-side test through.
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol);
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json'
      }
    });
    if (resp.getResponseCode() === 200) {
      const json = JSON.parse(resp.getContentText());
      const chartResult = json.chart && json.chart.result && json.chart.result[0];
      if (chartResult && chartResult.meta) {
        const meta = chartResult.meta;
        result = {
          price: meta.regularMarketPrice != null ? meta.regularMarketPrice : null,
          previousClose: meta.chartPreviousClose != null ? meta.chartPreviousClose : (meta.previousClose != null ? meta.previousClose : null),
          currency: meta.currency || '',
          name: meta.longName || meta.shortName || ''
        };
      }
    }
  } catch (e) {
    result = null;
  }
  cache.put(cacheKey, result ? JSON.stringify(result) : 'null', 60);
  return result;
}
