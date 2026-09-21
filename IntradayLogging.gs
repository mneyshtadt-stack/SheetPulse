// TASE moved from Sunday-Thursday to Monday-Friday trading on 2026-01-05, and Friday runs a
// shortened day (~9:55-13:50) instead of the usual 10:00-17:30.
function isTaseOpenNow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const hm = Utilities.formatDate(now, 'Asia/Jerusalem', 'HHmm');
  if (weekday === 'Fri') return hm >= '1000' && hm <= '1400';
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu'];
  return openDays.indexOf(weekday) !== -1 && hm >= '1000' && hm <= '1730';
}

// Standard US market hours, Monday-Friday.
function isUsMarketOpenNow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'America/New_York', 'EEE');
  const hm = Utilities.formatDate(now, 'America/New_York', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm >= '0930' && hm <= '1600';
}

// A narrow window just before TASE opens, used only to capture the real USD/ILS rate as of the
// actual start of today's session (see logIntradayValue). isTaseOpenNow() only turns true right
// at 10:00 itself -- by the time it fires, "the session-open rate" would really mean whatever rate
// happened to be live at that exact trigger run, not a rate genuinely captured ahead of the open.
function isTaseOpenGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const hm = Utilities.formatDate(now, 'Asia/Jerusalem', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm >= '0958' && hm < '1000';
}

// Same idea, just ahead of USA's own open, so the USA chart's own 16:30 point can carry its own
// "rate vs. previous close" marker instead of only the Portfolio chart's 10:00 point having one.
function isUsOpenGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'America/New_York', 'EEE');
  const hm = Utilities.formatDate(now, 'America/New_York', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm >= '0928' && hm < '0930';
}

// Apps Script's "every minute" triggers aren't perfectly precise — they can drift by a few
// minutes under load, so the last run before the exact close sometimes lands a few minutes
// early (e.g. 22:56 Israel time instead of 23:00) and the next run doesn't fire until after the
// open-market guard has already turned off, silently dropping the actual closing reading. These
// two functions define a short window right after each close where, once per day, we log one
// final reading explicitly stamped at the true close time (not whenever this happened to run).
function isTaseCloseGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const hm = Utilities.formatDate(now, 'Asia/Jerusalem', 'HHmm');
  if (weekday === 'Fri') return hm > '1350' && hm <= '1400';
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu'];
  return openDays.indexOf(weekday) !== -1 && hm > '1731' && hm <= '1741';
}
function taseCloseInstant() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const dateStr = Utilities.formatDate(now, 'Asia/Jerusalem', 'yyyy-MM-dd');
  const closeTime = weekday === 'Fri' ? '13:50:00' : '17:30:00';
  return Utilities.parseDate(dateStr + ' ' + closeTime, 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm:ss');
}
function isUsCloseGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'America/New_York', 'EEE');
  const hm = Utilities.formatDate(now, 'America/New_York', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm > '1600' && hm <= '1610';
}
function usCloseInstant() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'America/New_York', 'yyyy-MM-dd');
  return Utilities.parseDate(dateStr + ' 16:00:00', 'America/New_York', 'yyyy-MM-dd HH:mm:ss');
}

// The dashboard's chart only ever plots today's rows, so anything older is dead weight that
// just makes the sheet slower and every chart load download more than it needs. Clearing
// yesterday's rows the first time we log each new day keeps the sheet permanently small. This
// works the same whether rows are appended (ascending) or inserted at the top (descending),
// since row 2 always belongs to whatever day the current block of un-pruned rows is from.
function pruneIfNewDay(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // just the header, or empty — nothing to prune
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const firstRowDate = sheet.getRange(2, 1).getValue();
  const firstRowDay = Utilities.formatDate(new Date(firstRowDate), timeZone, 'yyyy-MM-dd');
  if (firstRowDay !== today) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

// A single ticker's price feed can glitch for one read (e.g. GOOGLEFINANCE briefly returning
// #N/A or 0), which shows up as a sharp V-shaped spike in the logged total that self-corrects
// the very next minute. Rather than logging every reading as-is, treat a jump of more than
// THRESHOLD from the last *logged* value as suspect and hold it back for one run — only commit
// it once the SAME unusual value shows up again on the following run, which means it's a real
// move rather than a one-off misread. propKey keeps the IL and US segments' pending state separate.
function isOutlierAndUnconfirmed(newValue, propKey) {
  const props = PropertiesService.getScriptProperties();
  const lastValue = Number(props.getProperty(propKey + '_last'));
  if (!lastValue) { props.setProperty(propKey + '_last', String(newValue)); return false; }

  const pctChange = Math.abs(newValue - lastValue) / lastValue;
  const THRESHOLD = 0.02; // 2% in a single minute is unusual for a diversified segment

  if (pctChange <= THRESHOLD) {
    props.deleteProperty(propKey + '_pending');
    props.setProperty(propKey + '_last', String(newValue));
    return false; // normal reading, log it
  }

  const pending = Number(props.getProperty(propKey + '_pending'));
  if (pending && Math.abs(newValue - pending) / pending <= THRESHOLD) {
    // matches what we saw last run too -> confirmed real move, let it through
    props.deleteProperty(propKey + '_pending');
    props.setProperty(propKey + '_last', String(newValue));
    return false;
  }

  // unconfirmed big jump -> remember it, skip logging this run
  props.setProperty(propKey + '_pending', String(newValue));
  return true;
}

// Google Sheets renders a raw Date value using the spreadsheet's locale, which for this
// spreadsheet defaults to US-style m/d/yyyy. Stamping the cell's own number format explicitly,
// every time a row is written, keeps the Timestamp column showing dd/mm/yyyy regardless of that
// default — and matches the explicit dd/mm/yyyy parser the dashboard's chart code now expects.
const TIMESTAMP_FORMAT = 'dd/mm/yyyy hh:mm:ss';

// LastClose keeps a rolling window of the last 7 trading days, one row per calendar date (Israel
// time), with columns Date / TASE market / USA market / USA market$ / USD/ILS Rate. A given day's
// row often gets filled in twice (TASE's columns when it closes, USA's columns later, since they
// close at different times) rather than all at once. USA market$ is read from the Tracker sheet's
// own "Current Value ($)" column rather than derived by dividing the ILS total by the FX rate.
// USD/ILS Rate is only ever written at USA's close (23:00), not TASE's -- it's meant to represent
// the day's final rate, and TASE closing hours earlier isn't that. Pruned by row COUNT, not
// calendar span -- since the trigger never fires Saturday/Sunday, this naturally holds the last 7
// actual trading days.
function setLastClose(market, timestamp, value, fxRate, usdValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('LastClose');
  if (!sheet) {
    sheet = ss.insertSheet('LastClose');
    sheet.appendRow(['Date', 'TASE market', 'USA market', 'USA market$', 'USD/ILS Rate']);
  }
  const dateStr = Utilities.formatDate(timestamp, 'Asia/Jerusalem', 'dd/MM/yyyy');
  const col = market === 'IL' ? 2 : 3;
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === dateStr) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) {
    sheet.appendRow([dateStr]);
    rowIdx = sheet.getLastRow();
  }
  sheet.getRange(rowIdx, col).setValue(value);
  if (market === 'US' && !isNaN(usdValue)) sheet.getRange(rowIdx, 4).setValue(usdValue);
  if (market === 'US' && !isNaN(fxRate)) sheet.getRange(rowIdx, 5).setValue(fxRate);

  const dataRows = sheet.getLastRow() - 1;
  const excess = dataRows - 7;
  if (excess > 0) sheet.deleteRows(2, excess);
}

// One continuous log spanning both markets' full active hours (~09:59-23:10 Israel time),
// replacing the old split of IntradayLogIL (TASE hours) + IntradayLogUS (USA hours) into two
// separate sheets. Current Value (ILS) is live on the Tracker sheet for BOTH segments regardless
// of which market is actually open right now -- GOOGLEFINANCE just holds a closed market's last
// price steady -- so reading both every tick naturally produces the right value everywhere: IL
// flat once TASE closes, USA flat-but-FX-refreshed before it opens (yesterday's US close re-priced
// at today's live rate), both genuinely live during the 16:30-17:30 overlap, USA alone live once
// TASE has closed for the day.
function logIntradayValue() {
  const taseOpen = isTaseOpenNow();
  const usOpen = isUsMarketOpenNow();
  const taseGraceWindow = !taseOpen && isTaseCloseGraceWindow();
  const usGraceWindow = !usOpen && isUsCloseGraceWindow();
  const isOpen = taseOpen || usOpen;

  const props = PropertiesService.getScriptProperties();
  const todayIL = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd');
  const todayET = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');

  if (!isOpen && !taseGraceWindow && !usGraceWindow) {
    // A couple of minutes before TASE actually opens, stash the live USD/ILS rate in a script
    // property (once per day) -- this becomes column C below, and is also what the frontend
    // compares against LastClose's own recorded rate for the "FX move since previous close"
    // indicator at the Portfolio chart's 10:00 point.
    if (isTaseOpenGraceWindow() && props.getProperty('il_open_rate_date') !== todayIL) {
      const tracker = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tracker');
      const rate = Number(tracker.getRange(1, 17).getValue()); // Q1
      if (!isNaN(rate)) {
        props.setProperty('il_open_rate_date', todayIL);
        props.setProperty('il_open_rate_value', String(rate));
      }
    }
    return;
  }

  // Each grace window only needs to fire once a day; once logged, later ticks still inside the
  // same (multi-minute, drift-tolerant) window are treated as normal readings instead.
  const taseGrace = taseGraceWindow && props.getProperty('tase_close_logged_date') !== todayIL;
  const usGrace = usGraceWindow && props.getProperty('us_close_logged_date') !== todayET;
  if (!isOpen && !taseGrace && !usGrace) return; // both grace windows already logged today

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getSheetByName('Tracker');
  const data = tracker.getDataRange().getValues();

  let headerRowIdx = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].includes('Ticker') && data[i].includes('Shares')) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) throw new Error('Could not find the holdings header row');
  const headers = data[headerRowIdx];
  const cTicker = headers.indexOf('Ticker');
  const cValue = headers.indexOf('Current Value (ILS)');
  if (cTicker === -1 || cValue === -1) throw new Error('Expected columns not found');

  const israeliTickers = ['IBI.F35', 'IBI.FK4'];
  let ilValue = 0;
  let usValue = 0;
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const ticker = String(data[i][cTicker] || '').trim();
    if (!ticker) break;
    const v = Number(data[i][cValue]);
    if (isNaN(v)) continue;
    if (israeliTickers.indexOf(ticker) !== -1) ilValue += v;
    else usValue += v;
  }

  // A glitch in either segment's price feed shouldn't quietly corrupt the row -- hold the whole
  // row back for one tick if EITHER side looks like an unconfirmed spike; it logs next run once
  // confirmed (or once the feed self-corrects).
  const ilOutlier = isOutlierAndUnconfirmed(ilValue, 'il');
  const usOutlier = isOutlierAndUnconfirmed(usValue, 'us');
  if (ilOutlier || usOutlier) return;

  let logSheet = ss.getSheetByName('IntradayLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('IntradayLog');
    logSheet.appendRow(['Timestamp', 'IL Value', 'USD/ILS Rate (session open)', 'USA Value (ILS, live)']);
  }
  pruneIfNewDay(logSheet, 'Asia/Jerusalem');
  // Whatever's left after pruning is just the header row (1) the first time today's data gets
  // logged -- captured here, on that one row only, so the dashboard has a real recorded rate for
  // "the start of today's session" instead of reapplying whatever rate happens to be live at
  // whatever later moment someone views the chart.
  const isFirstRowToday = logSheet.getLastRow() <= 1;
  logSheet.insertRowBefore(2);
  logSheet.getRange(2, 1).setNumberFormat(TIMESTAMP_FORMAT);
  const tsValue = taseGrace ? taseCloseInstant() : (usGrace ? usCloseInstant() : new Date());
  logSheet.getRange(2, 1, 1, 2).setValues([[tsValue, ilValue]]);
  logSheet.getRange(2, 4).setValue(usValue);

  if (isFirstRowToday) {
    // Prefer the rate captured ahead of the open (see the isTaseOpenGraceWindow branch above) --
    // only reads live right now as a fallback, for the rare day a deploy or a missed trigger run
    // means that earlier capture never happened.
    let openFxRate = props.getProperty('il_open_rate_date') === todayIL
      ? Number(props.getProperty('il_open_rate_value'))
      : NaN;
    if (isNaN(openFxRate)) openFxRate = Number(tracker.getRange(1, 17).getValue()); // Q1
    if (!isNaN(openFxRate)) logSheet.getRange(2, 3).setValue(openFxRate);
  }

  // USA's own session-open rate reuses this same column, on whichever row happens to log during
  // that narrow window (~16:28-16:30 Israel time) instead -- the trigger is already running
  // continuously through the whole combined window, so no separate early-return branch is needed
  // the way TASE's own capture has above. The row's own Timestamp is what tells the two captures
  // apart on the frontend (an afternoon hour means it's USA's), so a second column isn't needed to
  // disambiguate them. Guarded separately so it only ever writes once per day.
  if (isUsOpenGraceWindow() && props.getProperty('il_us_open_rate_date') !== todayIL) {
    const usOpenRate = Number(tracker.getRange(1, 17).getValue()); // Q1
    if (!isNaN(usOpenRate)) {
      logSheet.getRange(2, 3).setValue(usOpenRate);
      props.setProperty('il_us_open_rate_date', todayIL);
    }
  }

  if (taseGrace) {
    props.setProperty('tase_close_logged_date', todayIL);
    setLastClose('IL', taseCloseInstant(), ilValue); // no fxRate -- USD/ILS Rate is only written at USA's close
  }
  if (usGrace) {
    props.setProperty('us_close_logged_date', todayET);
    const fxRate = Number(tracker.getRange(1, 17).getValue()); // Q1

    // USA market$ is only ever needed once a day here, at the close — this second pass over the
    // already-fetched `data` sums Current Value ($) for the same US tickers, kept entirely
    // separate from the main loop above so the per-tick chart-logging path is untouched.
    const cValueUsd = headers.indexOf('Current Value ($)');
    let usValueUsd = NaN;
    if (cValueUsd !== -1) {
      usValueUsd = 0;
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const ticker = String(data[i][cTicker] || '').trim();
        if (!ticker) break;
        if (israeliTickers.indexOf(ticker) === -1) {
          const v = Number(data[i][cValueUsd]);
          if (!isNaN(v)) usValueUsd += v;
        }
      }
    } else {
      console.warn('Current Value ($) column not found. Sheet headers were:', headers);
    }
    setLastClose('US', usCloseInstant(), usValue, fxRate, usValueUsd);
  }
}
