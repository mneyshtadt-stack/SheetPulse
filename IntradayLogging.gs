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

// A window just before TASE opens, in which logIntradayValue writes the day's opening reading
// (stamped exactly 10:00:00, see taseOpenInstant) -- so the TASE and Portfolio charts have a point
// the moment TASE opens instead of waiting for the first regular tick. TASE prices can't move
// before the open, so the reading is exactly the opening value (USA re-priced at the live rate).
// Six minutes wide so a ~5-minute trigger is sure to run inside it at least once.
function isTaseOpenGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const hm = Utilities.formatDate(now, 'Asia/Jerusalem', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm >= '0954' && hm < '1000';
}
function taseOpenInstant() {
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd');
  return Utilities.parseDate(dateStr + ' 10:00:00', 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm:ss');
}

// Just ahead of USA's own open (09:24-09:30 ET, ~16:24-16:30 Israel time): the row logged here is
// the one the dashboard takes USA's opening USD/ILS rate from (the USA chart's 16:30 FX marker) and
// uses as that chart's 16:30 starting point. Six minutes wide so a ~5-minute trigger is sure to run
// inside it.
function isUsOpenGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'America/New_York', 'EEE');
  const hm = Utilities.formatDate(now, 'America/New_York', 'HHmm');
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  return openDays.indexOf(weekday) !== -1 && hm >= '0924' && hm < '0930';
}

// Apps Script's "every minute" triggers aren't perfectly precise — they can drift by a few
// minutes under load, so the last run before the exact close sometimes lands a few minutes
// early (e.g. 22:56 Israel time instead of 23:00) and the next run doesn't fire until after the
// open-market guard has already turned off, silently dropping the actual closing reading. These
// two functions define a short window right after each close where, once per day, we log one
// final reading explicitly stamped at the true close time (not whenever this happened to run).
// Friday's session (and isTaseOpenNow) runs to 14:00, so its close window starts right after that.
function isTaseCloseGraceWindow() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const hm = Utilities.formatDate(now, 'Asia/Jerusalem', 'HHmm');
  if (weekday === 'Fri') return hm > '1400' && hm <= '1410';
  const openDays = ['Mon', 'Tue', 'Wed', 'Thu'];
  return openDays.indexOf(weekday) !== -1 && hm > '1731' && hm <= '1741';
}
function taseCloseInstant() {
  const now = new Date();
  const weekday = Utilities.formatDate(now, 'Asia/Jerusalem', 'EEE');
  const dateStr = Utilities.formatDate(now, 'Asia/Jerusalem', 'yyyy-MM-dd');
  const closeTime = weekday === 'Fri' ? '14:00:00' : '17:30:00';
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

// Keeps the sheet down to the most recent INTRADAY_KEEP_DATES distinct trading dates found in it
// (counted by dates actually present, so weekends need no special-casing: Saturday/Sunday never
// log anything, so Friday's rows stay valid as "the previous date" straight through Monday). Anything older than that gets deleted. Rows are
// always newest-first (each write does insertRowBefore(2)), so once a row older than the kept
// dates is found, everything below it is old too -- one contiguous block.
// This used to just wipe everything the moment a new day's first row logged, but the USA-only
// chart's own session doesn't start until ~16:30 -- wiping at TASE's 10:00 open destroyed
// yesterday's data (including yesterday's full USA session) hours before the USA chart's own
// "show the last completed session" fallback needed it, leaving that chart with nothing to show
// for the whole TASE-only morning. The dashboard's 5D chart range replots the last 5 trading
// dates straight from this sheet, so 7 are kept: 5 plus a margin in case the trigger misses a
// whole day's runs (observed before, see README). About 160 rows a day, so ~1,100 rows total.
const INTRADAY_KEEP_DATES = 7;
function pruneOldRows(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // just the header, or empty — nothing to prune
  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => Utilities.formatDate(new Date(r[0]), timeZone, 'yyyy-MM-dd'));
  const distinctDates = [...new Set(dates)]; // newest-first, since the rows themselves are
  if (distinctDates.length <= INTRADAY_KEEP_DATES) return;
  const keepDates = new Set(distinctDates.slice(0, INTRADAY_KEEP_DATES));
  let cutoffRow = -1;
  for (let i = 0; i < dates.length; i++) {
    if (!keepDates.has(dates[i])) { cutoffRow = i + 2; break; } // +2: sheet row number, 1-indexed past the header
  }
  if (cutoffRow !== -1) sheet.deleteRows(cutoffRow, lastRow - cutoffRow + 1);
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

// IntradayLog's columns. Each market's close is its own row, stamped exactly at the close (TASE
// 17:30:00, 14:00:00 on Fridays; USA 16:00:00 New York time, normally 23:00:00 in Israel) and filled
// green -- the dashboard reads each day's closes straight from those rows (they replaced the old
// LastClose tab). USA Value ($) is USA's holdings in dollars (Tracker's "Current Value ($)"), which
// the dashboard's Market Value (USD) tiles compare against.
const INTRADAY_HEADER = ['Timestamp', 'IL Value', 'USD/ILS Rate', 'USA Value (ILS, live)', 'USA Value ($)'];
const CLOSE_ROW_COLOR = '#00ff00';

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

  // Once a day, just before TASE opens: the opening reading, stamped 10:00:00. Its USD/ILS rate is
  // the one the dashboard compares against the previous close's for the FX marker at 10:00.
  const taseOpenGrace = !isOpen && isTaseOpenGraceWindow() && props.getProperty('tase_open_logged_date') !== todayIL;
  // Normally TASE is still open when USA's pre-open window comes round, so the logger is running
  // anyway; on Fridays TASE has closed at 14:00, so let one regular row through in that window.
  const usPreOpen = !isOpen && isUsOpenGraceWindow() && props.getProperty('us_preopen_logged_date') !== todayIL;
  if (!isOpen && !taseGraceWindow && !usGraceWindow && !taseOpenGrace && !usPreOpen) return;

  // Each grace window only needs to fire once a day; once logged, later ticks still inside the
  // same (multi-minute, drift-tolerant) window are treated as normal readings instead.
  const taseGrace = taseGraceWindow && props.getProperty('tase_close_logged_date') !== todayIL;
  const usGrace = usGraceWindow && props.getProperty('us_close_logged_date') !== todayET;
  if (!isOpen && !taseGrace && !usGrace && !taseOpenGrace && !usPreOpen) return; // grace windows already logged today

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
  const cValueUsd = headers.indexOf('Current Value ($)');
  if (cTicker === -1 || cValue === -1) throw new Error('Expected columns not found');
  if (cValueUsd === -1) console.warn('Current Value ($) column not found. Sheet headers were:', headers);

  const israeliTickers = ['IBI.F35', 'IBI.FK4'];
  let ilValue = 0;
  let usValue = 0;
  let usValueUsd = cValueUsd === -1 ? NaN : 0;
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const ticker = String(data[i][cTicker] || '').trim();
    if (!ticker) break;
    const v = Number(data[i][cValue]);
    if (isNaN(v)) continue;
    if (israeliTickers.indexOf(ticker) !== -1) { ilValue += v; continue; }
    usValue += v;
    if (cValueUsd !== -1) {
      const u = Number(data[i][cValueUsd]);
      if (!isNaN(u)) usValueUsd += u;
    }
  }

  // A glitch in either segment's price feed shouldn't quietly corrupt the row -- hold the whole
  // row back for one tick if EITHER side looks like an unconfirmed spike; it logs next run once
  // confirmed (or once the feed self-corrects).
  const ilOutlier = isOutlierAndUnconfirmed(ilValue, 'il');
  const usOutlier = isOutlierAndUnconfirmed(usValue, 'us');
  if (ilOutlier || usOutlier) return;

  const logSheet = intradayLogSheet_(ss);
  pruneOldRows(logSheet, 'Asia/Jerusalem');
  logSheet.insertRowBefore(2);
  logSheet.getRange(2, 1).setNumberFormat(TIMESTAMP_FORMAT);
  const tsValue = taseGrace ? taseCloseInstant() : (usGrace ? usCloseInstant() : (taseOpenGrace ? taseOpenInstant() : new Date()));
  // Rounded to 4 decimals, the precision the dashboard shows USD/ILS at everywhere.
  const liveRate = Math.round(Number(tracker.getRange(1, 17).getValue()) * 10000) / 10000; // Q1
  logSheet.getRange(2, 1, 1, 5).setValues([[tsValue, ilValue, isNaN(liveRate) ? '' : liveRate, usValue, isNaN(usValueUsd) ? '' : usValueUsd]]);
  logSheet.getRange(2, 3).setNumberFormat('0.0000');
  // Explicitly set every time: a newly inserted row picks up the formatting of the row below it,
  // which would otherwise spread a close row's green onto the next day's first rows.
  logSheet.getRange(2, 1, 1, INTRADAY_HEADER.length).setBackground(taseGrace || usGrace ? CLOSE_ROW_COLOR : null);

  if (taseOpenGrace) props.setProperty('tase_open_logged_date', todayIL);
  if (usPreOpen) props.setProperty('us_preopen_logged_date', todayIL);
  if (taseGrace) props.setProperty('tase_close_logged_date', todayIL);
  if (usGrace) props.setProperty('us_close_logged_date', todayET);
}

function intradayLogSheet_(ss) {
  let sheet = ss.getSheetByName('IntradayLog');
  if (!sheet) {
    sheet = ss.insertSheet('IntradayLog');
    sheet.appendRow(INTRADAY_HEADER);
  } else if (sheet.getRange(1, 1, 1, INTRADAY_HEADER.length).getValues()[0].join('|') !== INTRADAY_HEADER.join('|')) {
    sheet.getRange(1, 1, 1, INTRADAY_HEADER.length).setValues([INTRADAY_HEADER]);
  }
  return sheet;
}
