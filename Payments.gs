// Monthly bills from Gmail into the Payments tab, for SheetPulse's Payments window.
// Lives in its own spreadsheet ("SheetPulse Payments", separate from the Tracker): the script is
// bound to it (Extensions → Apps Script there), so the tab is created in that spreadsheet.
//
// Each source is a Gmail search (the user's "payments-bills" label plus the sender). The bill's
// figures are read from its PDF attachment: Drive converts the PDF to a temporary Google Doc (OCR,
// Hebrew), the text is read and the temporary file deleted. Only the extracted figures are stored,
// never the email or the PDF. Each email is processed once (its message id is kept in the tab).
//
// Payments tab: Date | Payee | Category | Amount | Period | Invoice date | Due date | Card | Invoice no | Message id | Added
// (Date = the email's date; Due date = when it's charged; Card = the card's last 4 digits.)
//
// Setup (once):
//   1. Apps Script editor → Services (+) → add "Drive API" (the advanced Drive service).
//   2. Run testPaymentsSource() and check the log: the extracted text and the parsed figures.
//   3. When the figures are right: run updatePayments() once, then installPaymentsTrigger().
//
// Adding a source later: one more entry in PAY_SOURCES with its own parse function.

const PAY_SHEET = 'Payments';
const PAY_HEADER = ['Date', 'Payee', 'Category', 'Amount', 'Period', 'Invoice date', 'Due date', 'Card', 'Invoice no', 'Message id', 'Added'];
const PAY_COL_MSG = 10;   // Message id column (1-based)
const PAY_SOURCES = [
  {
    payee: 'Partner Power', category: 'Electricity',
    // Partner also sends phone/internet invoices; the parser keeps only the Partner Power ones.
    query: 'label:payments-bills from:thankyou@partner.net.il has:attachment filename:pdf',
    parse: payParsePartnerPower_
  },
  {
    payee: 'NEXT TV', category: 'TV & Internet',
    query: 'label:payments-bills from:next has:attachment filename:pdf',
    parse: payParseNextTv_
  },
  {
    // The municipality's receipt (with PDF). cityPay's separate "אישור תשלום" confirmation of the same
    // payment has no PDF, so it's never picked up and nothing is counted twice.
    payee: 'Ramat Gan', category: 'Arnona',
    query: 'label:payments-bills from:citypay@onecity.co.il has:attachment filename:pdf',
    parse: payParseRamatGan_
  },
  {
    // The same receipts from cityPay's older address, up to Jul 2023 (onecity.co.il took over in Nov 2023).
    // A receipt whose transaction no. is already in the tab is skipped, so nothing is counted twice.
    payee: 'Ramat Gan', category: 'Arnona',
    query: 'label:payments-bills from:citypay@ladpc.co.il has:attachment filename:pdf',
    parse: payParseRamatGan_
  },
  {
    // Water (Mei Ramat Gan) paid through MAST: the amount is in the email text, there's no PDF.
    // All three MAST sender addresses over the years.
    payee: 'Mei Ramat Gan', category: 'Water', body: true,
    query: 'label:payments-bills (from:send.vpcontact.com OR from:mast.co.il OR from:mastdonotreply@gmail.com)',
    parse: payParseMast_
  }
];

// MAST confirmation: "שולם סכום של 350.00 שקלים חדשים, מספר תשלומים: 1. מספר אישור תשלום: 0123456."
// The payment date is the email's date; no card or period in the email. Null if it isn't a payment.
function payParseMast_(text) {
  const paid = text.match(/שולם\s*סכום\s*של\s*([\d,]+(?:\.\d+)?)/);
  if (!paid) return null;
  const approval = text.match(/מספר\s*אישור\s*תשלום:?\s*(\d+)/);
  const inst = text.match(/מספר\s*תשלומים:?\s*(\d+)/);
  return {amount: Number(paid[1].replace(/,/g, '')), period: inst && Number(inst[1]) > 1 ? inst[1] + ' installments' : '',
          invoiceDate: '', dueDate: '', card: '', invoiceNo: approval ? approval[1] : ''};
}

// Ramat Gan receipt ("קבלה למשלם"): arnona + guard levy. "בוצעה בתאריך 28/09/2026" (paid),
// "אשראי . 1234 08/10/2026 1,000.00" (card and charge date), "סה"כ שולם בש"ח 1,000.00", "מספר עסקה 1234567".
// Null if it isn't a receipt.
function payParseRamatGan_(text) {
  if (!/קבלה\s*למשלם/.test(text)) return null;
  const total = text.match(/סה["״]כ\s*שולם\s*בש["״]ח\s*([\d,]+\.\d{2})/) || text.match(/סה["״]כ\s*לתשלום\s*בש["״]ח\s*([\d,]+\.\d{2})/);
  const paid = text.match(/בוצעה\s*בתאריך\s*(\d{2}\/\d{2}\/\d{4})/);
  const charge = text.match(/אשראי[\s.]*?(\d{4})\s+(\d{2}\/\d{2}\/\d{4})/);
  const deal = text.match(/מספר\s*עסקה\s*(\d{5,})/);
  const year = text.match(/ארנונה/) && text.match(/\b(20\d{2})\s*ארנונה/);
  return {amount: total ? Number(total[1].replace(/,/g, '')) : NaN, period: year ? year[1] : '',
          invoiceDate: paid ? paid[1] : '', dueDate: charge ? charge[2] : '', card: charge ? charge[1] : '',
          invoiceNo: deal ? deal[1] : ''};
}

// NEXT TV (HOT) statement: "תקופת החשבון: 09.2026", "סה"כ חשבון נוכחי 120.00" (with VAT). The total is
// read from its own label: the largest number on the page can be a line item (it was once).
// No due date or card in the PDF. Null if it isn't a NEXT statement.
function payParseNextTv_(text) {
  if (!/NEXT/i.test(text)) return null;
  const total = text.match(/סה["״]כ\s*חשבון\s*נוכחי\s*([\d,]+\.\d{2})/);
  const period = text.match(/תקופת\s*החשבון:?\s*(\d{2}\.\d{4})/);
  return {amount: total ? Number(total[1].replace(/,/g, '')) : NaN, period: period ? period[1] : '',
          invoiceDate: '', dueDate: '', card: '', invoiceNo: ''};
}
const PAY_MAX_PER_RUN = 10;   // OCR is slow (a few seconds a PDF); older bills fill in over later runs

// Partner Power invoice: "חשבונית מס-מקור לתקופת החשבון 01.08.26-31.08.26", "סה"כ לתשלום 400.00",
// "תאריך עריכת חשבון 13/09/26", "מס' חשבונית 1234567", "חיוב כרטיס אשראי המסתיים בספרות 1234 לא לפני
// 24/09/26". Null if it isn't a Partner Power invoice (the logo is an image, so "חשמל" identifies it).
function payParsePartnerPower_(text) {
  if (!/Partner\s*Power/i.test(text) && text.indexOf('חשמל') === -1) return null;
  const period = text.match(/(\d{2}\.\d{2}\.\d{2,4})\s*[-–]\s*(\d{2}\.\d{2}\.\d{2,4})/);
  const invDate = text.match(/(\d{2}\/\d{2}\/\d{2,4})/);
  const invNo = text.match(/חשבונית[^\d\n]{0,12}(\d{5,})/) || text.match(/(\d{7})\s*\n?\s*:?\s*מס['׳]?\s*חשבונית/);
  const due = text.match(/לא\s*לפני\s*\.?\s*(\d{2}\/\d{2}\/\d{2,4})/);
  const card = text.match(/בספרות\s*(\d{4})/);
  return {amount: payTotalToPay_(text), period: period ? period[1] + '-' + period[2] : '',
          invoiceDate: invDate ? invDate[1] : '', dueDate: due ? due[1] : '', card: card ? card[1] : '',
          invoiceNo: invNo ? invNo[1] : ''};
}

// The total to pay: the amount on a line with "לתשלום" (to pay), else the largest amount on the page.
function payTotalToPay_(text) {
  const amounts = s => (s.match(/-?\d{1,3}(?:,\d{3})*\.\d{2}/g) || []).map(a => Number(a.replace(/,/g, '')));
  const lines = text.split(/\n/);
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].indexOf('לתשלום') === -1) continue;
    const here = amounts(lines[i]).concat(amounts(lines[i + 1] || '')).filter(a => a > 0);
    if (here.length) return Math.max.apply(null, here);
  }
  const all = amounts(text).filter(a => a > 0);
  return all.length ? Math.max.apply(null, all) : NaN;
}

// The PDF's text via Drive's OCR (a temporary Google Doc, deleted right after).
// Works with the Drive API service in v3 (the default) or v2. Drive's code for Hebrew is the old
// "iw" ("he" is rejected as "Invalid Value"); if that's refused too, OCR detects the language itself.
function payPdfText_(blob) {
  blob = blob.setContentType('application/pdf');   // Gmail may label the attachment octet-stream
  const v3 = typeof Drive.Files.create === 'function';
  const make = opts => v3
    ? Drive.Files.create({name: 'tmp-payment-ocr', mimeType: MimeType.GOOGLE_DOCS}, blob, opts).id
    : Drive.Files.insert({title: 'tmp-payment-ocr', mimeType: 'application/pdf'}, blob, Object.assign({ocr: true}, opts)).id;
  let id;
  try { id = make({ocrLanguage: 'iw'}); } catch (e) { id = make({}); }
  try {
    return DocumentApp.openById(id).getBody().getText();
  } finally {
    Drive.Files.remove(id);
  }
}

function paySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(PAY_SHEET) || ss.insertSheet(PAY_SHEET);
  if (String(sh.getRange(1, 1).getValue()).trim() !== PAY_HEADER[0]) {
    if (sh.getLastRow() > 0) sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, PAY_HEADER.length).setValues([PAY_HEADER]);
  }
  return sh;
}

// Emails read but not a bill (a Partner phone invoice, a receipt already stored) are remembered in a hidden
// tab, so they're read only once and the Payments tab keeps only real rows.
const PAY_SEEN_SHEET = 'Payments seen';
function paySeenSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PAY_SEEN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PAY_SEEN_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['Message id', 'Added']]);
    sh.getRange('A:A').setNumberFormat('@');
    sh.hideSheet();
  }
  return sh;
}
// Moves the empty "skipped" rows of earlier runs out of the Payments tab into the hidden tab, then sorts
// the Payments tab by date, newest first (rows pasted by hand, anywhere, land in their place).
function payTidy_(sh, seenSh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const tz = Session.getScriptTimeZone();
  const all = sh.getRange(2, 1, last - 1, PAY_HEADER.length).getValues();
  // Column A is a real date; Period, Invoice date, Due date, Card, Invoice no, Message id and Added are
  // text as the bills print them. A rewrite once turned "08/10/2026" into a date read the US way
  // (Aug 10): such a cell becomes text again, day and month swapped back ("08/10/2026").
  const asText = (v, isDate) => v instanceof Date
    ? (isDate ? Utilities.formatDate(v, tz, 'MM/dd/yyyy') : Utilities.formatDate(v, tz, 'yyyy-MM-dd'))
    : String(v);
  const fixed = all.map(r => r.map((v, c) => c === 0 ? v : c === 3 ? v : c === 5 || c === 6 ? asText(v, true) : c === 1 || c === 2 ? v : asText(v, false)));
  const keep = fixed.filter(r => String(r[1]).trim());
  const skipped = fixed.filter(r => !String(r[1]).trim() && String(r[PAY_COL_MSG - 1]).trim());
  if (skipped.length) seenSh.getRange(seenSh.getLastRow() + 1, 1, skipped.length, 2).setValues(skipped.map(r => [String(r[PAY_COL_MSG - 1]), r[PAY_HEADER.length - 1]]));
  sh.getRange(2, 1, last - 1, PAY_HEADER.length).clearContent();
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:K').setNumberFormat('@');
  if (keep.length) sh.getRange(2, 1, keep.length, PAY_HEADER.length).setValues(keep);
  if (keep.length > 1) sh.getRange(2, 1, keep.length, PAY_HEADER.length).sort({column: 1, ascending: false});
}

// New bills from every source into the Payments tab (up to PAY_MAX_PER_RUN PDFs a run). At the end the
// tab is tidied and sorted (payTidy_).
function updatePayments() {
  try { updateContracts(); } catch (e) { Logger.log('Contracts: ' + e); }
  try { updateReceipts(); } catch (e) { Logger.log('Receipts: ' + e); }
  const sh = paySheet_(), seenSh = paySeenSheet_();
  payTidy_(sh, seenSh);
  const last = sh.getLastRow();
  const seen = {};
  const seenNo = {};   // payee|invoice no. already in the tab (the same receipt from two senders)
  const noKey = (payee, no) => payee + '|' + String(no).replace(/^0+/, '');   // Sheets drops leading zeros
  if (last > 1) sh.getRange(2, 1, last - 1, PAY_COL_MSG).getValues().forEach(r => {
    seen[String(r[PAY_COL_MSG - 1])] = true;
    if (r[1] && r[8]) seenNo[noKey(r[1], r[8])] = true;
  });
  if (seenSh.getLastRow() > 1) seenSh.getRange(2, 1, seenSh.getLastRow() - 1, 1).getValues().forEach(r => { seen[String(r[0])] = true; });
  const tz = Session.getScriptTimeZone();
  let added = 0, skipped = 0, ocr = 0;
  PAY_SOURCES.forEach(src => {
    const threads = GmailApp.search(src.query, 0, 100);
    const msgs = [].concat.apply([], threads.map(t => t.getMessages()))
      .filter(m => !seen[m.getId()]).sort((a, b) => a.getDate() - b.getDate());
    msgs.forEach(m => {
      let p = null;
      if (src.body) {
        // The figures are in the email's own text (no PDF): quick, no per-run limit.
        p = src.parse(m.getPlainBody());
      } else {
        if (ocr >= PAY_MAX_PER_RUN) return;   // PDF conversions are slow: up to PAY_MAX_PER_RUN a run
        const pdf = m.getAttachments().filter(a => /pdf$/i.test(a.getContentType()) || /\.pdf$/i.test(a.getName()))[0];
        if (!pdf) return;
        ocr++;
        try { p = src.parse(payPdfText_(pdf.copyBlob())); } catch (e) { Logger.log('OCR failed for ' + m.getSubject() + ': ' + e); return; }
      }
      // Not this source's kind of bill (e.g. a Partner phone invoice), or a receipt already stored from
      // another sender: remember the email, store nothing.
      if (!p || !(p.amount > 0) || (p.invoiceNo && seenNo[noKey(src.payee, p.invoiceNo)])) {
        seenSh.appendRow([m.getId(), 'skipped ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')]);
        skipped++; return;
      }
      sh.appendRow([Utilities.formatDate(m.getDate(), tz, 'yyyy-MM-dd'), src.payee, src.category, p.amount,
                    p.period, p.invoiceDate, p.dueDate, p.card, p.invoiceNo, m.getId(), Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')]);
      if (p.invoiceNo) seenNo[noKey(src.payee, p.invoiceNo)] = true;
      added++;
    });
  });
  payTidy_(sh, seenSh);
  Logger.log('Payments: ' + added + ' added, ' + skipped + ' skipped (not a matching bill, or already in the tab).');
}

// ---- Contracts tab (warranties, service plans, insurance, subscriptions), read by SheetPulse's Payments
// window → Contracts. updateContracts() adds one row per email thread in the Gmail label
// payments-contracts (your own notes like "Bituah Haim - Starting 11/2024 100 NIS", forwarded documents
// and the providers' own emails), guessing what it can from the subject; then correct or complete the row
// by hand: the script never changes a row again. A thread is imported once (remembered in the script's
// properties), so deleting a row keeps it out. To hide a contract and keep its row, set its Type to Hidden.
// Columns: Name | Provider | Type | Reference | Start | End | Cost | Per (once / month / year) | Covers |
// Notes | Link (the Gmail thread) | Added.
const CONTRACTS_SHEET = 'Contracts';
const CONTRACTS_LABEL = 'payments-contracts';
const CONTRACTS_HEADER = ['Name', 'Provider', 'Type', 'Reference', 'Start', 'End', 'Cost', 'Per', 'Covers', 'Notes', 'Link', 'Added'];
const CONTRACT_TYPES = [
  [/bituah|ביטוח|insurance|polic|פוליס/i, 'Insurance'],
  [/guarantee|warrant|אחריות|הארכת/i, 'Warranty'],
  [/energy|electric|חשמל|power|partner/i, 'Utility plan'],
  [/psagot|פסגות|pension|פנסי|גמל|השתלמות/i, 'Investment'],
  [/\btv\b|internet|fiber|סיבים|טלוויזי|hot|cellular|סלולר/i, 'Subscription']
];
const CONTRACT_NAMES = [[/^bituah dira$/i, 'Home insurance'], [/^bituah haim$/i, 'Life insurance']];
const CONTRACT_PROVIDERS = [[/hot energy/i, 'HOT Energy'], [/\bhot\b/i, 'HOT'], [/partner/i, 'Partner'], [/psagot|פסגות/i, 'Psagot'],
  [/טופ פיקס|top ?fix/i, 'Top Fix Home']];

function contractsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONTRACTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONTRACTS_SHEET);
    sh.getRange(1, 1, 1, CONTRACTS_HEADER.length).setValues([CONTRACTS_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('E:F').setNumberFormat('yyyy-mm-dd');
  }
  return sh;
}
function setupContractsTab() { contractsSheet_(); updateContracts(); }
// Deletes every row of the Contracts tab (hand edits too) and imports the label again.
function rebuildContracts() {
  const sh = contractsSheet_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  PropertiesService.getScriptProperties().deleteProperty('contractsSeen');
  updateContracts();
}

// "01/12/2024" → 1 Dec 2024; "11/2024" → 1 Nov 2024 (day/month, Israeli order).
function contractDate_(s) {
  let m = String(s).match(/(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})/);
  if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  m = String(s).match(/(\d{1,2})[\/.](\d{4})/);
  return m ? new Date(+m[2], +m[1] - 1, 1) : '';
}

function contractFromEmail_(t, me) {
  const msgs = t.getMessages(), m = msgs[0];
  const subject = m.getSubject().replace(/^\s*(fwd?|fw|הועבר)\s*:\s*/i, '').trim();
  const body = m.getPlainBody() || '';
  const fromAddr = m.getFrom().replace(/.*<([^>]+)>.*/, '$1').toLowerCase();
  const mine = fromAddr === me;
  // A forwarded email: the original sender's name ("From: טופ פיקס לבית <...>").
  const fwd = (body.match(/(?:From|מאת)\s*:\s*\*?\s*([^<\n*]*?)\s*<([^>]+)>/i) || []);
  const fromName = mine ? (fwd[1] || (fwd[2] || '').replace(/^.*@|\..*$/g, '')) : m.getFrom().replace(/\s*<.*$/, '').replace(/"/g, '');
  const hay = subject + ' ' + fromName + ' ' + fromAddr + ' ' + (fwd[2] || '');
  // Your own subjects: "Name - details" or "Name---details".
  const parts = mine ? subject.split(/\s*(?:-{2,}|\s-\s)\s*/) : [subject];
  let name = mine ? parts[0] : '';
  const own = mine ? body.split(/-{5,}\s*(?:Forwarded message|הודעה שהועברה)|\n\s*(?:From|מאת)\s*:/i)[0].replace(/\s+/g, ' ').trim().slice(0, 300) : '';
  const rest = mine ? [parts.slice(1).join(' - '), own].filter(String).join(' ') : subject;
  CONTRACT_NAMES.forEach(([re, n]) => { if (re.test(name)) name = n; });
  const prov = (CONTRACT_PROVIDERS.find(([re]) => re.test(hay)) || [])[1] || fromName || '';
  if (!name) name = prov || subject.slice(0, 60);
  const type = (CONTRACT_TYPES.find(([re]) => re.test(hay)) || [])[1] || 'Other';
  const st = rest.match(/start(?:ing|s)?\s+([\d\/.]+)/i), till = rest.match(/(?:till|until|ends?|עד)\s+([\d\/.]+)/i);
  // A forwarded document starts on its original date ("Date: Thu, Mar 7, 2024 at 10:12 AM").
  const fd = mine && (body.match(/(?:^|\n)\s*\*?(?:Date|תאריך)\s*:\s*\*?\s*([^\n]+)/i) || [])[1];
  const fdMs = fd ? new Date(fd.replace(/\s+at\s+/, ' ').replace(/\s*[AP]M\s*$/i, '')) : null;
  let start = st ? contractDate_(st[1]) : mine ? (fdMs && !isNaN(fdMs) ? fdMs : '') : m.getDate();
  let end = till ? contractDate_(till[1]) : '';
  const yrs = rest.match(/(\d+)\s*years?/i);
  if (!end && start && yrs) { end = new Date(start); end.setFullYear(end.getFullYear() + +yrs[1]); }
  const costs = [...rest.matchAll(/(\d+(?:\.\d+)?)\s*(?:NIS|ILS|₪|ש"?ח)/gi)].map(x => +x[1]);
  const cost = costs.length ? costs[costs.length - 1] : '';
  const covers = mine ? rest.replace(/start(?:ing|s)?\s+[\d\/.]+|(?:till|until)\s+[\d\/.]+|\d+(?:\.\d+)?\s*(?:NIS|ILS|₪)/gi, '').replace(/\s{2,}/g, ' ').trim() : '';
  return {name, prov, type, start, end, cost, per: cost !== '' ? 'month' : '', covers,
    notes: (mine ? 'Your note: ' + [subject, own].filter(String).join(' · ') : 'Email: ' + subject) + (msgs.length > 1 ? ` (${msgs.length} emails)` : ''),
    link: 'https://mail.google.com/mail/u/0/#all/' + t.getId(), date: m.getDate(), mine, body,
    refNo: (subject.match(/\d{5,}/g) || [])};
}

function updateContracts() {
  const sh = contractsSheet_();
  const props = PropertiesService.getScriptProperties();
  const seen = new Set(JSON.parse(props.getProperty('contractsSeen') || '[]'));
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const threads = GmailApp.search('label:' + CONTRACTS_LABEL, 0, 200).filter(t => !seen.has(t.getId()));
  let found = threads.map(t => contractFromEmail_(t, me)).sort((a, b) => a.date - b.date);
  const myText = found.filter(c => c.mine).map(c => c.body).join(' ');
  found = found.filter(c => c.mine || !c.refNo.some(n => myText.includes(n)));
  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, CONTRACTS_HEADER.length).getValues() : [];
  let added = 0, merged = 0;
  found.forEach(c => {
    // "Bituah Dira - Till 30/11/2026": an end date for a contract already in the tab goes into that row.
    const same = data.findIndex(r => String(r[0]).toLowerCase() === c.name.toLowerCase());
    if (same >= 0 && c.end && !c.cost && !data[same][5]) {
      sh.getRange(same + 2, 6).setValue(c.end);
      sh.getRange(same + 2, 10).setValue([data[same][9], c.notes].filter(String).join(' · '));
      data[same][5] = c.end;
      merged++;
    } else {
      const row = [c.name, c.prov, c.type, '', c.start, c.end, c.cost, c.per, c.covers, c.notes, c.link, today];
      sh.appendRow(row);
      data.push(row);
      added++;
    }
    Logger.log(`${Utilities.formatDate(c.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')} | ${c.name} | ${c.prov} | ${c.type} | start ${c.start && Utilities.formatDate(c.start, Session.getScriptTimeZone(), 'yyyy-MM-dd')} | end ${c.end && Utilities.formatDate(c.end, Session.getScriptTimeZone(), 'yyyy-MM-dd')} | ${c.cost} ${c.per} | ${c.covers}`);
  });
  threads.forEach(t => seen.add(t.getId()));
  props.setProperty('contractsSeen', JSON.stringify([...seen]));
  if (sh.getLastRow() > 1) sh.getRange(2, 5, sh.getLastRow() - 1, 2).setNumberFormat('yyyy-mm-dd');
  Logger.log(`Contracts: ${added} added, ${merged} end dates filled in, ${threads.length} new threads in ${CONTRACTS_LABEL}.`);
}

// Nothing written: each thread in the label with its sender, attachments and the first part of the
// email's text, to see what a contract email holds (send this log to Claude to refine the rows).
function listContracts() {
  GmailApp.search('label:' + CONTRACTS_LABEL, 0, 50).forEach(t => {
    const m = t.getMessages()[0];
    Logger.log('=== ' + Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' | ' + m.getFrom() + ' | ' + m.getSubject()
      + ' | ' + t.getMessageCount() + ' emails | attachments: ' + m.getAttachments().map(a => a.getName()).join(', '));
    Logger.log((m.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 700));
  });
}

// ---- Receipts tab: every email in the Gmail label payments-receipts (purchase receipts, invoices,
// guarantees of things you bought), read by SheetPulse's Payments window → Receipts. One row per email:
// Date | From | Subject | Amount | Currency | Attachments | Link (the Gmail thread) | Message id | Notes |
// Product (what it paid for) | Kind (Product, Service or Don't count): the last two typed by you, here or
// in SheetPulse's Receipts tab (✎).
// The amount is a guess from the subject, then the email's text (a total, ₪ / NIS / $); correct it by hand,
// the script never changes a row again. Runs with updatePayments (daily), or run updateReceipts() any time.
const RECEIPTS_SHEET = 'Receipts';
const RECEIPTS_LABEL = 'payments-receipts';
const RECEIPTS_HEADER = ['Date', 'From', 'Subject', 'Amount', 'Currency', 'Attachments', 'Link', 'Message id', 'Notes', 'Product', 'Kind'];

function receiptAmount_(text) {
  const t = String(text);
  const pats = [
    [/(?:סה"?כ(?: לתשלום)?|total|amount paid|סכום)\D{0,20}?(₪|NIS|ILS|\$|USD|€|EUR)?\s*([\d,]+(?:\.\d{1,2})?)\s*(₪|NIS|ILS|ש"?ח|\$|USD|€|EUR)?/i],
    [/(₪|\$|€)\s*([\d,]+(?:\.\d{1,2})?)/],
    [/()([\d,]+(?:\.\d{1,2})?)\s*(₪|NIS|ILS|ש"?ח|USD|EUR)/i]
  ];
  for (const [re] of pats) {
    const m = t.match(re);
    if (!m) continue;
    const v = +String(m[2]).replace(/,/g, '');
    if (!(v > 0)) continue;
    const c = (m[1] || m[3] || '₪').toUpperCase();
    return {amount: v, cur: /\$|USD/.test(c) ? 'USD' : /€|EUR/.test(c) ? 'EUR' : 'ILS'};
  }
  return {amount: '', cur: ''};
}

// A receipt PDF's total. OCR of these Hebrew documents separates labels from their numbers, so instead of
// "total → number": the largest amount written with agorot (1,234.00; not a percentage) that appears at
// least twice, a document's total being printed more than once; else the largest one. Tested 12 Oct 2026
// on Home Center, Talcar (invoice and receipt) and Comp4U PDFs.
function receiptPdfAmount_(text) {
  const t = String(text).replace(/\s+/g, ' ');
  const counts = {};
  for (const m of t.matchAll(/(?<![\d.,])(\d{1,3}(?:,\d{3})+|\d+)\.(\d{2})(?![\d%])(?!\s*%)/g)) {
    const v = +(m[1].replace(/,/g, '') + '.' + m[2]);
    if (v > 0) counts[v] = (counts[v] || 0) + 1;
  }
  const vals = Object.keys(counts).map(Number);
  const twice = vals.filter(v => counts[v] >= 2);
  const v = Math.max(...(twice.length ? twice : vals), 0);
  return v > 0 ? {amount: v, cur: /\$|USD/.test(t) && !/₪|ש"?ח|ILS/.test(t) ? 'USD' : 'ILS'} : {amount: '', cur: ''};
}

function updateReceipts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RECEIPTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RECEIPTS_SHEET);
    sh.getRange(1, 1, 1, RECEIPTS_HEADER.length).setValues([RECEIPTS_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sh.getRange('H:H').setNumberFormat('@');
  }
  // Tabs made before Product / Kind existed get the two header cells.
  if (sh.getLastColumn() < RECEIPTS_HEADER.length) sh.getRange(1, 1, 1, RECEIPTS_HEADER.length).setValues([RECEIPTS_HEADER]).setFontWeight('bold');
  const ids = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 8, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const rows = [];
  let start = 0, batch;
  do {
    batch = GmailApp.search('label:' + RECEIPTS_LABEL, start, 100);
    batch.forEach(t => t.getMessages().forEach(m => {
      if (ids.has(m.getId())) return;
      const subject = m.getSubject().replace(/^\s*(fwd?|fw|הועבר)\s*:\s*/i, '').trim();
      const a = receiptAmount_(subject).amount ? receiptAmount_(subject) : receiptAmount_((m.getPlainBody() || '').slice(0, 6000));
      rows.push([m.getDate(), m.getFrom().replace(/\s*<.*$/, '').replace(/"/g, '').trim() || m.getFrom().replace(/[<>]/g, ''), subject, a.amount, a.cur,
        m.getAttachments().map(x => x.getName()).join(', '), 'https://mail.google.com/mail/u/0/#all/' + t.getId(), m.getId(), '', '', '']);
    }));
    start += batch.length;
  } while (batch.length === 100 && start < 1000);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, RECEIPTS_HEADER.length).setValues(rows);
  const props = PropertiesService.getScriptProperties();
  const tried = new Set(JSON.parse(props.getProperty('receiptsPdfTried') || '[]'));
  let ocr = 0, filled = 0;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, RECEIPTS_HEADER.length).getValues().forEach((r, i) => {
    const id = String(r[7]);
    if (r[3] !== '' || !id || tried.has(id) || ocr >= 15 || !/pdf/i.test(String(r[5]))) return;
    tried.add(id); ocr++;
    try {
      const m = GmailApp.getMessageById(id);
      const pdf = m.getAttachments().filter(x => /pdf$/i.test(x.getContentType()) || /\.pdf$/i.test(x.getName()))[0];
      if (!pdf) return;
      const a = receiptPdfAmount_(payPdfText_(pdf.copyBlob()));
      if (a.amount) { sh.getRange(i + 2, 4, 1, 2).setValues([[a.amount, a.cur]]); filled++; }
    } catch (e) { Logger.log('Receipt PDF ' + id + ': ' + e); }
  });
  props.setProperty('receiptsPdfTried', JSON.stringify([...tried]));
  if (ocr) Logger.log(`Receipts: ${ocr} PDFs read, ${filled} amounts found.`);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, RECEIPTS_HEADER.length).sort({column: 1, ascending: false});
  Logger.log(`Receipts: ${rows.length} added from ${RECEIPTS_LABEL}.`);
}

// Once: rows whose amount is a VAT rate (17 or 18, written by the first, label-based rule) get their PDF
// read again with the current rule. Logs each change.
function recheckReceiptAmounts() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIPTS_SHEET);
  sh.getRange(2, 1, sh.getLastRow() - 1, RECEIPTS_HEADER.length).getValues().forEach((r, i) => {
    if (![17, 18].includes(Number(r[3])) || !/pdf/i.test(String(r[5]))) return;
    try {
      const m = GmailApp.getMessageById(String(r[7]));
      const pdf = m.getAttachments().filter(x => /pdf$/i.test(x.getContentType()) || /\.pdf$/i.test(x.getName()))[0];
      const a = pdf ? receiptPdfAmount_(payPdfText_(pdf.copyBlob())) : {amount: ''};
      sh.getRange(i + 2, 4, 1, 2).setValues([[a.amount, a.amount ? a.cur : r[4]]]);
      Logger.log(`${r[2]}: ${r[3]} → ${a.amount || 'empty'}`);
    } catch (e) { Logger.log(r[2] + ': ' + e); }
  });
}

// Nothing written: for each receipt with a PDF and no amount (up to 4), the PDF's text as read by OCR
// and the amount found in it, to tune receiptPdfAmount_(). Also clears the "tried" list, so the next
// updateReceipts() reads those PDFs again.
function testReceiptPdfs() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIPTS_SHEET);
  let n = 0;
  sh.getRange(2, 1, sh.getLastRow() - 1, RECEIPTS_HEADER.length).getValues().forEach(r => {
    if (n >= 4 || r[3] !== '' || !/pdf/i.test(String(r[5]))) return;
    n++;
    try {
      const m = GmailApp.getMessageById(String(r[7]));
      const pdf = m.getAttachments().filter(x => /pdf$/i.test(x.getContentType()) || /\.pdf$/i.test(x.getName()))[0];
      const text = pdf ? payPdfText_(pdf.copyBlob()) : '';
      Logger.log('=== ' + r[1] + ' | ' + r[2] + ' | PDF: ' + (pdf ? pdf.getName() : 'none') + ' | found: ' + JSON.stringify(receiptPdfAmount_(text)));
      Logger.log(text.replace(/\s+/g, ' ').slice(0, 1800));
    } catch (e) { Logger.log('=== ' + r[2] + ': ' + e); }
  });
  PropertiesService.getScriptProperties().deleteProperty('receiptsPdfTried');
}

// ---- Routine bills tab: charges entered by hand that repeat on a schedule (insurance premiums, standing
// orders) rather than arriving by email. SheetPulse's Bills tab makes one charge per due month from each row
// (a month that already has a Gmail bill from the same payee keeps that bill). Rows are typed or pasted from
// SheetPulse ("+ Add a routine bill"); this script never changes them. Run setupRoutineBillsTab() once.
// Columns: Payee | Category | Amount | Every (month, 2 months, 3 months, year) | Day (of the month) |
// Start (yyyy-mm) | End (yyyy-mm, empty = still running) | Card (last 4) | Notes | Added.
const ROUTINE_SHEET = 'Routine bills';
const ROUTINE_HEADER = ['Payee', 'Category', 'Amount', 'Every', 'Day', 'Start', 'End', 'Card', 'Notes', 'Added'];
function setupRoutineBillsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ROUTINE_SHEET)) { Logger.log('The Routine bills tab already exists; nothing changed.'); return; }
  const sh = ss.insertSheet(ROUTINE_SHEET);
  sh.getRange(1, 1, 1, ROUTINE_HEADER.length).setValues([ROUTINE_HEADER]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('D:D').setNumberFormat('@');
  sh.getRange('F:H').setNumberFormat('@');   // Start, End as text (yyyy-mm), Card keeps leading zeros
  Logger.log('Routine bills tab created. Add one row per routine charge, or paste rows from SheetPulse.');
}

// Daily at 07:00 (Israel time of the script). Safe to run again: replaces only this script's trigger.
function installPaymentsTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'updatePayments') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('updatePayments').timeBased().everyDays(1).atHour(7).create();
  Logger.log('updatePayments now runs daily at about 07:00.');
}

// Test, nothing written: the latest email of the first source, its PDF text and the parsed figures.
function testPaymentsSource() {
  const src = PAY_SOURCES[0];
  const t = GmailApp.search(src.query, 0, 1)[0];
  if (!t) { Logger.log('No email found for: ' + src.query); return; }
  const m = t.getMessages().pop();
  const pdf = m.getAttachments().filter(a => /pdf$/i.test(a.getContentType()) || /\.pdf$/i.test(a.getName()))[0];
  Logger.log('Email: ' + m.getDate() + ' | ' + m.getSubject() + ' | PDF: ' + (pdf ? pdf.getName() : 'none'));
  if (!pdf) return;
  const text = payPdfText_(pdf.copyBlob());
  Logger.log('--- PDF text (first 1500 characters) ---\n' + text.slice(0, 1500));
  Logger.log('--- Parsed ---\n' + JSON.stringify(src.parse(text)));
}

// ---- Helpers for adding new sources ----

// Every sender in the label: number of emails, whether they carry a PDF, and the latest subject.
// Nothing written. Use it to choose the next source.
function listPaymentSenders() {
  const by = {};
  let start = 0, batch;
  do {
    batch = GmailApp.search('label:payments-bills', start, 100);
    batch.forEach(t => t.getMessages().forEach(m => {
      const from = m.getFrom().replace(/.*<([^>]+)>.*/, '$1').toLowerCase();
      const e = by[from] || (by[from] = {n: 0, pdf: 0, last: null, subject: ''});
      e.n++;
      if (m.getAttachments().some(a => /\.pdf$/i.test(a.getName()) || /pdf/i.test(a.getContentType()))) e.pdf++;
      if (!e.last || m.getDate() > e.last) { e.last = m.getDate(); e.subject = m.getSubject(); }
    }));
    start += batch.length;
  } while (batch.length === 100 && start < 1000);
  Object.keys(by).sort((a, b) => by[b].n - by[a].n).forEach(k => {
    const e = by[k];
    Logger.log(`${k} | ${e.n} emails, ${e.pdf} with PDF | latest ${Utilities.formatDate(e.last, Session.getScriptTimeZone(), 'yyyy-MM-dd')} | ${e.subject}`);
  });
}

// Where could arnona and water get a billing period from? The payment receipts don't state one
// (Ramat Gan: "ארנונה חיוב תקופתי" and a voucher number; MAST: amount, approval and stub number).
// This logs (nothing written): the latest two PDFs from the older Ramat Gan sender citypay@ladpc.co.il,
// and every sender in the whole mailbox (not only the label) whose emails mention the water company
// or an arnona bill, so the actual bills, which state their period, can be found.
function findPeriodSources() {
  GmailApp.search('label:payments-bills from:citypay@ladpc.co.il has:attachment', 0, 2).forEach(t => {
    const m = t.getMessages().pop();
    const pdf = m.getAttachments().filter(a => /pdf$/i.test(a.getContentType()) || /\.pdf$/i.test(a.getName()))[0];
    Logger.log('=== ladpc: ' + m.getDate() + ' | ' + m.getSubject() + ' | PDF: ' + (pdf ? pdf.getName() : 'none'));
    if (pdf) Logger.log(payPdfText_(pdf.copyBlob()).slice(0, 2500));
  });
  [['water', '"מי רמת גן" OR "מי-רמת-גן" OR "תאגיד המים"'], ['arnona', '(ארנונה OR "שובר תשלום" OR "הודעת תשלום") ("רמת גן" OR "רמת-גן")']].forEach(([what, q]) => {
    const by = {};
    GmailApp.search(q, 0, 200).forEach(t => t.getMessages().forEach(m => {
      const from = m.getFrom().replace(/.*<([^>]+)>.*/, '$1').toLowerCase();
      const e = by[from] || (by[from] = {n: 0, pdf: 0, last: null, subject: ''});
      e.n++;
      if (m.getAttachments().some(a => /\.pdf$/i.test(a.getName()) || /pdf/i.test(a.getContentType()))) e.pdf++;
      if (!e.last || m.getDate() > e.last) { e.last = m.getDate(); e.subject = m.getSubject(); }
    }));
    Logger.log('=== ' + what + ' senders, whole mailbox ===');
    Object.keys(by).sort((a, b) => by[b].n - by[a].n).forEach(k => {
      const e = by[k];
      Logger.log(`${k} | ${e.n} emails, ${e.pdf} with PDF | latest ${Utilities.formatDate(e.last, Session.getScriptTimeZone(), 'yyyy-MM-dd')} | ${e.subject}`);
    });
  });
}

// The latest email for a search: its PDF's text, or the email's own text when there's no PDF.
// Change TEST_QUERY to the sender to look at, then run. Nothing written.
const TEST_QUERY = 'label:payments-bills from:next';
function testPaymentsQuery() {
  const t = GmailApp.search(TEST_QUERY, 0, 1)[0];
  if (!t) { Logger.log('No email found for: ' + TEST_QUERY); return; }
  const m = t.getMessages().pop();
  const pdf = m.getAttachments().filter(a => /pdf$/i.test(a.getContentType()) || /\.pdf$/i.test(a.getName()))[0];
  Logger.log('Email: ' + m.getDate() + ' | ' + m.getFrom() + ' | ' + m.getSubject() + ' | PDF: ' + (pdf ? pdf.getName() : 'none'));
  const text = pdf ? payPdfText_(pdf.copyBlob()) : m.getPlainBody();
  Logger.log('--- ' + (pdf ? 'PDF' : 'email') + ' text (first 2000 characters) ---\n' + text.slice(0, 2000));
}
