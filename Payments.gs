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

// New bills from every source into the Payments tab (oldest first, up to PAY_MAX_PER_RUN a run).
function updatePayments() {
  const sh = paySheet_();
  const last = sh.getLastRow();
  const seen = {};
  if (last > 1) sh.getRange(2, PAY_COL_MSG, last - 1, 1).getValues().forEach(r => { seen[String(r[0])] = true; });
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
      // Not this source's kind of bill (e.g. a Partner phone invoice): remember it, store nothing.
      if (!p || !(p.amount > 0)) {
        sh.appendRow(['', '', '', '', '', '', '', '', '', m.getId(), 'skipped ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')]);
        skipped++; return;
      }
      sh.appendRow([Utilities.formatDate(m.getDate(), tz, 'yyyy-MM-dd'), src.payee, src.category, p.amount,
                    p.period, p.invoiceDate, p.dueDate, p.card, p.invoiceNo, m.getId(), Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')]);
      added++;
    });
  });
  Logger.log('Payments: ' + added + ' added, ' + skipped + ' skipped (not a matching bill).');
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
