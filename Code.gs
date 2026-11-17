/**
 * Overdue Invoice Reminders — Google Sheets + Apps Script
 *
 * Reads the "Invoices" sheet, finds unpaid invoices that are past due, and emails
 * each client a polite reminder (firmer after 14 days). Every send is logged, and
 * a client is never reminded about the same invoice more than once every N days.
 *
 * Sheets used:
 *   Invoices  — Invoice #, Client, Email, Amount, Due date, Status, Last reminder, Reminders sent
 *   Settings  — key/value pairs (see DEFAULT_SETTINGS)
 *   Log       — created automatically; one row per reminder (or preview)
 */

const SHEET_INVOICES = 'Invoices';
const SHEET_SETTINGS = 'Settings';
const SHEET_LOG = 'Log';

const DEFAULT_SETTINGS = {
  'Your name': 'Your Company',
  'Days between reminders': 7,
  'Firm tone after days overdue': 14,
  'Max reminders per invoice': 3,
  'Currency': 'EUR',
  'Preview only (TRUE = log, do not send)': true,
};

// Column positions in the Invoices sheet (1-based)
const COL = {
  invoice: 1, client: 2, email: 3, amount: 4, due: 5,
  status: 6, lastReminder: 7, count: 8,
};

/** Adds the "Invoice reminders" menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Invoice reminders')
    .addItem('Send reminders now', 'sendReminders')
    .addItem('Turn on daily run (9:00)', 'installDailyTrigger')
    .addItem('Turn off daily run', 'removeDailyTrigger')
    .addSeparator()
    .addItem('Set up sheets', 'setup')
    .addToUi();
}

/** Creates the Settings and Log sheets if they don't exist yet. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(SHEET_SETTINGS);
    const rows = Object.entries(DEFAULT_SETTINGS);
    settings.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
    settings.getRange(2, 1, rows.length, 2).setValues(rows);
    settings.autoResizeColumns(1, 2);
  }
  getLogSheet_();
  SpreadsheetApp.getActive().toast('Setup complete. Check the Settings sheet.');
}

/** Main job: find overdue invoices and remind clients. */
function sendReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // another run is in progress

  try {
    const settings = readSettings_();
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_INVOICES);
    if (!sheet) throw new Error('Sheet "' + SHEET_INVOICES + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, COL.count).getValues();

    const today = startOfDay_(new Date());
    const preview = settings['Preview only (TRUE = log, do not send)'] === true;
    let sent = 0, skipped = 0;

    data.forEach(function (row, i) {
      const rowNum = i + 2;
      const inv = {
        number: String(row[COL.invoice - 1]).trim(),
        client: String(row[COL.client - 1]).trim(),
        email: String(row[COL.email - 1]).trim(),
        amount: Number(row[COL.amount - 1]),
        due: row[COL.due - 1],
        status: String(row[COL.status - 1]).trim().toLowerCase(),
        last: row[COL.lastReminder - 1],
        count: Number(row[COL.count - 1]) || 0,
      };

      if (!inv.number || inv.status === 'paid' || inv.status === 'cancelled') return;
      if (!(inv.due instanceof Date)) { log_(inv, 'SKIPPED', 'Due date is not a valid date'); skipped++; return; }
      if (!isValidEmail_(inv.email)) { log_(inv, 'SKIPPED', 'Invalid email address'); skipped++; return; }

      const daysOverdue = Math.floor((today - startOfDay_(inv.due)) / 86400000);
      if (daysOverdue <= 0) return;
      if (inv.count >= settings['Max reminders per invoice']) return;
      if (inv.last instanceof Date &&
          (today - startOfDay_(inv.last)) / 86400000 < settings['Days between reminders']) return;

      if (MailApp.getRemainingDailyQuota() < 1) {
        log_(inv, 'STOPPED', 'Daily email quota reached — will continue tomorrow');
        return;
      }

      const firm = daysOverdue >= settings['Firm tone after days overdue'];
      const msg = buildEmail_(inv, daysOverdue, firm, settings);

      if (preview) {
        log_(inv, 'PREVIEW', msg.subject);
      } else {
        MailApp.sendEmail({ to: inv.email, subject: msg.subject, body: msg.body, name: settings['Your name'] });
        sheet.getRange(rowNum, COL.lastReminder).setValue(new Date());
        sheet.getRange(rowNum, COL.count).setValue(inv.count + 1);
        log_(inv, 'SENT', msg.subject);
      }
      sent++;
    });

    SpreadsheetApp.getActive().toast(
      (preview ? 'Preview: ' : '') + sent + ' reminder(s), ' + skipped + ' skipped. See the Log sheet.');
  } finally {
    lock.releaseLock();
  }
}

function buildEmail_(inv, daysOverdue, firm, settings) {
  const amount = formatMoney_(inv.amount, settings['Currency']);
  const due = Utilities.formatDate(inv.due, Session.getScriptTimeZone(), 'd MMM yyyy');
  const greeting = 'Hi ' + (inv.client || 'there') + ',\n\n';
  const sign = '\n\nBest regards,\n' + settings['Your name'];

  if (firm) {
    return {
      subject: 'Second notice: invoice ' + inv.number + ' is ' + daysOverdue + ' days overdue',
      body: greeting +
        'Invoice ' + inv.number + ' for ' + amount + ' was due on ' + due +
        ' and is now ' + daysOverdue + ' days overdue.\n\n' +
        'Please arrange payment within the next 7 days, or reply to let me know if there is a problem with the invoice.' +
        sign,
    };
  }
  return {
    subject: 'Friendly reminder: invoice ' + inv.number + ' (' + amount + ')',
    body: greeting +
      'Just a quick reminder that invoice ' + inv.number + ' for ' + amount +
      ' was due on ' + due + '. If you have already paid, thank you, and please ignore this message.' +
      sign,
  };
}

/** Daily trigger management (avoids creating duplicates). */
function installDailyTrigger() {
  removeDailyTrigger(true);
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getActive().toast('Daily run turned on (around 9:00).');
}

function removeDailyTrigger(silent) {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'sendReminders'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  if (silent !== true) SpreadsheetApp.getActive().toast('Daily run turned off.');
}

// ---------- helpers ----------

function readSettings_() {
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return settings;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] !== '') settings[String(r[0]).trim()] = r[1];
  });
  return settings;
}

function getLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.appendRow(['Time', 'Invoice #', 'Client', 'Email', 'Result', 'Details']);
    log.getRange(1, 1, 1, 6).setFontWeight('bold');
    log.setFrozenRows(1);
  }
  return log;
}

function log_(inv, result, details) {
  getLogSheet_().appendRow([new Date(), inv.number, inv.client, inv.email, result, details]);
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatMoney_(amount, currency) {
  const symbols = { EUR: '€', USD: '$', GBP: '£' };
  return (symbols[currency] || currency + ' ') + Number(amount).toFixed(2);
}
