# Automatic overdue-invoice reminders in Google Sheets

A Google Apps Script automation for small businesses that track invoices in a spreadsheet. Every morning it finds unpaid invoices that are past due and emails each client a polite reminder. It switches to a firmer tone after 14 days, never sends two reminders too close together, and logs every email.

No extra software or paid tools: it runs inside Google Sheets with the business's own Gmail.

## What it does

- **Finds overdue invoices**: unpaid, with a due date in the past. Paid and cancelled invoices are ignored.
- **Picks the right tone**: a friendly reminder first, then a "second notice" once an invoice is 14+ days overdue.
- **Avoids spamming clients**: at most one reminder every 7 days per invoice, and at most 3 in total. Both limits are configurable.
- **Logs everything**: a *Log* sheet records each reminder, plus skipped rows with the reason, such as an invalid email address.
- **Preview mode (on by default)**: writes what *would* be sent to the Log without emailing anyone, so the owner can check before going live.
- **Safe to run anytime**: a lock prevents double runs, it respects Gmail's daily sending quota, and it updates "Last reminder" and "Reminders sent" in the sheet.
- **Easy controls**: an *Invoice reminders* menu in the spreadsheet: send now, turn the daily 9:00 run on or off, set up sheets.

## Example run (sample data)

| Invoice | Client | Due | Status | Result |
|---|---|---|---|---|
| INV-1001 | Baltic Bakery | 15 Oct | Paid | ignored |
| INV-1002 | Riga Auto Service | 28 Oct | Unpaid | **Second notice** (20 days overdue) |
| INV-1003 | Green Cafe | 10 Nov | Unpaid | **Friendly reminder** |
| INV-1004 | North Studio | 1 Dec | Unpaid | not due yet |
| INV-1005 | Old Town Florist | 1 Nov | Unpaid | skipped: invalid email (logged) |
| INV-1006 | Sea Breeze Hotel | 3 Nov | Unpaid | waits: reminded 5 days ago |

Running it again the same day sends nothing.

## Setup (about 5 minutes)

1. Create a Google Sheet, rename the first tab to **Invoices**, and import `sample-invoices.csv` (File → Import → *Replace current sheet*). Make sure the *Due date* column is formatted as dates.
2. Open **Extensions → Apps Script**, replace the code with `Code.gs`, and save.
3. Optional: in Project Settings, tick *Show "appsscript.json"* and paste in `appsscript.json` so the script only asks for the permissions it needs.
4. Reload the spreadsheet. An **Invoice reminders** menu appears. Click **Set up sheets** and approve the permissions.
5. Click **Send reminders now**. Preview mode is on, so it only fills the *Log* sheet.
6. When it looks right, set *Preview only* to `FALSE` in **Settings**, then click **Turn on daily run (9:00)**.

## Settings

| Setting | Default |
|---|---|
| Your name (email signature and sender name) | Your Company |
| Days between reminders | 7 |
| Firm tone after days overdue | 14 |
| Max reminders per invoice | 3 |
| Currency | EUR |
| Preview only | TRUE |

## Files

```
Code.gs               the automation
appsscript.json       manifest with minimal permissions, Riga time zone
sample-invoices.csv   demo data to try it with
```

---

*Need something like this for your own spreadsheet, or have an Apps Script that stopped working? Send me the sheet structure and what should happen, and I'll quote a fixed price.*
