# Google Sheets App Script
A fully automated, serverless financial tracking system built with Google Apps Script. This project automates weekly expense archiving, dynamic ledger calculations, PDF report generation, and automated email notifications directly from a Google Workspace environment.

## Features

* **Automated Data Archiving:** Moves active transaction rows from daily tracking sheets to historical archives dynamically using `getLastRow()` limits.
* **Running Ledger Calculation:** Maintains a continuous balance for household expenses (capable of handling negative credit/advance payment states).
* **Automated Reimbursement Processing:** Scans for reimbursement inputs daily, adjusts the running ledger, clears the input cell, and dispatches a confirmation email.
* **Dynamic PDF Generation:** Compiles raw weekly household expenditure data into an HTML table, converts it to a PDF blob, and attaches it to the weekly summary email.
* **Cron-Triggered Reminders:** Dispatches daily email reminders to ensure consistent data logging.

## Sheet Architecture Structure

For the script to execute successfully, the bound Google Spreadsheet must contain exactly these five sheets with the following naming conventions:
1. `personal expenses` (Columns A-C: Expenses, Amount Spent, Date)
2. `house expenses` (Columns A-C: Expenses, Amount Spent, Date | Cell G2: Reimbursement Input)
3. `previous personal expenses` (Archive)
4. `previous household expenses` (Archive)
5. `Financial tracker` (Dashboard)
   - `A4`: Personal Weekly Total
   - `B4`: Amount Saved Formula (e.g., `=F4-A4`)
   - `F4`: Personal Target Spend
   - `A8`: Household Current Expenditure (Running Ledger)

## Setup & Deployment

1. Open your Google Sheet, navigate to **Extensions > Apps Script**.
2. Paste the contents of `Code.gs` into the editor and save.
3. **Execute manually** once (e.g., run `runEndOfWeekAutomation`) to grant Google OAuth permissions for `SpreadsheetApp` and `MailApp`.
4. Navigate to the **Triggers** menu (clock icon) and configure the following time-driven events:
   * `runEndOfWeekAutomation`: Time-driven -> Week timer -> Every Saturday -> 9pm to 10pm.
   * `processDailyReimbursement`: Time-driven -> Day timer -> 9pm to 10pm.
   * `sendDailyReminder`: Time-driven -> Day timer -> 7pm to 8pm.

## Author
**IaMdEvJaXoN**
