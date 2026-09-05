// 1. END-OF-WEEK AUTOMATION (Runs Every Saturday Night at 9:00 PM)
function runEndOfWeekAutomation() {
  var recipientEmail = Session.getActiveUser().getEmail();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Sheet References
    var personalSheet = ss.getSheetByName("personal expenses");
    var houseSheet = ss.getSheetByName("house expenses");
    var prevPersonalSheet = ss.getSheetByName("previous personal expenses");
    var prevHouseSheet = ss.getSheetByName("previous household expenses");
    var trackerSheet = ss.getSheetByName("Financial tracker");

    // Throw custom error if a sheet is missing/renamed
    if (!personalSheet || !houseSheet || !prevPersonalSheet || !prevHouseSheet || !trackerSheet) {
      throw new Error("One or more required sheets are missing or renamed. Check sheet names.");
    }

    // A. PROCESS PERSONAL EXPENSES
    var lastPersonalRow = personalSheet.getLastRow();
    var totalPersonalSpent = 0;

    if (lastPersonalRow > 2) { 
      var personalData = personalSheet.getRange(3, 1, lastPersonalRow - 2, 3).getValues();
      
      for (var i = 0; i < personalData.length; i++) {
        totalPersonalSpent += Number(personalData[i][1]) || 0;
      }

      prevPersonalSheet.getRange(
        prevPersonalSheet.getLastRow() + 1, 1, personalData.length, 3
      ).setValues(personalData);

      personalSheet.getRange(3, 1, lastPersonalRow - 2, 3).clearContent();
    }

    // B. PROCESS HOUSEHOLD EXPENSES & GENERATE PDF
    var lastHouseRow = houseSheet.getLastRow();
    var totalHouseSpent = 0;
    var pdfAttachment = null;

    if (lastHouseRow > 2) {
      var houseData = houseSheet.getRange(3, 1, lastHouseRow - 2, 3).getValues();

      var htmlContent = "<h2 style='color: #333; font-family: Arial, sans-serif;'>Weekly Household Expenses</h2>";
      htmlContent += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;'>";
      htmlContent += "<tr style='background-color: #f2f2f2;'><th align='left'>Date</th><th align='left'>Expense</th><th align='left'>Amount (KES)</th></tr>";

      for (var j = 0; j < houseData.length; j++) {
        var expName = houseData[j][0] || "Unknown";
        var expAmt = Number(houseData[j][1]) || 0;
        var rawDate = houseData[j][2];
        
        var expDate = (rawDate instanceof Date) 
          ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM dd, yyyy") 
          : (rawDate || "No Date");

        totalHouseSpent += expAmt;

        htmlContent += "<tr>";
        htmlContent += "<td>" + expDate + "</td>";
        htmlContent += "<td>" + expName + "</td>";
        htmlContent += "<td>" + expAmt.toFixed(2) + "</td>";
        htmlContent += "</tr>";
      }

      htmlContent += "<tr style='background-color: #e6e6e6;'><td colspan='2' align='right'><b>Total Spent:</b></td><td><b>" + totalHouseSpent.toFixed(2) + "</b></td></tr>";
      htmlContent += "</table>";

      var blob = Utilities.newBlob(htmlContent, MimeType.HTML, "Household_Expenses.html");
      pdfAttachment = blob.getAs(MimeType.PDF);
      pdfAttachment.setName("Weekly_Household_Expenses.pdf");

      prevHouseSheet.getRange(
        prevHouseSheet.getLastRow() + 1, 1, houseData.length, 3
      ).setValues(houseData);

      houseSheet.getRange(3, 1, lastHouseRow - 2, 3).clearContent();
    }

    // C. UPDATE FINANCIAL TRACKER SHEET (LEDGER MATH)
    trackerSheet.getRange("A4").setValue(totalPersonalSpent);

    var currentHouseTrackerRange = trackerSheet.getRange("A8");
    var previousHouseTotal = Number(currentHouseTrackerRange.getValue()) || 0;
    var newRunningHouseTotal = previousHouseTotal + totalHouseSpent;
    currentHouseTrackerRange.setValue(newRunningHouseTotal);

    var targetSpend = trackerSheet.getRange("F4").getValue();
    var amountSaved = trackerSheet.getRange("B4").getValue();

    // D. SEND SUCCESS EMAIL
    var subject = "📊 Weekly Expense Summary Report";
    var body = "Hello IaMdEvJaXoN,\n\nHere is your weekly financial summary:\n\n" +
               "• Personal Total Spent: KES " + totalPersonalSpent.toFixed(2) + "\n" +
               "• Weekly Target Spend: KES " + Number(targetSpend).toFixed(2) + "\n" +
               "• Amount Saved (Save this amount): KES " + Number(amountSaved).toFixed(2) + "\n" +
               "• Household Total Spent: KES " + totalHouseSpent.toFixed(2) + "\n\n" +
               "Your active expense sheets have been archived and cleared for the new week.\n\n" +
               "View your sheets here: " + ss.getUrl();

    var emailOptions = {};
    if (pdfAttachment !== null) {
      emailOptions.attachments = [pdfAttachment];
    }

    MailApp.sendEmail(recipientEmail, subject, body, emailOptions);

  } catch (error) {
    // FATAL ERROR CATCH: Notify the admin immediately
    var errorSubject = "⚠️ SYSTEM FAILURE: Weekly Expense Automation";
    var errorBody = "Hello IaMdEvJaXoN,\n\nYour automated weekly expense script encountered a fatal error and halted execution.\n\n" +
                    "Error Message: " + error.message + "\n\n" +
                    "Stack Trace:\n" + error.stack + "\n\n" +
                    "Please check your Google Sheet and Apps Script logs immediately.";
    MailApp.sendEmail(recipientEmail, errorSubject, errorBody);
  }
}

// 2. DAILY REIMBURSEMENT CHECK (Runs Every Day at 9:00 PM)
function processDailyReimbursement() {
  var recipientEmail = Session.getActiveUser().getEmail();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var houseSheet = ss.getSheetByName("house expenses");
    var trackerSheet = ss.getSheetByName("Financial tracker");

    if (!houseSheet || !trackerSheet) throw new Error("Required sheets missing for reimbursement processing.");

    var reimbursedRange = houseSheet.getRange("G2");
    var reimbursedAmount = reimbursedRange.getValue();

    if (reimbursedAmount !== "" && !isNaN(reimbursedAmount) && reimbursedAmount > 0) {
      var trackerRange = trackerSheet.getRange("A8");
      var oldExpenditure = Number(trackerRange.getValue()) || 0;
      var newExpenditure = oldExpenditure - Number(reimbursedAmount);
      
      trackerRange.setValue(newExpenditure);
      reimbursedRange.clearContent();

      var subject = "✅ Reimbursement Processed: Ledger Updated";
      var body = "Hello IaMdEvJaXoN,\n\n" +
                 "A reimbursement has been successfully processed and applied to your Financial Tracker.\n\n" +
                 "• Old Household Expenditure: KES " + oldExpenditure.toFixed(2) + "\n" +
                 "• Reimbursed Amount Subtracted: KES " + Number(reimbursedAmount).toFixed(2) + "\n" +
                 "• New Household Expenditure: KES " + newExpenditure.toFixed(2) + "\n\n" +
                 "Cell G2 has been automatically cleared.\n\n" +
                 "View your sheet here: " + ss.getUrl();

      MailApp.sendEmail(recipientEmail, subject, body);
    }
  } catch (error) {
    var errorSubject = "⚠️ SYSTEM FAILURE: Daily Reimbursement Check";
    var errorBody = "Hello IaMdEvJaXoN,\n\nThe daily reimbursement script failed.\n\nError: " + error.message + "\n\nStack: " + error.stack;
    MailApp.sendEmail(recipientEmail, errorSubject, errorBody);
  }
}

// ===========================================================================
// 3. DAILY EXPENSE REMINDER AUTOMATION (Runs Every Day at 7:00 PM)
// ===========================================================================
function sendDailyReminder() {
  var recipientEmail = Session.getActiveUser().getEmail();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subject = "📌 Daily Reminder: Log Your Expenses";
    var body = "Hello IaMdEvJaXoN,\n\n" +
               "Don't forget to log your personal and household expenses for today.\n\n" +
               "Click here to update your sheets: " + ss.getUrl();

    MailApp.sendEmail(recipientEmail, subject, body);
  } catch (error) {
    var errorSubject = "⚠️ SYSTEM FAILURE: Daily Reminder";
    var errorBody = "Hello IaMdEvJaXoN,\n\nThe daily reminder script failed to send.\n\nError: " + error.message;
    MailApp.sendEmail(recipientEmail, errorSubject, errorBody);
  }
}