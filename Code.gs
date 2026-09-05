//END-OF-WEEK AUTOMATION (Runs Every Saturday Night at 9:00 PM)
function runEndOfWeekAutomation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //Sheet References
  var personalSheet = ss.getSheetByName("personal expenses");
  var houseSheet = ss.getSheetByName("house expenses");
  var prevPersonalSheet = ss.getSheetByName("previous personal expenses");
  var prevHouseSheet = ss.getSheetByName("previous household expenses");
  var trackerSheet = ss.getSheetByName("Financial tracker");

  //PROCESS PERSONAL EXPENSES
  var lastPersonalRow = personalSheet.getLastRow();
  var totalPersonalSpent = 0;

  if (lastPersonalRow > 2) { 
    var personalData = personalSheet.getRange(3, 1, lastPersonalRow - 2, 3).getValues();
    
    //Calculate total spent
    for (var i = 0; i < personalData.length; i++) {
      totalPersonalSpent += Number(personalData[i][1]) || 0;
    }

    //Append individual rows to history archive
    prevPersonalSheet.getRange(
      prevPersonalSheet.getLastRow() + 1, 1, personalData.length, 3
    ).setValues(personalData);

    //Clear active personal sheet data (Rows 3 downwards)
    personalSheet.getRange(3, 1, lastPersonalRow - 2, 3).clearContent();
  }

  //PROCESS HOUSEHOLD EXPENSES
  var lastHouseRow = houseSheet.getLastRow();
  var totalHouseSpent = 0;
  var pdfAttachment = null;

  if (lastHouseRow > 2) {
    var houseData = houseSheet.getRange(3, 1, lastHouseRow - 2, 3).getValues();

    //Calculate total spent & build HTML for the PDF
    var htmlContent = "<h2 style='color: #333; font-family: Arial, sans-serif;'>Weekly Household Expenses</h2>";
    htmlContent += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;'>";
    htmlContent += "<tr style='background-color: #f2f2f2;'><th align='left'>Date</th><th align='left'>Expense</th><th align='left'>Amount (KES)</th></tr>";

    for (var j = 0; j < houseData.length; j++) {
      var expName = houseData[j][0] || "Unknown";
      var expAmt = Number(houseData[j][1]) || 0;
      var rawDate = houseData[j][2];
      
      // Format the date properly for the PDF
      var expDate = (rawDate instanceof Date) 
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM dd, yyyy") 
        : (rawDate || "No Date");

      totalHouseSpent += expAmt;

      // Add row to HTML table
      htmlContent += "<tr>";
      htmlContent += "<td>" + expDate + "</td>";
      htmlContent += "<td>" + expName + "</td>";
      htmlContent += "<td>" + expAmt.toFixed(2) + "</td>";
      htmlContent += "</tr>";
    }

    // Add total row at the bottom of the table
    htmlContent += "<tr style='background-color: #e6e6e6;'><td colspan='2' align='right'><b>Total Spent:</b></td><td><b>" + totalHouseSpent.toFixed(2) + "</b></td></tr>";
    htmlContent += "</table>";

    // Convert HTML directly into a PDF Blob
    var blob = Utilities.newBlob(htmlContent, MimeType.HTML, "Household_Expenses.html");
    pdfAttachment = blob.getAs(MimeType.PDF);
    pdfAttachment.setName("Weekly_Household_Expenses.pdf");

    //Append individual rows to history archive
    prevHouseSheet.getRange(
      prevHouseSheet.getLastRow() + 1, 1, houseData.length, 3
    ).setValues(houseData);

    //Clear active house sheet data (Rows 3 downwards)
    houseSheet.getRange(3, 1, lastHouseRow - 2, 3).clearContent();
  }

  //UPDATE FINANCIAL TRACKER SHEET
  // Update Personal Expenses (Overwrites with the new week's total)
  trackerSheet.getRange("A4").setValue(totalPersonalSpent);

  // Update Household Expenses (Adds to the existing running total)
  var currentHouseTrackerRange = trackerSheet.getRange("A8");
  var previousHouseTotal = Number(currentHouseTrackerRange.getValue()) || 0;
  var newRunningHouseTotal = previousHouseTotal + totalHouseSpent;
  
  currentHouseTrackerRange.setValue(newRunningHouseTotal);

  var targetSpend = trackerSheet.getRange("F4").getValue();
  var amountSaved = trackerSheet.getRange("B4").getValue();

  //SEND EMAIL SUMMARY REPORT WITH ATTACHMENT
  var recipientEmail = Session.getActiveUser().getEmail();
  var subject = "📊 Weekly Expense Summary Report";
  
  var body = "Hello, IamJaxon,\n\nHere is your weekly financial summary:\n\n" +
             "• Personal Total Spent: KES " + totalPersonalSpent.toFixed(2) + "\n" +
             "• Weekly Target Spend: KES " + Number(targetSpend).toFixed(2) + "\n" +
             "• Amount Saved(Save this amount): KES " + Number(amountSaved).toFixed(2) + "\n" +
             "• Household Total Spent: KES " + totalHouseSpent.toFixed(2) + "\n\n" +
             "Your active expense sheets have been archived and cleared for the new week.\n\n" +
             "View your sheets here: " + ss.getUrl();

  // Configure email options to include the PDF if data existed
  var emailOptions = {};
  if (pdfAttachment !== null) {
    emailOptions.attachments = [pdfAttachment];
  }

  MailApp.sendEmail(recipientEmail, subject, body, emailOptions);
}

//DAILY EXPENSE REMINDER AUTOMATION (Runs Every Day at 7:00 PM)
function sendDailyReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recipientEmail = Session.getActiveUser().getEmail();
  
  var subject = "📌 Daily Reminder: Log Your Expenses";
  var body = "Hi, IamJaxon,\n\n" +
             "Don't forget to log your personal and household expenses for today.\n\n" +
             "Click here to update your sheets: " + ss.getUrl();

  MailApp.sendEmail(recipientEmail, subject, body);
}

// DAILY REIMBURSEMENT CHECK (Runs Every Day at 9:00 PM)
function processDailyReimbursement() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet References
  var houseSheet = ss.getSheetByName("house expenses");
  var trackerSheet = ss.getSheetByName("Financial tracker");

  // Read the reimbursed amount from cell G2
  var reimbursedRange = houseSheet.getRange("G2");
  var reimbursedAmount = reimbursedRange.getValue();

  // Check if G2 has a valid number greater than 0
  if (reimbursedAmount !== "" && !isNaN(reimbursedAmount) && reimbursedAmount > 0) {
    
    // Read current household expenditure from A8 in the Financial tracker
    var trackerRange = trackerSheet.getRange("A8");
    var oldExpenditure = Number(trackerRange.getValue()) || 0;
    
    // Calculate the new expenditure after reimbursement
    var newExpenditure = oldExpenditure - Number(reimbursedAmount);
    
    // Update cell A8 with the new calculated value
    trackerRange.setValue(newExpenditure);
    
    // Clear cell G2 in house expenses so it isn't subtracted again tomorrow
    reimbursedRange.clearContent();

    // Send the email notification
    var recipientEmail = Session.getActiveUser().getEmail();
    var subject = "✅ Reimbursement Processed: Household Expenses Updated";
    
    var body = "Hello IaMdEvJaXoN,\n\n" +
               "A reimbursement has been successfully processed and applied to your Financial Tracker.\n\n" +
               "• Old Household Expenditure: KES " + oldExpenditure.toFixed(2) + "\n" +
               "• Reimbursed Amount Subtracted: KES " + Number(reimbursedAmount).toFixed(2) + "\n" +
               "• New Household Expenditure: KES " + newExpenditure.toFixed(2) + "\n\n" +
               "The reimbursement cell (G2) has been automatically cleared.\n\n" +
               "View your updated sheet here: " + ss.getUrl();

    MailApp.sendEmail(recipientEmail, subject, body);
  }
  
  // If G2 is empty or 0, the script simply finishes without doing anything.
}