function sendEmails() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var startRow = 2;  // First row of data to process
    var numRows = 2;   // Number of rows to process
    var emailColumn = 4;
    // Fetch the range of cells A2:B3
    var dataRange = sheet.getRange(startRow, 1, numRows, 5);
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    for (i in data) {
        var row = data[i];
        var emailAddress = row[emailColumn];  // First column
        var message = row[1];       // Second column
        var subject = "Sending emails from a Spreadsheet";
        if (emailAddress) {
            MailApp.sendEmail(emailAddress, subject, message);
        }
    }
}
