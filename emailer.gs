function sendEmails() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var allData = sheet.getDataRange().getValues();
    var headings = allData[0];
    var employeeEmailColumn = 1;
    var managerEmailColumn = 2;
    for (var i = 1; i < allData.length; i++) {
        emailRow(headings, allData[i], [employeeEmailColumn, managerEmailColumn]);
    }
}


function emailRow(headingRow, dataRow, addressIndexes) {
    var emails = [];
    for (var i = 0; i < addressIndexes.length; i++) {
        emails.push(dataRow[addressIndexes[i]]);
    }
    var subject = "Sending emails from a Spreadsheet";
    if (emails.length) {
        MailApp.sendEmail({
            to: emails[0],
            cc: emails.slice(1).join(),
            subject: subject,
            htmlBody: formatMessage(headingRow, dataRow)
        })
    }
}


function formatMessage(headings, values) {
    function fmtElement(heading, value) {
        return '<p><strong>{heading}</strong></p><p>{value}</p>'.replace('{heading}', heading).replace('{value}', value);
    }
    message = '';
    for (var i = 0; i < headings.length; i++) {
        message += fmtElement(headings[i], values[i]);
    }
    return message;
}
