var EMAIL_CHECK_HEADING = 'Email manager?';
var EMAIL_CHECK_SEND = 'yes';
var EMAIL_CHECK_SENT = 'sent';
var EMAIL_COLUMNS = [1, 2];


function sendEmails() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var allData = sheet.getDataRange().getValues();
    var headings = allData[0];
    var emailCheckColumn = headings.indexOf(EMAIL_CHECK_HEADING);
    for (var i = 1; i < allData.length; i++) {
        var row = allData[i];
        if (row[emailCheckColumn] == EMAIL_CHECK_SEND) {
            emailRow(headings, allData[i], EMAIL_COLUMNS);
            // mark the cell as sent
            sheet.getRange(i + 1, emailCheckColumn + 1).setValue(EMAIL_CHECK_SENT);
        }
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
        return '<p><strong>{heading}</strong></p><p><pre>{value}</pre></p>'.replace('{heading}', heading).replace('{value}', value);
    }
    message = '';
    for (var i = 0; i < headings.length; i++) {
        message += fmtElement(headings[i], values[i]);
    }
    return message;
}
