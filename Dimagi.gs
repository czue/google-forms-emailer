var PRIMARY_EMAIL_COLUMN = 1;

// self evals
var SELF_EVAL_FORM_URL = 'https://docs.google.com/a/dimagi.com/forms/d/1SOAiTqQbr5jQcs2jpN_jNuak0B8EIW9xXbn7Cdi-Cuk/viewform';
var SELF_EVAL_SHEET_INDEX = 0;
var SELF_EVAL_CALC_SHEET_INDEX = 5;
var EMAIL_CHECK_HEADING = 'Finalize and Submit to manager?';
var EMAIL_CHECK_SEND = 'Finalize and send to manager';
var EMAIL_CHECK_SENT = 'sent';
var CALC_EMAILS = [2, 3];

function sendAll() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sendSelfEvals(ss);
    sendManagerReviews(ss);
    sendPeerFeedback(ss);
}


function sendSelfEvals(ss) {
    var mainSheetInfo = getSheetData(ss, SELF_EVAL_SHEET_INDEX);
    var calcSheetInfo = getSheetData(ss, SELF_EVAL_CALC_SHEET_INDEX);
    var emailCheckColumn = mainSheetInfo.headings.indexOf(EMAIL_CHECK_HEADING);
    for (var i = 1; i < mainSheetInfo.data.length; i++) {
        var mainRow = mainSheetInfo.data[i];
        if (mainRow[emailCheckColumn] === EMAIL_CHECK_SEND) {
            var calcRow = calcSheetInfo.data[i];
            var primaryEmail = mainRow[PRIMARY_EMAIL_COLUMN];
            var managerEmails = extractEmailsFromRow(calcRow, CALC_EMAILS);
            var selfEvalBody = formatSelfEval(
                formatMessage(calcSheetInfo.headings, calcRow , HEADING_VALUE_TEMPLATE_COMPACT),
                formatMessage(mainSheetInfo.headings, mainRow, HEADING_VALUE_TEMPLATE_FULL, 1)
            );
            sendEmail(primaryEmail, managerEmails, 'Self eval for ' + primaryEmail, selfEvalBody);
            // mark the cell as sent
            mainSheetInfo.sheet.getRange(i + 1, emailCheckColumn + 1).setValue(EMAIL_CHECK_SENT);
        }
    }
}

function sendManagerReviews(ss) {
    // config data
    var MANAGER_REVIEW_SHEET_INDEX = 1;
    var MANAGER_REVIEW_CALC_SHEET_INDEX = 6;
    var MANAGER_EMAIL_CHECK_HEADING = 'internal_email_status';
    var MANAGER_CALC_EMAILS = [2];
    var MANAGER_CALC_MANAGER_COLUMN = 1;

    var mainSheetInfo = getSheetData(ss, MANAGER_REVIEW_SHEET_INDEX);
    var calcSheetInfo = getSheetData(ss, MANAGER_REVIEW_CALC_SHEET_INDEX);
    var emailCheckColumn = mainSheetInfo.headings.indexOf(MANAGER_EMAIL_CHECK_HEADING);
    for (var i = 1; i < mainSheetInfo.data.length; i++) {
        var mainRow = mainSheetInfo.data[i];
        if (mainRow[emailCheckColumn] !== EMAIL_CHECK_SENT) {
            var calcRow = calcSheetInfo.data[i];
            var primaryEmail = mainRow[PRIMARY_EMAIL_COLUMN];
            var managerEmail = calcRow[MANAGER_CALC_MANAGER_COLUMN];
            var ccEmails = extractEmailsFromRow(calcRow, MANAGER_CALC_EMAILS);
            var managerReviewBody = formatMessage(mainSheetInfo.headings, mainRow , HEADING_VALUE_TEMPLATE_FULL, 1);
            sendEmail(primaryEmail, ccEmails, 'Manager review for ' + managerEmail + ' from ' + primaryEmail, managerReviewBody);
            // mark the cell as sent
            mainSheetInfo.sheet.getRange(i + 1, emailCheckColumn + 1).setValue(EMAIL_CHECK_SENT);
        }
    }
}


function sendPeerFeedback(ss) {
    // todo: this is basically exactly the same as manager review
    // config data
    var PEER_REVIEW_SHEET_INDEX = 2;
    var PEER_REVIEW_CALC_SHEET_INDEX = 7;
    var PEER_EMAIL_CHECK_HEADING = 'internal_email_status';
    var PEER_MANAGER_EMAIL_COLUMN_INDEX = 1;
    var PEER_AUTHOR_INDEX = 2;

    var mainSheetInfo = getSheetData(ss, PEER_REVIEW_SHEET_INDEX);
    var calcSheetInfo = getSheetData(ss, PEER_REVIEW_CALC_SHEET_INDEX);
    var emailCheckColumn = mainSheetInfo.headings.indexOf(PEER_EMAIL_CHECK_HEADING);
    for (var i = 1; i < mainSheetInfo.data.length; i++) {
        var mainRow = mainSheetInfo.data[i];
        if (mainRow[emailCheckColumn] !== EMAIL_CHECK_SENT) {
            var calcRow = calcSheetInfo.data[i];
            var peerName = mainRow[PRIMARY_EMAIL_COLUMN];
            var managerEmail = calcRow[PEER_MANAGER_EMAIL_COLUMN_INDEX];
            var peerReviewBody = formatMessage(mainSheetInfo.headings, mainRow , HEADING_VALUE_TEMPLATE_FULL, 1);
            sendEmail(managerEmail, [], 'Peer review for ' + peerName + ' from ' + mainRow[PEER_AUTHOR_INDEX], peerReviewBody);
            // mark the cell as sent
            mainSheetInfo.sheet.getRange(i + 1, emailCheckColumn + 1).setValue(EMAIL_CHECK_SENT);
        }
    }
}


function getSheetData(spreadsheet, sheetIndex, rows) {
    var sheet = spreadsheet.getSheets()[sheetIndex];
    var data = sheet.getDataRange().getValues();
    if (rows) {
        var colCount = data[0].length;
        data = sheet.getRange(1, 1, colCount, rows).getValues();
    }
    return {
        'sheet': sheet,
        'data': data,
        'headings': data[0]
    };
}


function sendEmail(primaryEmail, ccEmails, subject, messageBody) {
    if (isValidEmail(primaryEmail)) {
        MailApp.sendEmail({
            to: primaryEmail,
            subject: subject,
            cc: ccEmails.join(),
            htmlBody: messageBody
        });
    }
}


function extractEmailsFromRow(row, emailIndices) {
    var emails = [];
    for (var i = 0; i < emailIndices.length; i++) {
        var email = row[emailIndices[i]]
        if (isValidEmail(email)) {
            emails.push(email);
        }
    }
    return emails;
}


function isValidEmail(email) {
    return email.indexOf('@') !== -1;
}


HEADING_VALUE_TEMPLATE_COMPACT = '<p><strong>{heading}</strong>: {value}</p>';
HEADING_VALUE_TEMPLATE_FULL = '<p><strong>{heading}</strong></p><p><pre>{value}</pre></p>';

function formatMessage(headings, values, template, stripFromEnd) {

    function fmtElement(heading, value) {
        return template.replace('{heading}', heading).replace('{value}', value);
    }
    message = '';
    for (var i = 0; i < headings.length - stripFromEnd; i++) {
        message += fmtElement(headings[i], values[i]);
    }
    return message;
}

function formatSelfEval(calcSummary, selfEval) {
    return (
        "<p>Self eval submitted! You can <a href='{form}'>return to the form</a> to edit your responses.</p>" +
        "<h3>Employee Information</h3>{summary}" +
        "<h3>Self Eval</h3>{eval}"
    ).replace('{form}', SELF_EVAL_FORM_URL).replace('{summary}', calcSummary).replace('{eval}', selfEval);


}
