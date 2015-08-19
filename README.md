# google-forms-emailer

A google apps script to send emails from a google sheet.

This script can be installed on any google form to setup automatic emails from that form.

## How it works

- First you setup a [google form](https://www.google.com/forms/about/).
- In this form you should collect email addresses somewhere. These can either be automatically collected from the logged in user, questions in the form, or both.
- You should also add a question that triggers when the form should be sent.
- Then whenever the trigger matches, the entire form will be sent to all the configured emails.

## Configuration

This script has a few configuration options which are set at the top.
These correspond to values in your form that you use, and are described here in more detail.

These are defined here:

Config Option       | Description
------------------- | ------------
EMAIL_CHECK_HEADING | The heading of the column that will be used to check whether to send the email. This is typically the name of the question.
EMAIL_CHECK_SEND    | The value of that column that triggers a sent email.
EMAIL_CHECK_SENT    | The value that will be saved to that column to flag that the email has been sent. This defaults to `"sent"` and does not need to be changed.
EMAIL_COLUMNS       | A list of the column IDs that contain email addresses to send to. These start from 1.

The built in configuration is setup where the first two columns of the spreadsheet contain the email addresses.
It is also setup with a question trigger assuming the form has a question called "Email manager?" which, when set to "yes" will send the email.

## Installation

- Make a google form
- Load the responses spreadsheet
- Click "Tools" --> "Script Editor"
- Paste the contents of `Code.gs` into the text area.
- In the script editor menu chose the `sendEmails` function from the dropdown.
- Click the clock to open the triggers.
- Add a trigger to run the `sendEmails` function from the spreadsheet on form submit.
- Submitting a form should now send emails as configured (see above)
