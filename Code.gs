// ---------------------------------------------------------------- //
// ----------------- LETTERS MUST SET VARIABLES  ------------------ //
// ---------------------------------------------------------------- //

const LETTERS_TEMPLATE_ID = 'ID OF YOUR LETTER TEMPLATE';
const LETTERS_DESTINATION_FOLDER_ID = 'ID OF DESTINATION FOLDER FOR YOUR GENERATED';
const LETTER_DEFAULT_FILE_TITLE = 'TITLE PREFIX OR SUFIX FOR YOUR LETTER';

// ---------------------------------------------------------------- //
// ------------------ EMAILS MUST SET VARIABLES  ------------------ //
// ---------------------------------------------------------------- //

const EMAILS_TEMPLATE_ID = 'ID OF YOUR EMAIL TEMPLATE';
const EMAILS_SUBJECT = 'YOUR EMAIL SUBJECT';
const ATTACHMENT_FILES_IDS = ['ARRAY', 'LIST', 'OF', 'FILES', 'IDS']; // File must be uploaded on Google Drive.
const IS_SIGNATURE_TRUE = true; // change to false if you don't have an Gmail Signature set.

//////////////////////////////////////////////////////////////////////

// ---------------------------------------------------------------- //
// ----------------------------- MENU ----------------------------- //
// ---------------------------------------------------------------- //

// Build Menu on Google Sheets
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Report and Email Automation');
    menu.addItem('Create Letters', 'createLetters');
    menu.addItem('Send Emails', 'sendEmails');
    menu.addToUi();
}

// ---------------------------------------------------------------- //
// ---------------------- GENERAL FUNCTIONS ----------------------- //
// ---------------------------------------------------------------- //

// Get Date
function getTodaysDate() {
    return new Date().toLocaleDateString('en-GB', {day: '2-digit', month: 'long', year: '2-digit'});
}

// ---------------------------------------------------------------- //
// -------------------- GENERATE LETTERS FUNC --------------------- //
// ---------------------------------------------------------------- //

function createLetters() {
    // Get letter template and final destination folder
    const letterTemplate = DriveApp.getFileById(LETTERS_TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(LETTERS_DESTINATION_FOLDER_ID);

    // Open Sheet and Get Rows
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();

    // Get current date to be used on letter
    const date = getTodaysDate();

    // forEach loop through list of rows
    rows.forEach((row, index) => {
        if (index === 0) return; // jump header
        if (row[3]) return; // check for existing doc files

        // Create a copy of the letter template
        const copy = letterTemplate.makeCopy(`${row[1]} - ${LETTER_DEFAULT_FILE_TITLE}`, destinationFolder);
        // uncomment line above for prefix
        // const copy = letterTemplate.makeCopy(`${LETTER_DEFAULT_FILE_TITLE} - ${row[1]}`, destinationFolder);
        // uncomment line above for sufix

        // Open the copied letter
        const doc = DocumentApp.openById(copy.getId());

        // Get letter body (so we can edit)
        const body = doc.getBody();

        // Replace / Edit fields according to google sheets list
        const city = row[0];
        const eventDate = row[1];
        const eventTime = row[2];
        const destinatary = row[3];

        body.replaceText('{{Date}}', date); // add's today date on the letter
        body.replaceText('{{City}}', city);
        body.replaceText('{{Event Date}}', eventDate);
        body.replaceText('{{Event Time}}', eventTime);
        body.replaceText('{{Destinatary}}', destinatary);

        // save edited letter
        doc.saveAndClose();

        // Get google doc url
        const url = doc.getUrl();

        // update google sheets list with link to new letter file
        // index + 1 - Because JS Arrays starts at Zero and Google Sheets rows starts at One
        sheet.getRange(`D${index + 1}`).setValue(url);
    });
}

// ---------------------------------------------------------------- //
// ----------------------- SEND EMAILS FUNC ----------------------- //
// ---------------------------------------------------------------- //

function sendEmails() {
    // Open Email Template File and get Email Body
    const emailTemplate = DocumentApp.openById(EMAILS_TEMPLATE_ID);

    // Open Sheet and Get Rows
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();

    // Get Gmail Signature - > MUST HAVE Gmail API Services activated
    // If you don't have a Gmail signature comment the line below.
    const signature = Gmail.Users.Settings.SendAs.list('me').sendAs.find(acc => acc.isDefault).signature;

    // Get Attatchment Files Array (Map Array Method)
    const files = ATTACHMENT_FILES_IDS.map(id => {
        return DriveApp.getFileById(id);
    });

    // Get current date for future reference
    const date = getTodaysDate();

    // forEach loop through list of rows
    rows.forEach((row, index) => {
        if (index === 0) return; // jump header
        if (row[3]) return; // check if email has been sent

        // Convert body to a String
        let body = emailTemplate.getBody().copy();
        body.editAsText();

        // Replace / Edit fields according to google sheets list
        const city = row[0];
        const destinatary = row[1];
        const destinataryEmail = row[2];

        body.replaceText('{{Destinatary}}', destinatary);
        body.replaceText('{{City}}', city);

        // getText() called here preserve the \n spaces. If called to put on a different var it loses format
        IS_SIGNATURE_TRUE
            ? (htmlBody = `<p>${body.getText()}</p><br><br>${signature}`)
            : (htmlBody = `<p>${body.getText()}</p>`);

        // Send Email
        GmailApp.sendEmail(destinataryEmail, EMAILS_SUBJECT, null, {
            htmlBody: htmlbody,
            attachments: ATTACHMENT_FILES,
        });

        // Mark email as Sent
        sheet.getRange(`D${index + 1}`).setValue('Sent');
        sheet.getRange(`E${index + 1}`).setValue(date);
    });
}
