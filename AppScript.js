/**
 * Google Sheets Mail Merge with Unique Passwords + Google Drive Attachments
 * Author: HTPN ICT 1980 (adapted)
 * Version: 3.1
 */

const RECIPIENT_COL = "Email";
const NAME_COL = "Name";
const IC_NAME_COL = "Name";
const PASSWORD_COL = "Password";
const EMAIL_SENT_COL = "Sent";
const EMAIL_SUBJECT_COL = "Subject";
const EMAIL_ATTACHMENT_COL = "FileID";
const DATE_APPPOINT_COL = "DateAppointment";
const TIME_APPPOINT_COL = "TimeAppointment";
const LOC_APPPOINT_COL = "LocationAppointment";

const FOLDER_ID = "1fbcxeC0YQpJT6wRT_6NCssQa0Y4ThoIF"; // Google Drive folder containing files

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Mail Merge")
    .addItem("Send Emails", "sendEmailsWithAttachment")
    .addToUi();
}

/**
 * Sends personalized emails with attachment from Google Drive
 */
function sendEmailsWithAttachment() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const headers = data.shift();

  const recipientIdx = headers.indexOf(RECIPIENT_COL);
  const nameIdx = headers.indexOf(NAME_COL);
  const passwordIdx = headers.indexOf(PASSWORD_COL);
  const emailSentIdx = headers.indexOf(EMAIL_SENT_COL);
  const fileIdIdx = headers.indexOf(EMAIL_ATTACHMENT_COL);
  const subjectIdx = headers.indexOf(EMAIL_SUBJECT_COL);

  // Load all files in the folder into a map for fast lookup
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const filesMap = {};
  while (files.hasNext()) {
    const file = files.next();
    filesMap[file.getName()] = file;
  }

  const out = [];

  data.forEach((row) => {
    if (!row[emailSentIdx]) {
      const recipient = row[recipientIdx];
      const name = row[nameIdx];
      const password = row[passwordIdx];
      const fileName = row[fileIdIdx];
      const subjectLine = row[subjectIdx] || "No Subject";

      const bodyText = `Hello ${name},\n\n` +
                       `Your File Password is: ${password}\n\n` +
                       `Do not share your code with anyone.\n\n` +
                       `Kind regards,\nYour Team`;

      const bodyHtml = `<p>Hello <b>${name}</b>,</p>` +
                       `<p>Your File Password is: <b>${password}</b></p>` +
                       `<p>Do not share your code with anyone.</p>` +
                       `<p>Kind regards,<br>Your Team</p>`;

      try {
        const emailOptions = { htmlBody: bodyHtml };

        // Attach the file from Drive if it exists
        if (fileName && filesMap[fileName]) {
          emailOptions.attachments = [filesMap[fileName].getAs(MimeType.PDF)];
        } else if (fileName) {
          // Log a warning if file is missing and skip this email
          out.push([`File not found: ${fileName}`]);
          return;
        }

        GmailApp.sendEmail(recipient, subjectLine, bodyText, emailOptions);
        out.push([new Date()]);
      } catch (e) {
        out.push([`Error for ${recipient}: ${e.message}`]);
      }

      // 10-second delay between emails to avoid Gmail quota issues
      Utilities.sleep(10000);
    } else {
      out.push([row[emailSentIdx]]);
    }
  });

  // Update "Sent" column with timestamp or error
  sheet.getRange(2, emailSentIdx + 1, out.length, 1).setValues(out);
  SpreadsheetApp.getUi().alert("Emails sent successfully!");
}
