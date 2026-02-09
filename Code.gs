/**
 * Google Sheets Mail Merge with Unique Passwords + Google Drive Attachments
 * Author: HTPN ICT 1980 (adapted)
 * Version: 4.1b
 */


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Merge Tools")
    //.addItem("Setup Template & Folder", "setupConfig")
    .addItem("Setup Template & Folder", "openConfigDialog")
    .addItem("Show Config", "showConfig")
    .addItem("Check Template & Folder", "testAccess")
    .addSeparator()
    .addItem("Sync", "syncData")
    .addItem("Generate Merge Doc", "generateCertificates")
    .addItem("Edit Email Template", "openEmailTemplateEditor") // NEW
    .addItem("Preview Email for Selected Row", "previewEmailForSelectedRow") // NEW
    .addItem("Send Emails", "sendCertificatesEmail")
    .addToUi();
}


function openConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ConfigDialog")
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Setup Template & Folder");
}

function getHeaderRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Return only non-empty headers
  return headers;
}

function saveConfig(templateId, folderId, options, mapping) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("TEMPLATE_ID", templateId);
  props.setProperty("FOLDER_ID", folderId);
  props.setProperty("OPTIONS", JSON.stringify(options));
  props.setProperty("MAPPING", JSON.stringify(mapping)); // save dValue -> {{valueX}} mapping
}

function showConfig() {
  const props = PropertiesService.getDocumentProperties();

  const templateId = props.getProperty("TEMPLATE_ID") || "Not set";
  const folderId = props.getProperty("FOLDER_ID") || "Not set";
  const options = props.getProperty("OPTIONS") ? JSON.parse(props.getProperty("OPTIONS")) : [];
  const mapping = props.getProperty("MAPPING") ? JSON.parse(props.getProperty("MAPPING")) : {};

  // Build HTML dynamically
  let html = '<div style="font-family:Arial; padding:12px;">';
  html += '<h3>Saved Configuration</h3>';
  html += `<p><strong>Template ID:</strong> ${templateId}</p>`;
  html += `<p><strong>Folder ID:</strong> ${folderId}</p>`;

  html += `<p><strong>Options:</strong> ${options.length ? options.join(", ") : "None"}</p>`;

  html += '<h4>Placeholder Mapping</h4>';
  if (Object.keys(mapping).length) {
    html += '<ul>';
    for (let key in mapping) {
      html += `<li><strong>${key}:</strong> ${mapping[key] || "(not mapped)"}</li>`;
    }
    html += '</ul>';
  } else {
    html += '<p>No mapping saved.</p>';
  }

  html += '<button onclick="google.script.host.close()" style="margin-top:15px;padding:6px 12px;">Close</button>';
  html += '</div>';

  // Show modal dialog
  const uiHtml = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(uiHtml, "Saved Config");
}

function syncData() {
  const ui = SpreadsheetApp.getUi();

  // 1️⃣ Ask for SOURCE Spreadsheet ID
  const idResponse = ui.prompt(
    "Sync Data",
    "Enter SOURCE Spreadsheet ID:",
    ui.ButtonSet.OK_CANCEL
  );

  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const SOURCE_SPREADSHEET_ID = idResponse.getResponseText().trim();

  if (!SOURCE_SPREADSHEET_ID) {
    ui.alert("Spreadsheet ID cannot be empty.");
    return;
  }

  // 2️⃣ Ask for SOURCE Sheet Name
  const sheetResponse = ui.prompt(
    "Sync Data",
    "Enter SOURCE Sheet Name:",
    ui.ButtonSet.OK_CANCEL
  );

  if (sheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const SOURCE_SHEET_NAME = sheetResponse.getResponseText().trim();

  if (!SOURCE_SHEET_NAME) {
    ui.alert("Sheet name cannot be empty.");
    return;
  }

  try {
    // 3️⃣ Open source
    const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      ui.alert("Source sheet not found: " + SOURCE_SHEET_NAME);
      return;
    }

    // 4️⃣ Target = current active sheet
    const targetSheet = SpreadsheetApp.getActiveSheet();

    // 5️⃣ Confirm overwrite
    const confirm = ui.alert(
      "Confirm Sync",
      `This will REPLACE ALL data in sheet:\n\n"${targetSheet.getName()}"\n\nContinue?`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    // 6️⃣ Sync data
    const data = sourceSheet.getDataRange().getValues();

    targetSheet.clearContents();
    targetSheet
      .getRange(1, 1, data.length, data[0].length)
      .setValues(data);

    ui.alert("✅ Data synced successfully!");

  } catch (err) {
    ui.alert("❌ Error:\n" + err.message);
  }
}




// ---------------- Setup Config ----------------
function setupConfig() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  let templatePrompt = ui.prompt(
    "Template ID",
    "Enter Google Docs Template ID:",
    ui.ButtonSet.OK_CANCEL
  );
  if (templatePrompt.getSelectedButton() !== ui.Button.OK) return;
  let templateId = templatePrompt.getResponseText();

  let folderPrompt = ui.prompt(
    "Folder ID",
    "Enter Google Drive Folder ID:",
    ui.ButtonSet.OK_CANCEL
  );
  if (folderPrompt.getSelectedButton() !== ui.Button.OK) return;
  let folderId = folderPrompt.getResponseText();

  props.setProperty("TEMPLATE_ID", templateId);
  props.setProperty("FOLDER_ID", folderId);

  ui.alert("Configuration saved successfully.");
}

// ---------------- Test Access ----------------
function testAccess() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const templateId = props.getProperty("TEMPLATE_ID");
  const folderId = props.getProperty("FOLDER_ID");

  if (!templateId || !folderId) {
    ui.alert("Please run 'Setup Template & Folder' first.");
    return;
  }

  let templateOk = false;
  let folderOk = false;
  let msg = "";

  // Test template access
  try {
    const file = DriveApp.getFileById(templateId);
    msg += "Template Access OK:\n" + file.getName() + "\n\n";
    templateOk = true;
  } catch (err) {
    msg += "Template Access Error:\n" + err.message + "\n\n";
  }

  // Test folder access
  try {
    const folder = DriveApp.getFolderById(folderId);
    msg += "Folder Access OK:\n" + folder.getName();
    folderOk = true;
  } catch (err) {
    msg += "Folder Access Error:\n" + err.message;
  }

   // Show result
  ui.alert("Access Test Result", msg, ui.ButtonSet.OK);
}


// ---------------- Generate Certificates Only (DYNAMIC) ----------------
function generateCertificates() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const templateId = props.getProperty("TEMPLATE_ID");
  const folderId   = props.getProperty("FOLDER_ID");
  const mappingRaw = props.getProperty("MAPPING");

  if (!templateId || !folderId || !mappingRaw) {
    ui.alert("Please run 'Setup Template & Folder' first.");
    return;
  }

  const mapping = JSON.parse(mappingRaw); // { dValue1: "{{value2}}", ... }

  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data   = sheet.getDataRange().getValues();
  const headers = data[0]; // header row
  const folder = DriveApp.getFolderById(folderId);

  let countGenerated = 0;
  let errorRows = [];

  // Malay month names
  const months = ["Jan","Feb","Mac","Apr","Mei","Jun","Jul","Ogos","Sep","Okt","Nov","Dis"];

  // Build header -> column index map (skip first column = index)
  const headerIndexMap = {};
  for (let c = 1; c < headers.length; c++) {
    if (headers[c]) headerIndexMap[headers[c]] = c;
  }

  for (let i = 1; i < data.length; i++) { // skip header row
    const rowNumber = i + 1;
    const row = data[i];

    const janaColIndex = headerIndexMap["Jana"]; // MUST exist
    if (!janaColIndex && janaColIndex !== 0) {
      ui.alert("Column 'Jana' not found.");
      return;
    }

    if (!row[1]) continue; // skip empty name rows
    if (row[janaColIndex] && !row[janaColIndex].toString().startsWith("ERR")) continue;

    try {
      // Copy template
      const copyDoc = DriveApp.getFileById(templateId)
        .makeCopy("Certificate - " + row[1], folder);

      const doc  = DocumentApp.openById(copyDoc.getId());
      const body = doc.getBody();

      // Loop mapping dynamically
      for (const dKey in mapping) {
        const placeholder = mapping[dKey]; // {{valueX}}
        if (!placeholder) continue;

        // dValueN -> column index (skip index column)
        const colNumber = parseInt(dKey.replace("dValue", ""), 10);
        const colIndex  = colNumber; // already offset because col 0 = index

        let value = row[colIndex] ?? "";

        // Auto-format Date → Malay DD MMM YYYY
        if (value instanceof Date && !isNaN(value)) {
          const day   = value.getDate();
          const month = months[value.getMonth()];
          const year  = value.getFullYear();
          value = `${day} ${month} ${year}`;
        }

        body.replaceText(
          "\\{\\{\\s*" + placeholder.replace(/[{}]/g, "") + "\\s*\\}\\}",
          value.toString()
        );
      }

      doc.saveAndClose();

      // Convert to PDF
      const pdfBlob = copyDoc.getAs(MimeType.PDF);

      // Unique filename
      const random12 = Math.floor(100000000000 + Math.random() * 900000000000);
      const now = new Date();
      const yy = String(now.getFullYear()).slice(-2);
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      const dd = String(now.getDate()).padStart(2, "0");

      const pdfName = `DOC_${random12}_${yy}${mm}${dd}.pdf`;
      folder.createFile(pdfBlob).setName(pdfName);

      // Trash temp doc
      copyDoc.setTrashed(true);

      // Update Jana + FileID (by header name)
      if (headerIndexMap["Jana"] !== undefined)
        sheet.getRange(rowNumber, headerIndexMap["Jana"] + 1).setValue(new Date());

      if (headerIndexMap["FileID"] !== undefined)
        sheet.getRange(rowNumber, headerIndexMap["FileID"] + 1).setValue(pdfName);

      countGenerated++;

    } catch (e) {
      sheet.getRange(rowNumber, headerIndexMap["Jana"] + 1)
        .setValue("ERR: " + e.message);

      errorRows.push(`Row ${rowNumber}: ${e.message}`);
      Logger.log(e);
    }
  }

  let msg = `Merge generation complete!\nTotal generated: ${countGenerated}`;
  if (errorRows.length) msg += "\n\nErrors:\n" + errorRows.join("\n");

  ui.alert(msg);
}

// Open the HTML editor popup
function openEmailTemplateEditor() {
  const html = HtmlService.createHtmlOutputFromFile("EmailTemplateEditor")
      .setWidth(600)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Email Template Editor");
}

// Get current template from Document Properties
function getEmailTemplate() {
  const props = PropertiesService.getDocumentProperties();
  return {
    subject: props.getProperty("EMAIL_SUBJECT") || "",
    body: props.getProperty("EMAIL_BODY") || ""
  };
}

// Save template to Document Properties
function saveEmailTemplate(subject, body) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("EMAIL_SUBJECT", subject);
  props.setProperty("EMAIL_BODY", body);
}

function getEmailAndFileIdFromRow(row) {
  const props = PropertiesService.getDocumentProperties();
  const mapping = JSON.parse(props.getProperty("MAPPING") || "{}");

  let email = "";
  let fileId = "";

  for (const dKey in mapping) {
    // Get the zero-based column index from dValueN
    const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
    const value = row[colIndex] ?? "";

    // Decide which one is email/fileId based on the **exact mapping keys**
    // Replace these numbers with the actual dValue for email/fileId from your config
    if (dKey === "dValue6") email = value;    // your email column
    if (dKey === "dValue9") fileId = value;   // your fileId column
  }

  return { email, fileId };
}

// ---------------- Preview Email (Mapping-driven Email/FileID) ----------------
function previewEmailForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  if (!selection) return SpreadsheetApp.getUi().alert("Please select a row to preview.");

  const rowNumber = selection.getRow();
  if (rowNumber === 1) return SpreadsheetApp.getUi().alert("Please select a data row, not the header.");

  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const props = PropertiesService.getDocumentProperties();
  const mapping = JSON.parse(props.getProperty("MAPPING"));
  const subjectTpl = props.getProperty("EMAIL_SUBJECT");
  const bodyTpl    = props.getProperty("EMAIL_BODY");

  if (!mapping || !subjectTpl || !bodyTpl)
    return SpreadsheetApp.getUi().alert("Missing Mapping or Email Template.");

  // --- Get column indexes dynamically from mapping ---
  const emailColIndex = parseInt("6".replace("dValue", "")) - 1;   // Hardcoded example: dValue6 is email
  const fileIdColIndex = parseInt("9".replace("dValue", "")) - 1;  // Hardcoded example: dValue9 is fileId

  // Better: read from mapping keys dynamically (from your config)
  // Find the mapping keys for email/fileid
  let emailKey = Object.keys(mapping).find(k => k === "dValue6");   // replace with your actual dValue for email
  let fileIdKey = Object.keys(mapping).find(k => k === "dValue9");  // replace with your actual dValue for fileId

  const email = emailKey ? row[parseInt(emailKey.replace("dValue",""))] : "(missing)";
  const fileId = fileIdKey ? row[parseInt(fileIdKey.replace("dValue",""))] : "(missing)";

  // Replace placeholders in subject/body dynamically
  const subject = Object.keys(mapping).reduce((acc, dKey) => {
    const value = row[parseInt(dKey.replace("dValue", ""))] ?? "";
    return acc.replace(new RegExp("\\{\\{\\s*" + mapping[dKey].replace(/[{}]/g,"") + "\\s*\\}\\}", "g"), value);
  }, subjectTpl);

  const body = Object.keys(mapping).reduce((acc, dKey) => {
    const value = row[parseInt(dKey.replace("dValue", ""))] ?? "";
    return acc.replace(new RegExp("\\{\\{\\s*" + mapping[dKey].replace(/[{}]/g,"") + "\\s*\\}\\}", "g"), value);
  }, bodyTpl);

  const htmlContent = `
    <div style="font-family:Arial; padding:12px;">
      <h3>Email Preview</h3>
      <p><strong>To:</strong> ${email}</p>
      <p><strong>Subject:</strong> ${subject}</p>
      <h4>Body:</h4>
      <div style="border:1px solid #ccc; padding:10px; max-height:300px; overflow:auto;">${body}</div>
      <button onclick="google.script.host.close()" style="margin-top:10px;padding:6px 12px;">Close</button>
    </div>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(600).setHeight(400),
    `Preview Email - Row ${rowNumber}`
  );
}


// ---------------- Send Emails (Mapping-driven Email/FileID) ----------------
function sendCertificatesEmail() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const mapping = JSON.parse(props.getProperty("MAPPING"));
  const subjectTpl = props.getProperty("EMAIL_SUBJECT");
  const bodyTpl    = props.getProperty("EMAIL_BODY");

  if (!mapping || !subjectTpl || !bodyTpl)
    return ui.alert("Missing Mapping or Email Template.");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  const sentColIndex = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].indexOf("Sent");

  // Find the column indexes from mapping
  const emailColKey = Object.keys(mapping).find(dKey => mapping[dKey].toLowerCase().includes("email"));
  const fileIdColKey = Object.keys(mapping).find(dKey => mapping[dKey].toLowerCase().includes("fileid"));
  const emailColIndex = emailColKey ? parseInt(emailColKey.replace("dValue", ""), 10) : null;
  const fileIdColIndex = fileIdColKey ? parseInt(fileIdColKey.replace("dValue", ""), 10) : null;

  let sentCount = 0, errorRows = [];

  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;
    const row = data[i];

    if (sentColIndex !== -1 && row[sentColIndex] && !row[sentColIndex].toString().startsWith("ERR")) continue;

    try {
      const email = emailColIndex !== null ? row[emailColIndex] : null;
      const fileId = fileIdColIndex !== null ? row[fileIdColIndex] : null;
      if (!email || !fileId) throw new Error("Email or FileID missing");

      const subject = Object.keys(mapping).reduce((acc, dKey) => {
        const placeholder = mapping[dKey];
        const value = row[parseInt(dKey.replace("dValue", ""), 10)] ?? "";
        return acc.replace(new RegExp("\\{\\{\\s*" + placeholder.replace(/[{}]/g, "") + "\\s*\\}\\}", "g"), value);
      }, subjectTpl);

      const body = Object.keys(mapping).reduce((acc, dKey) => {
        const placeholder = mapping[dKey];
        const value = row[parseInt(dKey.replace("dValue", ""), 10)] ?? "";
        return acc.replace(new RegExp("\\{\\{\\s*" + placeholder.replace(/[{}]/g, "") + "\\s*\\}\\}", "g"), value);
      }, bodyTpl);

      const pdfFile = DriveApp.getFileById(fileId).getAs(MimeType.PDF);

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body,
        attachments: [pdfFile]
      });

      if (sentColIndex !== -1) sheet.getRange(rowNumber, sentColIndex + 1).setValue(new Date());
      sentCount++;

    } catch (e) {
      if (sentColIndex !== -1) sheet.getRange(rowNumber, sentColIndex + 1).setValue("ERR: " + e.message);
      errorRows.push(`Row ${rowNumber}: ${e.message}`);
      Logger.log(e);
    }
  }

  let msg = `Email sending complete!\nTotal sent: ${sentCount}`;
  if (errorRows.length) msg += "\n\nErrors:\n" + errorRows.join("\n");
  ui.alert(msg);
}
