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
    //.addItem("tESTT", "TESTaWSD")
    .addSeparator()
    

    .addToUi();
}



function TESTaWSD() {
  const templateId = PropertiesService.getDocumentProperties().getProperty("TEMPLATE_ID");
  const file = DriveApp.getFileById(templateId);
  Logger.log(file.getName());
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

  const mapping = JSON.parse(mappingRaw); // { dValue12: "{{value12}}", ... }

  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data   = sheet.getDataRange().getValues();
  const headers = data[0]; // header row
  const folder = DriveApp.getFolderById(folderId);

  let countGenerated = 0;
  let errorRows = [];

  // Malay month names
  const months = ["Jan","Feb","Mac","Apr","Mei","Jun","Jul","Ogos","Sep","Okt","Nov","Dis"];

  // Build header -> column index map (all headers)
  const headerIndexMap = {};
  headers.forEach((h, i) => {
    if (h) headerIndexMap[h] = i;
  });

  // Build mapping dValue -> column index dynamically
  const mappingIndexes = {};
  for (const dKey in mapping) {
    const placeholder = mapping[dKey]; // e.g., {{value11}}
    const valueNum = parseInt(placeholder.replace(/[^0-9]/g, ""), 10); // get number from valueXX
    const colIndex = valueNum - 1; // zero-based index
    mappingIndexes[dKey] = colIndex;
  }

  // Find Jana and FileID columns by header
  const janaColIndex = headerIndexMap["Jana"];
  const fileIdColIndex = headerIndexMap["FileID"];

  if (janaColIndex === undefined || fileIdColIndex === undefined) {
    ui.alert("Required column 'Jana' or 'FileID' not found.");
    return;
  }

  for (let i = 1; i < data.length; i++) { // skip header row
    const rowNumber = i + 1;
    const row = data[i];

    if (!row[1]) continue; // skip empty name rows
    if (row[janaColIndex] && !row[janaColIndex].toString().startsWith("ERR")) continue;

    try {
      // Copy template
      const copyDoc = DriveApp.getFileById(templateId)
        .makeCopy("TEMP - " + row[1], folder);

      const doc  = DocumentApp.openById(copyDoc.getId());
      const body = doc.getBody();

      // Loop mapping dynamically
      for (const dKey in mapping) {
        const placeholder = mapping[dKey];
        if (!placeholder) continue;

        const colIndex  = mappingIndexes[dKey];
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

      // Update Jana + FileID in sheet
      sheet.getRange(rowNumber, janaColIndex + 1).setValue(new Date());
      sheet.getRange(rowNumber, fileIdColIndex + 1).setValue(pdfName);

      countGenerated++;

    } catch (e) {
      sheet.getRange(rowNumber, janaColIndex + 1)
        .setValue("ERR: " + e.message);

      errorRows.push(`Row ${rowNumber}: ${e.message}`);
      Logger.log(e);
    }
  }

  let msg = `✅ Merge generation complete!\nTotal generated: ${countGenerated}`;
  if (errorRows.length) msg += "\n\nErrors:\n" + errorRows.join("\n");

  ui.alert(msg);
}


// Open the HTML editor popup
function openEmailTemplateEditor() {
  const html = HtmlService.createHtmlOutputFromFile("EmailTemplateEditor")
      .setWidth(800)
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

// ---------------- Preview Email (Fixed - Mapping-driven) ----------------
function previewEmailForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  if (!selection) return SpreadsheetApp.getUi().alert("Please select a row to preview.");

  const rowNumber = selection.getRow();
  if (rowNumber === 1) return SpreadsheetApp.getUi().alert("Please select a data row, not the header.");

  const props = PropertiesService.getDocumentProperties();
  const mappingProp = props.getProperty("MAPPING");
  const subjectTpl = props.getProperty("EMAIL_SUBJECT");
  const bodyTpl = props.getProperty("EMAIL_BODY");

  if (!mappingProp || !subjectTpl || !bodyTpl)
    return SpreadsheetApp.getUi().alert("Missing Mapping or Email Template. Please run Setup first.");

  const mapping = JSON.parse(mappingProp);
  
  // Get the row data
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Auto-detect Email and FileID columns (same logic as send function)
  let emailColIndex = null;
  let fileNameColIndex = null;

  // First try: Check first data row for patterns
  if (rowNumber > 1) {
    for (const dKey in mapping) {
      const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
      if (colIndex >= row.length) continue;
      
      const value = String(row[colIndex] || "");
      
      // Detect email
      if (value.includes("@") && value.includes(".")) {
        emailColIndex = colIndex;
      }
      // Detect PDF filename
      else if (value.startsWith("DOC_") && value.endsWith(".pdf")) {
        fileNameColIndex = colIndex;
      }
    }
  }

  // Second try: Search headers
  if (emailColIndex === null) {
    emailColIndex = headers.findIndex(h => /email|e-mail|emel/i.test(String(h)));
  }
  if (fileNameColIndex === null) {
    fileNameColIndex = headers.findIndex(h => /fileid|file.*name|pdf|sijil/i.test(String(h)));
  }

  // Third try: Look for specific placeholder names in mapping
  if (emailColIndex === null) {
    // Find mapping value that looks like an email placeholder {{email}} or {{emel}}
    for (const dKey in mapping) {
      if (/email|emel/i.test(mapping[dKey])) {
        emailColIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
        break;
      }
    }
  }

  const email = emailColIndex !== null && emailColIndex !== -1 ? row[emailColIndex] : "(email not detected)";
  const fileId = fileNameColIndex !== null && fileNameColIndex !== -1 ? row[fileNameColIndex] : "(filename not detected)";

  // Replace placeholders in subject/body using mapping
  const subject = Object.keys(mapping).reduce((acc, dKey) => {
    const placeholder = mapping[dKey].replace(/[{}]/g, "");
    const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
    const value = row[colIndex] ?? "";
    return acc.replace(new RegExp("\\{\\{\\s*" + placeholder + "\\s*\\}\\}", "g"), value);
  }, subjectTpl);

  const body = Object.keys(mapping).reduce((acc, dKey) => {
    const placeholder = mapping[dKey].replace(/[{}]/g, "");
    const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
    const value = row[colIndex] ?? "";
    return acc.replace(new RegExp("\\{\\{\\s*" + placeholder + "\\s*\\}\\}", "g"), value);
  }, bodyTpl);

  // Build preview HTML
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .field { margin: 10px 0; padding: 10px; background: #f8f9fa; border-radius: 4px; }
        .label { font-weight: bold; color: #555; font-size: 12px; text-transform: uppercase; }
        .value { margin-top: 4px; color: #333; }
        .body-box { border: 1px solid #ddd; padding: 15px; border-radius: 4px; background: white; max-height: 300px; overflow: auto; }
        .attachment { color: #4285f4; }
        button { margin-top: 15px; padding: 10px 20px; background: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
        button:hover { background: #3367d6; }
        .warning { color: #f44336; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h3>📧 Email Preview - Row ${rowNumber}</h3>
        
        <div class="field">
          <div class="label">To:</div>
          <div class="value">${email} ${emailColIndex === null ? '<span class="warning">⚠️ Could not detect email column</span>' : ''}</div>
        </div>
        
        <div class="field">
          <div class="label">Attachment:</div>
          <div class="value attachment">${fileId} ${fileNameColIndex === null ? '<span class="warning">⚠️ Could not detect PDF column</span>' : ''}</div>
        </div>
        
        <div class="field">
          <div class="label">Subject:</div>
          <div class="value">${subject}</div>
        </div>
        
        <div class="field">
          <div class="label">Body:</div>
          <div class="body-box">${body}</div>
        </div>
        
        <button onclick="google.script.host.close()">Close Preview</button>
      </div>
    </body>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(650).setHeight(500),
    `Preview - Row ${rowNumber}`
  );
}


// ---------------- Send Emails (Mapping-driven Email/FileID/Dynamic Mapping) ----------------
function sendCertificatesEmail() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const mappingProp = props.getProperty("MAPPING");
  const folderId = props.getProperty("FOLDER_ID");
  const subjectTpl = props.getProperty("EMAIL_SUBJECT") || "Your Certificate";
  const bodyTpl = props.getProperty("EMAIL_BODY") || "Dear {{value1}},\n\nPlease find your certificate attached.";

  if (!mappingProp || !folderId) {
    return ui.alert("Mapping or Folder ID not found. Please run 'Setup Template & Folder' first.");
  }

  const mapping = JSON.parse(mappingProp);
  const folder = DriveApp.getFolderById(folderId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find "Hantar" column
  const sentColIndex = headers.indexOf("Hantar");
  if (sentColIndex === -1) return ui.alert("Column 'Sent' not found!");

  // Auto-detect Email and FileID (PDF Name) columns
  let emailColIndex = null;
  let fileNameColIndex = null;
  
  if (data.length > 1) {
    const firstRow = data[1];
    
    for (const dKey in mapping) {
      const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
      if (colIndex >= firstRow.length) continue;
      
      const value = String(firstRow[colIndex] || "");
      
      // Detect email
      if (value.includes("@") && value.includes(".")) {
        emailColIndex = colIndex;
      }
      // Detect PDF filename (starts with DOC_ and ends with .pdf)
      else if (value.startsWith("DOC_") && value.endsWith(".pdf")) {
        fileNameColIndex = colIndex;
      }
    }
  }

  // Fallback to header names if auto-detect fails
  if (emailColIndex === null) {
    emailColIndex = headers.findIndex(h => /email|e-mail|emel/i.test(String(h)));
  }
  if (fileNameColIndex === null) {
    fileNameColIndex = headers.findIndex(h => /fileid|file.*name|pdf/i.test(String(h)));
  }

  if (emailColIndex === null || fileNameColIndex === null) {
    return ui.alert(
      `Could not detect columns:\n` +
      `Email: ${emailColIndex !== null ? 'Found' : 'NOT FOUND'}\n` +
      `PDF Filename: ${fileNameColIndex !== null ? 'Found' : 'NOT FOUND'}\n\n` +
      `Please ensure you have run "Generate Merge Doc" first.`
    );
  }

  let sentCount = 0;
  let errorRows = [];

  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;
    const row = data[i];

    // Skip already sent
    if (row[sentColIndex] && !row[sentColIndex].toString().startsWith("ERR")) continue;
    if (!row[1]) continue;

    try {
      const email = row[emailColIndex];
      const pdfName = row[fileNameColIndex];
      
      if (!email) throw new Error("Email missing");
      if (!pdfName) throw new Error("PDF filename missing");

      // Find PDF file by name in the folder
      const files = folder.getFilesByName(pdfName);
      if (!files.hasNext()) {
        throw new Error(`PDF not found in Drive: ${pdfName}`);
      }
      const pdfFile = files.next().getAs(MimeType.PDF);

      // Replace placeholders in subject/body
      const subject = Object.keys(mapping).reduce((acc, dKey) => {
        const placeholder = mapping[dKey].replace(/[{}]/g, "");
        const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
        const value = row[colIndex] ?? "";
        return acc.replace(new RegExp("\\{\\{\\s*" + placeholder + "\\s*\\}\\}", "g"), value);
      }, subjectTpl);

      const body = Object.keys(mapping).reduce((acc, dKey) => {
        const placeholder = mapping[dKey].replace(/[{}]/g, "");
        const colIndex = parseInt(dKey.replace("dValue", ""), 10) - 1;
        const value = row[colIndex] ?? "";
        return acc.replace(new RegExp("\\{\\{\\s*" + placeholder + "\\s*\\}\\}", "g"), value);
      }, bodyTpl);

      // Send email
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body.replace(/\n/g, "<br>"),
        attachments: [pdfFile]
      });

      // Mark as sent
      sheet.getRange(rowNumber, sentColIndex + 1).setValue(new Date());
      sentCount++;

    } catch (e) {
      sheet.getRange(rowNumber, sentColIndex + 1).setValue("ERR: " + e.message);
      errorRows.push(`Row ${rowNumber}: ${e.message}`);
      Logger.log("Error on row " + rowNumber + ": " + e);
    }
  }

  let msg = `Email sending complete!\nTotal sent: ${sentCount}`;
  if (errorRows.length) msg += "\n\nErrors:\n" + errorRows.join("\n");
  ui.alert(msg);
}
