/**
 * ===================================================================
 * CONFIGURATION
 * ===================================================================
 * A. Set your preferences here.
 * B. Set up your sheet names and column *numbers* (A=1, B=2, etc.).
 */
const CONFIG = {
  // 1. Sheet Names
  SHEET_MONITORING: "Task_Monitoring",    // Sheet to watch for edits
  SHEET_CONTACTS: "Contact_List",         // Sheet with email contacts

  // 2. Monitoring Sheet Column Numbers
  MONITOR_CONFIG: {
    START_ROW: 3,         // First row of data
    REGION_COL: 57,       // BE: Column for "Region" dropdown
    COUNTRY_COL: 58,      // BF: Column for "Country" dropdown
    BRANCH_COL: 59,       // BG: Column for "Branch" (Email Trigger)
    STATUS_COL: 60,       // BH: Column to write "Sent" status
    
    // Data columns to include in the email table
    DATA_START_COL: 6,    // F: "Originally Made Out" start
    DATA_END_COL: 21,     // U: "To Be Amended To Read" end
    
    ATTACHMENT_COL: 22,   // V: Column with Google Drive URL
    KEY_DATA_COL: 3       // C: Column with unique ID (e.g., 'REQ-1001')
  },

  // 3. Contact List Sheet Column Numbers
  CONTACT_CONFIG: {
    START_ROW: 2,         // First row of contact data
    REGION_COL: 2,        // B: Column with Region (e.g., 'APAC', 'EMEA')
    COUNTRY_COL: 3,       // C: Column with Country (e.g., 'Japan', 'Germany')
    BRANCH_COL: 4,        // D: Column with Branch (e.g., 'Tokyo', 'Berlin')
    EMAIL_COL: 6          // F: Column with Email Address
  },
  
  // 4. Headers for the generated email table
  //    This array MUST have 8 items, matching the 16 columns
  //    defined in MONITOR_CONFIG.DATA_START_COL/DATA_END_COL.
  CHANGE_HEADERS: [
    "Item 1", "Item 2", "Item 3", "Item 4",
    "Item 5", "Item 6", "Item 7", "Item 8"
  ],

  // 5. Generic email signature
  EMAIL_SIGNATURE: `
    <br><br>
    Thank you and best regards,<br>
    <b>[Your Name / Team]</b><br>
    <i>This is an automated message.</i>
  `
};
// ===================================================================
// END OF CONFIGURATION - DO NOT EDIT BELOW THIS LINE
// ===================================================================


/**
 * Creates the "Automation Menu" when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Setup Dropdowns', 'setupDropdowns')
    .addItem('Log Sheet Names', 'logSheetNames')
    .addToUi();
}

/**
 * Sets up the dynamic dropdowns in the Monitoring sheet.
 */
function setupDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monitoringSheet = ss.getSheetByName(CONFIG.SHEET_MONITORING);
  const contactSheet = ss.getSheetByName(CONFIG.SHEET_CONTACTS);

  if (!monitoringSheet || !contactSheet) {
    const errorMsg = `Error: Missing sheets. Expected '${CONFIG.SHEET_MONITORING}' and '${CONFIG.SHEET_CONTACTS}'.`;
    SpreadsheetApp.getUi().alert(errorMsg);
    Logger.log(errorMsg);
    return;
  }

  const contactData = contactSheet.getRange(
    CONFIG.CONTACT_CONFIG.START_ROW,
    CONFIG.CONTACT_CONFIG.REGION_COL,
    contactSheet.getLastRow() - (CONFIG.CONTACT_CONFIG.START_ROW - 1),
    3 // Region, Country, Branch
  ).getValues();
  
  const regions = [...new Set(contactData.map(row => row[0]))].filter(Boolean);

  const monitorEndRow = monitoringSheet.getLastRow();
  if (monitorEndRow < CONFIG.MONITOR_CONFIG.START_ROW) {
    Logger.log("Monitoring sheet is empty, no dropdowns to set.");
    return;
  }
  
  const regionRange = monitoringSheet.getRange(CONFIG.MONITOR_CONFIG.START_ROW, CONFIG.MONITOR_CONFIG.REGION_COL, monitorEndRow - (CONFIG.MONITOR_CONFIG.START_ROW - 1));
  regionRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(regions, true)
    .build());

  // Clear existing validations for dependent dropdowns
  monitoringSheet.getRange(CONFIG.MONITOR_CONFIG.START_ROW, CONFIG.MONITOR_CONFIG.COUNTRY_COL, monitorEndRow - (CONFIG.MONITOR_CONFIG.START_ROW - 1)).clearDataValidations();
  monitoringSheet.getRange(CONFIG.MONITOR_CONFIG.START_ROW, CONFIG.MONITOR_CONFIG.BRANCH_COL, monitorEndRow - (CONFIG.MONITOR_CONFIG.START_ROW - 1)).clearDataValidations();

  // Set headers
  monitoringSheet.getRange(1, CONFIG.MONITOR_CONFIG.REGION_COL).setValue("Region");
  monitoringSheet.getRange(1, CONFIG.MONITOR_CONFIG.COUNTRY_COL).setValue("Country");
  monitoringSheet.getRange(1, CONFIG.MONITOR_CONFIG.BRANCH_COL).setValue("Branch");
  monitoringSheet.getRange(1, CONFIG.MONITOR_CONFIG.STATUS_COL).setValue("Send Status");
  
  SpreadsheetApp.getUi().alert("Dropdown setup complete.");
}

/**
 * Logs all sheet names for debugging.
 */
function logSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name: " + ss.getName());
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  Logger.log("Sheet names in this spreadsheet: " + sheetNames.join(", "));
  SpreadsheetApp.getUi().alert("Sheet names logged. Check View > Logs in Apps Script editor.");
}

/**
 * Main trigger function that runs on any edit.
 */
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Only run on the Monitoring sheet and below the header
  if (sheet.getName() === CONFIG.SHEET_MONITORING && row >= CONFIG.MONITOR_CONFIG.START_ROW) {

    // If Region or Country is edited, update dependent dropdowns
    if (col === CONFIG.MONITOR_CONFIG.REGION_COL || col === CONFIG.MONITOR_CONFIG.COUNTRY_COL) {
      updateDependentDropdowns(row);
    }

    // If the Branch (trigger) column is edited, send the email
    if (col === CONFIG.MONITOR_CONFIG.BRANCH_COL && range.getValue() !== "") {
      Logger.log("Trigger column edited at row " + row + ": " + range.getValue());
      sendAutomatedNotice(row);
    }
  }
}

/**
 * Updates Country and Branch dropdowns based on Region/Country selection.
 */
function updateDependentDropdowns(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monitoringSheet = ss.getSheetByName(CONFIG.SHEET_MONITORING);
  const contactSheet = ss.getSheetByName(CONFIG.SHEET_CONTACTS);

  if (!monitoringSheet || !contactSheet) return;

  const region = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.REGION_COL).getValue();
  const country = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.COUNTRY_COL).getValue();
  
  const contactData = contactSheet.getRange(
    CONFIG.CONTACT_CONFIG.START_ROW,
    CONFIG.CONTACT_CONFIG.REGION_COL,
    contactSheet.getLastRow() - (CONFIG.CONTACT_CONFIG.START_ROW - 1),
    3 // Region, Country, Branch
  ).getValues();

  // Update Country dropdown based on Region
  if (region) {
    const countries = [...new Set(contactData
      .filter(r => r[0] === region)
      .map(r => r[1]))].filter(Boolean);
      
    const countryCell = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.COUNTRY_COL);
    countryCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(countries.length > 0 ? countries : ["N/A"], true)
      .build());
    
    // Clear Branch validation if Region changes
    if (range.getColumn() === CONFIG.MONITOR_CONFIG.REGION_COL) {
        monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.BRANCH_COL).clearDataValidations().clearContent();
    }
  }

  // Update Branch dropdown based on Country
  if (country) {
    const branches = contactData
      .filter(r => r[0] === region && r[1] === country)
      .map(r => r[2])
      .filter(Boolean);
      
    const branchCell = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.BRANCH_COL);
    branchCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(branches.length > 0 ? branches : ["N/A"], true)
      .build());
  }
}

/**
 * Main function to send the automated email notice.
 */
function sendAutomatedNotice(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monitoringSheet = ss.getSheetByName(CONFIG.SHEET_MONITORING);
  const contactSheet = ss.getSheetByName(CONFIG.SHEET_CONTACTS);

  if (!monitoringSheet || !contactSheet) {
    Logger.log("Error: Missing required sheets in sendAutomatedNotice.");
    return;
  }

  const statusCell = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.STATUS_COL);
  statusCell.setValue("Sending...");

  try {
    // Get data from Monitoring sheet
    const keyData = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.KEY_DATA_COL).getValue();
    const region = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.REGION_COL).getValue();
    const country = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.COUNTRY_COL).getValue();
    const branch = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.BRANCH_COL).getValue();
    
    // Get email address from Contact sheet
    const contactData = contactSheet.getRange(
      CONFIG.CONTACT_CONFIG.START_ROW,
      CONFIG.CONTACT_CONFIG.REGION_COL,
      contactSheet.getLastRow() - (CONFIG.CONTACT_CONFIG.START_ROW - 1),
      CONFIG.CONTACT_CONFIG.EMAIL_COL - CONFIG.CONTACT_CONFIG.REGION_COL + 1
    ).getValues();

    let email = "";
    for (let i = 0; i < contactData.length; i++) {
      // Find the row matching Region, Country, and Branch
      if (contactData[i][0] === region && contactData[i][1] === country && contactData[i][2] === branch) {
        email = contactData[i][CONFIG.CONTACT_CONFIG.EMAIL_COL - CONFIG.CONTACT_CONFIG.REGION_COL];
        break;
      }
    }

    if (!email) {
      throw new Error(`No email found for ${region}/${country}/${branch}`);
    }

    // Get the row data for the email table
    const dataWidth = CONFIG.MONITOR_CONFIG.DATA_END_COL - CONFIG.MONITOR_CONFIG.DATA_START_COL + 1;
    const changeData = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.DATA_START_COL, 1, dataWidth).getValues()[0];
    const attachmentUrl = monitoringSheet.getRange(row, CONFIG.MONITOR_CONFIG.ATTACHMENT_COL).getValue();

    // Generate email content
    const tableHtml = generateChangesTable(changeData, CONFIG.CHANGE_HEADERS);
    const emailContent = generateEmailContent(keyData, country, tableHtml);

    let emailOptions = {
      to: email,
      subject: emailContent.subject,
      htmlBody: emailContent.body
    };

    // Add attachment if URL exists
    if (attachmentUrl) {
      try {
        const fileId = extractFileIdFromUrl(attachmentUrl);
        const file = DriveApp.getFileById(fileId);
        emailOptions.attachments = [file.getBlob()];
        Logger.log("Attaching file: " + file.getName());
      } catch (err) {
        Logger.log("Failed to attach file: " + err.message);
        // Continue without attachment
      }
    }

    Logger.log(`Sending email to: ${email}, Subject: ${emailOptions.subject}`);
    MailApp.sendEmail(emailOptions);

    statusCell.setValue("Sent " + new Date().toLocaleString());
  } catch (error) {
    statusCell.setValue("Failed: " + error.message);
    SpreadsheetApp.getUi().alert("Email failed: " + error.message);
    Logger.log("Error in sendAutomatedNotice: " + error.message);
  }
}

/**
 * Generates the subject and body for the email.
 * CUSTOMIZE THIS FUNCTION for your own templates.
 * @param {string} keyData - The unique ID, e.g., "REQ-1001".
 * @param {string} country - The selected country, e.g., "USA".
 * @param {string} tableHtml - The generated HTML table of changes.
 * @returns {object} An object with {subject, body}.
 */
function generateEmailContent(keyData, country, tableHtml) {
  let subject = "";
  let body = "";

  // Example: Create a different template for a specific 'country' or 'region'
  // This logic is generic and can be adapted to any use case
  if (country === "USA") {
    subject = `Request for Review: ${keyData}`;
    body = `
      Dear Team,<br><br>
      Please find below changes for your reference. We have updated the data as per the customer's request.<br><br>
      <b>Reference ID:</b> ${keyData}<br><br>
      <b style="color: red;">Correction Details:</b><br>
      ${tableHtml}
      ${CONFIG.EMAIL_SIGNATURE}
    `;
  } else {
    // Default template for all other countries
    subject = `Data Change Notification: ${keyData}`;
    body = `
      Dear Colleagues,<br><br>
      We have updated the data for ${keyData}. Please find the correction details below:<br><br>
      ${tableHtml}
      ${CONFIG.EMAIL_SIGNATURE}
    `;
  }

  return { subject: subject, htmlBody: body };
}


/**
 * Extracts a Google Drive file ID from a URL.
 * @param {string} url - The Google Drive share URL.
 * @returns {string} The file ID.
 */
function extractFileIdFromUrl(url) {
  const regex = /\/file\/d\/(.+?)\/(?:view|edit|usp=sharing)?/;
  const match = url.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  // Fallback for URLs that might just be the ID
  if (url.length > 40 && !url.includes('/')) {
    return url;
  }
  throw new Error("Invalid Google Drive URL format. URL: " + url);
}

/**
 * Generates an HTML table from a row of data.
 * Assumes data is in pairs [old1, new1, old2, new2, ...]
 * @param {Array} dataRow - The array of data from the sheet.
 * @param {Array} headers - The array of headers (e.g., "Item 1", "Item 2").
 * @returns {string} An HTML table.
 */
function generateChangesTable(dataRow, headers) {
  let table = `
    <table border='1' cellpadding='5' style='border-collapse: collapse;'>
      <tr style="background-color: #f0f0f0;">
        <th>Item</th>
        <th>Old Value</th>
        <th>New Value</th>
      </tr>`;

  // Ensure we don't go out of bounds
  const pairCount = Math.min(headers.length, Math.floor(dataRow.length / 2));

  for (let i = 0; i < pairCount; i++) {
    const header = headers[i];
    const oldVal = dataRow[i * 2] || "";
    const newVal = dataRow[i * 2 + 1] || "";

    // Only add a row if there is data for it
    if (oldVal || newVal) {
      table += `<tr><td><b>${header}</b></td><td>${oldVal}</td><td>${newVal}</td></tr>`;
    }
  }
  
  table += "</table>";
  return table;
}
