# Email Sending Workflow Automation

A Google Apps Script solution for sending automated, data-driven email notifications based on user selections in a Google Sheet.

This script monitors a sheet for changes. It dynamically populates dependent dropdowns (e.g., Region -> Country -> Branch) from a central contact list. When a user makes a final selection, it automatically finds the correct email address, generates a templated email with data from that row, and sends it.

## Features

* **Dynamic Dependent Dropdowns:** Populates 3-level dropdowns (e.g., Region, Country, Branch) that update based on the previous selection.
* **Central `CONFIG` Object:** Easily configure all sheet names, column numbers, and settings in one place at the top of the script. No hardcoding.
* **Data-Driven Email Automation:** Automatically triggers an email `onEdit` when a final "trigger" column is populated.
* **HTML Table Generation:** Generates a clean HTML table of changes (from "Old Value" to "New Value") to include in the email.
* **Google Drive Attachment Support:** Automatically finds and attaches a file from a Google Drive link specified in a column.
* **Custom Menu:** Adds an "Automation Menu" to your Google Sheet to easily set up the dropdowns.

## Repository Structure
├── Code.gs # The main Google Apps Script file 
└── README.md # This file

## Setup Instructions

### 1. Sheets Required

You need two (2) sheets in your Google Spreadsheet. You can name them whatever you want and update the names in the `CONFIG` object.

* **Monitoring Sheet (Default: "Task_Monitoring"):** This is the main sheet where you track data and trigger emails.
    * **Data Columns:** You need columns for your "before" and "after" data (e.g., "Old Value 1", "New Value 1", "Old Value 2", "New Value 2", etc.).
    * **Key ID Column:** A column with a unique ID (e.g., "Request-ID" or "Task-1001").
    * **Dropdown Columns:** Three columns for the dynamic dropdowns (e.g., "Region", "Country", "Branch").
    * **Status Column:** A column where the script will write "Sent" or "Error."
    * **Attachment Column:** (Optional) A column containing a full Google Drive URL to a file.
* **Contact List Sheet (Default: "Contact_List"):** This sheet acts as your database for the dropdowns and emails.
    * **Column 1:** Region (e.g., "AMER", "APAC", "EMEA")
    * **Column 2:** Country (e.g., "USA", "Japan", "Germany")
    * **Column 3:** Branch (e.g., "New York", "Tokyo", "Berlin")
    * **Column 4:** Email Address (e.g., "team@example.com")

### 2. Configure the Script

1.  Copy the contents of `Code.gs` into the Apps Script editor of your Google Sheet.
2.  At the top of `Code.gs`, **update the `CONFIG` object** to match your sheet names and column numbers (A=1, B=2, C=3, etc.). This is the most important step.
    ```javascript
    const CONFIG = {
      SHEET_MONITORING: "Task_Monitoring",
      SHEET_CONTACTS: "Contact_List",
      
      MONITOR_CONFIG: {
        START_ROW: 3,         // First row of data in Monitoring sheet
        REGION_COL: 57,       // Column for "Region" dropdown
        COUNTRY_COL: 58,      // Column for "Country" dropdown
        BRANCH_COL: 59,       // Column for "Branch" (trigger)
        STATUS_COL: 60,       // Column for "Sent" status
        DATA_START_COL: 6,    // First data column for email table
        DATA_END_COL: 21,     // Last data column for email table
        ATTACHMENT_COL: 22,   // Column with Google Drive URL
        KEY_DATA_COL: 3       // Column with unique ID (e.g., 'REQ-1001')
      },
      
      // Define the headers for your data pairs
      CHANGE_HEADERS: [
        "Item 1", "Item 2", "Item 3", "Item 4",
        "Item 5", "Item 6", "Item 7", "Item 8"
      ]
    };
    ```

### 3. Authorize the Script

1.  Save the script.
2.  Reload your Google Sheet. A new **"Automation Menu"** will appear.
3.  Click **Automation Menu > Setup Dropdowns**.
4.  A popup will ask for authorization. Grant the script permission to access your Sheets, Drive (for attachments), and Mail services.

## Usage

1.  Fill in your `Contact_List` sheet with your regions, countries, branches, and corresponding email addresses.
2.  Run **Automation Menu > Setup Dropdowns** to apply the validation rules to your `Task_Monitoring` sheet.
3.  In the `Task_Monitoring` sheet, fill in a row with data.
4.  Use the dropdowns to select a **Region**, then **Country**.
5.  When you select a **Branch** (the trigger column), the `onEdit` function will fire.
6.  The script will:
    * Find the correct email from the `Contact_List`.
    * Find all data for that row.
    * Generate an HTML table of changes (based on `CHANGE_HEADERS`).
    * Attach any file from the attachment column.
    * Send the email.
    * Update the "Status" column to "Sent" or "Failed."

## Customization

To change the email content, edit the `generateEmailContent` function inside `Code.gs`. You can create different subjects and bodies based on the data (like `country` or `region`) passed to it.
