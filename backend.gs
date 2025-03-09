/**
 * Enhanced Mail Merge for Google Docs
 * This script creates a sidebar for sending personalized emails directly from Google Docs
 * with advanced features for email account selection, content validation, and more.
 */

"use strict";

// Create a menu when the document is opened
function onOpen() {
  DocumentApp.getUi()
      .createMenu('📧 Mail Merge')
      .addItem('Open Mail Merge Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Shows the mail merge sidebar.
 */
function showSidebar() {
  const ui = HtmlService.createTemplateFromFile('TabMailMergeSidebar')
      .evaluate()
      .setTitle('Mail Merge')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Gets HTML content from a file.
 * @param {string} filename - The name of the HTML file.
 * @return {string} The HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the content of the current document.
 * Consolidated function for document text retrieval.
 * @return {string} The document content as raw text.
 */
function getDocContent() {
  return DocumentApp.getActiveDocument().getBody().getText();
}

/**
 * Gets all available email addresses that the user can send from.
 * @return {Object[]} Array of email addresses and display names.
 */
function getAvailableFromAddresses() {
  const primaryEmail = Session.getActiveUser().getEmail();
  const results = [
    {
      email: primaryEmail,
      name: getUserName(),
      isPrimary: true
    }
  ];
  try {
    const delegatedAddresses = getDelegatedAddresses();
    if (delegatedAddresses && delegatedAddresses.length > 0) {
      delegatedAddresses.forEach(addr => {
        results.push({
          email: addr.email,
          name: addr.name || '',
          isPrimary: false
        });
      });
    }
  } catch (e) {
    Logger.log("Couldn't retrieve delegated addresses: " + e.message);
  }
  return results;
}

/**
 * Gets delegated email addresses from Gmail settings.
 * Note: This is a placeholder function. Actual implementation would require OAuth
 * and Gmail API access.
 * @return {Object[]} Array of delegated email addresses.
 */
function getDelegatedAddresses() {
  return [];
}

/**
 * Gets the user's display name from their Google Account.
 * @return {string} The user's name or email if name is not available.
 */
function getUserName() {
  const email = Session.getActiveUser().getEmail();
  const namePart = email.split('@')[0];
  return formatNameFromEmail(namePart);
}

/**
 * Formats a name from an email address part.
 * @param {string} namePart - The part of the email before @
 * @return {string} Formatted name.
 */
function formatNameFromEmail(namePart) {
  return namePart
    .split(/[._-]/)
    .map(part => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

/**
 * Gets the list of sheets from a spreadsheet.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @return {string[]} Array of sheet names.
 */
function getSheetNames(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    return sheets.map(sheet => sheet.getName());
  } catch (e) {
    return ["Error: " + e.message];
  }
}

/**
 * Gets the column headers from a sheet along with recipient count for the email column.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {string} emailColumn - Optional email column name to count recipients.
 * @return {Object} Object with headers array and recipientCount.
 */
function getColumnsAndRecipientCount(spreadsheetId, sheetName, emailColumn) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0].filter(header => header !== "");
    let recipientCount = 0;
    if (emailColumn && values.length > 1) {
      const emailColIndex = headers.indexOf(emailColumn);
      if (emailColIndex !== -1) {
        for (let i = 1; i < values.length; i++) {
          if (values[i][emailColIndex] && String(values[i][emailColIndex]).trim() !== '') {
            recipientCount++;
          }
        }
      }
    }
    return {
      headers: headers,
      recipientCount: recipientCount
    };
  } catch (e) {
    return {
      headers: ["Error: " + e.message],
      recipientCount: 0
    };
  }
}

/**
 * Gets the column headers from a sheet.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @return {string[]} Array of column headers.
 */
function getColumnHeaders(spreadsheetId, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.filter(header => header !== "");
  } catch (e) {
    return ["Error: " + e.message];
  }
}

/**
 * Gets the recipient count for a specific column.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {string} emailColumn - The column name containing email addresses.
 * @return {number} The number of recipients.
 */
function getRecipientCount(spreadsheetId, sheetName, emailColumn) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIndex = headers.indexOf(emailColumn);
    if (emailColIndex === -1) {
      throw new Error("Email column not found: " + emailColumn);
    }
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailColIndex] && String(data[i][emailColIndex]).trim() !== '') {
        count++;
      }
    }
    return count;
  } catch (e) {
    Logger.log("Error getting recipient count: " + e.message);
    return 0;
  }
}

/**
 * Extracts a spreadsheet ID from a URL or returns the ID directly.
 * @param {string} spreadsheetUrl - The spreadsheet URL or ID.
 * @return {string} The extracted spreadsheet ID.
 */
function extractSpreadsheetId(spreadsheetUrl) {
  if (/^[a-zA-Z0-9_-]{40,}$/.test(spreadsheetUrl)) {
    return spreadsheetUrl;
  }
  const matches = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (matches && matches[1]) {
    return matches[1];
  }
  return spreadsheetUrl;
}

/**
 * Gets the spreadsheet data.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @return {Object} An object with headers and rows.
 */
function getSpreadsheetData(spreadsheetId, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    const range = sheet.getDataRange();
    const displayValues = range.getDisplayValues();
    const headers = displayValues[0];
    const rows = displayValues.slice(1);
    return {
      headers: headers,
      rows: rows
    };
  } catch (e) {
    throw new Error("Error getting spreadsheet data: " + e.message);
  }
}

/**
 * Gets a specific row from a spreadsheet.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {number} rowIndex - The 0-based row index (0 returns the first data row).
 * @return {Object} An object mapping column names to values.
 */
function getSpreadsheetRow(spreadsheetId, sheetName, rowIndex) {
  try {
    const data = getSpreadsheetData(spreadsheetId, sheetName);
    if (rowIndex >= data.rows.length) {
      throw new Error(`Row index ${rowIndex} is out of bounds. Max index is ${data.rows.length - 1}`);
    }
    const row = data.rows[rowIndex];
    const result = {};
    for (let i = 0; i < data.headers.length; i++) {
      if (data.headers[i]) {
        result[data.headers[i]] = row[i];
      }
    }
    return result;
  } catch (e) {
    throw new Error("Error getting spreadsheet row: " + e.message);
  }
}

/**
 * Analyzes document content and validates placeholders against spreadsheet headers.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @return {Object} Validation results.
 */
function validatePlaceholders(spreadsheetId, sheetName) {
  try {
    const docText = getDocContent();
    const placeholderRegex = /\{\{([^{}]+)\}\}/g;
    const docPlaceholders = [];
    let match;
    while ((match = placeholderRegex.exec(docText)) !== null) {
      docPlaceholders.push(match[1].trim());
    }
    const uniqueDocPlaceholders = [...new Set(docPlaceholders)];
    const headers = getColumnHeaders(spreadsheetId, sheetName);
    const matched = [];
    const unmatched = [];
    const unused = [];
    uniqueDocPlaceholders.forEach(placeholder => {
      if (headers.includes(placeholder)) {
        matched.push(placeholder);
      } else {
        unmatched.push(placeholder);
      }
    });
    headers.forEach(header => {
      if (!uniqueDocPlaceholders.includes(header)) {
        unused.push(header);
      }
    });
    const exampleValues = {};
    if (matched.length > 0) {
      try {
        const data = getSpreadsheetData(spreadsheetId, sheetName);
        if (data.rows.length > 0) {
          const firstRow = data.rows[0];
          matched.forEach(placeholder => {
            const index = data.headers.indexOf(placeholder);
            if (index !== -1) {
              exampleValues[placeholder] = firstRow[index];
            }
          });
        }
      } catch (e) {
        Logger.log("Error getting example values: " + e.message);
      }
    }
    return {
      success: true,
      matched: matched,
      unmatched: unmatched,
      unused: unused,
      examples: exampleValues
    };
  } catch (e) {
    return {
      success: false,
      message: "Error validating placeholders: " + e.message
    };
  }
}

/**
 * Replaces placeholders in the text with values from the data row.
 * @param {string} text - The template text.
 * @param {string[]} headers - The column headers.
 * @param {string[]} row - The data row.
 * @return {string} The text with placeholders replaced.
 */
function replacePlaceholders(text, headers, row) {
  let result = text;
  for (let i = 0; i < headers.length; i++) {
    const placeholder = '{{' + headers[i] + '}}';
    const value = row[i] !== undefined && row[i] !== null ? row[i].toString() : '';
    result = result.split(placeholder).join(value);
  }
  return result;
}

/**
 * Prepares email options for sending emails.
 * @param {string} fromEmail - The sender's email address.
 * @param {string} fromName - The sender's display name.
 * @param {Object} options - Additional options like cc and bcc.
 * @return {Object} Email options object.
 */
function prepareEmailOptions(fromEmail, fromName, options = {}) {
  const emailOptions = { name: fromName || undefined };
  if (options.cc) {
    emailOptions.cc = options.cc;
  }
  if (options.bcc) {
    emailOptions.bcc = options.bcc;
  }
  if (fromEmail && fromEmail !== Session.getActiveUser().getEmail()) {
    emailOptions.from = fromEmail;
  }
  return emailOptions;
}

/**
 * Logs mail merge actions to a spreadsheet if logging is enabled.
 * @param {Sheet|null} logSheet - The log sheet or null if logging is disabled.
 * @param {string} email - The recipient email address.
 * @param {string} status - The status of the email (Success/Error).
 * @param {string} message - Additional information about the status.
 */
function logMailMergeAction(logSheet, email, status, message) {
  if (!logSheet) return;
  logSheet.appendRow([ new Date(), email, status, message ]);
}

/**
 * Creates a log sheet for mail merge operations.
 * @return {Sheet} The log sheet.
 */
function createOrGetLogSheet() {
  const ss = SpreadsheetApp.create('Mail Merge Log - ' + new Date().toISOString().split('T')[0]);
  const sheet = ss.getActiveSheet();
  sheet.setName('Mail Merge Log');
  sheet.appendRow(['Timestamp', 'Email', 'Status', 'Message']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  return sheet;
}

/**
 * Processes a single row for mail merge.
 * @param {Object} params - Parameters containing all necessary data for processing a row.
 * @return {Object} Result object with success flag and error message if any.
 */
function processMailMergeRow(params) {
  const { emailAddress, subjectLine, templateHtml, headers, row, emailOptions, createDrafts } = params;
  try {
    if (!emailAddress) {
      return { success: false, message: 'No email address provided' };
    }
    const personalizedSubject = replacePlaceholders(subjectLine, headers, row);
    const personalizedBody = replacePlaceholders(templateHtml, headers, row);
    const rowEmailOptions = Object.assign({}, emailOptions, { htmlBody: personalizedBody });
    if (createDrafts) {
      GmailApp.createDraft(emailAddress, personalizedSubject, "", rowEmailOptions);
    } else {
      GmailApp.sendEmail(emailAddress, personalizedSubject, "", rowEmailOptions);
    }
    return { success: true, message: createDrafts ? 'Draft created' : 'Email sent' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Sends a test email.
 * @param {string} recipients - Comma-separated list of email addresses.
 * @param {string} subject - The email subject.
 * @param {string} fromEmail - The sender's email address.
 * @param {string} fromName - The sender's display name.
 * @param {string} cc - Optional CC addresses.
 * @param {string} bcc - Optional BCC addresses.
 * @param {Object} options - Additional options for test email.
 * @return {Object} Status object with success flag and message.
 */
function sendTestEmailWithData(recipients, subject, fromEmail, fromName, cc, bcc, options) {
  try {
    Logger.log('Sending test email with parameters:');
    Logger.log('Recipients: ' + recipients);
    Logger.log('Subject: ' + subject);
    Logger.log('From: ' + fromEmail);
    const body = getDocContent();
    let htmlBody = body;
    try {
      if (options && options.replacePlaceholders && options.spreadsheetId && options.sheetName) {
        Logger.log('Attempting to replace placeholders');
        const data = getSpreadsheetData(options.spreadsheetId, options.sheetName);
        if (data.rows.length > 0) {
          const firstRow = data.rows[0];
          subject = replacePlaceholders(subject, data.headers, firstRow);
          htmlBody = replacePlaceholders(body, data.headers, firstRow);
        }
      }
    } catch (placeholderError) {
      Logger.log('Error replacing placeholders: ' + placeholderError.message);
    }
    const emailList = recipients.split(',').map(email => email.trim());
    const emailOptions = prepareEmailOptions(fromEmail, fromName, { cc, bcc });
    emailOptions.htmlBody = htmlBody;
    Logger.log('Sending email with options: ' + JSON.stringify(emailOptions));
    for (const email of emailList) {
      if (email) {
        GmailApp.sendEmail(email, subject, "", emailOptions);
      }
    }
    Logger.log('Test email sent successfully');
    return {
      success: true,
      message: `Test email sent to: ${recipients}`
    };
  } catch (e) {
    Logger.log('Error in sendTestEmailWithData: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return {
      success: false,
      message: "Error sending test email: " + e.message
    };
  }
}

/**
 * Executes the mail merge.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {string} emailColumn - The column containing email addresses.
 * @param {string} subjectLine - The email subject.
 * @param {string} fromEmail - The sender's email address.
 * @param {string} fromName - The sender's display name.
 * @param {Object} options - Additional options.
 * @return {Object} Status object with success flag, message, and counts.
 */
function executeMailMerge(spreadsheetId, sheetName, emailColumn, subjectLine, fromEmail, fromName, options = {}) {
  try {
    options = {
      cc: '',
      bcc: '',
      enableLogging: false,
      createDrafts: false,
      ...options
    };
    let logSheet = null;
    if (options.enableLogging) {
      logSheet = createOrGetLogSheet();
    }
    const templateHtml = getDocContent();
    const data = getSpreadsheetData(spreadsheetId, sheetName);
    const headers = data.headers;
    const rows = data.rows;
    const emailIndex = headers.indexOf(emailColumn);
    if (emailIndex === -1) {
      throw new Error(`Email column "${emailColumn}" not found in spreadsheet`);
    }
    const emailOptions = prepareEmailOptions(fromEmail, fromName, options);
    let sentCount = 0;
    let errorCount = 0;
    let errorEmails = [];
    for (const row of rows) {
      const emailAddress = row[emailIndex];
      const result = processMailMergeRow({
        emailAddress,
        subjectLine,
        templateHtml,
        headers,
        row,
        emailOptions,
        createDrafts: options.createDrafts
      });
      if (result.success) {
        sentCount++;
        logMailMergeAction(logSheet, emailAddress, 'Success', result.message);
      } else {
        errorCount++;
        errorEmails.push(emailAddress);
        logMailMergeAction(logSheet, emailAddress, 'Error', result.message);
        Logger.log(`Error sending to ${emailAddress}: ${result.message}`);
      }
      Utilities.sleep(100);
    }
    return {
      success: true,
      message: options.createDrafts 
        ? `Mail merge complete. Drafts created: ${sentCount}, Errors: ${errorCount}`
        : `Mail merge complete. Sent: ${sentCount}, Errors: ${errorCount}`,
      sent: sentCount,
      errors: errorCount,
      errorEmails: errorEmails.join(", ")
    };
  } catch (e) {
    return {
      success: false,
      message: "Error executing mail merge: " + e.message,
      sent: 0,
      errors: 0
    };
  }
}

/**
 * Generates a preview of an email with merged data.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {number} rowIndex - The row index to use for the preview.
 * @return {Object} Status object with success flag, message, and preview HTML.
 */
function generateEmailPreview(spreadsheetId, sheetName, rowIndex) {
  try {
    const templateHtml = getDocContent();
    const data = getSpreadsheetData(spreadsheetId, sheetName);
    if (rowIndex >= data.rows.length) {
      throw new Error(`Row index ${rowIndex} is out of bounds. Max index is ${data.rows.length - 1}`);
    }
    const row = data.rows[rowIndex];
    const personalizedBody = replacePlaceholders(templateHtml, data.headers, row);
    return {
      success: true,
      message: "Preview generated successfully",
      previewHtml: personalizedBody
    };
  } catch (e) {
    return {
      success: false,
      message: "Error generating preview: " + e.message
    };
  }
}

/**
 * Validates a spreadsheet ID/URL by trying to open it.
 * @param {string} spreadsheetUrl - The spreadsheet URL or ID.
 * @return {Object} Validation result with success flag and ID if successful.
 */
function validateSpreadsheet(spreadsheetUrl) {
  try {
    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return {
      success: true,
      id: spreadsheetId,
      name: spreadsheet.getName()
    };
  } catch (e) {
    return {
      success: false,
      message: "Invalid spreadsheet: " + e.message
    };
  }
}
