/**
 * Enhanced Mail Merge for Google Docs
 * This script creates a sidebar for sending personalized emails directly from Google Docs
 * with advanced features for email account selection, content validation, and more.
 */

"use strict";

/**
 * Gets HTML content from a file.
 * @param {string} filename - The name of the HTML file.
 * @return {string} The HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the content of the current document, extracting only the template content
 * between the template markers if they exist.
 * @return {string} The template content as raw text.
 */
function getDocContent() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const fullText = body.getText();
  
  // Check for template markers
  const contentStartMarker = "===TEMPLATE CONTENT===";
  const contentEndMarker = "===TEMPLATE END===";
  
  const startIndex = fullText.indexOf(contentStartMarker);
  const endIndex = fullText.indexOf(contentEndMarker);
  
  // If both markers exist, extract only the content between them
  if (startIndex !== -1 && endIndex !== -1 && startIndex < endIndex) {
    // Extract the content between markers (add marker length to get position after marker)
    const contentStartPos = startIndex + contentStartMarker.length;
    return fullText.substring(contentStartPos, endIndex).trim();
  }
  
  // Fallback: If markers not found, return the full document content (for backward compatibility)
  return fullText;
}

/**
 * Gets the configuration section from the document if it exists
 * @return {Object} Extracted configuration values or null if not found
 */
function getDocConfig() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const fullText = body.getText();
  
  // Check for config markers
  const configStartMarker = "===TEMPLATE CONFIG===";
  const configEndMarker = "===TEMPLATE CONTENT===";
  
  const startIndex = fullText.indexOf(configStartMarker);
  const endIndex = fullText.indexOf(configEndMarker);
  
  // If both markers exist, extract the configuration
  if (startIndex !== -1 && endIndex !== -1 && startIndex < endIndex) {
    // Extract the config section
    const configStartPos = startIndex + configStartMarker.length;
    const configText = fullText.substring(configStartPos, endIndex).trim();
    
    // Parse the configuration
    const config = {};
    const lines = configText.split('\n');
    
    for (const line of lines) {
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0) {
        const key = line.substring(0, colonIndex).trim();
        const value = line.substring(colonIndex + 1).trim();
        
        // Handle required fields specially
        if (key === "Required Fields") {
          config.requiredFields = value.split(',').map(field => field.trim());
        } else {
          config[key] = value;
        }
      }
    }
    
    return config;
  }
  
  // No configuration section found
  return null;
}

/**
 * Sets document content with proper template boundaries
 * @param {string} content - The template content to set
 * @param {Object} config - Optional configuration to include
 * @return {boolean} Success status
 */
function setDocContent(content, config = null) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    // Clear existing content
    body.clear();
    
    // Add configuration section with clear visual boundaries
    if (config) {
      // Add config header with distinct formatting
      const configHeaderPara = body.appendParagraph("===TEMPLATE CONFIG===");
      configHeaderPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      configHeaderPara.setBackgroundColor("#e6f2ff");
      
      // Add separator line
      const separatorPara = body.appendParagraph("----------------------------------------");
      separatorPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      // Add all configuration properties with bold attribute names
      for (const [key, value] of Object.entries(config)) {
        if (key === 'requiredFields' && Array.isArray(value)) {
          const para = body.appendParagraph("");
          const text = para.appendText(key + ": ");
          text.setBold(true);
          para.appendText(value.join(', '));
        } 
        else if (key !== 'documentContent') {
          const para = body.appendParagraph("");
          const text = para.appendText(key + ": ");
          text.setBold(true);
          para.appendText(String(value));
        }
      }
      
      // Add bottom separator
      body.appendParagraph("----------------------------------------")
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
    
    // Add content section with clear visual marker
    const contentHeaderPara = body.appendParagraph("===TEMPLATE CONTENT===");
    contentHeaderPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contentHeaderPara.setBackgroundColor("#f0f7e6");
    
    // Add actual content
    body.appendParagraph(content);
    
    // Add end marker
    const endMarkerPara = body.appendParagraph("===TEMPLATE END===");
    endMarkerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    endMarkerPara.setBackgroundColor("#fff2e6");
    
    return true;
  } catch (e) {
    Logger.log("Error setting document content: " + e.message);
    return false;
  }
}

/**
 * Gets all available email addresses that the user can send from.
 * Updated with better error handling for permission issues.
 * @return {Object[]} Array of email addresses and display names.
 */
function getAvailableFromAddresses() {
  try {
    // Try to get the active user's email
    const primaryEmail = Session.getActiveUser().getEmail();
    
    // Check if we actually got an email (empty means no permission)
    if (!primaryEmail) {
      throw new Error("Unable to get user email. Check permissions.");
    }
    
    const results = [
      {
        email: primaryEmail,
        name: getUserName(),
        isPrimary: true
      }
    ];
    
    // Still try to get delegated addresses, but handle errors better
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
      // Just continue with primary email
    }
    
    return results;
  } catch (e) {
    // If we can't get the user's email at all, return a fallback response
    Logger.log("Error getting user email: " + e.message);
    return [
      {
        email: "",
        name: "Default Sender",
        isPrimary: true,
        error: "Permission denied. Authorize the app or contact the administrator."
      }
    ];
  }
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
 * Gets all available email addresses that the user can send from.
 * Ultra-resilient version that works with heavily restricted Gmail features.
 * @return {Object[]} Array of email addresses and display names.
 */
function getAvailableFromAddresses() {
  try {
    // First try to get basic user email - this should work based on your test
    let userEmail = null;
    let userName = null;
    
    try {
      userEmail = Session.getActiveUser().getEmail();
      // If we got an email, try to get a name
      if (userEmail) {
        userName = formatNameFromEmail(userEmail.split('@')[0]);
      }
    } catch (emailError) {
      console.warn("Couldn't get user email:", emailError.message);
    }
    
    // If we couldn't get the email, try to get effective user as fallback
    if (!userEmail) {
      try {
        userEmail = Session.getEffectiveUser().getEmail();
        if (userEmail) {
          userName = formatNameFromEmail(userEmail.split('@')[0]); 
        }
      } catch (effError) {
        console.warn("Couldn't get effective user:", effError.message);
      }
    }
    
    // If we still don't have an email, return a placeholder
    if (!userEmail) {
      return [{
        email: "",
        name: "Mail Merge User",
        isPrimary: true,
        error: "Unable to detect email. Please enter manually."
      }];
    }
    
    // We got a valid email, return just the primary user
    return [{
      email: userEmail,
      name: userName || formatNameFromEmail(userEmail.split('@')[0]),
      isPrimary: true
    }];
    
    // Important: Skip delegation checks completely
    // This appears to be what's failing in your domain environment
    
  } catch (e) {
    console.error("Error in getAvailableFromAddresses:", e.message);
    return [{
      email: "",
      name: "Mail Merge User",
      isPrimary: true,
      error: "Error detecting email address."
    }];
  }
}

/**
 * Gets the user's display name from their Google Account.
 * Now with better error handling.
 * @return {string} The user's name or default value if not available.
 */
function getUserName() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return "Mail Merge User";
    }
    const namePart = email.split('@')[0];
    return formatNameFromEmail(namePart);
  } catch (e) {
    Logger.log("Error getting user name: " + e.message);
    return "Mail Merge User";
  }
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
 * Gets the current email quota information.
 * @return {Object} An object with quota details.
 */
function getEmailQuotaInfo() {
  try {
    // Get remaining quota from MailApp
    const remaining = MailApp.getRemainingDailyQuota();
    
    // Log for debugging
    Logger.log("Remaining email quota: " + remaining);
    
    // Determine account type for quota limit
    let total = 100; // Default for consumer Gmail
    const email = Session.getActiveUser().getEmail();
    
    // Check account type
    if (email && !email.endsWith('@gmail.com')) {
      total = 1500; // Typical Google Workspace limit
    }
    
    // Calculate used quota
    const used = total - remaining;
    
    return {
      remaining: remaining,
      total: total,
      used: used,
      hasPermission: true
    };
  } catch (e) {
    Logger.log("Error getting quota info: " + e.message);
    
    // Check if this is a permissions error
    const isPermissionError = e.message.includes("permissions are not sufficient");
    
    // Return error state
    return {
      hasPermission: false,
      error: e.message,
      isPermissionError: isPermissionError
    };
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
    
    // Try to check quota, but proceed even if there's an error
    let quotaLimited = false;
    let rowsToProcess = rows;
    
    try {
      const remaining = MailApp.getRemainingDailyQuota();
      
      // If we have more recipients than quota, limit to available quota
      if (remaining < rows.length) {
        rowsToProcess = rows.slice(0, remaining);
        quotaLimited = true;
        Logger.log(`Limited mail merge to ${remaining} recipients due to quota restrictions`);
      }
    } catch (quotaError) {
      // If we can't check quota, log the issue but proceed with all rows
      Logger.log(`Could not check quota: ${quotaError.message}. Proceeding with full recipient list.`);
    }
    
    let sentCount = 0;
    let errorCount = 0;
    let errorEmails = [];
    
    for (const row of rowsToProcess) {
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
    
    // Create appropriate message based on quota limitation
    let quotaMessage = "";
    if (quotaLimited) {
      quotaMessage = ` (Limited by quota: ${rowsToProcess.length}/${rows.length})`;
    }
    
    return {
      success: true,
      message: options.createDrafts 
        ? `Mail merge complete. Drafts created: ${sentCount}, Errors: ${errorCount}${quotaMessage}`
        : `Mail merge complete. Sent: ${sentCount}, Errors: ${errorCount}${quotaMessage}`,
      sent: sentCount,
      errors: errorCount,
      errorEmails: errorEmails.join(", "),
      quotaLimited: quotaLimited
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

/**
 * Shows the configuration management dialog.
 */
function showConfigDialog() {
  const ui = HtmlService.createTemplateFromFile('ConfigurationDialog')
      .evaluate()
      .setWidth(500)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  DocumentApp.getUi().showModalDialog(ui, 'Mail Merge Configuration');
}

/**
 * Gets all available configurations from document properties.
 * @return {Object} Configurations object.
 */
function getAvailableConfigurations() {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    return JSON.parse(configsJson);
  } catch (e) {
    Logger.log("Error getting configurations: " + e.message);
    return {};
  }
}

/**
 * Gets current values from the sidebar UI.
 * Used by the configuration dialog to show current values.
 * @return {Object} Current sidebar values.
 */
function getCurrentSidebarValues() {
  try {
    // Create a temporary object to hold the values we can access
    const values = {};
    
    // Try to access document values that are readily available
    const doc = DocumentApp.getActiveDocument();
    if (doc) {
      // Could potentially get subject line from document title or other metadata
      values.subjectLine = doc.getName().replace(/\.docx?$/i, '');
    }
    
    // Get user email if available
    try {
      values.fromEmail = Session.getActiveUser().getEmail();
      // Format display name from email
      if (values.fromEmail) {
        const namePart = values.fromEmail.split('@')[0];
        values.fromName = formatNameFromEmail(namePart);
      }
    } catch (emailError) {
      Logger.log("Couldn't get user email: " + emailError.message);
    }
    
    // At this point, this is basically a placeholder function
    // In a full implementation, you'd need a more complex mechanism to 
    // retrieve the current values from the sidebar
    
    return values;
  } catch (e) {
    Logger.log("Error getting sidebar values: " + e.message);
    return {};
  }
}

/**
 * Saves a configuration with specific values.
 * Modified to support template boundaries and metadata.
 * @param {string} name - The configuration name.
 * @param {Object} values - The values to save.
 * @return {Object} Result with success flag and message.
 */
function saveConfigurationWithValues(name, values) {
  try {
    if (!name) {
      return { success: false, message: "Configuration name is required" };
    }
    
    const docProperties = PropertiesService.getDocumentProperties();
    const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const configs = JSON.parse(configsJson);
    
    // Include document content if requested
    if (values.includeDocumentContent) {
      // Create a configuration object for the document
      const docConfig = {
        Template: name,
        Version: values.templateVersion || '1.0',
        Description: values.templateDescription || '',
        Spreadsheet: values.spreadsheetUrl ? shortenUrl(values.spreadsheetUrl) : '',
        Sheet: values.sheetName || '',
        'Email Column': values.emailColumn || '',
        Subject: values.subjectLine || '',
        'Required Fields': values.requiredFields ? values.requiredFields.join(', ') : '',
        'Last Updated': new Date().toLocaleDateString()
      };
      
      // First clear the document and add the config and markers
      const doc = DocumentApp.getActiveDocument();
      const body = doc.getBody();
      body.clear();
      
      // Add configuration section
      let configText = "===TEMPLATE CONFIG===\n";
      for (const [key, value] of Object.entries(docConfig)) {
        if (value) configText += `${key}: ${value}\n`;
      }
      body.appendParagraph(configText);
      
      // Get the original document content
      const origContent = getOriginalDocContent();
      
      // Add content markers and the content
      body.appendParagraph("===TEMPLATE CONTENT===");
      body.appendParagraph(origContent || "");
      body.appendParagraph("===TEMPLATE END===");
      
      // Now capture the complete document with markers
      values.documentContent = DocumentApp.getActiveDocument().getBody().getText();
      values.documentContentTimestamp = new Date().toISOString();
      values.documentName = DocumentApp.getActiveDocument().getName();
    }
    
    // Add to configs
    configs[name] = values;
    
    // Save back to document properties
    docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
    
    return { success: true, message: 'Configuration saved successfully!' };
  } catch (e) {
    Logger.log("Error saving configuration: " + e.message);
    return { success: false, message: 'Error saving configuration: ' + e.message };
  }
}

/**
 * Gets the original document content before adding template markers.
 * Used to preserve content when adding configuration.
 * @return {string} The original document content.
 */
function getOriginalDocContent() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const text = body.getText();
    
    // Check if document already has template markers
    const contentStartMarker = "===TEMPLATE CONTENT===";
    const contentEndMarker = "===TEMPLATE END===";
    
    const startIndex = text.indexOf(contentStartMarker);
    const endIndex = text.indexOf(contentEndMarker);
    
    // If both markers exist, extract only the content between them
    if (startIndex !== -1 && endIndex !== -1 && startIndex < endIndex) {
      // Get position after start marker plus newline
      const contentStartPos = startIndex + contentStartMarker.length;
      
      // Extract ONLY the content, make sure we're not including any CONFIG markers
      const content = text.substring(contentStartPos, endIndex).trim();
      
      // Check if content contains another CONFIG marker (which would cause duplication)
      if (content.includes("===TEMPLATE CONFIG===")) {
        // Return only the portion after the last CONFIG marker
        const lastConfigIndex = content.lastIndexOf("===TEMPLATE CONFIG===");
        const lastContentIndex = content.lastIndexOf("===TEMPLATE CONTENT===");
        
        if (lastContentIndex > lastConfigIndex) {
          // If there's a CONTENT marker after the CONFIG marker, return text after CONTENT marker
          return content.substring(lastContentIndex + "===TEMPLATE CONTENT===".length).trim();
        } else {
          // Just return empty string to avoid duplication
          return "";
        }
      }
      
      return content;
    }
    
    // If no markers, return the full content
    return text;
  } catch (e) {
    Logger.log("Error getting original document content: " + e.message);
    return "";
  }
}

/**
 * Shortens a URL for display in configuration
 * @param {string} url - The URL to shorten
 * @return {string} Shortened URL or original if not recognized
 */
function shortenUrl(url) {
  if (!url) return '';
  
  // If it's a Google Sheets URL, shorten it
  if (url.includes('docs.google.com/spreadsheets')) {
    const parts = url.split('/');
    // Get the file name if possible, otherwise just the ID
    let fileName = url.match(/\/[^\/]+-([^\/]+)\//) || [];
    if (fileName.length > 1) {
      return fileName[1] + ' (Google Sheet)';
    }
    // Return just the ID portion
    for (let i = 0; i < parts.length; i++) {
      if (parts[i] === 'd' && i+1 < parts.length) {
        return parts[i+1].substring(0, 10) + '... (Google Sheet)';
      }
    }
  }
  
  // Otherwise return truncated URL
  return url.length > 40 ? url.substring(0, 37) + '...' : url;
}

/**
 * Modified function to load a configuration and respect template boundaries.
 * @param {string} name - The configuration name.
 * @param {boolean} loadDocumentContent - Whether to load document content.
 * @return {Object} Result with success flag and loaded config.
 */
function loadConfiguration(name, loadDocumentContent = false) {
  try {
    const configs = getAvailableConfigurations();
    const config = configs[name];
    
    if (!config) {
      return { success: false, message: 'Configuration not found' };
    }
    
    // Store active configuration
    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.setProperty('activeMailMergeConfig', JSON.stringify(config));
    
    // Handle document content if present and requested
    let documentContentLoaded = false;
    if (config.documentContent && loadDocumentContent) {
      // Check if the saved content has template markers
      if (config.documentContent.includes("===TEMPLATE CONTENT===") && 
          config.documentContent.includes("===TEMPLATE END===")) {
        
        // Set the document content directly with markers
        const doc = DocumentApp.getActiveDocument();
        const body = doc.getBody();
        body.clear();
        body.setText(config.documentContent);
        documentContentLoaded = true;
      } else {
        // Legacy content without markers - wrap it with markers
        documentContentLoaded = setDocContent(config.documentContent, {
          Template: name,
          Version: config.templateVersion || '1.0',
          Description: config.templateDescription || '',
          Spreadsheet: config.spreadsheetUrl ? shortenUrl(config.spreadsheetUrl) : '',
          Sheet: config.sheetName || '',
          'Email Column': config.emailColumn || '',
          Subject: config.subjectLine || '',
          'Required Fields': config.requiredFields ? config.requiredFields.join(', ') : '',
          'Last Updated': new Date().toLocaleDateString()
        });
      }
      
      if (!documentContentLoaded) {
        return { 
          success: false, 
          message: 'Error loading document content. Settings were not applied.' 
        };
      }
    }
    
    return { 
      success: true, 
      config: config,
      documentContentLoaded: documentContentLoaded
    };
  } catch (e) {
    Logger.log("Error loading configuration: " + e.message);
    return { success: false, message: 'Error loading configuration: ' + e.message };
  }
}

/**
 * Deletes a configuration.
 * @param {string} name - The configuration name.
 * @return {Object} Result with success flag and message.
 */
function deleteConfiguration(name) {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const configs = JSON.parse(configsJson);
    
    if (!configs[name]) {
      return { success: false, message: 'Configuration not found' };
    }
    
    delete configs[name];
    docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
    
    return { success: true, message: 'Configuration deleted successfully' };
  } catch (e) {
    Logger.log("Error deleting configuration: " + e.message);
    return { success: false, message: 'Error deleting configuration: ' + e.message };
  }
}

/**
 * Add menu items to onOpen function.
 * Modify the existing onOpen() function in backend.gs to add the configuration menu item:
 */
function onOpen() {
  DocumentApp.getUi()
      .createMenu('📧 Mail Merge')
      .addItem('Open Mail Merge Sidebar', 'showSidebar')
      .addSeparator()
      .addItem('Manage Configurations...', 'showConfigDialog')
      .addToUi();
}

/**
 * Shows the configuration management dialog with improved size and handling.
 */
function showConfigDialog() {
  const ui = HtmlService.createTemplateFromFile('ConfigurationDialog')
      .evaluate()
      .setWidth(550)  // Increased width to prevent cutoff
      .setHeight(650) // Increased height to accommodate the details view
      .setTitle('Mail Merge Configuration')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  DocumentApp.getUi().showModalDialog(ui, 'Mail Merge Configuration');
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
 * Sets a flag to indicate configurations have been updated.
 * This will be checked when the dialog closes to refresh the sidebar.
 */
function setConfigurationRefreshFlag() {
  PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
}

/**
 * Checks if configurations have been updated and need refreshing.
 * @return {boolean} True if the sidebar should refresh configurations.
 */
function checkConfigurationsNeedRefresh() {
  const needsRefresh = PropertiesService.getUserProperties().getProperty('configurationUpdated') === 'true';
  
  // Clear the flag if it was set
  if (needsRefresh) {
    PropertiesService.getUserProperties().deleteProperty('configurationUpdated');
  }
  
  return needsRefresh;
}