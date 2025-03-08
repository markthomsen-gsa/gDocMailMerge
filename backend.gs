/**
 * Enhanced Mail Merge for Google Docs
 * This script creates a sidebar for sending personalized emails directly from Google Docs
 * with advanced features for email account selection, content validation, and more.
 */

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
   * @return {string} The document content as HTML.
   */
  function getDocContent() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    // Get the document as HTML
    const bodyText = body.getText();
    
    // Create a new clean HTML document to simulate simple HTML from the Google Doc
    let html = '';
    let currentParagraph = '';
    
    for (let i = 0; i < bodyText.length; i++) {
      if (bodyText[i] === '\n' || i === bodyText.length - 1) {
        // Add the last character if we're at the end
        if (i === bodyText.length - 1) {
          currentParagraph += bodyText[i];
        }
        
        // Only add paragraph if not empty
        if (currentParagraph.trim() !== '') {
          html += '<p>' + currentParagraph + '</p>\n';
        }
        currentParagraph = '';
      } else {
        currentParagraph += bodyText[i];
      }
    }
    
    return html;
  }
  
  /**
   * Gets the raw text content of the current document.
   * @return {string} The document content as plain text.
   */
  function getDocPlainText() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    return body.getText();
  }
  
  /**
   * Gets the subject line from the document (assumes first line is subject).
   * @return {string} The subject line.
   */
  function getSubjectLine() {
    const doc = DocumentApp.getActiveDocument();
    const paragraphs = doc.getBody().getParagraphs();
    
    if (paragraphs.length > 0) {
      return paragraphs[0].getText();
    }
    
    return "No subject";
  }
  
  /**
   * Gets all available email addresses that the user can send from.
   * @return {Object[]} Array of email addresses and display names.
   */
  function getAvailableFromAddresses() {
    // Get the user's primary email address
    const primaryEmail = Session.getActiveUser().getEmail();
    
    // Create a result array with primary email
    const results = [
      {
        email: primaryEmail,
        name: getUserName(),
        isPrimary: true
      }
    ];
    
    // Try to get delegated addresses using Gmail API (if available and authorized)
    try {
      // This is a simplified example. In a production app, you'd need proper Gmail API auth
      const delegatedAddresses = getDelegatedAddresses();
      if (delegatedAddresses && delegatedAddresses.length > 0) {
        // Add each delegated address to the results
        delegatedAddresses.forEach(addr => {
          results.push({
            email: addr.email,
            name: addr.name || '',
            isPrimary: false
          });
        });
      }
    } catch (e) {
      // Skip delegated addresses if there's an error or no access
      Logger.log('Couldn\'t retrieve delegated addresses: ' + e.message);
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
    // This would use Gmail API in a real implementation
    // Since we don't have full Gmail API access here, this returns dummy data
    
    // In a real implementation, you'd use Gmail API's users.settings.sendAs.list method
    // See: https://developers.google.com/gmail/api/reference/rest/v1/users.settings.sendAs/list
    
    return [];
    
    /* Example of real implementation with Gmail API:
    
    const gmailService = Gmail.Users.Settings.SendAs.list('me');
    const sendAsAddresses = gmailService.sendAs || [];
    
    return sendAsAddresses
      .filter(addr => !addr.isPrimary)
      .map(addr => ({
        email: addr.sendAsEmail,
        name: addr.displayName
      }));
    */
  }
  
  /**
   * Gets the user's display name from their Google Account.
   * @return {string} The user's name or email if name is not available.
   */
  function getUserName() {
    // This is a simplified example
    // In a production app with proper OAuth scope, you'd get the user's profile info
    const email = Session.getActiveUser().getEmail();
    
    // Extract a name from the email address if actual name not available
    const namePart = email.split('@')[0];
    return formatNameFromEmail(namePart);
  }
  
  /**
   * Formats a name from an email address part.
   * @param {string} namePart - The part of the email before @
   * @return {string} Formatted name
   */
  function formatNameFromEmail(namePart) {
    // Convert something like "john.doe" to "John Doe"
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
      
      // Get headers from first row
      const headers = values[0].filter(header => header !== "");
      
      // Calculate recipient count if email column is specified
      let recipientCount = 0;
      if (emailColumn && values.length > 1) {
        const emailColIndex = headers.indexOf(emailColumn);
        if (emailColIndex !== -1) {
          // Count non-empty cells in the email column (excluding header row)
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
      
      // Get all data
      const data = sheet.getDataRange().getValues();
      
      // Find email column index
      const headers = data[0];
      const emailColIndex = headers.indexOf(emailColumn);
      
      if (emailColIndex === -1) {
        throw new Error("Email column not found: " + emailColumn);
      }
      
      // Count non-empty cells in the email column (excluding header row)
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
    // Check if it's already an ID (simple alphanumeric string)
    if (/^[a-zA-Z0-9_-]{40,}$/.test(spreadsheetUrl)) {
      return spreadsheetUrl;
    }
    
    // Extract ID from URL
    const matches = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (matches && matches[1]) {
      return matches[1];
    }
    
    // Return original if no match (might be a direct ID)
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
      const values = range.getValues();
      
      const headers = values[0];
      const rows = values.slice(1);
      
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
   * @param {number} rowIndex - The 0-based row index (0 would return the first data row, not headers).
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
   * Generates a preview document with merged data.
   * @param {string} spreadsheetId - The ID of the spreadsheet.
   * @param {string} sheetName - The name of the sheet.
   * @param {number} rowIndex - The row index to use for the preview.
   * @return {Object} Status object with success flag, message, and doc URL.
   */
  function generatePreviewDoc(spreadsheetId, sheetName, rowIndex) {
    try {
      // If called from menu without params, show UI to select params
      if (!spreadsheetId || !sheetName) {
        const ui = DocumentApp.getUi();
        const response = ui.alert(
          'Generate Preview Document',
          'This will create a new document with merged data.\n\n' +
          'Please open the Mail Merge sidebar to configure and use this feature.',
          ui.ButtonSet.OK
        );
        return {
          success: false,
          message: "Please use the Mail Merge sidebar to configure and preview"
        };
      }
      
      // Get source document
      const sourceDoc = DocumentApp.getActiveDocument();
      const sourceDocText = sourceDoc.getBody().getText();
      
      // Get row data
      const rowData = getSpreadsheetRow(spreadsheetId, sheetName, rowIndex);
      
      // Create new document
      const newDoc = DocumentApp.create(`${sourceDoc.getName()} - Preview`);
      const body = newDoc.getBody();
      
      // Replace placeholders in content
      let mergedContent = sourceDocText;
      
      for (const [key, value] of Object.entries(rowData)) {
        const placeholder = `{{${key}}}`;
        const valueStr = value !== undefined && value !== null ? value.toString() : '';
        
        // Global replace all occurrences
        mergedContent = mergedContent.split(placeholder).join(valueStr);
      }
      
      // Set content to new document
      body.setText(mergedContent);
      
      // Save the document
      newDoc.saveAndClose();
      
      return {
        success: true,
        message: "Preview document created successfully",
        docUrl: newDoc.getUrl()
      };
    } catch (e) {
      return {
        success: false,
        message: "Error generating preview document: " + e.message
      };
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
      // Get document content
      const docText = getDocPlainText();
      
      // Extract placeholders from document using regex
      const placeholderRegex = /\{\{([^{}]+)\}\}/g;
      const docPlaceholders = [];
      let match;
      
      while ((match = placeholderRegex.exec(docText)) !== null) {
        docPlaceholders.push(match[1].trim());
      }
      
      // Get unique placeholders
      const uniqueDocPlaceholders = [...new Set(docPlaceholders)];
      
      // Get spreadsheet headers
      const headers = getColumnHeaders(spreadsheetId, sheetName);
      
      // Create validation results
      const matched = [];
      const unmatched = [];
      const unused = [];
      
      // Find matched and unmatched placeholders
      uniqueDocPlaceholders.forEach(placeholder => {
        if (headers.includes(placeholder)) {
          matched.push(placeholder);
        } else {
          unmatched.push(placeholder);
        }
      });
      
      // Find unused headers
      headers.forEach(header => {
        if (!uniqueDocPlaceholders.includes(header)) {
          unused.push(header);
        }
      });
      
      // Get an example row for the matched placeholders
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
          // If we can't get example values, just continue without them
          Logger.log('Error getting example values: ' + e.message);
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
      
      // Global replace all occurrences of the placeholder
      result = result.split(placeholder).join(value);
    }
    
    return result;
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
      // Get document content as HTML
      const body = getDocContent();
      let htmlBody = body;
      
      // Replace placeholders if requested and spreadsheet data is provided
      if (options && options.replacePlaceholders && options.spreadsheetId && options.sheetName) {
        try {
          // Get first row of data for placeholders
          const data = getSpreadsheetData(options.spreadsheetId, options.sheetName);
          if (data.rows.length > 0) {
            const firstRow = data.rows[0];
            
            // Replace placeholders in subject and body
            subject = replacePlaceholders(subject, data.headers, firstRow);
            htmlBody = replacePlaceholders(body, data.headers, firstRow);
          }
        } catch (e) {
          Logger.log('Error replacing placeholders: ' + e.message);
          // Continue with original content if placeholder replacement fails
        }
      }
      
      // Split comma-separated emails and trim whitespace
      const emailList = recipients.split(',').map(email => email.trim());
      
      // Create email options
      const emailOptions = {
        htmlBody: htmlBody,
        name: fromName || undefined
      };
      
      // Set optional fields
      if (cc) {
        emailOptions.cc = cc;
      }
      
      if (bcc) {
        emailOptions.bcc = bcc;
      }
      
      // Set from address if different from the user's address and if it's a delegated address
      if (fromEmail && fromEmail !== Session.getActiveUser().getEmail()) {
        emailOptions.from = fromEmail;
      }
      
      for (const email of emailList) {
        if (email) {
          GmailApp.sendEmail(email, subject, "", emailOptions);
        }
      }
      
      return {
        success: true,
        message: `Test email sent to: ${recipients}`
      };
    } catch (e) {
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
      // Set default options
      options = {
        cc: '',
        bcc: '',
        enableLogging: false,
        createDrafts: false,
        ...options
      };
      
      // Initialize logging if enabled
      let logSheet = null;
      if (options.enableLogging) {
        logSheet = createOrGetLogSheet();
        logSheet.appendRow(['Timestamp', 'Email', 'Status', 'Message']);
      }
      
      // Get document content as HTML
      const templateHtml = getDocContent();
      
      // Get spreadsheet data
      const data = getSpreadsheetData(spreadsheetId, sheetName);
      const headers = data.headers;
      const rows = data.rows;
      
      // Find email column index
      const emailIndex = headers.indexOf(emailColumn);
      if (emailIndex === -1) {
        throw new Error(`Email column "${emailColumn}" not found in spreadsheet`);
      }
      
      // Create email options
      const emailOptions = {
        name: fromName || undefined
      };
      
      // Add cc and bcc if provided
      if (options.cc) {
        emailOptions.cc = options.cc;
      }
      
      if (options.bcc) {
        emailOptions.bcc = options.bcc;
      }
      
      // Set from address if different from the user's address and if it's a delegated address
      if (fromEmail && fromEmail !== Session.getActiveUser().getEmail()) {
        emailOptions.from = fromEmail;
      }
      
      let sentCount = 0;
      let errorCount = 0;
      let errorEmails = [];
      
      // Process each row
      for (const row of rows) {
        const emailAddress = row[emailIndex];
        
        // Skip rows with no email
        if (!emailAddress) {
          continue;
        }
        
        try {
          // Replace placeholders in subject and body
          const personalizedSubject = replacePlaceholders(subjectLine, headers, row);
          const personalizedBody = replacePlaceholders(templateHtml, headers, row);
          
          // Copy options and add HTML body
          const rowEmailOptions = Object.assign({}, emailOptions, { htmlBody: personalizedBody });
          
          // Send email or create draft
          if (options.createDrafts) {
            GmailApp.createDraft(emailAddress, personalizedSubject, "", rowEmailOptions);
          } else {
            GmailApp.sendEmail(emailAddress, personalizedSubject, "", rowEmailOptions);
          }
          
          sentCount++;
          
          // Log success if logging is enabled
          if (logSheet) {
            logSheet.appendRow([
              new Date(),
              emailAddress,
              'Success',
              options.createDrafts ? 'Draft created' : 'Email sent'
            ]);
          }
          
          // Add a small delay to avoid rate limits
          Utilities.sleep(100);
        } catch (emailError) {
          errorCount++;
          errorEmails.push(emailAddress);
          
          // Log error if logging is enabled
          if (logSheet) {
            logSheet.appendRow([
              new Date(),
              emailAddress,
              'Error',
              emailError.message
            ]);
          }
          
          Logger.log(`Error sending to ${emailAddress}: ${emailError.message}`);
        }
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
   * Creates or gets a log sheet for mail merge operations.
   * @return {Sheet} The log sheet.
   */
  function createOrGetLogSheet() {
    const ss = SpreadsheetApp.create('Mail Merge Log - ' + new Date().toISOString().split('T')[0]);
    const sheet = ss.getActiveSheet();
    sheet.setName('Mail Merge Log');
    return sheet;
  }
  
  /**
   * Gets information about the current document.
   * @return {Object} Document information.
   */
  function getDocumentInfo() {
    const doc = DocumentApp.getActiveDocument();
    
    return {
      name: doc.getName(),
      subject: getSubjectLine()
    };
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
      // Get document content as HTML
      const templateHtml = getDocContent();
      
      // Get row data
      const data = getSpreadsheetData(spreadsheetId, sheetName);
      
      if (rowIndex >= data.rows.length) {
        throw new Error(`Row index ${rowIndex} is out of bounds. Max index is ${data.rows.length - 1}`);
      }
      
      const row = data.rows[rowIndex];
      
      // Generate personalized content
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