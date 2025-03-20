/**
 * Comprehensive permissions diagnostic function with enhanced console output.
 * Tests all permissions required by the Mail Merge add-on.
 * @return {Object} Diagnostic results with permission status
 */

/*
function runPermissionDiagnostics() {
  console.log('=====================================================');
  console.log('        MAIL MERGE PERMISSION DIAGNOSTICS           ');
  console.log('=====================================================');
  console.log(`Timestamp: ${new Date().toLocaleString()}`);
  console.log('');
  
  // Run each test with enhanced console output
  console.log('1️⃣ TESTING USER INFORMATION ACCESS...');
  const userInfo = testUserInfo();
  logTestResult('User Information', userInfo);
  
  console.log('2️⃣ TESTING GMAIL API ACCESS...');
  const gmail = testGmailAccess();
  logTestResult('Gmail API', gmail);
  
  console.log('3️⃣ TESTING SPREADSHEET ACCESS...');
  const spreadsheets = testSpreadsheetAccess();
  logTestResult('Spreadsheet Access', spreadsheets);
  
  console.log('4️⃣ TESTING EMAIL SENDING CAPABILITIES...');
  const emailSending = testEmailSending();
  logTestResult('Email Sending', emailSending);
  
  console.log('5️⃣ TESTING DOCUMENT ACCESS...');
  const document = testDocumentAccess();
  logTestResult('Document Access', document);
  
  console.log('6️⃣ TESTING SCRIPT PROPERTIES STORAGE...');
  const storage = testScriptProperties();
  logTestResult('Script Properties', storage);
  
  // Assemble the results
  const results = {
    timestamp: new Date().toISOString(),
    userInfo: userInfo,
    gmail: gmail,
    spreadsheets: spreadsheets,
    emailSending: emailSending,
    document: document,
    storage: storage
  };
  
  // Calculate overall status
  results.overallStatus = Object.values(results)
    .filter(val => typeof val === 'object' && val !== null && 'success' in val)
    .every(val => val.success) ? 'PASS' : 'FAIL';
  
  // Log overall status
  console.log('=====================================================');
  console.log(`OVERALL STATUS: ${results.overallStatus === 'PASS' ? 
              '✅ PASS - All permissions available' : 
              '❌ FAIL - Some permissions are restricted'}`);
  console.log('=====================================================');
  
  return results;
}

/**
 * Helper function to log test results in a formatted way
 * @param {string} testName - Name of the test
 * @param {Object} result - Test result object
 */
function logTestResult(testName, result) {
  const separator = '-----------------------------------------------------';
  console.log(separator);
  
  if (result.success) {
    console.info(`✅ PASS: ${testName}`);
    console.log(`Message: ${result.message}`);
    
    // Log additional details if present
    if (result.email) console.log(`Email: ${result.email}`);
    if (result.effectiveUser) console.log(`Effective User: ${result.effectiveUser}`);
    if (result.quotaRemaining !== undefined) console.log(`Quota Remaining: ${result.quotaRemaining}`);
    if (result.documentName) console.log(`Document Name: ${result.documentName}`);
  } else {
    console.error(`❌ FAIL: ${testName}`);
    console.warn(`Message: ${result.message}`);
    
    if (result.error) {
      console.error(`Error Details: ${result.error}`);
    }
  }
  
  console.log(''); // Empty line for better readability
}

/**
 * Tests access to user information.
 * @return {Object} Test results
 */
function testUserInfo() {
  try {
    const email = Session.getActiveUser().getEmail();
    const effectiveUser = Session.getEffectiveUser().getEmail();
    
    return {
      success: Boolean(email),
      email: email || 'NOT AVAILABLE',
      effectiveUser: effectiveUser || 'NOT AVAILABLE',
      message: email ? 'Successfully retrieved user email' : 'Could not retrieve user email'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error accessing user information'
    };
  }
}

/**
 * Tests Gmail API access.
 * @return {Object} Test results
 */
function testGmailAccess() {
  try {
    // Test if we can access Gmail settings (doesn't actually modify anything)
    const remaining = MailApp.getRemainingDailyQuota();
    
    return {
      success: true,
      quotaRemaining: remaining,
      message: 'Successfully accessed Gmail API'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error accessing Gmail API'
    };
  }
}

/**
 * Tests Spreadsheet access without requiring Drive API permissions.
 * @return {Object} Test results
 */
function testSpreadsheetAccess() {
  try {
    // Instead of creating a temporary spreadsheet, just check if we can access the Spreadsheet service
    const canAccess = Boolean(SpreadsheetApp);
    
    // For an additional test, try to list available spreadsheets (but catch any errors)
    let extraInfo = "";
    try {
      // This just tests if we can list spreadsheets without actually creating any
      const spreadsheets = SpreadsheetApp.getActiveSpreadsheet();
      if (spreadsheets) {
        extraInfo = "Successfully accessed active spreadsheet";
      }
    } catch (minorError) {
      // This is fine - we might not have an active spreadsheet
      extraInfo = "Note: No active spreadsheet available (this is normal if running from the script editor)";
    }
    
    return {
      success: canAccess,
      message: canAccess ? 
        'Successfully accessed Spreadsheet service' + (extraInfo ? '. ' + extraInfo : '') : 
        'Could not access Spreadsheet service',
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error accessing spreadsheets'
    };
  }
}

/**
 * Tests email sending capabilities.
 * @return {Object} Test results
 */
function testEmailSending() {
  try {
    // We don't actually send an email, just check if we can access the service
    const remaining = MailApp.getRemainingDailyQuota();
    
    return {
      success: remaining > 0,
      quotaRemaining: remaining,
      message: remaining > 0 ? 
        'Can send emails (quota available)' : 
        'Quota check succeeded but no emails remaining today'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error checking email sending capabilities'
    };
  }
}

/**
 * Tests document access.
 * @return {Object} Test results
 */
function testDocumentAccess() {
  try {
    // Try to access the current document
    const doc = DocumentApp.getActiveDocument();
    const docName = doc.getName();
    const docId = doc.getId();
    
    return {
      success: true,
      documentName: docName,
      documentId: docId,
      message: 'Successfully accessed the current document'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error accessing document'
    };
  }
}

/**
 * Tests script properties storage.
 * @return {Object} Test results
 */
function testScriptProperties() {
  try {
    // Test writing and reading from properties
    const testKey = 'test_key_' + new Date().getTime();
    const testValue = 'test_value_' + new Date().getTime();
    
    // Use script properties (doesn't persist between users)
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(testKey, testValue);
    const readValue = scriptProperties.getProperty(testKey);
    
    // Clean up
    scriptProperties.deleteProperty(testKey);
    
    return {
      success: readValue === testValue,
      message: 'Successfully wrote and read from script properties'
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Error accessing script properties'
    };
  }
}

/**
 * UI function to run diagnostics and display results.
 * Also outputs formatted results to console log.
 */
function showPermissionDiagnostics() {
  // Run the diagnostic tests
  const results = runPermissionDiagnostics();
  
  // Create a simple UI to show results
  const ui = DocumentApp.getUi();
  
  // Format results as human-readable text
  let message = 'Mail Merge Permission Diagnostics\n\n';
  message += 'Overall Status: ' + (results.overallStatus === 'PASS' ? '✅ PASS' : '❌ FAIL') + '\n\n';
  
  // Add detailed results
  message += '1. User Information: ' + (results.userInfo.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.userInfo.message + '\n\n';
  
  message += '2. Gmail Access: ' + (results.gmail.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.gmail.message + '\n\n';
  
  message += '3. Spreadsheet Access: ' + (results.spreadsheets.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.spreadsheets.message + '\n\n';
  
  message += '4. Email Sending: ' + (results.emailSending.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.emailSending.message + '\n\n';
  
  message += '5. Document Access: ' + (results.document.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.document.message + '\n\n';
  
  message += '6. Script Storage: ' + (results.storage.success ? '✅ PASS' : '❌ FAIL') + '\n';
  message += '   ' + results.storage.message + '\n\n';
  
  message += 'For detailed error messages, please check the Apps Script logs.';
  
  // Show the results
  ui.alert('Diagnostics Results', message, ui.ButtonSet.OK);
  
  // Log a message telling them to check the console
  console.log('Detailed diagnostic results are available in the Apps Script logs.');
}

/**
 * Tests permissions by checking the available scopes in appscript.json
 * @return {Object} Test results with permission details
 */
function testPermissionScopes() {
  try {
    // Create a list of permissions that might be needed by Mail Merge
    const requiredPermissions = [
      { scope: "https://www.googleapis.com/auth/gmail.addons.current.action.compose", name: "Gmail Compose Add-ons" },
      { scope: "https://www.googleapis.com/auth/gmail.compose", name: "Gmail Compose" },
      { scope: "https://www.googleapis.com/auth/gmail.send", name: "Gmail Send" },
      { scope: "https://www.googleapis.com/auth/gmail.settings.basic", name: "Gmail Settings" },
      { scope: "https://www.googleapis.com/auth/script.container.ui", name: "UI Containers" },
      { scope: "https://www.googleapis.com/auth/spreadsheets", name: "Spreadsheets" },
      { scope: "https://www.googleapis.com/auth/documents.currentonly", name: "Current Document" },
      { scope: "https://www.googleapis.com/auth/userinfo.email", name: "User Email" },
      { scope: "https://www.googleapis.com/auth/script.send_mail", name: "Script Mail" },
      { scope: "https://www.googleapis.com/auth/script.storage", name: "Script Storage" }
    ];
    
    // We can't directly check OAuth scopes, but we can run simple tests
    // for each permission to see what works
    const results = {};
    
    // Test user email access
    try {
      const email = Session.getActiveUser().getEmail();
      results["userinfo.email"] = {
        name: "User Email",
        status: email ? "Available" : "Limited",
        details: email ? `Email: ${email}` : "Email not available"
      };
    } catch (e) {
      results["userinfo.email"] = {
        name: "User Email",
        status: "Denied",
        details: e.message
      };
    }
    
    // Test Gmail quota check (tests gmail.send)
    try {
      const quota = MailApp.getRemainingDailyQuota();
      results["gmail.send"] = {
        name: "Gmail Send",
        status: "Available",
        details: `Quota remaining: ${quota}`
      };
    } catch (e) {
      results["gmail.send"] = {
        name: "Gmail Send",
        status: "Denied",
        details: e.message
      };
    }
    
    // Test spreadsheet access
    try {
      const test = SpreadsheetApp.getActiveSpreadsheet();
      results["spreadsheets"] = {
        name: "Spreadsheets",
        status: "Available",
        details: test ? "Active spreadsheet accessible" : "Service available but no active spreadsheet"
      };
    } catch (e) {
      results["spreadsheets"] = {
        name: "Spreadsheets",
        status: "Denied or Limited",
        details: e.message
      };
    }
    
    // Test document access
    try {
      const doc = DocumentApp.getActiveDocument();
      results["documents"] = {
        name: "Documents",
        status: "Available",
        details: doc ? `Document name: ${doc.getName()}` : "No active document"
      };
    } catch (e) {
      results["documents"] = {
        name: "Documents",
        status: "Denied or Limited",
        details: e.message
      };
    }
    
    // Check script properties
    try {
      const props = PropertiesService.getScriptProperties();
      const testKey = "permission_test";
      props.setProperty(testKey, "test");
      props.deleteProperty(testKey);
      results["script.storage"] = {
        name: "Script Storage",
        status: "Available",
        details: "Successfully wrote and read script properties"
      };
    } catch (e) {
      results["script.storage"] = {
        name: "Script Storage",
        status: "Denied",
        details: e.message
      };
    }
    
    // Calculate overall status
    const deniedPermissions = Object.values(results).filter(p => p.status === "Denied" || p.status === "Limited");
    
    return {
      success: deniedPermissions.length === 0,
      permissions: results,
      deniedCount: deniedPermissions.length,
      message: deniedPermissions.length === 0 ? 
        "All tested permissions are available" : 
        `${deniedPermissions.length} permissions are denied or limited`
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: "Error testing permissions"
    };
  }
}

/**
 * Gets and logs all email addresses the user can send from.
 * This includes the primary email and any configured "Send As" addresses
 * (which includes delegated addresses).
 * 
 * Run this function directly to see the results in the Apps Script logs.
 */
function listAllSendFromAddresses() {
  Logger.log("Starting to retrieve send-from addresses...");
  
  // Get the user's primary email
  const primaryEmail = Session.getActiveUser().getEmail();
  Logger.log(`Primary email: ${primaryEmail}`);
  
  // Array to store all available addresses
  const availableAddresses = [{
    email: primaryEmail,
    name: getUserName(),
    isPrimary: true,
    type: 'primary'
  }];
  
  try {
    // Get "send as" addresses using Gmail API
    Logger.log("Retrieving 'Send As' addresses from Gmail API...");
    const sendAsResponse = Gmail.Users.Settings.SendAs.list('me');
    const sendAsAddresses = sendAsResponse.sendAs || [];
    
    Logger.log(`Found ${sendAsAddresses.length} SendAs configurations`);
    
    for (const sendAs of sendAsAddresses) {
      // Skip the primary email if it's also in SendAs list to avoid duplicates
      if (sendAs.isPrimary && sendAs.sendAsEmail === primaryEmail) {
        Logger.log(`Skipping primary email in SendAs: ${sendAs.sendAsEmail}`);
        continue;
      }
      
      Logger.log(`Adding SendAs address: ${sendAs.sendAsEmail} (${sendAs.displayName || 'No display name'})`);
      availableAddresses.push({
        email: sendAs.sendAsEmail,
        name: sendAs.displayName || '',
        isPrimary: false,
        type: 'sendAs',
        verified: sendAs.isVerified || false
      });
    }
    
    // Try to get delegate addresses
    Logger.log("Checking for delegate addresses...");
    try {
      // Note: This will list accounts that have delegated access TO you, not FROM you
      const delegatesResponse = Gmail.Users.Settings.Delegates.list('me');
      const delegates = delegatesResponse.delegates || [];
      
      Logger.log(`Found ${delegates.length} delegate configurations`);
      
      for (const delegate of delegates) {
        Logger.log(`Found delegate: ${delegate.delegateEmail} (${delegate.verificationStatus})`);
        // These are accounts that have delegated TO you, not accounts you can send FROM
      }
    } catch (delegateErr) {
      // This is expected in many cases as delegates API can be restricted
      Logger.log(`Note on delegates: ${delegateErr.message}`);
    }
  } catch (e) {
    Logger.log(`Error retrieving addresses: ${e.message}`);
    Logger.log(`Error stack: ${e.stack}`);
  }
  
  // Log summary to console
  Logger.log("\n=== AVAILABLE SEND-FROM ADDRESSES SUMMARY ===");
  for (const addr of availableAddresses) {
    Logger.log(`- ${addr.email} (${addr.name || 'No name'}) [${addr.type}]${addr.isPrimary ? ' (Primary)' : ''}${addr.verified === false ? ' (unverified)' : ''}`);
  }
  
  return availableAddresses;
}

/**
 * Lists all email sending identities available to the current user.
 * 
 * This function retrieves:
 * 1. The user's primary Gmail address
 * 2. Any "Send As" addresses configured in Gmail settings
 * 
 * "Send As" addresses are alternate identities you've configured in Gmail settings
 * that allow you to send emails from different addresses through your Gmail account.
 * 
 * Run this function directly to see results in the Apps Script logs.
 */
function listEmailSendingIdentities() {
  Logger.log("==== EMAIL SENDING IDENTITIES REPORT ====");
  Logger.log("Retrieving all available email addresses you can send from...");
  
  // Get the user's primary email
  const primaryEmail = Session.getActiveUser().getEmail();
  Logger.log(`Primary Gmail address: ${primaryEmail}`);
  
  // Array to store all available addresses
  const availableAddresses = [{
    email: primaryEmail,
    name: getUserNameOrDefault(primaryEmail),
    isPrimary: true,
    type: 'primary'
  }];
  
  try {
    // Get "send as" addresses using Gmail API
    Logger.log("\nRetrieving 'Send As' addresses from Gmail settings...");
    const sendAsResponse = Gmail.Users.Settings.SendAs.list('me');
    const sendAsAddresses = sendAsResponse.sendAs || [];
    
    Logger.log(`Found ${sendAsAddresses.length} SendAs configurations`);
    
    // Loop through and log each SendAs address
    for (const sendAs of sendAsAddresses) {
      // Skip the primary email if it's also in SendAs list to avoid duplicates
      if (sendAs.isPrimary && sendAs.sendAsEmail === primaryEmail) {
        Logger.log(`  - Skipping primary email in SendAs list: ${sendAs.sendAsEmail}`);
        continue;
      }
      
      // Add to our collection
      availableAddresses.push({
        email: sendAs.sendAsEmail,
        name: sendAs.displayName || '',
        isPrimary: false,
        isDefault: sendAs.isDefault || false,
        type: 'sendAs',
        verified: sendAs.isVerified || false
      });
      
      // Log detailed information
      Logger.log(`  - Found SendAs address: ${sendAs.sendAsEmail}`);
      Logger.log(`    Display name: ${sendAs.displayName || '(None)'}`);
      Logger.log(`    Verified: ${sendAs.isVerified ? 'Yes' : 'No'}`);
      Logger.log(`    Is default: ${sendAs.isDefault ? 'Yes' : 'No'}`);
      
      // If we have reply-to info, log it
      if (sendAs.replyToAddress && sendAs.replyToAddress !== sendAs.sendAsEmail) {
        Logger.log(`    Reply-To address: ${sendAs.replyToAddress}`);
      }
      
      // If we have signature info, log it
      if (sendAs.signature) {
        Logger.log(`    Has signature: Yes (${sendAs.signature.length} characters)`);
      } else {
        Logger.log(`    Has signature: No`);
      }
    }
  } catch (e) {
    Logger.log(`\nError retrieving SendAs addresses: ${e.message}`);
    Logger.log(`Error stack: ${e.stack}`);
  }
  
  // Log summary to console
  Logger.log("\n==== SUMMARY OF AVAILABLE SENDING IDENTITIES ====");
  for (const addr of availableAddresses) {
    const displayName = addr.name ? `"${addr.name}" ` : '';
    const flags = [
      addr.isPrimary ? 'PRIMARY' : '',
      addr.type === 'sendAs' ? 'SENDAS' : '',
      addr.isDefault ? 'DEFAULT' : '',
      addr.verified === false ? 'UNVERIFIED' : ''
    ].filter(Boolean).join(', ');
    
    Logger.log(`- ${displayName}<${addr.email}> [${flags}]`);
  }
  
  return availableAddresses;
}

/**
 * Extracts a display name from an email address if no name is provided.
 * Used as a helper function.
 * @param {string} email - Email address to extract name from
 * @return {string} Formatted name
 */
function getUserNameOrDefault(email) {
  try {
    const namePart = email.split('@')[0];
    return namePart
      .split(/[._-]/)
      .map(part => part.charAt(0).toUpperCase() + part.slice(1))
      .join(' ');
  } catch (e) {
    return "";
  }
}

/**
 * Updated implementation for the getDelegatedAddresses function in your mail merge app.
 * This correctly retrieves all SendAs addresses (which include delegates configured in Gmail).
 * 
 * @return {Object[]} Array of SendAs email addresses
 */
function updateGetDelegatedAddressesFunction() {
  try {
    // This is how your existing getDelegatedAddresses function should be implemented
    const sendAsResponse = Gmail.Users.Settings.SendAs.list('me');
    const sendAsAddresses = sendAsResponse.sendAs || [];
    const delegatedAddresses = [];
    
    for (const sendAs of sendAsAddresses) {
      // Skip the primary email
      if (sendAs.isPrimary) continue;
      
      delegatedAddresses.push({
        email: sendAs.sendAsEmail,
        name: sendAs.displayName || '',
        verified: sendAs.isVerified || false
      });
    }
    
    Logger.log(`Found ${delegatedAddresses.length} SendAs addresses to use in mail merge`);
    return delegatedAddresses;
  } catch (e) {
    Logger.log(`Error in getDelegatedAddresses: ${e.message}`);
    return [];
  }
}

/**
 * Copies all document properties as formatted JSON to the clipboard.
 * This function creates a dialog that displays the JSON and provides copy functionality.
 * @return {Object} Result with success flag and message
 */
function copyDocumentPropertiesToClipboard() {
  try {
    // Get all document properties
    const docProperties = PropertiesService.getDocumentProperties();
    const allProperties = docProperties.getProperties();
    
    // Format the JSON with pretty printing (indentation)
    const formattedJson = JSON.stringify(allProperties, null, 2);
    
    // Calculate size
    const sizeKB = (formattedJson.length / 1024).toFixed(2);
    
    // Create a dialog to display the JSON and allow copying
    const html = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 16px;
            }
            pre {
              background-color: #f5f5f5;
              padding: 12px;
              border: 1px solid #ddd;
              border-radius: 4px;
              overflow: auto;
              max-height: 400px;
              white-space: pre-wrap;
              font-size: 12px;
            }
            .info {
              margin-bottom: 12px;
              color: #666;
            }
            .buttons {
              margin-top: 16px;
            }
            button {
              padding: 8px 16px;
              cursor: pointer;
              background-color: #4285f4;
              color: white;
              border: none;
              border-radius: 4px;
            }
            .success {
              color: green;
              font-weight: bold;
              display: none;
              margin-left: 12px;
            }
          </style>
        </head>
        <body>
          <h3>Document Properties JSON</h3>
          <div class="info">
            Total properties: ${Object.keys(allProperties).length}<br>
            Size: ${sizeKB} KB
          </div>
          <pre id="jsonContent">${formattedJson.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>
          <div class="buttons">
            <button id="copyBtn" onclick="copyToClipboard()">Copy to Clipboard</button>
            <span id="successMsg" class="success">✓ Copied!</span>
          </div>
          
          <script>
            function copyToClipboard() {
              const textarea = document.createElement('textarea');
              textarea.value = document.getElementById('jsonContent').textContent;
              document.body.appendChild(textarea);
              textarea.select();
              document.execCommand('copy');
              document.body.removeChild(textarea);
              
              // Show success message
              const successMsg = document.getElementById('successMsg');
              successMsg.style.display = 'inline';
              setTimeout(() => {
                successMsg.style.display = 'none';
              }, 2000);
            }
          </script>
        </body>
      </html>
    `)
    .setWidth(600)
    .setHeight(500);
    
    DocumentApp.getUi().showModalDialog(html, 'Document Properties JSON');
    
    return {
      success: true,
      message: "Document properties displayed in dialog"
    };
  } catch (e) {
    Logger.log("Error copying document properties: " + e.message);
    return {
      success: false,
      message: "Error copying document properties: " + e.message
    };
  }
}

/**
 * Shows the document properties in a dialog and includes a button to copy to clipboard.
 * Wrapper function that shows a UI message after execution.
 */
function showDocumentPropertiesDialog() {
  const result = copyDocumentPropertiesToClipboard();
  
  if (!result.success) {
    const ui = DocumentApp.getUi();
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}

/**
 * Dumps all PropertiesService stores to the console log
 * Shows contents of DocumentProperties, UserProperties, and ScriptProperties
 */
function dumpAllPropertyStores() {
  const docProps = PropertiesService.getDocumentProperties().getProperties();
  const userProps = PropertiesService.getUserProperties().getProperties();
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  
  console.log('=== DOCUMENT PROPERTIES ===');
  console.log(JSON.stringify(docProps, null, 2));
  
  console.log('=== USER PROPERTIES ===');
  console.log(JSON.stringify(userProps, null, 2));
  
  console.log('=== SCRIPT PROPERTIES ===');
  console.log(JSON.stringify(scriptProps, null, 2));
  
  // Return summary for UI display
  return {
    documentPropertiesCount: Object.keys(docProps).length,
    userPropertiesCount: Object.keys(userProps).length,
    scriptPropertiesCount: Object.keys(scriptProps).length,
    
    // Details about mail merge configs specifically
    documentConfigsFound: Object.keys(docProps).filter(key => key === 'mailMergeConfigs').length > 0,
    userConfigsFound: Object.keys(userProps).filter(key => key === 'mailMergeConfigs').length > 0,
    scriptConfigsFound: Object.keys(scriptProps).filter(key => key === 'mailMergeConfigs').length > 0
  };
}

/**
 * Extracts and returns just the mail merge configurations from all property stores
 * for easy comparison
 */
function getMailMergeConfigs() {
  // Helper to safely parse JSON
  const safeJsonParse = (jsonStr) => {
    try {
      if (!jsonStr) return null;
      return JSON.parse(jsonStr);
    } catch (e) {
      return { error: e.message };
    }
  };
  
  // Get configs from each store
  const docConfigs = safeJsonParse(
    PropertiesService.getDocumentProperties().getProperty('mailMergeConfigs')
  );
  
  const userConfigs = safeJsonParse(
    PropertiesService.getUserProperties().getProperty('mailMergeConfigs')
  );
  
  const scriptConfigs = safeJsonParse(
    PropertiesService.getScriptProperties().getProperty('mailMergeConfigs')
  );
  
  // List all template names across all stores
  const allTemplateNames = new Set();
  
  if (docConfigs) {
    Object.keys(docConfigs).forEach(name => allTemplateNames.add(name));
  }
  
  if (userConfigs) {
    Object.keys(userConfigs).forEach(name => allTemplateNames.add(name));
  }
  
  if (scriptConfigs) {
    Object.keys(scriptConfigs).forEach(name => allTemplateNames.add(name));
  }
  
  return {
    documentConfigurations: docConfigs,
    userConfigurations: userConfigs,
    scriptConfigurations: scriptConfigs,
    allTemplateNames: Array.from(allTemplateNames)
  };
}

/**
 * Creates a UI dialog to display the configurations and debug information
 */
function showDebugDialog() {
  const configs = getMailMergeConfigs();
  const propertyCounts = dumpAllPropertyStores();
  
  // Format results for display
  let html = '<html><head>';
  html += '<style>';
  html += 'body { font-family: Arial, sans-serif; margin: 20px; }';
  html += 'h2 { color: #4285f4; }';
  html += 'h3 { color: #5f6368; margin-top: 20px; }';
  html += '.store { border: 1px solid #dadce0; border-radius: 4px; padding: 10px; margin-bottom: 20px; }';
  html += '.error { background-color: #fce8e6; border: 1px solid #ea4335; padding: 10px; }';
  html += '.not-found { color: #5f6368; font-style: italic; }';
  html += 'pre { background-color: #f8f9fa; padding: 10px; overflow: auto; }';
  html += '.summary { background-color: #e8f0fe; padding: 10px; margin-bottom: 20px; }';
  html += '</style>';
  html += '</head><body>';
  
  // Summary section
  html += '<div class="summary">';
  html += '<h2>Mail Merge Configuration Stores</h2>';
  html += '<p>Document Properties Count: ' + propertyCounts.documentPropertiesCount + 
          (propertyCounts.documentConfigsFound ? ' <strong>(Contains Mail Merge Configs)</strong>' : '') + '</p>';
  html += '<p>User Properties Count: ' + propertyCounts.userPropertiesCount + 
          (propertyCounts.userConfigsFound ? ' <strong>(Contains Mail Merge Configs)</strong>' : '') + '</p>';
  html += '<p>Script Properties Count: ' + propertyCounts.scriptPropertiesCount + 
          (propertyCounts.scriptConfigsFound ? ' <strong>(Contains Mail Merge Configs)</strong>' : '') + '</p>';
  
  // All template names
  html += '<p>All Template Names: ';
  if (configs.allTemplateNames.length === 0) {
    html += '<span class="not-found">No templates found in any store</span>';
  } else {
    html += configs.allTemplateNames.join(', ');
  }
  html += '</p>';
  html += '</div>';
  
  // Document Properties
  html += '<div class="store">';
  html += '<h3>Document Properties</h3>';
  if (configs.documentConfigurations) {
    html += '<pre>' + JSON.stringify(configs.documentConfigurations, null, 2) + '</pre>';
  } else {
    html += '<p class="not-found">No Mail Merge configurations found in Document Properties</p>';
  }
  html += '</div>';
  
  // User Properties
  html += '<div class="store">';
  html += '<h3>User Properties</h3>';
  if (configs.userConfigurations) {
    html += '<pre>' + JSON.stringify(configs.userConfigurations, null, 2) + '</pre>';
  } else {
    html += '<p class="not-found">No Mail Merge configurations found in User Properties</p>';
  }
  html += '</div>';
  
  // Script Properties
  html += '<div class="store">';
  html += '<h3>Script Properties</h3>';
  if (configs.scriptConfigurations) {
    html += '<pre>' + JSON.stringify(configs.scriptConfigurations, null, 2) + '</pre>';
  } else {
    html += '<p class="not-found">No Mail Merge configurations found in Script Properties</p>';
  }
  html += '</div>';
  
  html += '<button onclick="google.script.host.close()">Close</button>';
  html += '</body></html>';
  
  const ui = HtmlService.createHtmlOutput(html)
      .setWidth(800)
      .setHeight(600)
      .setTitle('Mail Merge Debug Information');
  
  DocumentApp.getUi().showModalDialog(ui, 'Mail Merge Debug Information');
}

/**
 * Shows a more compact dialog with options to copy between stores
 */
function showConfigurationFixerDialog() {
  const configs = getMailMergeConfigs();
  
  let html = '<html><head>';
  html += '<style>';
  html += 'body { font-family: Arial, sans-serif; margin: 20px; }';
  html += 'h2, h3 { color: #4285f4; }';
  html += '.panel { border: 1px solid #dadce0; border-radius: 4px; padding: 10px; margin-bottom: 20px; }';
  html += 'button { background-color: #4285f4; color: white; border: none; padding: 8px 16px; margin: 5px; cursor: pointer; }';
  html += 'button.warning { background-color: #ea4335; }';
  html += 'select { padding: 8px; margin: 5px; min-width: 200px; }';
  html += '</style>';
  html += '</head><body>';
  
  html += '<h2>Mail Merge Configuration Fixer</h2>';
  
  // Show template counts
  const docCount = configs.documentConfigurations ? Object.keys(configs.documentConfigurations).length : 0;
  const userCount = configs.userConfigurations ? Object.keys(configs.userConfigurations).length : 0;
  const scriptCount = configs.scriptConfigurations ? Object.keys(configs.scriptConfigurations).length : 0;
  
  html += '<div class="panel">';
  html += '<h3>Configuration Counts</h3>';
  html += '<p>Document Properties: ' + docCount + ' templates</p>';
  html += '<p>User Properties: ' + userCount + ' templates</p>';
  html += '<p>Script Properties: ' + scriptCount + ' templates</p>';
  html += '</div>';
  
  // Copy templates section
  html += '<div class="panel">';
  html += '<h3>Copy Templates</h3>';
  html += '<p>Copy configurations from one store to another:</p>';
  
  // Source selection
  html += '<div><label for="source">Source: </label>';
  html += '<select id="source">';
  html += '<option value="document">Document Properties (' + docCount + ')</option>';
  html += '<option value="user">User Properties (' + userCount + ')</option>';
  html += '<option value="script">Script Properties (' + scriptCount + ')</option>';
  html += '</select></div>';
  
  // Destination selection
  html += '<div><label for="destination">Destination: </label>';
  html += '<select id="destination">';
  html += '<option value="document">Document Properties</option>';
  html += '<option value="user">User Properties</option>';
  html += '<option value="script">Script Properties</option>';
  html += '</select></div>';
  
  // Action buttons
  html += '<div style="margin-top: 15px;">';
  html += '<button onclick="copyConfigs()">Copy Configurations</button>';
  html += '<button class="warning" onclick="clearDestination()">Clear Destination First</button>';
  html += '</div>';
  html += '</div>';
  
  // Add action for force refreshing sidebar config list
  html += '<div class="panel">';
  html += '<h3>Force Sidebar to Refresh</h3>';
  html += '<p>If the sidebar is not showing your configurations, click this:</p>';
  html += '<button onclick="forceRefresh()">Force Configuration Refresh</button>';
  html += '</div>';
  
  // Add scripts
  html += '<script>';
  html += 'function copyConfigs() {';
  html += '  const source = document.getElementById("source").value;';
  html += '  const destination = document.getElementById("destination").value;';
  html += '  if (source === destination) {';
  html += '    alert("Source and destination cannot be the same");';
  html += '    return;';
  html += '  }';
  html += '  google.script.run';
  html += '    .withSuccessHandler(function(result) {';
  html += '      alert(result.message);';
  html += '    })';
  html += '    .withFailureHandler(function(error) {';
  html += '      alert("Error: " + error.message);';
  html += '    })';
  html += '    .copyConfigsBetweenStores(source, destination, false);';
  html += '}';
  
  html += 'function clearDestination() {';
  html += '  const destination = document.getElementById("destination").value;';
  html += '  if (confirm("Are you sure you want to CLEAR all configurations in " + destination + "?")) {';
  html += '    google.script.run';
  html += '      .withSuccessHandler(function(result) {';
  html += '        alert(result.message);';
  html += '      })';
  html += '      .withFailureHandler(function(error) {';
  html += '        alert("Error: " + error.message);';
  html += '      })';
  html += '      .clearConfigStore(destination);';
  html += '  }';
  html += '}';
  
  html += 'function forceRefresh() {';
  html += '  google.script.run';
  html += '    .withSuccessHandler(function() {';
  html += '      alert("Refresh flag set. Close and reopen sidebar to see changes.");';
  html += '    })';
  html += '    .setConfigurationRefreshFlag();';
  html += '}';
  html += '</script>';
  
  html += '<div style="margin-top: 20px;">';
  html += '<button onclick="google.script.host.close()">Close</button>';
  html += '<button onclick="window.location.reload()">Refresh Dialog</button>';
  html += '</div>';
  
  html += '</body></html>';
  
  const ui = HtmlService.createHtmlOutput(html)
      .setWidth(500)
      .setHeight(500)
      .setTitle('Mail Merge Configuration Fixer');
  
  DocumentApp.getUi().showModalDialog(ui, 'Mail Merge Configuration Fixer');
}

/**
 * Copies configurations between property stores
 * @param {string} source - The source store ("document", "user", or "script")
 * @param {string} destination - The destination store ("document", "user", or "script")
 * @param {boolean} clearFirst - Whether to clear destination before copying
 * @return {Object} Result with success/failure info
 */
function copyConfigsBetweenStores(source, destination, clearFirst) {
  try {
    // Get source properties
    let sourceProps;
    switch(source) {
      case "document":
        sourceProps = PropertiesService.getDocumentProperties();
        break;
      case "user":
        sourceProps = PropertiesService.getUserProperties();
        break;
      case "script":
        sourceProps = PropertiesService.getScriptProperties();
        break;
      default:
        throw new Error("Invalid source store: " + source);
    }
    
    // Get destination properties
    let destProps;
    switch(destination) {
      case "document":
        destProps = PropertiesService.getDocumentProperties();
        break;
      case "user":
        destProps = PropertiesService.getUserProperties();
        break;
      case "script":
        destProps = PropertiesService.getScriptProperties();
        break;
      default:
        throw new Error("Invalid destination store: " + destination);
    }
    
    // Get mail merge configs from source
    const configsJson = sourceProps.getProperty('mailMergeConfigs');
    if (!configsJson) {
      return {
        success: false,
        message: "No mail merge configurations found in " + source + " properties"
      };
    }
    
    // Clear destination if requested
    if (clearFirst) {
      destProps.deleteProperty('mailMergeConfigs');
    }
    
    // Copy to destination
    destProps.setProperty('mailMergeConfigs', configsJson);
    
    // Set refresh flag for UI
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return {
      success: true,
      message: "Successfully copied configurations from " + source + " to " + destination
    };
  } catch (e) {
    return {
      success: false,
      message: "Error copying configurations: " + e.message
    };
  }
}

/**
 * Clears mail merge configurations from a specific store
 * @param {string} store - The store to clear ("document", "user", or "script")
 * @return {Object} Result with success/failure info
 */
function clearConfigStore(store) {
  try {
    // Get properties service
    let props;
    switch(store) {
      case "document":
        props = PropertiesService.getDocumentProperties();
        break;
      case "user":
        props = PropertiesService.getUserProperties();
        break;
      case "script":
        props = PropertiesService.getScriptProperties();
        break;
      default:
        throw new Error("Invalid store: " + store);
    }
    
    // Delete mail merge configs
    props.deleteProperty('mailMergeConfigs');
    
    // Set refresh flag for UI
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return {
      success: true,
      message: "Successfully cleared configurations from " + store + " properties"
    };
  } catch (e) {
    return {
      success: false,
      message: "Error clearing configurations: " + e.message
    };
  }
}

/**
 * Repairs malformed template structures in document properties
 * Fixes issues with duplicated config sections and missing content
 */
function repairTemplateStructures() {
  try {
    // Get current configurations
    const docProperties = PropertiesService.getDocumentProperties();
    const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const configs = JSON.parse(configsJson);
    
    let repairedCount = 0;
    let failures = [];
    
    // Process each configuration
    for (const [name, config] of Object.entries(configs)) {
      if (config.documentContent) {
        try {
          // Check for duplicated config section
          const hasDuplicateConfig = 
            (config.documentContent.match(/===TEMPLATE CONFIG===/g) || []).length > 1;
          
          if (hasDuplicateConfig) {
            console.log(`Repairing template structure for "${name}"`);
            
            // Extract the correct parts
            const firstConfigStart = config.documentContent.indexOf("===TEMPLATE CONFIG===");
            const contentMarkerStart = config.documentContent.indexOf("===TEMPLATE CONTENT===");
            const endMarkerStart = config.documentContent.indexOf("===TEMPLATE END===");
            
            if (firstConfigStart !== -1 && contentMarkerStart !== -1 && endMarkerStart !== -1) {
              // Get configuration section (only once)
              const configSection = config.documentContent.substring(
                firstConfigStart, 
                contentMarkerStart
              );
              
              // Create proper content (placeholder if none exists)
              let contentSection = "";
              
              // Try to find any actual content between markers
              const secondConfigStart = config.documentContent.indexOf(
                "===TEMPLATE CONFIG===", 
                contentMarkerStart + 1
              );
              
              if (secondConfigStart !== -1 && secondConfigStart < endMarkerStart) {
                // Found duplicated config, just use placeholder content
                contentSection = "\n\nThis is a repaired email template. Please add your content here.\n\n";
              } else {
                // Try to extract content between content marker and end marker
                contentSection = config.documentContent.substring(
                  contentMarkerStart + "===TEMPLATE CONTENT===".length,
                  endMarkerStart
                ).trim();
                
                // If content is empty, add placeholder
                if (!contentSection) {
                  contentSection = "\n\nThis is a repaired email template. Please add your content here.\n\n";
                }
              }
              
              // Rebuild proper structure
              const repairedContent = 
                configSection + 
                "===TEMPLATE CONTENT===\n" + 
                contentSection + 
                "\n===TEMPLATE END===";
              
              // Update the configuration
              config.documentContent = repairedContent;
              configs[name] = config;
              repairedCount++;
            } else {
              failures.push({name, reason: "Missing required markers"});
            }
          }
        } catch (templateError) {
          console.error(`Error repairing template "${name}": ${templateError.message}`);
          failures.push({name, reason: templateError.message});
        }
      }
    }
    
    // Save the updated configurations back if any were repaired
    if (repairedCount > 0) {
      docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
      
      // Force sidebar to update
      PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    }
    
    return {
      success: true,
      repairedCount: repairedCount,
      totalTemplates: Object.keys(configs).length,
      failures: failures
    };
  } catch (e) {
    return {
      success: false,
      message: "Error repairing templates: " + e.message
    };
  }
}

/**
 * Shows a UI dialog to run the repair function and see results
 */
function showTemplateRepairDialog() {
  const ui = DocumentApp.getUi();
  
  // Confirm before proceeding
  const response = ui.alert(
    'Repair Template Structures', 
    'This will scan and attempt to fix malformed template structures. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Run the repair function
  const result = repairTemplateStructures();
  
  if (result.success) {
    // Show results
    const message = `Template Repair Complete\n\n` +
                   `Templates scanned: ${result.totalTemplates}\n` +
                   `Templates repaired: ${result.repairedCount}\n` +
                   `Failed repairs: ${result.failures.length}\n\n` +
                   (result.failures.length > 0 ? 
                    `Failed templates: ${result.failures.map(f => f.name).join(", ")}` : 
                    `All repairs successful!`);
    
    ui.alert('Repair Results', message, ui.ButtonSet.OK);
  } else {
    // Show error
    ui.alert('Repair Error', result.message, ui.ButtonSet.OK);
  }
}

/**
 * Clears all Mail Merge configurations from Document Properties
 * @return {Object} Result object with success and message
 */
function clearMailMergeConfigurations() {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.deleteProperty('mailMergeConfigs');
    
    // Set refresh flag for sidebar
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return {
      success: true,
      message: "Successfully cleared all Mail Merge configurations from Document Properties."
    };
  } catch (e) {
    return {
      success: false,
      message: "Error clearing configurations: " + e.message
    };
  }
}

/**
 * Clears ALL properties from Document Properties (not just Mail Merge)
 * Use with caution - this will remove ALL document-specific settings
 * @return {Object} Result object with success and message
 */
function clearAllDocumentProperties() {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.deleteAllProperties();
    
    // Set refresh flag for sidebar
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return {
      success: true,
      message: "Successfully cleared ALL properties from Document Properties."
    };
  } catch (e) {
    return {
      success: false,
      message: "Error clearing properties: " + e.message
    };
  }
}

/**
 * Shows a dialog to confirm and clear Document Properties
 */
function showClearPropertiesDialog() {
  const ui = DocumentApp.getUi();
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2 { color: #ea4335; }
        .warning { color: #ea4335; font-weight: bold; }
        .panel { border: 1px solid #dadce0; border-radius: 4px; padding: 15px; margin: 15px 0; }
        button { padding: 10px 15px; margin: 5px; cursor: pointer; }
        .danger { background-color: #ea4335; color: white; border: none; }
        .secondary { background-color: #f1f3f4; border: 1px solid #dadce0; }
      </style>
    </head>
    <body>
      <h2>Clear Document Properties</h2>
      
      <div class="panel">
        <p><span class="warning">⚠️ Warning:</span> This operation cannot be undone!</p>
        <p>Choose an option:</p>
        <ul>
          <li><strong>Clear Mail Merge only</strong>: Removes only Mail Merge configurations</li>
          <li><strong>Clear ALL properties</strong>: Removes ALL document properties (use with caution)</li>
        </ul>
      </div>
      
      <button class="danger" onclick="clearMailMergeOnly()">Clear Mail Merge Only</button>
      <button class="danger" onclick="clearAllProperties()" style="background-color: #b31412;">Clear ALL Properties</button>
      <button class="secondary" onclick="google.script.host.close()">Cancel</button>
      
      <script>
        function clearMailMergeOnly() {
          if (confirm("Are you sure you want to clear all Mail Merge configurations?")) {
            google.script.run
              .withSuccessHandler(function(result) {
                alert(result.message);
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert("Error: " + error.message);
              })
              .clearMailMergeConfigurations();
          }
        }
        
        function clearAllProperties() {
          if (confirm("⚠️ WARNING: This will delete ALL document properties, not just Mail Merge! Are you absolutely sure?")) {
            google.script.run
              .withSuccessHandler(function(result) {
                alert(result.message);
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert("Error: " + error.message);
              })
              .clearAllDocumentProperties();
          }
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(450)
  .setHeight(350)
  .setTitle('Clear Document Properties');
  
  ui.showModalDialog(htmlOutput, 'Clear Document Properties');
}
*/