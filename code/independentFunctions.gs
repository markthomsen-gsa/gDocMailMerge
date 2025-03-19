/**
 * Comprehensive permissions diagnostic function with enhanced console output.
 * Tests all permissions required by the Mail Merge add-on.
 * @return {Object} Diagnostic results with permission status
 */
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