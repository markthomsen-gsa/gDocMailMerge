

/**
 * Checks and displays the user's email quota information
 * Shows different values based on whether it's a Workspace or personal Gmail account
 */
function checkEmailQuota() {
    const ui = DocumentApp.getUi();
    const email = Session.getActiveUser().getEmail();
    
    // Determine if it's a Workspace or personal Gmail account
    const isWorkspace = !email.toLowerCase().endsWith('@gmail.com');
    
    // Set quota limits based on account type
    const dailyLimit = isWorkspace ? 1500 : 500; // Workspace has higher limit than personal Gmail
    
    // Get remaining Gmail quota for today using Apps Script quotas
    // Note: This is a rough estimate as Google doesn't provide direct API for this
    let remainingToday;
    let accountType;
    
    if (isWorkspace) {
      accountType = "Google Workspace";
      // Try to calculate based on Workspace settings
      // This is an estimate since exact quota can vary by Workspace plan
      remainingToday = dailyLimit;
    } else {
      accountType = "Personal Gmail";
      remainingToday = dailyLimit;
    }
    
    // Check quota history if we store it in user properties
    // Note: This relies on the mail merge actually tracking sent emails
    const userProperties = PropertiesService.getUserProperties();
    const emailsSentToday = parseInt(userProperties.getProperty('emailsSentToday') || '0', 10);
    
    if (emailsSentToday > 0) {
      // If we're tracking sent emails, show accurate remaining count
      remainingToday = Math.max(0, dailyLimit - emailsSentToday);
    }
    
    // Create message for user
    const message = 
      `Email Account Type: ${accountType}\n\n` +
      `Daily Sending Limit: ${dailyLimit} emails\n` +
      `Emails Sent Today: ${emailsSentToday || "0"}\n` +
      `Remaining Quota: ${remainingToday} emails\n\n` +
      `Account: ${email}\n\n` +
      "Note: These limits reset daily at midnight Pacific Time."
    
    // Show the quota information in an alert
    ui.alert('Email Quota Information', message, ui.ButtonSet.OK);
  }
  
  /**
   * Utility function to update the count of emails sent today
   * This should be called after each successful mail merge operation
   * 
   * @param {number} count - Number of emails just sent
   */
  function updateEmailsSentCount(count) {
    const userProperties = PropertiesService.getUserProperties();
    const currentCount = parseInt(userProperties.getProperty('emailsSentToday') || '0', 10);
    const newCount = currentCount + count;
    
    // Update the count in user properties
    userProperties.setProperty('emailsSentToday', newCount.toString());
    
    // Also update the date to ensure tracking is for today
    const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD format
    userProperties.setProperty('lastCountDate', today);
  }
  
  /**
   * Resets the daily email counter
   * Automatically called at midnight Pacific Time
   * Can also be called manually if needed
   */
  function resetEmailsSentCount() {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('emailsSentToday', '0');
    const today = new Date().toISOString().split('T')[0];
    userProperties.setProperty('lastCountDate', today);
  }

  function testQuota() {
    const result = getEmailQuotaInfo();
    Logger.log("QUOTA TEST RESULT: " + JSON.stringify(result));
    return result;
  }