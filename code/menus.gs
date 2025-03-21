/**
 * Master onOpen function with emojified menu
 * Includes all regular functions and troubleshooting tools
 */
function onOpen() {
  DocumentApp.getUi()
      .createMenu('📧 Mail Merge')
      
      // Main Functions
      .addItem('📋 Open Mail Merge Sidebar', 'showSidebar')
      .addSeparator()
      
      // Template Management
      .addItem('📝 Create New Template', 'createNewTemplate')
      .addItem('📂 Load Template to Document', 'showLoadTemplateDialog')
      .addItem('💾 Save Template to Storage', 'saveTemplateAndShowResult')
      .addSeparator()
      
      // Testing & Utilities
      .addItem('✉️ Send Test Email', 'showTestEmailDialog')
      .addItem('📤 Backup Templates to Email', 'backupTemplatesToEmailAndShowResult')
      .addSeparator()
      
      // Diagnostics
      .addItem('🔍 Run Permission Diagnostics', 'showPermissionDiagnostics')
      
      // Troubleshooting Submenu - Using Function Names That Definitely Exist
      .addSubMenu(DocumentApp.getUi().createMenu('🛠️ Troubleshooting')
          .addItem('🔎 Debug Storage Issues', 'showDebugStorage')
          .addItem('🔄 Fix Configuration Storage', 'showFixerDialog')
          .addItem('🔧 Repair Template Structures', 'showRepairDialog')
          .addItem('🧹 Clear Document Properties', 'showClearDialog')
          .addItem('📊 Show Storage Usage', 'showStorageUsage')
      )
      .addToUi();
}

/**
* Shows a dialog to send a test email directly.
*/
function showTestEmailDialog() {
  const html = `
  <html>
  <head>
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; }
      h2 { color: #4285f4; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; color: #5f6368; }
      input, textarea { width: 100%; padding: 8px; border: 1px solid #dadce0; border-radius: 4px; box-sizing: border-box; }
      button { padding: 8px 16px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      .actions { margin-top: 20px; }
    </style>
  </head>
  <body>
    <h2>✉️ Send Test Email</h2>
    
    <div class="form-group">
      <label for="recipient">Recipient Email:</label>
      <input type="email" id="recipient" placeholder="you@example.com" required>
    </div>
    
    <div class="form-group">
      <label for="subject">Subject:</label>
      <input type="text" id="subject" placeholder="Test Email Subject" required>
    </div>
    
    <div class="form-group">
      <label for="body">Email Body:</label>
      <textarea id="body" rows="8" placeholder="Enter your email content here..."></textarea>
    </div>
    
    <div class="actions">
      <button onclick="sendTestEmail()">Send Test Email</button>
      <button onclick="google.script.host.close()" style="background-color: #f1f3f4; color: #202124; border: 1px solid #dadce0; margin-left: 8px;">Cancel</button>
    </div>
    
    <script>
      // Pre-fill with document content if available
      google.script.run
        .withSuccessHandler(function(content) {
          document.getElementById('body').value = content;
        })
        .getDocContent();
      
      function sendTestEmail() {
        const recipient = document.getElementById('recipient').value.trim();
        const subject = document.getElementById('subject').value.trim();
        const body = document.getElementById('body').value.trim();
        
        if (!recipient) {
          alert('Please enter a recipient email address');
          return;
        }
        
        if (!subject) {
          alert('Please enter a subject line');
          return;
        }
        
        // Disable button and show sending state
        const buttons = document.querySelectorAll('button');
        buttons.forEach(btn => btn.disabled = true);
        buttons[0].innerHTML = 'Sending...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              alert('Test email sent successfully!');
              google.script.host.close();
            } else {
              alert('Error sending email: ' + result.message);
              // Re-enable buttons
              buttons.forEach(btn => btn.disabled = false);
              buttons[0].innerHTML = 'Send Test Email';
            }
          })
          .withFailureHandler(function(error) {
            alert('Error: ' + error.message);
            // Re-enable buttons
            buttons.forEach(btn => btn.disabled = false);
            buttons[0].innerHTML = 'Send Test Email';
          })
          .sendSimpleTestEmail(recipient, subject, body);
      }
    </script>
  </body>
  </html>
  `;
  
  const ui = DocumentApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600)
    .setTitle('Send Test Email');
  
  ui.showModalDialog(htmlOutput, 'Send Test Email');
}

/**
* Sends a simple test email without mail merge functionality
*/
function sendSimpleTestEmail(recipient, subject, body) {
try {
  // Basic validation
  if (!recipient || !subject) {
    return { 
      success: false, 
      message: 'Recipient and subject are required'
    };
  }
  
  // Get email content
  const htmlBody = body || getDocContent();
  
  // Get sender information
  const senderDisplayName = APP_DEFAULTS.senderDisplayName;
  
  // Send the email
  GmailApp.sendEmail(recipient, subject, "", {
    name: senderDisplayName,
    htmlBody: htmlBody
  });
  
  return {
    success: true,
    message: 'Email sent successfully'
  };
} catch (e) {
  return {
    success: false,
    message: e.message
  };
}
}

/**
* Shows a simple dialog with storage usage information
*/
function showStorageUsage() {
try {
  // Get configurations
  const docProps = PropertiesService.getDocumentProperties();
  const mailMergeConfigsJson = docProps.getProperty('mailMergeConfigs') || '{}';
  const configs = JSON.parse(mailMergeConfigsJson);
  
  // Calculate sizes
  const totalSize = mailMergeConfigsJson.length;
  const templateCount = Object.keys(configs).length;
  
  // Calculate size per template
  const templateSizes = {};
  let largestTemplate = { name: '', size: 0 };
  
  for (const [name, config] of Object.entries(configs)) {
    const configJson = JSON.stringify(config);
    const size = configJson.length;
    templateSizes[name] = size;
    
    if (size > largestTemplate.size) {
      largestTemplate = { name, size };
    }
  }
  
  // Create HTML for dialog
  let html = `
  <html>
  <head>
    <style>
      body { font-family: Arial, sans-serif; margin: 20px; }
      h2 { color: #4285f4; }
      .panel { border: 1px solid #dadce0; border-radius: 4px; padding: 15px; margin: 15px 0; }
      .progress-outer { background-color: #f1f3f4; height: 8px; border-radius: 4px; margin: 8px 0; }
      .progress-inner { background-color: #4285f4; height: 100%; border-radius: 4px; }
      table { width: 100%; border-collapse: collapse; }
      th, td { text-align: left; padding: 8px; border-bottom: 1px solid #dadce0; }
      th { background-color: #f8f9fa; }
      .size-bar { background-color: #4285f4; height: 6px; border-radius: 3px; }
    </style>
  </head>
  <body>
    <h2>📊 Mail Merge Storage Usage</h2>
    
    <div class="panel">
      <h3>Summary</h3>
      <p>Total Templates: ${templateCount}</p>
      <p>Total Storage Used: ${formatBytes(totalSize)}</p>
      <p>Largest Template: "${largestTemplate.name}" (${formatBytes(largestTemplate.size)})</p>
      
      <p>Storage Usage:</p>
      <div class="progress-outer">
        <div class="progress-inner" style="width: ${Math.min(100, (totalSize / 100000) * 100)}%;"></div>
      </div>
      <p style="text-align: right; font-size: 12px; color: #5f6368;">${formatBytes(totalSize)} of 100KB limit</p>
    </div>
  `;
  
  // Add template details if there are any
  if (templateCount > 0) {
    html += `
    <div class="panel">
      <h3>Template Details</h3>
      <table>
        <tr>
          <th>Template Name</th>
          <th>Size</th>
          <th>Percentage</th>
        </tr>
    `;
    
    // Sort templates by size (largest first)
    const sortedTemplates = Object.entries(templateSizes)
      .sort((a, b) => b[1] - a[1])
      .map(([name, size]) => ({name, size}));
    
    for (const template of sortedTemplates) {
      const percentage = (template.size / totalSize) * 100;
      html += `
        <tr>
          <td>${template.name}</td>
          <td>${formatBytes(template.size)}</td>
          <td>
            <div style="display: flex; align-items: center;">
              <div class="size-bar" style="width: ${percentage}%; margin-right: 8px;"></div>
              ${percentage.toFixed(1)}%
            </div>
          </td>
        </tr>
      `;
    }
    
    html += `
      </table>
    </div>
    `;
  }
  
  html += `
    <button onclick="google.script.host.close()" style="padding: 8px 16px;">Close</button>
  </body>
  </html>
  `;
  
  // Show dialog
  const ui = DocumentApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(500)
    .setTitle('Mail Merge Storage Usage');
  
  ui.showModalDialog(htmlOutput, 'Mail Merge Storage Usage');
} catch (e) {
  DocumentApp.getUi().alert('Error showing storage usage: ' + e.message);
}
}

/**
* Helper function to format bytes in a human-readable way
*/
function formatBytes(bytes) {
if (bytes < 1024) return bytes + " bytes";
else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
else return (bytes / 1048576).toFixed(1) + " MB";
}

/**
 * Shows permission diagnostics dialog.
 */
function showPermissionDiagnostics() {
  const permissionStatus = {};
  
  try {
    // Test document access
    permissionStatus.document = {
      access: true,
      message: "Document access granted"
    };
    DocumentApp.getActiveDocument().getBody().getText();
  } catch (e) {
    permissionStatus.document = {
      access: false,
      message: "Document access error: " + e.message
    };
  }
  
  try {
    // Test email access
    permissionStatus.email = {
      access: true,
      message: "Email access granted. Quota: " + MailApp.getRemainingDailyQuota()
    };
  } catch (e) {
    permissionStatus.email = {
      access: false,
      message: "Email access error: " + e.message
    };
  }
  
  try {
    // Test user info access
    const email = Session.getActiveUser().getEmail();
    permissionStatus.userInfo = {
      access: email && email.length > 0,
      message: email ? "User info access granted: " + email : "Could not get user email"
    };
  } catch (e) {
    permissionStatus.userInfo = {
      access: false,
      message: "User info access error: " + e.message
    };
  }
  
  try {
    // Test properties access
    PropertiesService.getDocumentProperties().setProperty('testProperty', 'test');
    PropertiesService.getDocumentProperties().deleteProperty('testProperty');
    permissionStatus.properties = {
      access: true,
      message: "Properties access granted"
    };
  } catch (e) {
    permissionStatus.properties = {
      access: false,
      message: "Properties access error: " + e.message
    };
  }
  
  try {
    // Test Gmail API access - advanced service
    try {
      Gmail.Users.getProfile('me');
      permissionStatus.gmailApi = {
        access: true,
        message: "Gmail API advanced service access granted"
      };
    } catch (e) {
      permissionStatus.gmailApi = {
        access: false,
        message: "Gmail API access not available: " + e.message
      };
    }
  } catch (e) {
    // Gmail API is not enabled, this is normal
    permissionStatus.gmailApi = {
      access: false,
      message: "Gmail API advanced service not enabled"
    };
  }
  
  try {
    // Test spreadsheet access
    const ss = SpreadsheetApp.create("Test Spreadsheet");
    const ssId = ss.getId();
    
    permissionStatus.spreadsheet = {
      access: true,
      message: "Spreadsheet access granted. Test spreadsheet created."
    };
    
    // Clean up
    DriveApp.getFileById(ssId).setTrashed(true);
  } catch (e) {
    permissionStatus.spreadsheet = {
      access: false,
      message: "Spreadsheet access error: " + e.message
    };
  }
  
  // Create HTML to display results
  let html = '<html><head>';
  html += '<style>';
  html += 'body { font-family: Arial, sans-serif; margin: 20px; }';
  html += 'h2 { color: #4285f4; }';
  html += '.status-card { margin-bottom: 20px; border: 1px solid #dadce0; border-radius: 8px; overflow: hidden; }';
  html += '.status-header { padding: 10px 15px; font-weight: bold; }';
  html += '.status-content { padding: 15px; }';
  html += '.success { background-color: #e6f4ea; color: #137333; }';
  html += '.error { background-color: #fce8e6; color: #b31412; }';
  html += '.warning { background-color: #fef7e0; color: #994c00; }';
  html += '.actions { margin-top: 20px; }';
  html += 'button { padding: 8px 16px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }';
  html += '.secondary { background-color: #f1f3f4; color: #202124; border: 1px solid #dadce0; }';
  html += '</style>';
  html += '</head><body>';
  
  html += '<h2>Mail Merge Permission Diagnostics</h2>';
  
  // Generate status cards
  for (const [key, status] of Object.entries(permissionStatus)) {
    html += '<div class="status-card">';
    html += `<div class="status-header ${status.access ? 'success' : 'error'}">`;
    html += `${key.charAt(0).toUpperCase() + key.slice(1)}: ${status.access ? 'OK' : 'ERROR'}`;
    html += '</div>';
    html += `<div class="status-content">${status.message}</div>`;
    html += '</div>';
  }
  
  // Add overall summary
  const requiredPermissions = ['document', 'email', 'properties'];
  const allRequiredPermissionsGranted = requiredPermissions.every(perm => 
    permissionStatus[perm] && permissionStatus[perm].access);
  
  const allPermissionsGranted = Object.values(permissionStatus).every(status => status.access);
  
  let summaryClass = 'error';
  let summaryTitle = 'Missing critical permissions';
  
  if (allRequiredPermissionsGranted) {
    summaryClass = allPermissionsGranted ? 'success' : 'warning';
    summaryTitle = allPermissionsGranted ? 
      'All permissions granted' : 
      'Core functionality available, but some features limited';
  }
  
  html += '<div class="status-card">';
  html += `<div class="status-header ${summaryClass}">`;
  html += `Overall Status: ${summaryTitle}`;
  html += '</div>';
  html += '<div class="status-content">';
  
  if (allRequiredPermissionsGranted) {
    if (allPermissionsGranted) {
      html += 'Mail Merge has all permissions needed for full functionality.';
    } else {
      html += 'Mail Merge can operate with basic functionality, but some advanced features may be limited.';
      html += '<ul>';
      if (!permissionStatus.userInfo.access) {
        html += '<li>Cannot automatically detect your email address and name</li>';
      }
      if (!permissionStatus.gmailApi.access) {
        html += '<li>Cannot use advanced Gmail features like drafts</li>';
      }
      if (!permissionStatus.spreadsheet.access) {
        html += '<li>Limited spreadsheet creation capabilities</li>';
      }
      html += '</ul>';
    }
  } else {
    html += '<strong>Critical permissions are missing. Mail Merge may not function correctly.</strong>';
    html += '<ul>';
    if (!permissionStatus.document.access) {
      html += '<li>Cannot access document content - required for template creation</li>';
    }
    if (!permissionStatus.email.access) {
      html += '<li>Cannot send emails - core Mail Merge functionality</li>';
    }
    if (!permissionStatus.properties.access) {
      html += '<li>Cannot save templates or configurations</li>';
    }
    html += '</ul>';
  }
  
  html += '</div>';
  html += '</div>';
  
  // Add suggested actions
  html += '<div class="actions">';
  html += '<h3>Suggested Actions</h3>';
  html += '<p>If you\'re experiencing permission issues:</p>';
  html += '<ol>';
  html += '<li>Click the refresh button below to re-check permissions</li>';
  html += '<li>Reload the add-on from the Add-ons menu</li>';
  html += '<li>Reinstall the add-on if permissions are still missing</li>';
  html += '</ol>';
  html += '</div>';
  
  // Add buttons
  html += '<div class="actions">';
  html += '<button onclick="refreshPermissions()">Refresh Permissions</button>';
  html += '<button onclick="google.script.host.close()" class="secondary" style="margin-left: 10px;">Close</button>';
  html += '</div>';
  
  // Add script for refresh
  html += '<script>';
  html += 'function refreshPermissions() {';
  html += '  google.script.run';
  html += '    .withSuccessHandler(function() {';
  html += '      window.location.reload();';
  html += '    })';
  html += '    .showPermissionDiagnostics();';
  html += '}';
  html += '</script>';
  
  html += '</body></html>';
  
  const ui = DocumentApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600)
    .setTitle('Mail Merge Permission Diagnostics');
  
  ui.showModalDialog(htmlOutput, 'Mail Merge Permission Diagnostics');
}