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
   * Shows a simple dialog with storage usage information
   */
  function showStorageUsageDialog() {
    try {
      // Get configurations
      const docProps = PropertiesService.getDocumentProperties();
      const userProps = PropertiesService.getUserProperties();
      const scriptProps = PropertiesService.getScriptProperties();
      
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
        
        <script>
          // Helper function to format bytes
          function formatBytes(bytes) {
            return bytes + " bytes";
          }
        </script>
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
   * Shows a dialog to send a test email directly
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
      
      // Send the email
      GmailApp.sendEmail(recipient, subject, '', {
        htmlBody: body
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