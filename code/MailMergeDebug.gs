
  
  /**
   * Shows a dialog with all property stores for debugging
   */
  function showDebugStorage() {
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
   * Shows a dialog with tools to fix configuration storage
   */
  function showFixerDialog() {
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
   * Shows the clear properties dialog
   */
  function showClearDialog() {
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
   * Shows template repair dialog
   */
  function showRepairDialog() {
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
   * Shows dialog with storage usage information
   */
  function showStorageUsage() {
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