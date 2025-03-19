# Mail Merge Configuration Management Specification

## 1. Overview

This specification outlines the implementation of a configuration management system for the Mail Merge Google Docs add-on. The system will allow users to save, load, and manage mail merge configurations using document properties, while maintaining full access to modify all settings after loading.

## 2. Goals

- Enable users to save complete mail merge configurations in document properties
- Allow loading of saved configurations without restricting subsequent modifications
- Provide dedicated UI for managing configurations
- Support multiple named configurations per document
- Create an extensible system for future enhancements

## 3. User Experience

### 3.1 Configuration Selection and Loading

1. Users can select a saved configuration from a dropdown in the sidebar
2. Clicking "Load" applies the configuration to all relevant fields
3. All mail merge panels remain accessible and all fields remain editable
4. A status indicator shows which configuration is currently active
5. Users can modify any loaded configuration settings as needed

### 3.2 Configuration Management

1. Clicking "Manage" opens a modal dialog
2. Dialog allows users to:
   - View all saved configurations
   - Create new configurations
   - Edit existing configurations
   - Delete unwanted configurations
3. When creating/editing, users can:
   - Specify which settings to include in the configuration
   - Group settings by category for easier management
   - Preview what will be applied when loaded

## 4. UI Components

### 4.1 Sidebar Configuration Section

**Location:** Added as the first section (00) in the Mail Merge sidebar

**Components:**
- **Section Header:** "Configuration" (collapsible like other sections)
- **Configuration Selector:** Dropdown listing all saved configurations
- **Load Button:** Applies the selected configuration
- **Manage Button:** Opens the configuration management dialog
- **Status Indicator:** Shows currently active configuration (if any)

**Example HTML:**
```html
<div id="configuration" class="section">
  <div class="section-indicator">
    <div class="section-indicator-bar" style="height: 60px;"></div>
  </div>
  
  <div class="section-header" onclick="toggleSection('configuration')" aria-expanded="false">
    <div class="section-title">
      <span class="section-number">00</span>
      Configuration
    </div>
    <div class="section-toggle">
      <div class="toggle-indicator" aria-hidden="true">+</div>
      <div class="toggle-bar"></div>
    </div>
  </div>
  
  <div id="configurationContent" class="section-content">
    <div class="form-group">
      <label for="configSelect">Saved Configurations</label>
      <div class="input-group">
        <select id="configSelect" aria-label="Select configuration">
          <option value="">Select a configuration...</option>
          <!-- Options populated by JavaScript -->
        </select>
      </div>
    </div>
    
    <div class="form-group flex gap-2">
      <button id="loadConfigBtn" onclick="loadSelectedConfiguration()" class="flex-1">
        <span style="margin-right: 6px;">▶</span> Load
      </button>
      <button id="manageConfigBtn" onclick="openConfigDialog()" class="flex-1 secondary">
        <span style="margin-right: 6px;">⚙️</span> Manage
      </button>
    </div>
    
    <div id="activeConfigIndicator" class="alert alert-info hidden">
      <div class="alert-icon">ℹ</div>
      <div>Using: <span id="activeConfigName"></span></div>
    </div>
  </div>
</div>
```

### 4.2 Configuration Management Dialog

**Components:**
- **Dialog Header:** "Manage Mail Merge Configurations"
- **List View:**
  - List of saved configurations
  - New configuration button
  - Edit and Delete buttons for each configuration
- **Edit/Create View:**
  - Configuration name field
  - Grouped settings with checkboxes
  - Categorized sections for better organization
  - Save/Cancel buttons

**Example HTML (to be created via HtmlService):**
```html
<div class="dialog">
  <div class="dialog-header">
    <h3>Manage Mail Merge Configurations</h3>
    <button onclick="closeDialog()" class="close-btn">×</button>
  </div>
  
  <div class="dialog-content">
    <!-- List View -->
    <div id="configListView">
      <div class="dialog-toolbar">
        <h4>Saved Configurations</h4>
        <button onclick="showCreateView()" class="create-btn">+ New Configuration</button>
      </div>
      
      <ul id="configList" class="config-list">
        <!-- List items populated by JavaScript -->
      </ul>
    </div>
    
    <!-- Edit/Create View (initially hidden) -->
    <div id="configEditView" class="hidden">
      <div class="dialog-toolbar">
        <h4 id="editViewTitle">Create New Configuration</h4>
      </div>
      
      <div class="form-group">
        <label for="configName">Configuration Name</label>
        <input type="text" id="configName" placeholder="e.g., Monthly Newsletter">
      </div>
      
      <!-- Data Source Settings -->
      <div class="settings-group">
        <div class="settings-header">
          <input type="checkbox" id="dataSourceGroup" checked>
          <label for="dataSourceGroup">Data Source Settings</label>
        </div>
        <div class="settings-content">
          <div class="setting-item">
            <input type="checkbox" id="includeSpreadsheetUrl" checked>
            <div class="setting-input">
              <label for="spreadsheetUrlValue">Spreadsheet URL</label>
              <input type="text" id="spreadsheetUrlValue">
            </div>
          </div>
          <!-- Additional settings... -->
        </div>
      </div>
      
      <!-- Email Settings -->
      <div class="settings-group">
        <div class="settings-header">
          <input type="checkbox" id="emailSettingsGroup" checked>
          <label for="emailSettingsGroup">Email Settings</label>
        </div>
        <div class="settings-content">
          <!-- Email settings... -->
        </div>
      </div>
    </div>
  </div>
  
  <div class="dialog-footer">
    <button onclick="closeDialog()" class="secondary-btn">Cancel</button>
    <button id="saveConfigBtn" onclick="saveConfiguration()" class="primary-btn">Save</button>
  </div>
</div>
```

## 5. Data Model and Storage

### 5.1 Configuration Object Model

```javascript
/**
 * Represents a complete Mail Merge configuration.
 */
class MailMergeConfig {
  constructor(name) {
    this.version = "1.0";
    this.name = name || "Unnamed Configuration";
    this.timestamp = new Date().toISOString();
    this.settings = {
      // Data source settings
      spreadsheetUrl: { enabled: false, value: "" },
      sheetName: { enabled: false, value: "" },
      recipientColumn: { enabled: false, value: "" },
      ccColumn: { enabled: false, value: "" },
      bccColumn: { enabled: false, value: "" },
      
      // Sender settings
      fromName: { enabled: false, value: "" },
      fromEmail: { enabled: false, value: "" },
      
      // Content settings
      subjectLine: { enabled: false, value: "" },
      
      // Override settings
      ccOverride: { enabled: false, value: "" },
      bccOverride: { enabled: false, value: "" }
    };
  }
  
  /**
   * Creates a configuration from the current UI state.
   * @param {string} name - The name for the configuration.
   * @return {MailMergeConfig} The created configuration.
   */
  static fromCurrentState(name) {
    // Implementation details...
  }
  
  /**
   * Validates the configuration.
   * @return {Object} Validation results.
   */
  validate() {
    // Implementation details...
  }
}
```

### 5.2 Document Properties Storage

Configurations will be stored in document properties as stringified JSON:

```javascript
/**
 * Saves a configuration to document properties.
 * @param {MailMergeConfig} config - The configuration to save.
 */
function saveConfigToProperties(config) {
  const docProperties = PropertiesService.getDocumentProperties();
  
  // Get existing configurations
  const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
  const configs = JSON.parse(configsJson);
  
  // Add or update the configuration
  configs[config.name] = config;
  
  // Save back to document properties
  docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
}

/**
 * Gets all configurations from document properties.
 * @return {Object} Object mapping configuration names to configurations.
 */
function getConfigsFromProperties() {
  const docProperties = PropertiesService.getDocumentProperties();
  const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
  return JSON.parse(configsJson);
}

/**
 * Deletes a configuration from document properties.
 * @param {string} name - The name of the configuration to delete.
 */
function deleteConfigFromProperties(name) {
  const docProperties = PropertiesService.getDocumentProperties();
  const configsJson = docProperties.getProperty('mailMergeConfigs') || '{}';
  const configs = JSON.parse(configsJson);
  
  if (configs[name]) {
    delete configs[name];
    docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
  }
}
```

## 6. Implementation Details

### 6.1 Backend (backend.gs)

#### 6.1.1 Document Property Management

```javascript
/**
 * Gets all available configurations from document properties.
 * @return {Object} Configurations object.
 */
function getAvailableConfigurations() {
  return getConfigsFromProperties();
}

/**
 * Saves the current UI state as a new configuration.
 * @param {string} name - The name for the configuration.
 * @return {boolean} Success flag.
 */
function saveCurrentAsConfiguration(name) {
  try {
    const config = MailMergeConfig.fromCurrentState(name);
    saveConfigToProperties(config);
    return true;
  } catch (e) {
    Logger.log('Error saving configuration: ' + e.message);
    return false;
  }
}

/**
 * Gets a specific configuration by name.
 * @param {string} name - The configuration name.
 * @return {MailMergeConfig|null} The configuration or null if not found.
 */
function getConfiguration(name) {
  const configs = getConfigsFromProperties();
  return configs[name] || null;
}

/**
 * Deletes a configuration.
 * @param {string} name - The configuration name.
 * @return {boolean} Success flag.
 */
function deleteConfiguration(name) {
  try {
    deleteConfigFromProperties(name);
    return true;
  } catch (e) {
    Logger.log('Error deleting configuration: ' + e.message);
    return false;
  }
}

/**
 * Updates an existing configuration.
 * @param {MailMergeConfig} config - The updated configuration.
 * @return {boolean} Success flag.
 */
function updateConfiguration(config) {
  try {
    saveConfigToProperties(config);
    return true;
  } catch (e) {
    Logger.log('Error updating configuration: ' + e.message);
    return false;
  }
}
```

#### 6.1.2 Configuration Dialog Management

```javascript
/**
 * Shows the configuration management dialog.
 */
function showConfigDialog() {
  const html = HtmlService.createTemplateFromFile('ConfigDialog')
      .evaluate()
      .setWidth(500)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  DocumentApp.getUi().showModalDialog(html, 'Manage Mail Merge Configurations');
}

/**
 * Gets the current UI state for the dialog.
 * @return {Object} The current UI state.
 */
function getCurrentStateForDialog() {
  const state = {
    spreadsheetUrl: getElementValue('spreadsheetUrl'),
    sheetName: getElementValue('sheetSelect'),
    recipientColumn: getElementValue('emailColumnSelect'),
    ccColumn: getElementValue('ccColumnSelect'),
    bccColumn: getElementValue('bccColumnSelect'),
    fromName: getElementValue('fromName'),
    fromEmail: getElementValue('fromEmailSelect'),
    subjectLine: getElementValue('subjectLine'),
    ccOverride: getElementValue('ccField') || '',
    bccOverride: getElementValue('bccField') || ''
  };
  
  return state;
}
```

### 6.2 Frontend (JavaScriptFile.html)

#### 6.2.1 Configuration List Management

```javascript
// Function to load configuration list
function loadConfigurationList() {
  const configSelect = document.getElementById('configSelect');
  configSelect.innerHTML = '<option value="">Select a configuration...</option>';
  
  google.script.run
    .withSuccessHandler(function(configs) {
      // Store configurations in memory
      window.availableConfigurations = configs;
      
      // Add to dropdown
      for (const name in configs) {
        if (configs.hasOwnProperty(name)) {
          const option = document.createElement('option');
          option.value = name;
          option.textContent = name;
          configSelect.appendChild(option);
        }
      }
      
      if (Object.keys(configs).length === 0) {
        showNotification('No saved configurations found', 'info', config.toastDuration);
      }
    })
    .withFailureHandler(function(error) {
      showNotification('Error loading configurations: ' + error.message, 'error', config.toastDuration);
    })
    .getAvailableConfigurations();
}
```

#### 6.2.2 Configuration Loading

```javascript
// Function to load a selected configuration
function loadSelectedConfiguration() {
  const configName = document.getElementById('configSelect').value;
  if (!configName) {
    showNotification('Please select a configuration', 'error', config.toastDuration);
    return;
  }
  
  showNotification(`Loading configuration "${configName}"...`, 'info', config.toastDuration);
  
  google.script.run
    .withSuccessHandler(function(loadedConfig) {
      if (!loadedConfig) {
        showNotification(`Configuration "${configName}" not found`, 'error', config.toastDuration);
        return;
      }
      
      // Apply configuration to UI
      applyConfigurationToUI(loadedConfig);
      
      // Show active configuration indicator
      const activeIndicator = document.getElementById('activeConfigIndicator');
      const activeConfigName = document.getElementById('activeConfigName');
      
      if (activeIndicator && activeConfigName) {
        activeConfigName.textContent = configName;
        activeIndicator.classList.remove('hidden');
      }
      
      showNotification(`Configuration "${configName}" loaded successfully`, 'success', config.toastDuration);
    })
    .withFailureHandler(function(error) {
      showNotification('Error loading configuration: ' + error.message, 'error', config.toastDuration);
    })
    .getConfiguration(configName);
}

// Function to apply configuration to UI
function applyConfigurationToUI(config) {
  // Apply settings in the correct sequence, respecting dependencies
  
  // First, apply spreadsheet URL if enabled
  if (config.settings.spreadsheetUrl.enabled && config.settings.spreadsheetUrl.value) {
    document.getElementById('spreadsheetUrl').value = config.settings.spreadsheetUrl.value;
    validateSpreadsheet();
    
    // Create a queue for subsequent operations
    applyRemainingSettingsInSequence(config);
  } else {
    // If no spreadsheet URL, just apply other settings directly
    applyNonSpreadsheetSettings(config);
  }
}

// Function to apply remaining settings in sequence
function applyRemainingSettingsInSequence(config) {
  // Wait for spreadsheet validation to complete
  setTimeout(function() {
    // Apply sheet name if enabled
    if (config.settings.sheetName.enabled && config.settings.sheetName.value) {
      const sheetSelect = document.getElementById('sheetSelect');
      selectOptionByValue(sheetSelect, config.settings.sheetName.value);
      loadSheetColumns();
      
      // Wait for columns to load before applying column selections
      setTimeout(function() {
        applyColumnSelections(config);
      }, 1500);
    }
  }, 1500);
}

// Function to apply column selections
function applyColumnSelections(config) {
  // Apply recipient column if enabled
  if (config.settings.recipientColumn.enabled && config.settings.recipientColumn.value) {
    selectOptionByValue(document.getElementById('emailColumnSelect'), config.settings.recipientColumn.value);
  }
  
  // Make CC/BCC visible if needed
  if ((config.settings.ccColumn.enabled && config.settings.ccColumn.value) || 
      (config.settings.bccColumn.enabled && config.settings.bccColumn.value)) {
    // Show CC/BCC fields if hidden
    if (document.getElementById('ccColumnGroup').classList.contains('hidden')) {
      toggleCcBcc();
    }
    
    // Apply CC column if enabled
    if (config.settings.ccColumn.enabled && config.settings.ccColumn.value) {
      selectOptionByValue(document.getElementById('ccColumnSelect'), config.settings.ccColumn.value);
    }
    
    // Apply BCC column if enabled
    if (config.settings.bccColumn.enabled && config.settings.bccColumn.value) {
      selectOptionByValue(document.getElementById('bccColumnSelect'), config.settings.bccColumn.value);
    }
  }
  
  // Apply remaining non-spreadsheet settings
  applyNonSpreadsheetSettings(config);
}

// Function to apply non-spreadsheet dependent settings
function applyNonSpreadsheetSettings(config) {
  // Apply subject line if enabled
  if (config.settings.subjectLine.enabled && config.settings.subjectLine.value) {
    document.getElementById('subjectLine').value = config.settings.subjectLine.value;
    updateCharCounter();
  }
  
  // Apply from name if enabled
  if (config.settings.fromName.enabled && config.settings.fromName.value) {
    document.getElementById('fromName').value = config.settings.fromName.value;
  }
  
  // Apply from email if enabled
  if (config.settings.fromEmail.enabled && config.settings.fromEmail.value) {
    selectOptionByValue(document.getElementById('fromEmailSelect'), config.settings.fromEmail.value);
  }
  
  // Apply CC override if enabled
  if (config.settings.ccOverride.enabled && document.getElementById('ccField')) {
    document.getElementById('ccField').value = config.settings.ccOverride.value;
  }
  
  // Apply BCC override if enabled
  if (config.settings.bccOverride.enabled && document.getElementById('bccField')) {
    document.getElementById('bccField').value = config.settings.bccOverride.value;
  }
  
  // Update summary to reflect all changes
  updateSummary();
}

// Helper function to select an option by value
function selectOptionByValue(selectElement, value) {
  if (!selectElement) return false;
  
  for (let i = 0; i < selectElement.options.length; i++) {
    if (selectElement.options[i].value === value) {
      selectElement.selectedIndex = i;
      // Trigger change event to activate any listeners
      const event = new Event('change');
      selectElement.dispatchEvent(event);
      return true;
    }
  }
  
  return false;
}
```

#### 6.2.3 Configuration Dialog Management

```javascript
// Function to open configuration management dialog
function openConfigDialog() {
  google.script.run.showConfigDialog();
}
```

### 6.3 Configuration Dialog (ConfigDialog.html)

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    /* Dialog styles */
  </style>
  <script>
    // Dialog initialization
    document.addEventListener('DOMContentLoaded', function() {
      loadConfigurations();
    });
    
    // Function to load configurations into dialog
    function loadConfigurations() {
      google.script.run
        .withSuccessHandler(function(configs) {
          // Populate configuration list
          const configList = document.getElementById('configList');
          configList.innerHTML = '';
          
          if (Object.keys(configs).length === 0) {
            configList.innerHTML = '<li class="empty-message">No saved configurations</li>';
            return;
          }
          
          for (const name in configs) {
            if (configs.hasOwnProperty(name)) {
              const config = configs[name];
              const listItem = document.createElement('li');
              listItem.className = 'config-item';
              
              listItem.innerHTML = `
                <span class="config-name">${name}</span>
                <div class="config-actions">
                  <button onclick="editConfiguration('${name}')" class="edit-btn">Edit</button>
                  <button onclick="confirmDeleteConfiguration('${name}')" class="delete-btn">Delete</button>
                </div>
              `;
              
              configList.appendChild(listItem);
            }
          }
        })
        .withFailureHandler(function(error) {
          showError('Failed to load configurations: ' + error.message);
        })
        .getAvailableConfigurations();
    }
    
    // Function to show the create view
    function showCreateView() {
      document.getElementById('configListView').classList.add('hidden');
      document.getElementById('configEditView').classList.remove('hidden');
      document.getElementById('editViewTitle').textContent = 'Create New Configuration';
      document.getElementById('configName').value = '';
      
      // Get current state for defaults
      google.script.run
        .withSuccessHandler(function(state) {
          // Populate form with current values
          populateFormWithState(state);
        })
        .getCurrentStateForDialog();
      
      // Update footer buttons
      document.getElementById('saveConfigBtn').textContent = 'Create';
      document.getElementById('saveConfigBtn').onclick = createConfiguration;
    }
    
    // Function to edit a configuration
    function editConfiguration(name) {
      document.getElementById('configListView').classList.add('hidden');
      document.getElementById('configEditView').classList.remove('hidden');
      document.getElementById('editViewTitle').textContent = 'Edit Configuration';
      document.getElementById('configName').value = name;
      
      // Load configuration data
      google.script.run
        .withSuccessHandler(function(config) {
          if (!config) {
            showError('Configuration not found');
            return;
          }
          
          // Populate form with configuration values
          populateFormWithConfig(config);
        })
        .getConfiguration(name);
      
      // Update footer buttons
      document.getElementById('saveConfigBtn').textContent = 'Save Changes';
      document.getElementById('saveConfigBtn').onclick = function() { updateConfiguration(name); };
    }
    
    // Function to populate form with state
    function populateFormWithState(state) {
      // Set all checkboxes to true for a new configuration
      // and populate values from current state
      document.getElementById('includeSpreadsheetUrl').checked = true;
      document.getElementById('spreadsheetUrlValue').value = state.spreadsheetUrl || '';
      
      // Additional field population...
    }
    
    // Function to populate form with config
    function populateFormWithConfig(config) {
      // Set checkboxes and values based on configuration
      document.getElementById('includeSpreadsheetUrl').checked = config.settings.spreadsheetUrl.enabled;
      document.getElementById('spreadsheetUrlValue').value = config.settings.spreadsheetUrl.value;
      
      // Additional field population...
    }
    
    // Function to create a new configuration
    function createConfiguration() {
      const name = document.getElementById('configName').value.trim();
      if (!name) {
        showError('Please enter a configuration name');
        return;
      }
      
      // Build configuration object from form
      const config = buildConfigFromForm(name);
      
      // Save to document properties
      google.script.run
        .withSuccessHandler(function(success) {
          if (success) {
            // Return to list view and refresh
            returnToListView();
            loadConfigurations();
            
            // Also refresh sidebar configuration list
            google.script.run.refreshSidebarConfigList();
          } else {
            showError('Failed to save configuration');
          }
        })
        .withFailureHandler(function(error) {
          showError('Error: ' + error.message);
        })
        .updateConfiguration(config);
    }
    
    // Function to update an existing configuration
    function updateConfiguration(name) {
      const newName = document.getElementById('configName').value.trim();
      if (!newName) {
        showError('Please enter a configuration name');
        return;
      }
      
      // Build configuration object from form
      const config = buildConfigFromForm(newName);
      
      // Save to document properties
      google.script.run
        .withSuccessHandler(function(success) {
          if (success) {
            // If name changed, delete old configuration
            if (name !== newName) {
              google.script.run.deleteConfiguration(name);
            }
            
            // Return to list view and refresh
            returnToListView();
            loadConfigurations();
            
            // Also refresh sidebar configuration list
            google.script.run.refreshSidebarConfigList();
          } else {
            showError('Failed to update configuration');
          }
        })
        .withFailureHandler(function(error) {
          showError('Error: ' + error.message);
        })
        .updateConfiguration(config);
    }
    
    // Function to build configuration from form
    function buildConfigFromForm(name) {
      const config = {
        version: "1.0",
        name: name,
        timestamp: new Date().toISOString(),
        settings: {
          spreadsheetUrl: {
            enabled: document.getElementById('includeSpreadsheetUrl').checked,
            value: document.getElementById('spreadsheetUrlValue').value
          },
          // Additional fields...
        }
      };
      
      return config;
    }
    
    // Function to confirm configuration deletion
    function confirmDeleteConfiguration(name) {
      if (confirm('Are you sure you want to delete "' + name + '"? This cannot be undone.')) {
        deleteConfiguration(name);
      }
    }
    
    // Function to delete a configuration
    function deleteConfiguration(name) {
      google.script.run
        .withSuccessHandler(function(success) {
          if (success) {
            loadConfigurations();
            
            // Also refresh sidebar configuration list
            google.script.run.refreshSidebarConfigList();
          } else {
            showError('Failed to delete configuration');
          }
        })
        .withFailureHandler(function(error) {
          showError('Error: ' + error.message);
        })
        .deleteConfiguration(name);
    }
    
    // Function to return to list view
    function returnToListView() {
      document.getElementById('configEditView').classList.add('hidden');
      document.getElementById('configListView').classList.remove('hidden');
    }
    
    // Function to show error message
    function showError(message) {
      // Implementation...
    }
    
    // Function to close the dialog
    function closeDialog() {
      google.script.host.close();
    }
  </script>
</head>
<body>
  <!-- Dialog content will go here -->
</body>
</html>
```

## 7. Feature Implementation Timeline

### Phase 1: Core Structure (1 week)
- Define configuration object model
- Implement document properties storage & retrieval
- Add basic sidebar configuration section
- Create loading functionality for applying configurations

### Phase 2: Configuration Management Dialog (1 week)
- Create dialog UI for configuration list view
- Implement create/edit functionality
- Add configuration deletion with confirmation
- Connect dialog with sidebar

### Phase 3: Enhanced User Experience (1 week)
- Add validation for configurations
- Implement better error handling
- Add configuration status indicators
- Ensure proper sequence for loading settings

### Phase 4: Testing & Refinement (1 week)
- Test with various configuration scenarios
- Fix edge cases and bugs
- Optimize performance
- Document feature for users

## 8. Success Criteria

1. Users can save the current mail merge settings as a named configuration
2. Users can view, edit, and delete saved configurations through a dialog
3. Users can load a configuration which applies all enabled settings
4. After loading, users can still modify any setting as needed
5. The UI clearly indicates which configuration is currently active

## 9. Future Enhancements

1. Configuration duplication
2. Import/export configurations between documents
3. Default configuration on sidebar open
4. Configuration sharing across users
5. Configuration descriptions/notes
6. Configuration versioning

## 10. Technical Considerations

### 10.1 Document Properties Limitations

- Maximum storage size (consider for many configurations)
- No automatic synchronization between users
- Property values must be strings (requiring serialization)

### 10.2 User Experience Considerations

- Async operations require careful sequencing
- UI should provide clear feedback during operations
- Cancellation/error recovery must be handled gracefully

### 10.3 Security Considerations

- Configurations are only accessible to users with document access
- No sensitive data should be included in configuration names
- Consider sanitizing inputs before storing in properties