# Mail Merge Configuration Implementation Plan

## Overview

This plan outlines the development of a configuration management system for the Mail Merge Google Docs add-on. The feature will allow users to save, load, and reuse mail merge configurations directly in their documents, improving workflow efficiency and reducing errors.

## Goals

- Enable users to save complete mail merge configurations in their documents
- Provide a UI to load saved configurations with a single click
- Allow selective application of settings via checkboxes
- Support multiple named configurations in a single document
- Create an extensible foundation for future enhancements

## Technical Architecture

### 1. Configuration Object Model

```javascript
/**
 * Configuration object representing a Mail Merge configuration.
 */
class MailMergeConfig {
  constructor(name) {
    this.version = "1.0";
    this.name = name || "Unnamed Configuration";
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
  
  // Methods for parsing, formatting, validation, and application
}
```

### 2. Document Format

```
=== Start Configuration: [Profile Name] ===
- Use: [✓] Spreadsheet URL: https://docs.google.com/spreadsheets/d/...
- Use: [✓] Sheet Name: Subscribers
- Use: [✓] Recipients Column: Email
- Use: [ ] CC Column: Manager
- Use: [ ] BCC Column: 
- Use: [✓] From Name: Newsletter Team
- Use: [✓] From Email: newsletter@company.com
- Use: [✓] Subject Line: Monthly Updates - {{Month}} {{Year}}
- Use: [ ] CC Override: team@company.com
- Use: [ ] BCC Override: archives@company.com
=== End Configuration: [Profile Name] ===
```

### 3. UI Components

- Configuration section in sidebar
- Configuration selector dropdown
- Load/Refresh/Save buttons
- Configuration status indicator

## Implementation Timeline

### Phase 1: Foundation (Week 1)

- Define configuration object model
- Implement document parsing
- Create basic UI elements
- Establish MVP validation

### Phase 2: Core Functionality (Week 2)

- Implement configuration loading
- Add configuration saving
- Integrate with existing mail merge flow
- Add basic error handling

### Phase 3: Refinement (Week 3)

- Enhance validation
- Improve error messages
- Add configuration management UI
- Optimize performance

### Phase 4: Testing & Documentation (Week 4)

- Comprehensive testing
- User documentation
- Bug fixes and refinements
- Final release preparation

## Detailed Tasks Breakdown

### Phase 1: Foundation (Week 1)

#### Day 1-2: Configuration Object Model
- [ ] Create `MailMergeConfig` class with necessary properties
- [ ] Implement static `fromDocument` method for parsing
- [ ] Implement `toDocument` method for formatting
- [ ] Add basic validation methods

#### Day 3-4: Document Parsing
- [ ] Implement regex patterns for extracting configurations
- [ ] Create parsing logic for settings with checkboxes
- [ ] Develop error handling for malformed configurations
- [ ] Write extraction method for configuration list

#### Day 5: Basic UI Elements
- [ ] Add configuration section to sidebar
- [ ] Create configuration dropdown component
- [ ] Add basic buttons (Load, Refresh, Save)
- [ ] Implement UI state management

### Phase 2: Core Functionality (Week 2)

#### Day 1-2: Configuration Loading
- [ ] Implement configuration selection handling
- [ ] Create sequential loading mechanism for settings
- [ ] Add validation before loading
- [ ] Implement feedback for loaded settings

#### Day 3-4: Configuration Saving
- [ ] Create UI for naming and saving configurations
- [ ] Implement current state extraction to configuration object
- [ ] Develop document insertion logic for saving
- [ ] Add confirmation and error handling

#### Day 5: Integration & Basic Error Handling
- [ ] Connect configuration system to existing mail merge flow
- [ ] Implement error toast notifications
- [ ] Add status indicators for current configuration
- [ ] Test end-to-end workflow

### Phase 3: Refinement (Week 3)

#### Day 1-2: Enhanced Validation
- [ ] Implement spreadsheet URL validation
- [ ] Add column existence checking
- [ ] Validate email addresses and subject lines
- [ ] Categorize errors (critical/warning/info)

#### Day 3-4: UI Improvements
- [ ] Add configuration status indicators
- [ ] Implement better error messages with guidance
- [ ] Create configuration preview before loading
- [ ] Add configuration list refresh functionality

#### Day 5: Performance Optimization
- [ ] Implement caching for parsed configurations
- [ ] Add incremental updates for document changes
- [ ] Optimize spreadsheet operations
- [ ] Add loading indicators for operations

### Phase 4: Testing & Documentation (Week 4)

#### Day 1-2: Comprehensive Testing
- [ ] Create test cases for all functionality
- [ ] Test with various document formats and edge cases
- [ ] Verify error handling and recovery
- [ ] Performance testing with large documents

#### Day 3-4: Documentation & Refinement
- [ ] Create user documentation
- [ ] Add inline help and tooltips
- [ ] Fix bugs identified during testing
- [ ] Implement final UI refinements

#### Day 5: Final Preparation
- [ ] Final testing
- [ ] Version control tagging
- [ ] Release notes preparation
- [ ] Deployment preparation

## Implementation Details

### Backend (backend.gs)

#### 1. Configuration Parsing

```javascript
/**
 * Extracts all configurations from the document.
 * @return {Object} Object mapping configuration names to MailMergeConfig objects.
 */
function extractConfigurationsFromDocument() {
  const docContent = getDocContent();
  const configPattern = /===\s*Start Configuration:\s*([^=]+)\s*===([\s\S]*?)===\s*End Configuration/g;
  
  const configurations = {};
  let match;
  while ((match = configPattern.exec(docContent)) !== null) {
    const name = match[1].trim();
    const configSection = match[2];
    
    const config = new MailMergeConfig(name);
    parseConfigSection(configSection, config);
    
    configurations[name] = config;
  }
  
  return configurations;
}

/**
 * Parses a configuration section into a MailMergeConfig object.
 * @param {string} configSection - The text content of the configuration section.
 * @param {MailMergeConfig} config - The configuration object to populate.
 */
function parseConfigSection(configSection, config) {
  const lines = configSection.split('\n');
  
  for (const line of lines) {
    const settingMatch = line.match(/^\s*-\s*Use:\s*\[([\s✓x])\]\s*([^:]+):\s*(.*)/);
    if (settingMatch) {
      const isEnabled = settingMatch[1].trim() === '✓' || settingMatch[1].trim() === 'x';
      const settingName = settingMatch[2].trim();
      const settingValue = settingMatch[3].trim();
      
      // Map document setting names to object properties
      const settingKey = mapSettingNameToKey(settingName);
      if (settingKey && config.settings.hasOwnProperty(settingKey)) {
        config.settings[settingKey] = {
          enabled: isEnabled,
          value: settingValue
        };
      }
    }
  }
}

/**
 * Maps a document setting name to the corresponding object property.
 * @param {string} settingName - The setting name from the document.
 * @return {string|null} The corresponding object property or null if not found.
 */
function mapSettingNameToKey(settingName) {
  const nameMap = {
    "Spreadsheet URL": "spreadsheetUrl",
    "Sheet Name": "sheetName",
    "Recipients Column": "recipientColumn",
    "CC Column": "ccColumn",
    "BCC Column": "bccColumn",
    "From Name": "fromName",
    "From Email": "fromEmail",
    "Subject Line": "subjectLine",
    "CC Override": "ccOverride",
    "BCC Override": "bccOverride"
  };
  
  return nameMap[settingName] || null;
}
```

#### 2. Configuration Saving

```javascript
/**
 * Appends a configuration to the document.
 * @param {MailMergeConfig} config - The configuration to append.
 */
function appendConfigurationToDocument(config) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Format the configuration text
  let configText = `=== Start Configuration: ${config.name} ===\n`;
  
  const keyMap = {
    "spreadsheetUrl": "Spreadsheet URL",
    "sheetName": "Sheet Name",
    "recipientColumn": "Recipients Column",
    "ccColumn": "CC Column",
    "bccColumn": "BCC Column",
    "fromName": "From Name",
    "fromEmail": "From Email",
    "subjectLine": "Subject Line",
    "ccOverride": "CC Override",
    "bccOverride": "BCC Override"
  };
  
  // Add each setting
  for (const [key, settingObj] of Object.entries(config.settings)) {
    const settingName = keyMap[key];
    if (settingName) {
      const checkMark = settingObj.enabled ? '✓' : ' ';
      configText += `- Use: [${checkMark}] ${settingName}: ${settingObj.value}\n`;
    }
  }
  
  configText += `=== End Configuration: ${config.name} ===\n\n`;
  
  // Add to document
  body.appendParagraph(configText);
}

/**
 * Creates a configuration from the current UI state.
 * @param {string} name - The name for the configuration.
 * @return {MailMergeConfig} The created configuration.
 */
function createConfigFromCurrentState(name) {
  const config = new MailMergeConfig(name);
  
  // Extract values from UI
  config.settings.spreadsheetUrl.value = document.getElementById('spreadsheetUrl').value;
  config.settings.spreadsheetUrl.enabled = !!config.settings.spreadsheetUrl.value;
  
  config.settings.sheetName.value = document.getElementById('sheetSelect').value;
  config.settings.sheetName.enabled = !!config.settings.sheetName.value;
  
  config.settings.recipientColumn.value = document.getElementById('emailColumnSelect').value;
  config.settings.recipientColumn.enabled = !!config.settings.recipientColumn.value;
  
  // Extract remaining values
  // [...]
  
  return config;
}
```

#### 3. MVP Validation

```javascript
/**
 * Performs basic validation of a configuration.
 * @param {MailMergeConfig} config - The configuration to validate.
 * @return {Object} Validation result with success flag and error messages.
 */
function validateConfiguration(config) {
  const result = {
    success: true,
    criticalErrors: [],
    warnings: []
  };
  
  // Check for critical required fields
  if (config.settings.spreadsheetUrl.enabled) {
    try {
      const spreadsheetId = extractSpreadsheetId(config.settings.spreadsheetUrl.value);
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      result.success = false;
      result.criticalErrors.push(`Invalid spreadsheet URL: ${e.message}`);
    }
  }
  
  // Additional validations for enabled settings
  // [...]
  
  return result;
}
```

### Frontend (JavaScriptFile.html)

#### 1. Configuration UI

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
      for (const [name, config] of Object.entries(configs)) {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        configSelect.appendChild(option);
      }
      
      if (Object.keys(configs).length === 0) {
        showNotification('No configurations found in document', 'info', config.toastDuration);
      } else {
        showNotification(`Found ${Object.keys(configs).length} configurations`, 'success', config.toastDuration);
      }
    })
    .withFailureHandler(function(error) {
      showNotification('Error loading configurations: ' + error.message, 'error', config.toastDuration);
    })
    .extractConfigurationsFromDocument();
}

// Function to load a selected configuration
function loadSelectedConfiguration() {
  const configName = document.getElementById('configSelect').value;
  if (!configName) {
    showNotification('Please select a configuration', 'error', config.toastDuration);
    return;
  }
  
  showNotification(`Loading configuration "${configName}"...`, 'info', config.toastDuration);
  
  google.script.run
    .withSuccessHandler(function(result) {
      if (result.success) {
        applyConfiguration(result.config);
        showNotification(`Configuration "${configName}" loaded successfully`, 'success', config.toastDuration);
      } else {
        showNotification('Error: ' + result.message, 'error', config.toastDuration);
      }
    })
    .withFailureHandler(function(error) {
      showNotification('Error loading configuration: ' + error.message, 'error', config.toastDuration);
    })
    .loadAndValidateConfiguration(configName);
}

// Function to apply a configuration
function applyConfiguration(config) {
  // Apply settings in sequence with proper delays for async operations
  
  // Apply spreadsheet URL first if enabled
  if (config.settings.spreadsheetUrl.enabled) {
    document.getElementById('spreadsheetUrl').value = config.settings.spreadsheetUrl.value;
    validateSpreadsheet();
  }
  
  // Wait for validation to complete before continuing
  setTimeout(function() {
    // Apply sheet name if enabled
    if (config.settings.sheetName.enabled) {
      // Set sheet selector and trigger change event
      // Then load columns
    }
    
    // Wait for columns to load
    setTimeout(function() {
      // Apply remaining settings
      // [...]
      
      // Update summary
      updateSummary();
    }, 1500);
  }, 1500);
}

// Function to save current configuration
function saveCurrentConfiguration() {
  // Prompt for a name
  const name = prompt('Enter a name for this configuration:');
  if (!name) return;
  
  showNotification(`Saving configuration "${name}"...`, 'info', config.toastDuration);
  
  google.script.run
    .withSuccessHandler(function() {
      showNotification(`Configuration "${name}" saved successfully`, 'success', config.toastDuration);
      loadConfigurationList(); // Refresh the list
    })
    .withFailureHandler(function(error) {
      showNotification('Error saving configuration: ' + error.message, 'error', config.toastDuration);
    })
    .saveCurrentConfiguration(name);
}
```

## Testing Strategy

### Unit Testing

1. **Configuration Object**:
   - Test creation, parsing, and formatting
   - Verify property mapping
   - Check validation logic

2. **Document Parsing**:
   - Test extraction with various formats
   - Verify handling of malformed configurations
   - Check behavior with multiple configurations

3. **UI Components**:
   - Test dropdown population
   - Verify button actions
   - Check error displays

### Integration Testing

1. **End-to-End Flow**:
   - Save configuration and verify document format
   - Load configuration and verify UI state
   - Test with actual Google Sheets

2. **Error Handling**:
   - Test with invalid spreadsheet URLs
   - Verify behavior with missing columns
   - Check recovery from validation failures

3. **Performance**:
   - Test with large documents
   - Verify operation with multiple configurations
   - Check response times for various operations

## Potential Challenges & Mitigations

### Timing Issues with Async Operations

**Challenge**: Many operations (spreadsheet validation, sheet loading) are asynchronous, making it difficult to sequence configuration loading properly.

**Mitigation**: 
- Implement a queue-based approach for sequential operations
- Use proper callback chaining
- Add appropriate delays and retries
- Show clear loading indicators during operations

### Document Format Changes

**Challenge**: If users manually edit configurations in the document, they may introduce format errors that break parsing.

**Mitigation**:
- Implement robust error handling in the parser
- Add clear formatting guidance in comments
- Provide a way to fix or recreate invalid configurations

### UI Complexity

**Challenge**: The configuration UI adds another layer of complexity to an already feature-rich add-on.

**Mitigation**:
- Keep the UI minimal and collapsible
- Use clear, contextual help text
- Implement progressive disclosure of advanced features

## Success Criteria

1. Users can save and load configurations with minimal clicks
2. Multiple configurations can coexist in a single document
3. Configuration loading is reliable, with clear error messages
4. The system can be extended with new settings in the future

## Future Enhancements (Post-MVP)

1. Configuration version tracking
2. Export/import configurations between documents
3. Template library with pre-configured settings
4. Configuration duplication and editing
5. Advanced validation rules

---

## Appendix: Backend Class Reference

```javascript
/**
 * Represents a Mail Merge configuration.
 */
class MailMergeConfig {
  /**
   * Creates a new configuration.
   * @param {string} name - The configuration name.
   */
  constructor(name) {
    this.version = "1.0";
    this.name = name || "Unnamed Configuration";
    this.settings = {
      // Complete settings object
    };
  }
  
  /**
   * Creates a configuration from document text.
   * @param {string} text - The configuration text from the document.
   * @return {MailMergeConfig} The created configuration.
   */
  static fromDocument(text) {
    // Implementation details
  }
  
  /**
   * Converts the configuration to document format.
   * @return {string} The formatted configuration text.
   */
  toDocument() {
    // Implementation details
  }
  
  /**
   * Validates the configuration.
   * @return {Object} Validation results.
   */
  validate() {
    // Implementation details
  }
}
```