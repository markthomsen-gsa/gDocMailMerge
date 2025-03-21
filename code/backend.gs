/**
 * Enhanced Mail Merge for Google Docs
 * This script creates a sidebar for sending personalized emails directly from Google Docs
 * with advanced features for email account selection, content validation, and more.
 */

"use strict";

/**
 * CONFIGURATION: Template fields definition
 * Single source of truth for field configuration throughout the application
 */
const TEMPLATE_FIELDS = {
  "templateName": {
    displayName: "Template Name",
    status: "Required",
    default: "{env}-{campign}-{moniker}",
    description: "Unique template identifier"
  },
  "description": {
    displayName: "Description",
    status: "Optional", 
    default: "",
    description: "Purpose of this template"
  },
  "spreadsheetUrl": {
    displayName: "Spreadsheet",
    status: "Required",
    default: "",
    description: "Data source spreadsheet"
  },
  "sheetName": {
    displayName: "Sheet Name",
    status: "Required",
    default: "",
    description: "Specific sheet to use"
  },
  "emailColumn": {
    displayName: "Email Column",
    status: "Required",
    default: "",
    description: "Column with recipient emails"
  },
  "subjectLine": {
    displayName: "Subject Line",
    status: "Required",
    default: "",
    description: "Email subject line"
  },
  "fromEmail": {
    displayName: "From Email",
    status: "Hardcoded",
    default: "",
    description: "Sender email address"
  },
  "senderDisplayName": {
    displayName: "Sender Display Name",
    status: "Required",
    default: "",
    description: "Name shown to recipients"
  },
  "contentFormat": {
    displayName: "Content Format",
    status: "Auto",
    default: "Plain text with HTML output",
    description: "Email content format"
  },
  "lastUpdated": {
    displayName: "Last Updated",
    status: "Auto",
    default: "",
    description: "Last modified date"
  }
};

/**
 * CONFIGURATION: Application Defaults
 * Default values used throughout the application
 */
const APP_DEFAULTS = {
  // User Information
  senderDisplayName: "Mail Merge Sender", // Changed from fromName
  fromEmail: Session.getActiveUser().getEmail() || "",
  testEmail: Session.getActiveUser().getEmail() || "",
  
  // Default template values
  defaultSubjectLine: "[Enter Subject Line]",
  defaultDescription: "Mail Merge Template",
  
  // Markers
  configStartMarker: "--- CONFIGURATION START ---",
  configEndMarker: "--- CONFIGURATION END ---",
  contentStartMarker: "--- EMAIL CONTENT START ---",
  contentEndMarker: "--- EMAIL CONTENT END ---",
  
  // Placeholders
  defaultPlaceholders: ["FirstName", "LastName", "Email", "Company"],
  
  // Validation - updated field names
  requiredFields: ["templateName", "spreadsheetUrl", "sheetName", "emailColumn", "subjectLine", "senderDisplayName"]
};

/**
 * CONFIGURATION: Table styling preferences
 * This object controls all table styling throughout the application
 */
const TABLE_STYLES = {
  tableProperties: {
    borderWidth: 0.5,            // Thin borders
    borderColor: "#e0e0e0"       // Light gray
  },
  
  columnWidths: {
    attributeColumn: 130,
    valueColumn: 270,
    statusColumn: 100,
    descriptionColumn: 200       // Added for description column
  },
  
  headerRow: {
    fontFamily: "Open Sans",
    fontSize: 9,
    bold: true,
    backgroundColor: "#f8f9fa",  // Very light gray
    foregroundColor: "#202124"   // Dark gray text
  },
  
  attributeColumn: {
    fontFamily: "Open Sans",
    fontSize: 9,
    bold: true,
    foregroundColor: "#202124"
  },
  
  valueColumn: {
    fontFamily: "Open Sans",
    fontSize: 9,
    bold: false,
    foregroundColor: "#5f6368"   // Medium gray text
  },
  
  statusColumn: {
    fontFamily: "Open Sans",
    fontSize: 8,
    bold: false,
    horizontalAlignment: DocumentApp.HorizontalAlignment.CENTER,
    requiredColor: "#fdefe9",    // Very light red
    optionalColor: "#e6f4ea",    // Very light green
    autoColor: "#e8f0fe",        // Very light blue
    hardcodedColor: "#f8f9fa"    // Light gray for hardcoded values
  },
  
  descriptionColumn: {
    fontFamily: "Open Sans",
    fontSize: 8,
    bold: false,
    foregroundColor: "#5f6368",  // Medium gray text
    italics: true
  }
};

/**
 * Gets HTML content from a file.
 * @param {string} filename - The name of the HTML file.
 * @return {string} The HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets field configuration for a specific field or all fields
 * @param {string} fieldName - Optional field name
 * @return {Object} Field configuration
 */
function getFieldConfiguration(fieldName) {
  if (fieldName) {
    return TEMPLATE_FIELDS[fieldName] || null;
  }
  return TEMPLATE_FIELDS;
}

/**
 * Generates a table row for each field in TEMPLATE_FIELDS
 * @param {Table} table - The table object
 */
function generateTemplateTableFromConfig(table) {
  for (const [fieldName, config] of Object.entries(TEMPLATE_FIELDS)) {
    let defaultValue = config.default;
    
    // Set special defaults
    if (fieldName === "lastUpdated") {
      defaultValue = new Date().toLocaleDateString();
    } else if (fieldName === "fromEmail") {
      defaultValue = Session.getActiveUser().getEmail() || "[Your Email]";
    } else if (fieldName === "senderDisplayName") {
      defaultValue = APP_DEFAULTS.senderDisplayName;
    }
    
    addConfigRowToTable(
      table,
      config.displayName,
      defaultValue,
      config.status,
      config.description
    );
  }
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
  
  return extractEmailContent(fullText);
}

/**
 * Extracts just the email content from a document or text
 * with minimal processing to preserve HTML exactly as written
 * 
 * @param {string} text - The document text or template content
 * @return {string} The email content between markers
 */
function extractEmailContent(text) {
  const contentStartMarker = APP_DEFAULTS.contentStartMarker;
  const contentEndMarker = APP_DEFAULTS.contentEndMarker;
  
  const startIndex = text.indexOf(contentStartMarker);
  const endIndex = text.indexOf(contentEndMarker);
  
  // If both markers exist, extract only the content between them
  if (startIndex !== -1 && endIndex !== -1 && startIndex < endIndex) {
    // Extract the content between markers (add marker length to get position after marker)
    const contentStartPos = startIndex + contentStartMarker.length;
    
    // Preserve all whitespace and formatting exactly as written
    return text.substring(contentStartPos, endIndex);
  }
  
  // Fallback: If markers not found, return empty string
  return "";
}

/**
 * Creates a new template in the current document.
 * Inserts configuration markers, table, and content markers.
 * Uses the TEMPLATE_FIELDS configuration for consistent structure.
 */
function createNewTemplate() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Check if document is empty or user wants to replace content
  let proceed = true;
  if (body.getText().trim().length > 0) {
    const ui = DocumentApp.getUi();
    const response = ui.alert(
      'Replace Document Content?',
      'This will replace the current document content with a new template. Continue?',
      ui.ButtonSet.YES_NO
    );
    proceed = (response === ui.Button.YES);
  }
  
  if (!proceed) return;
  
  // Clear the document
  body.clear();
  
  // Add configuration section header
  const configStartPara = body.appendParagraph(APP_DEFAULTS.configStartMarker);
  configStartPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  configStartPara.setFontFamily("Courier New");
  configStartPara.setBold(true);
  
  // Create configuration table
  const table = body.appendTable();
  
  // Apply table border if specified
  if (TABLE_STYLES.tableProperties.borderWidth) {
    table.setBorderWidth(TABLE_STYLES.tableProperties.borderWidth);
  }
  if (TABLE_STYLES.tableProperties.borderColor) {
    table.setBorderColor(TABLE_STYLES.tableProperties.borderColor);
  }
  
  // Add table headers with styling from configuration
  const headerRow = table.appendTableRow();
  
  // Create and style header cells
  const headerCells = [
    headerRow.appendTableCell("Attribute Name"),
    headerRow.appendTableCell("Value"),
    headerRow.appendTableCell("Status"),
    headerRow.appendTableCell("Description")
  ];
  
  // Set widths for header cells
  headerCells[0].setWidth(TABLE_STYLES.columnWidths.attributeColumn);
  headerCells[1].setWidth(TABLE_STYLES.columnWidths.valueColumn);
  headerCells[2].setWidth(TABLE_STYLES.columnWidths.statusColumn);
  headerCells[3].setWidth(TABLE_STYLES.columnWidths.descriptionColumn);
  
  // Apply header row styling to all header cells
  headerCells.forEach(cell => {
    cell.setFontFamily(TABLE_STYLES.headerRow.fontFamily);
    cell.setFontSize(TABLE_STYLES.headerRow.fontSize);
    cell.setBold(TABLE_STYLES.headerRow.bold);
    cell.setBackgroundColor(TABLE_STYLES.headerRow.backgroundColor);
  });
  
  // Generate template configuration rows from configuration
  generateTemplateTableFromConfig(table);
  
  // Add configuration end marker
  const configEndPara = body.appendParagraph(APP_DEFAULTS.configEndMarker);
  configEndPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  configEndPara.setFontFamily("Courier New");
  configEndPara.setBold(true);
  
  // Add some spacing
  body.appendParagraph("");
  
  // Add content section header
  const contentStartPara = body.appendParagraph(APP_DEFAULTS.contentStartMarker);
  contentStartPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contentStartPara.setFontFamily("Courier New");
  contentStartPara.setBold(true);
  
  // Add placeholder content
  body.appendParagraph("Dear {{FirstName}},");
  body.appendParagraph("");
  body.appendParagraph("This is a sample email template. Replace this text with your actual email content.");
  body.appendParagraph("");
  body.appendParagraph("You can use placeholders like {{FirstName}}, {{LastName}}, or any column name from your spreadsheet surrounded by double curly braces.");
  body.appendParagraph("");
  body.appendParagraph("Best regards,");
  body.appendParagraph("{{SenderName}}");
  
  // Add content end marker
  const contentEndPara = body.appendParagraph(APP_DEFAULTS.contentEndMarker);
  contentEndPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contentEndPara.setFontFamily("Courier New");
  contentEndPara.setBold(true);
  
  // Alert the user
  const ui = DocumentApp.getUi();
  ui.alert(
    'Template Created',
    'A new mail merge template has been created. Please fill in the configuration table and edit the email content.',
    ui.ButtonSet.OK
  );
}

/**
 * Helper function to add a configuration row to the template table.
 * Uses the TABLE_STYLES configuration for consistent styling.
 * 
 * @param {Table} table - The table to add the row to
 * @param {string} attribute - The attribute name
 * @param {string} value - The value
 * @param {string} status - The status (Required/Optional/Auto/Hardcoded)
 * @param {string} description - The description of this attribute
 */
function addConfigRowToTable(table, attribute, value, status, description) {
  const row = table.appendTableRow();
  
  // Create cells (4 columns)
  const attrCell = row.appendTableCell(attribute);
  const valueCell = row.appendTableCell(value);
  const statusCell = row.appendTableCell(status);
  const descCell = row.appendTableCell(description || "");
  
  // Set column widths
  attrCell.setWidth(TABLE_STYLES.columnWidths.attributeColumn);
  valueCell.setWidth(TABLE_STYLES.columnWidths.valueColumn);
  statusCell.setWidth(TABLE_STYLES.columnWidths.statusColumn);
  descCell.setWidth(TABLE_STYLES.columnWidths.descriptionColumn);
  
  // Apply attribute column styling
  attrCell.setFontFamily(TABLE_STYLES.attributeColumn.fontFamily);
  attrCell.setFontSize(TABLE_STYLES.attributeColumn.fontSize);
  attrCell.setBold(TABLE_STYLES.attributeColumn.bold);
  
  // Apply value column styling
  valueCell.setFontFamily(TABLE_STYLES.valueColumn.fontFamily);
  valueCell.setFontSize(TABLE_STYLES.valueColumn.fontSize);
  valueCell.setBold(TABLE_STYLES.valueColumn.bold);
  
  // Apply status column styling
  statusCell.setFontFamily(TABLE_STYLES.statusColumn.fontFamily);
  statusCell.setFontSize(TABLE_STYLES.statusColumn.fontSize);
  statusCell.setBold(TABLE_STYLES.statusColumn.bold);
  
  // Apply description column styling
  descCell.setFontFamily(TABLE_STYLES.descriptionColumn.fontFamily);
  descCell.setFontSize(TABLE_STYLES.descriptionColumn.fontSize);
  descCell.setItalic(TABLE_STYLES.descriptionColumn.italics);
  descCell.setForegroundColor(TABLE_STYLES.descriptionColumn.foregroundColor);
  
  // Set status cell background color based on status type
  if (status === "Required") {
    statusCell.setBackgroundColor(TABLE_STYLES.statusColumn.requiredColor);
  } else if (status === "Optional") {
    statusCell.setBackgroundColor(TABLE_STYLES.statusColumn.optionalColor);
  } else if (status === "Auto") {
    statusCell.setBackgroundColor(TABLE_STYLES.statusColumn.autoColor);
  } else if (status === "Hardcoded") {
    statusCell.setBackgroundColor(TABLE_STYLES.statusColumn.hardcodedColor);
  }
  
  // Set minimal row height
  row.setMinimumHeight(0);
}

/**
 * Extracts configuration data from the document template table.
 * @return {Object|null} The configuration data or null if not found
 */
function extractTemplateConfig() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();
  
  // Look for a table with "Template Name" in first column
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (table.getNumRows() > 0) {
      const firstCell = table.getCell(0, 0);
      // Skip if first cell is null (shouldn't happen, but just in case)
      if (!firstCell) continue;
      
      const firstCellText = firstCell.getText().trim();
      
      // Check for header row
      if (firstCellText === "Attribute Name") {
        // Check second row for attribute name matching Template Name display name
        if (table.getNumRows() > 1) {
          const secondRowFirstCell = table.getCell(1, 0);
          if (secondRowFirstCell && 
              secondRowFirstCell.getText().trim() === TEMPLATE_FIELDS.templateName.displayName) {
            return extractConfigFromTable(table);
          }
        }
      }
      // Check if this is the template table directly (no header row)
      else if (firstCellText === TEMPLATE_FIELDS.templateName.displayName) {
        return extractConfigFromTable(table);
      }
    }
  }
  
  // No configuration table found
  return null;
}

/**
 * Extracts configuration from a template table using TEMPLATE_FIELDS
 * @param {Table} table - The configuration table
 * @return {Object} The configuration object
 */
function extractConfigFromTable(table) {
  const config = {};
  const startRow = table.getCell(0, 0).getText().trim() === "Attribute Name" ? 1 : 0;
  
  // Create mapping of display names to field names
  const displayToFieldMap = {};
  for (const [fieldName, fieldConfig] of Object.entries(TEMPLATE_FIELDS)) {
    displayToFieldMap[fieldConfig.displayName] = fieldName;
  }
  
  // Debug with Logger instead of console.log
  Logger.log("Extracting configuration from table");
  
  for (let i = startRow; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    if (row.getNumCells() < 2) continue;
    
    const attributeCell = table.getCell(i, 0);
    const valueCell = table.getCell(i, 1);
    
    if (!attributeCell || !valueCell) continue;
    
    const displayName = attributeCell.getText().trim();
    let value = valueCell.getText().trim();
    
    // Skip empty attribute names and placeholder values
    if (!displayName || (value.startsWith('[') && value.endsWith(']'))) continue;
    
    // Special handling for Spreadsheet field to support URL chips
    if (displayName === "Spreadsheet") {
      try {
        // Get raw text first as a fallback
        let rawValue = value;
        
        // Try to get the URL from rich text format if it's a chip
        const richText = valueCell.getRichTextValue();
        if (richText) {
          // Check for URL in any text runs
          const runs = richText.getRuns();
          for (const run of runs) {
            const linkUrl = run.getLinkUrl();
            if (linkUrl) {
              // Use the URL from the chip
              rawValue = linkUrl;
              Logger.log("Found URL in chip: " + rawValue);
              break;
            }
          }
        }
        
        // Store the raw value first (important!)
        const fieldName = displayToFieldMap[displayName];
        if (fieldName) {
          config[fieldName] = rawValue;
        }
        
        // Also try to extract ID - but keep original value if this fails
        try {
          const spreadsheetId = extractSpreadsheetId(rawValue);
          if (spreadsheetId && spreadsheetId.length > 10) { // Reasonable ID length check
            // Only override if we got a valid-looking ID
            config[fieldName] = spreadsheetId;
            Logger.log("Extracted ID: " + spreadsheetId);
          }
        } catch (idError) {
          Logger.log("ID extraction error: " + idError.message);
          // Keep the original value - already saved above
        }
      } catch (e) {
        // If there's any error in rich text handling, log it but keep going
        Logger.log("Error handling rich text: " + e.message);
        
        // Make sure we still save something
        const fieldName = displayToFieldMap[displayName];
        if (fieldName && value) {
          config[fieldName] = value;
          Logger.log("Using fallback text value: " + value);
        }
      }
    } else {
      // Standard handling for other fields
      const fieldName = displayToFieldMap[displayName];
      if (fieldName) {
        config[fieldName] = value;
      } else {
        // For backward compatibility, store by display name
        config[displayName] = value;
      }
    }
  }
  
  return config;
}

/**
 * Saves the current document as a template to storage.
 * @return {Object} Result with success flag and message
 */
function saveTemplateToStorage() {
  try {
    // Extract template configuration from the document
    const config = extractTemplateConfig();
    
    if (!config) {
      return {
        success: false,
        message: "No template configuration table found. Please use 'Create New Template' first."
      };
    }
    
    // Check for required fields
    const missingFields = [];
    for (const fieldName of APP_DEFAULTS.requiredFields) {
      // Get the display name for this field for user-friendly error messages
      const displayName = fieldName in TEMPLATE_FIELDS ? 
                         TEMPLATE_FIELDS[fieldName].displayName : 
                         fieldName;
      
      // Check if field is missing or contains placeholder text
      if (!config[fieldName] || 
          (config[fieldName].startsWith('[') && config[fieldName].endsWith(']'))) {
        missingFields.push(displayName);
      }
    }
    
    if (missingFields.length > 0) {
      return {
        success: false,
        message: "Missing required fields: " + missingFields.join(", ")
      };
    }
    
    // Extract template content
    const templateContent = getDocContent();
    
    if (!templateContent) {
      return {
        success: false,
        message: "No template content found between EMAIL CONTENT markers."
      };
    }
    
    // Add the full document content (including markers) to the template
    config.documentContent = DocumentApp.getActiveDocument().getBody().getText();
    config.documentContentSize = config.documentContent.length;
    config.documentName = DocumentApp.getActiveDocument().getName();
    config.lastSaved = new Date().toISOString();
    
    // Update the Last Updated field in the table
    updateLastUpdatedInTable(new Date().toLocaleDateString());
    
    // Prepare template storage - use templateName field or display name field for backward compatibility
    const templateName = config.templateName || config[TEMPLATE_FIELDS.templateName.displayName] || "";
    if (!templateName) {
      return {
        success: false,
        message: "Template name is required."
      };
    }
    
    const docProperties = PropertiesService.getDocumentProperties();
    const templatesJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const templates = JSON.parse(templatesJson);
    
    // Check if template already exists and confirm overwrite
    if (templates[templateName]) {
      const ui = DocumentApp.getUi();
      const response = ui.alert(
        'Template Already Exists',
        `Template "${templateName}" already exists. Last modified on ${new Date(templates[templateName].lastSaved).toLocaleDateString()}. Overwrite?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        return {
          success: false,
          message: "Template save cancelled."
        };
      }
    }
    
    // Add to templates
    templates[templateName] = config;
    
    // Save back to document properties
    docProperties.setProperty('mailMergeConfigs', JSON.stringify(templates));
    
    // Set flag to refresh configurations when returning to sidebar
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    // Calculate template size in KB
    const templateSize = (JSON.stringify(config).length / 1024).toFixed(1);
    
    return {
      success: true,
      message: `Template "${templateName}" saved successfully! (${templateSize} KB)`
    };
  } catch (e) {
    Logger.log("Error saving template: " + e.message);
    return {
      success: false,
      message: "Error saving template: " + e.message
    };
  }
}

/**
 * Updates the Last Updated field in the template table.
 * @param {string} dateStr - The date string to set
 */
function updateLastUpdatedInTable(dateStr) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const tables = body.getTables();
    
    // Look for a table with Template Name display name in first column
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (table.getNumRows() > 0) {
        const firstCell = table.getCell(0, 0);
        // Skip if first cell is null
        if (!firstCell) continue;
        
        const firstCellText = firstCell.getText().trim();
        const headerRow = firstCellText === "Attribute Name" ? 1 : 0;
        
        // Check if this is the template table using display name from TEMPLATE_FIELDS
        if ((headerRow === 1 && table.getNumRows() > 1 && 
             table.getCell(1, 0).getText().trim() === TEMPLATE_FIELDS.templateName.displayName) ||
            (headerRow === 0 && firstCellText === TEMPLATE_FIELDS.templateName.displayName)) {
          
          // Find Last Updated row using display name
          for (let row = headerRow; row < table.getNumRows(); row++) {
            if (table.getCell(row, 0).getText().trim() === TEMPLATE_FIELDS.lastUpdated.displayName) {
              // Update the date
              table.getCell(row, 1).setText(dateStr);
              return;
            }
          }
          
          // If Last Updated row not found, add it
          if (table.getNumRows() > 0) {
            const newRow = table.appendTableRow();
            newRow.appendTableCell(TEMPLATE_FIELDS.lastUpdated.displayName).setBold(true);
            newRow.appendTableCell(dateStr);
            newRow.appendTableCell(TEMPLATE_FIELDS.lastUpdated.status).setBackgroundColor(TABLE_STYLES.statusColumn.autoColor);
            newRow.appendTableCell(TEMPLATE_FIELDS.lastUpdated.description);
          }
          
          return;
        }
      }
    }
  } catch (e) {
    Logger.log("Error updating Last Updated field: " + e.message);
  }
}

/**
 * Shows a dialog to select and load a template into the document.
 */
function showLoadTemplateDialog() {
  const html = HtmlService.createTemplateFromFile('LoadTemplateDialog')
    .evaluate()
    .setWidth(400)
    .setHeight(300)
    .setTitle('Load Template to Document');
  
  DocumentApp.getUi().showModalDialog(html, 'Load Template to Document');
}

/**
 * Gets available templates from storage.
 * @return {Object} Object with template names and descriptions
 */
function getAvailableTemplates() {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    const templatesJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const templates = JSON.parse(templatesJson);
    
    // Create a simplified list for the dropdown
    const templateList = {};
    for (const name in templates) {
      const template = templates[name];
      templateList[name] = {
        description: template.description || template.Description || '',
        lastSaved: template.lastSaved || '',
        size: (JSON.stringify(template).length / 1024).toFixed(1) + ' KB'
      };
    }
    
    return templateList;
  } catch (e) {
    Logger.log("Error getting templates: " + e.message);
    return {};
  }
}

/**
 * Checks if document has content.
 * @return {boolean} True if document has content
 */
function documentHasContent() {
  return DocumentApp.getActiveDocument().getBody().getText().trim().length > 0;
}

/**
 * Loads a template into the document.
 * @param {string} templateName - The name of the template to load
 * @return {Object} Result with success flag and message
 */
function loadTemplateToDocument(templateName) {
  try {
    // Get template from storage
    const docProperties = PropertiesService.getDocumentProperties();
    const templatesJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const templates = JSON.parse(templatesJson);
    
    if (!templates[templateName]) {
      return {
        success: false,
        message: "Template not found: " + templateName
      };
    }
    
    const template = templates[templateName];
    
    // Check if document has content and confirm overwrite
    if (documentHasContent()) {
      const ui = DocumentApp.getUi();
      const response = ui.alert(
        'Replace Document Content?',
        'This will replace the current document content with the selected template. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        return {
          success: false,
          message: "Template loading cancelled by user."
        };
      }
    }
    
    // Extract content from template if it exists
    let content = "";
    if (template.documentContent) {
      content = extractEmailContent(template.documentContent);
    }
    
    // Rebuild the document structure
    const rebuildSuccess = rebuildTemplateDocument(template, content);
    
    if (!rebuildSuccess) {
      return {
        success: false,
        message: "Error rebuilding document structure."
      };
    }
    
    return {
      success: true,
      message: `Template "${templateName}" loaded successfully!`,
      config: convertTemplateToUIConfig(template)
    };
  } catch (e) {
    Logger.log("Error loading template: " + e.message);
    return {
      success: false,
      message: "Error loading template: " + e.message
    };
  }
}

/**
 * Rebuilds a template document with proper structure
 * @param {Object} config - Template configuration data
 * @param {string} emailContent - Content to place between content markers
 * @return {boolean} Success status
 */
function rebuildTemplateDocument(config, emailContent) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    
    // Clear the document
    body.clear();
    
    // Add configuration section header
    const configStartPara = body.appendParagraph(APP_DEFAULTS.configStartMarker);
    configStartPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    configStartPara.setFontFamily("Courier New");
    configStartPara.setBold(true);
    
    // Create configuration table
    const table = body.appendTable();
    
    // Apply table styling
    if (TABLE_STYLES.tableProperties.borderWidth) {
      table.setBorderWidth(TABLE_STYLES.tableProperties.borderWidth);
    }
    if (TABLE_STYLES.tableProperties.borderColor) {
      table.setBorderColor(TABLE_STYLES.tableProperties.borderColor);
    }
    
    // Add table headers
    const headerRow = table.appendTableRow();
    const headerCells = [
      headerRow.appendTableCell("Attribute Name"),
      headerRow.appendTableCell("Value"),
      headerRow.appendTableCell("Status"),
      headerRow.appendTableCell("Description")
    ];
    
    // Set widths for header cells
    headerCells[0].setWidth(TABLE_STYLES.columnWidths.attributeColumn);
    headerCells[1].setWidth(TABLE_STYLES.columnWidths.valueColumn);
    headerCells[2].setWidth(TABLE_STYLES.columnWidths.statusColumn);
    headerCells[3].setWidth(TABLE_STYLES.columnWidths.descriptionColumn);
    
    // Apply header row styling to all header cells
    headerCells.forEach(cell => {
      cell.setFontFamily(TABLE_STYLES.headerRow.fontFamily);
      cell.setFontSize(TABLE_STYLES.headerRow.fontSize);
      cell.setBold(TABLE_STYLES.headerRow.bold);
      cell.setBackgroundColor(TABLE_STYLES.headerRow.backgroundColor);
    });
    
    // Add configuration rows
    for (const [fieldName, fieldConfig] of Object.entries(TEMPLATE_FIELDS)) {
      // Get value from config
      // Try modern field name first, then fallback to display name for backward compatibility
      let value;
      if (fieldName in config) {
        value = config[fieldName];
      } else if (fieldConfig.displayName in config) {
        value = config[fieldConfig.displayName];
      } else {
        // Use empty string as fallback
        value = "";
      }
      
      addConfigRowToTable(
        table, 
        fieldConfig.displayName, 
        value, 
        fieldConfig.status, 
        fieldConfig.description
      );
    }
    
    // Add configuration end marker
    const configEndPara = body.appendParagraph(APP_DEFAULTS.configEndMarker);
    configEndPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    configEndPara.setFontFamily("Courier New");
    configEndPara.setBold(true);
    
    // Add some spacing
    body.appendParagraph("");
    
    // Add content section header
    const contentStartPara = body.appendParagraph(APP_DEFAULTS.contentStartMarker);
    contentStartPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contentStartPara.setFontFamily("Courier New");
    contentStartPara.setBold(true);
    
    // Add email content
    body.appendParagraph(emailContent || "");
    
    // Add content end marker
    const contentEndPara = body.appendParagraph(APP_DEFAULTS.contentEndMarker);
    contentEndPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contentEndPara.setFontFamily("Courier New");
    contentEndPara.setBold(true);
    
    return true;
  } catch (e) {
    Logger.log("Error rebuilding template document: " + e.message);
    return false;
  }
}

/**
 * Converts a template configuration to the format used by the UI.
 * Updated to handle new field names.
 * @param {Object} template - The template configuration
 * @return {Object} Configuration compatible with the UI
 */
function convertTemplateToUIConfig(template) {
  // Create a mapping with proper fallbacks for field names
  const getFieldValue = (fieldName) => {
    // Try field name first
    if (fieldName in template) {
      return template[fieldName];
    }
    
    // Try using display name if field name fails
    if (TEMPLATE_FIELDS[fieldName] && 
        TEMPLATE_FIELDS[fieldName].displayName in template) {
      return template[TEMPLATE_FIELDS[fieldName].displayName];
    }
    
    // Try legacy field names for backward compatibility
    const legacyFields = {
      spreadsheetUrl: "Spreadsheet",
      sheetName: "Sheet Name",
      emailColumn: "Email Column",
      subjectLine: "Subject Line",
      fromEmail: "From Email",
      senderDisplayName: "From Name"
    };
    
    if (fieldName in legacyFields && 
        legacyFields[fieldName] in template) {
      return template[legacyFields[fieldName]];
    }
    
    return '';
  };
  
  return {
    templateName: getFieldValue('templateName'),
    spreadsheetUrl: getFieldValue('spreadsheetUrl'),
    sheetName: getFieldValue('sheetName'),
    emailColumn: getFieldValue('emailColumn'),
    ccColumn: getFieldValue('ccColumn') || template["CC Column"] || '',
    bccColumn: getFieldValue('bccColumn') || template["BCC Column"] || '',
    fromEmail: getFieldValue('fromEmail'),
    // Use senderDisplayName, but map it to fromName for backward compatibility
    fromName: getFieldValue('senderDisplayName'), 
    subjectLine: getFieldValue('subjectLine'),
    ccOverride: getFieldValue('ccOverride') || template["CC Override"] || '',
    bccOverride: getFieldValue('bccOverride') || template["BCC Override"] || ''
  };
}

/**
 * Backs up all templates to an email as a JSON file.
 * Creates a well-formatted JSON attachment and emails it to the current user.
 * @return {Object} Result with success flag and message
 */
function backupTemplatesToEmail() {
  try {
    // Get templates from storage
    const docProperties = PropertiesService.getDocumentProperties();
    const templatesJson = docProperties.getProperty('mailMergeConfigs') || '{}';
    const templates = JSON.parse(templatesJson);
    
    // Check if there are any templates to backup
    if (Object.keys(templates).length === 0) {
      return {
        success: false,
        message: "No templates found to backup."
      };
    }
    
    // Create a pretty-printed version of the JSON for better readability
    const prettyJson = JSON.stringify(templates, null, 2);
    
    // Create backup file with timestamp in filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupFilename = `mail_merge_templates_backup_${timestamp}.json`;
    
    // Get user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return {
        success: false,
        message: "Unable to determine your email address for sending the backup."
      };
    }
    
    // Create a Blob for the backup file
    const blob = Utilities.newBlob(prettyJson, 'application/json', backupFilename);
    
    // Send email with attachment
    MailApp.sendEmail({
      to: userEmail,
      subject: "Mail Merge Templates Backup",
      body: "Attached is a backup of your Mail Merge templates as of " + new Date().toLocaleString() + ".\n\n" +
            "Total templates: " + Object.keys(templates).length + "\n" +
            "Storage size: " + (prettyJson.length / 1024).toFixed(2) + " KB\n\n" +
            "To restore these templates, you can use the JSON file with an administrator.",
      attachments: [blob]
    });
    
    return {
      success: true,
      message: `Backup sent to ${userEmail} with ${Object.keys(templates).length} templates.`
    };
  } catch (e) {
    Logger.log("Error backing up templates: " + e.message);
    return {
      success: false,
      message: "Error backing up templates: " + e.message
    };
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
 * Replaces placeholders in the text with values from the data row.
 * Enhanced to properly handle HTML content.
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
    
    // Use a more careful replacement approach to avoid breaking HTML
    result = result.split(placeholder).join(value);
  }
  return result;
}

/**
 * Prepares email options for sending emails.
 * Updated to use modern field names.
 * @param {string} fromEmail - The sender's email address.
 * @param {string} senderDisplayName - The sender's display name.
 * @param {Object} options - Additional options like cc and bcc.
 * @return {Object} Email options object.
 */
function prepareEmailOptions(fromEmail, senderDisplayName, options = {}) {
  const emailOptions = { name: senderDisplayName || undefined };
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
 * Sends a test email.
 * Updated to use new field names.
 * @param {string} recipients - Comma-separated list of email addresses.
 * @param {string} subject - The email subject.
 * @param {string} fromEmail - The sender's email address.
 * @param {string} senderDisplayName - The sender's display name.
 * @param {string} cc - Optional CC addresses.
 * @param {string} bcc - Optional BCC addresses.
 * @param {Object} options - Additional options for test email.
 * @return {Object} Status object with success flag and message.
 */
function sendTestEmailWithData(recipients, subject, fromEmail, senderDisplayName, cc, bcc, options) {
  try {
    Logger.log('Sending test email with parameters:');
    Logger.log('Recipients: ' + recipients);
    Logger.log('Subject: ' + subject);
    Logger.log('From: ' + fromEmail);
    Logger.log('Sender: ' + senderDisplayName);
    
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
    const emailOptions = prepareEmailOptions(fromEmail, senderDisplayName, { cc, bcc });
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
 * Executes the mail merge with optimized batch processing.
 * Updated to use modern field names and include batch processing.
 * @param {string} spreadsheetId - The ID of the spreadsheet
 * @param {string} sheetName - The name of the sheet
 * @param {string} emailColumn - The column containing email addresses
 * @param {string} subjectLine - The email subject
 * @param {string} fromEmail - The sender's email address
 * @param {string} senderDisplayName - The sender's display name
 * @param {Object} options - Additional options
 * @return {Object} Status object with success flag, message, and counts
 */
function executeMailMerge(spreadsheetId, sheetName, emailColumn, subjectLine, fromEmail, senderDisplayName, options = {}) {
  try {
    options = {
      cc: '',
      bcc: '',
      enableLogging: false,
      createDrafts: false,
      batchSize: 50, // Process 50 emails at a time
      ...options
    };
    
    const templateHtml = getDocContent();
    const data = getSpreadsheetData(spreadsheetId, sheetName);
    const headers = data.headers;
    const rows = data.rows;
    const emailIndex = headers.indexOf(emailColumn);
    
    if (emailIndex === -1) {
      throw new Error(`Email column "${emailColumn}" not found in spreadsheet`);
    }
    
    // Create email options
    const emailOptions = {
      name: senderDisplayName || undefined
    };
    
    if (options.cc) emailOptions.cc = options.cc;
    if (options.bcc) emailOptions.bcc = options.bcc;
    if (fromEmail && fromEmail !== Session.getActiveUser().getEmail()) {
      emailOptions.from = fromEmail;
    }
    
    // Try to check quota
    let quotaLimited = false;
    let rowsToProcess = rows;
    
    try {
      const remaining = MailApp.getRemainingDailyQuota();
      
      if (remaining < rows.length) {
        rowsToProcess = rows.slice(0, remaining);
        quotaLimited = true;
        Logger.log(`Limited mail merge to ${remaining} recipients due to quota restrictions`);
      }
    } catch (quotaError) {
      Logger.log(`Could not check quota: ${quotaError.message}. Proceeding with full recipient list.`);
    }
    
    let sentCount = 0;
    let errorCount = 0;
    let errorEmails = [];
    
    // Process in batches
    const batchSize = options.batchSize;
    for (let i = 0; i < rowsToProcess.length; i += batchSize) {
      const batch = rowsToProcess.slice(i, i + batchSize);
      
      // Process this batch
      for (const row of batch) {
        const emailAddress = row[emailIndex];
        if (!emailAddress) {
          continue;
        }
        
        try {
          const personalizedSubject = replacePlaceholders(subjectLine, headers, row);
          const personalizedBody = replacePlaceholders(templateHtml, headers, row);
          
          const rowEmailOptions = Object.assign({}, emailOptions, { htmlBody: personalizedBody });
          
          if (options.createDrafts) {
            GmailApp.createDraft(emailAddress, personalizedSubject, "", rowEmailOptions);
          } else {
            GmailApp.sendEmail(emailAddress, personalizedSubject, "", rowEmailOptions);
          }
          
          sentCount++;
        } catch (e) {
          errorCount++;
          errorEmails.push(emailAddress);
          Logger.log(`Error sending to ${emailAddress}: ${e.message}`);
        }
      }
      
      // Add a small delay between batches to avoid quota issues
      if (i + batchSize < rowsToProcess.length) {
        Utilities.sleep(1000);
      }
    }
    
    return {
      success: true,
      message: `Mail merge complete. Sent: ${sentCount}, Errors: ${errorCount}${quotaLimited ? ' (Limited by quota)' : ''}`,
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
 * Performs a comprehensive validation of the data source in one step.
 * @param {string} spreadsheetUrl - The spreadsheet URL or ID
 * @param {string} sheetName - Optional sheet name to validate
 * @param {string} emailColumn - Optional email column to validate
 * @return {Object} Validation results
 */
function validateDataSource(spreadsheetUrl, sheetName, emailColumn) {
  try {
    // Add debug logs
    Logger.log(`validateDataSource called with: url=${spreadsheetUrl}, sheet=${sheetName}, emailCol=${emailColumn}`);
    
    // Step 1: Validate spreadsheet
    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Prepare response object
    const result = {
      success: true,
      spreadsheet: {
        id: spreadsheetId,
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl(),
        valid: true
      },
      sheet: { valid: false },
      emailColumn: { valid: false },
      recipientCount: 0
    };
    
    // Get all sheet names for reference
    const sheets = spreadsheet.getSheets();
    result.availableSheets = sheets.map(sheet => sheet.getName());
    
    // Step 2: Validate sheet if provided
    if (sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        result.sheet.valid = false;
        result.sheet.message = "Sheet not found: " + sheetName;
      } else {
        result.sheet.valid = true;
        result.sheet.name = sheetName;
        result.sheet.rows = sheet.getLastRow();
        result.sheet.columns = sheet.getLastColumn();
        
        // Get headers from the sheet
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        result.headers = headers.filter(header => header !== "");
        
        // Step 3: Validate email column if provided
        if (emailColumn) {
          const emailColIndex = result.headers.indexOf(emailColumn);
          if (emailColIndex === -1) {
            result.emailColumn.valid = false;
            result.emailColumn.message = "Email column not found: " + emailColumn;
          } else {
            result.emailColumn.valid = true;
            result.emailColumn.name = emailColumn;
            result.emailColumn.index = emailColIndex;
            
            // Count recipients in the email column
            if (result.sheet.rows > 1) {
              const data = sheet.getRange(2, emailColIndex + 1, result.sheet.rows - 1, 1).getValues();
              let count = 0;
              for (const row of data) {
                if (row[0] && String(row[0]).trim() !== '') {
                  count++;
                }
              }
              result.recipientCount = count;
            }
          }
        }
      }
    }
    
    Logger.log(`validateDataSource results: ${JSON.stringify(result)}`);
    return result;
  } catch (e) {
    Logger.log(`validateDataSource error: ${e.message}`);
    return {
      success: false,
      message: "Error validating data source: " + e.message
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
    
    return values;
  } catch (e) {
    Logger.log("Error getting sidebar values: " + e.message);
    return {};
  }
}

/**
 * Saves a configuration with specific values.
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
      values.documentContent = DocumentApp.getActiveDocument().getBody().getText();
      values.documentContentTimestamp = new Date().toISOString();
      values.documentName = DocumentApp.getActiveDocument().getName();
    }
    
    // Convert to new field structure
    const newConfig = {
      templateName: name,
      spreadsheetUrl: values.spreadsheetUrl || "",
      sheetName: values.sheetName || "",
      emailColumn: values.emailColumn || "",
      ccColumn: values.ccColumn || "",
      bccColumn: values.bccColumn || "",
      subjectLine: values.subjectLine || "",
      fromEmail: values.fromEmail || "",
      senderDisplayName: values.fromName || "", // Map to new field name
      description: values.templateDescription || "",
      ccOverride: values.ccOverride || "",
      bccOverride: values.bccOverride || "",
      contentFormat: "Plain text with HTML output",
      lastUpdated: new Date().toISOString(),
    };
    
    // Include document content if requested
    if (values.includeDocumentContent) {
      newConfig.documentContent = values.documentContent;
      newConfig.documentContentTimestamp = values.documentContentTimestamp;
      newConfig.documentName = values.documentName;
    }
    
    // Save configuration with template name as key
    configs[name] = newConfig;
    
    // Save back to document properties
    docProperties.setProperty('mailMergeConfigs', JSON.stringify(configs));
    
    // Set refresh flag
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return { success: true, message: 'Configuration saved successfully!' };
  } catch (e) {
    Logger.log("Error saving configuration: " + e.message);
    return { success: false, message: 'Error saving configuration: ' + e.message };
  }
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
      // Extract the content between markers
      const content = extractEmailContent(config.documentContent);
      
      // Rebuild the document with proper structure
      documentContentLoaded = rebuildTemplateDocument(config, content);
      
      if (!documentContentLoaded) {
        return { 
          success: false, 
          message: 'Error loading document content. Settings were not applied.' 
        };
      }
    }
    
    return { 
      success: true, 
      config: convertTemplateToUIConfig(config),
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
    
    // Set refresh flag
    PropertiesService.getUserProperties().setProperty('configurationUpdated', 'true');
    
    return { success: true, message: 'Configuration deleted successfully' };
  } catch (e) {
    Logger.log("Error deleting configuration: " + e.message);
    return { success: false, message: 'Error deleting configuration: ' + e.message };
  }
}

/**
 * Opens the spreadsheet in a new tab.
 * @param {string} spreadsheetUrl - The spreadsheet URL
 * @return {Object} Result with success flag
 */
function openSpreadsheet(spreadsheetUrl) {
  try {
    return {
      success: true,
      url: spreadsheetUrl
    };
  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
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

/**
 * Save template and show result to user.
 */
function saveTemplateAndShowResult() {
  const result = saveTemplateToStorage();
  const ui = DocumentApp.getUi();
  
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
}

/**
 * Backup templates to email and show result to user.
 */
function backupTemplatesToEmailAndShowResult() {
  const result = backupTemplatesToEmail();
  const ui = DocumentApp.getUi();
  
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
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
  const allPermissionsGranted = Object.values(permissionStatus).every(status => status.access);
  html += '<div class="status-card">';
  html += `<div class="status-header ${allPermissionsGranted ? 'success' : 'error'}">`;
  html += `Overall Status: ${allPermissionsGranted ? 'All permissions granted' : 'Missing permissions'}`;
  html += '</div>';
  html += '<div class="status-content">';
  if (allPermissionsGranted) {
    html += 'Mail Merge should work correctly with all required permissions.';
  } else {
    html += 'Mail Merge may not work correctly. Please check the errors above.';
  }
  html += '</div>';
  html += '</div>';
  
  html += '<button onclick="google.script.host.close()" style="padding: 8px 16px;">Close</button>';
  html += '</body></html>';
  
  const ui = DocumentApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(500)
    .setTitle('Mail Merge Permission Diagnostics');
  
  ui.showModalDialog(htmlOutput, 'Mail Merge Permission Diagnostics');
}