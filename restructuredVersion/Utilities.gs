/**
 * Utilities
 * 
 * This module contains utility functions used throughout the application:
 * - Data formatting and transformation
 * - Value extraction and manipulation
 * - Helper functions for common operations
 */

/**
 * Formats a column name for display in the UI
 * @param {string} name - Raw column name
 * @return {string} Formatted column name
 */
function formatColumnName(name) {
  if (!name) return '';
  
  // Replace underscores with spaces
  let formatted = name.replace(/_/g, ' ');
  
  // Capitalize first letter of each word
  formatted = formatted.replace(/\b\w/g, l => l.toUpperCase());
  
  return formatted;
}

/**
 * Gets a value from an object by path notation (e.g., "person.email.value")
 * @param {Object} obj - The object to extract value from
 * @param {string} path - The path to the value
 * @return {*} The value at the path or empty string if not found
 */
function getValueByPath(obj, path) {
  try {
    if (!obj || !path) return '';
    
    // Handle simple properties directly
    if (obj[path] !== undefined) {
      return obj[path];
    }
    
    // Split path into parts and traverse the object
    const parts = path.split('.');
    let current = obj;
    
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      
      if (current[part] === undefined) {
        // Special case for nested objects in Pipedrive API
        if (part === 'value' && current.value !== undefined) {
          return current.value;
        }
        
        // Special case for owner_id which is actually an object in Pipedrive
        if (part === 'owner_id' && current.owner && current.owner.id !== undefined) {
          return current.owner.id;
        }
        
        // For arrays, try to extract the first element
        if (Array.isArray(current) && current.length > 0) {
          if (i === 0) {
            // Start from the first array element
            current = current[0];
            i--; // Re-process the current part
            continue;
          }
        }
        
        return '';
      }
      
      current = current[part];
      
      // If we've reached a null or undefined value, return empty string
      if (current === null || current === undefined) {
        return '';
      }
    }
    
    // Special handling for Pipedrive API objects
    if (typeof current === 'object' && current !== null) {
      if (current.value !== undefined) {
        return current.value;
      }
      
      if (current.name !== undefined) {
        return current.name;
      }
    }
    
    return current;
  } catch (e) {
    Logger.log(`Error in getValueByPath for path ${path}: ${e.message}`);
    return '';
  }
}

/**
 * Formats a value based on its column and field type
 * @param {*} value - The value to format
 * @param {string} columnPath - The column path
 * @param {Object} optionMappings - Option mappings for enum/select fields
 * @return {string} Formatted value
 */
function formatValue(value, columnPath, optionMappings = {}) {
  try {
    // Handle null/undefined values
    if (value === null || value === undefined) {
      return '';
    }
    
    // Handle arrays
    if (Array.isArray(value)) {
      if (value.length === 0) return '';
      
      // Try to extract meaningful data from array items
      return value.map(item => {
        if (item === null || item === undefined) return '';
        
        if (typeof item === 'object') {
          // If object has name property, use it
          if (item.name !== undefined) return item.name;
          
          // If object has value property, use it
          if (item.value !== undefined) return item.value;
          
          // Fall back to stringification for other objects
          return JSON.stringify(item);
        }
        
        return item.toString();
      }).join(', ');
    }
    
    // Handle objects
    if (typeof value === 'object') {
      // If object has name property, use it
      if (value.name !== undefined) return value.name;
      
      // If object has value property, use it
      if (value.value !== undefined) return value.value;
      
      // If object has id property, try to find a label in option mappings
      if (value.id !== undefined && optionMappings[columnPath]) {
        const mapping = optionMappings[columnPath];
        const option = mapping[value.id.toString()];
        if (option) return option;
      }
      
      // Fall back to stringification for other objects
      return JSON.stringify(value);
    }
    
    // Handle boolean values
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    
    // Handle date fields
    if (columnPath.includes('date') || columnPath.includes('_at') || columnPath.includes('_time')) {
      // Try to parse as date if string
      if (typeof value === 'string') {
        try {
          const date = new Date(value);
          if (!isNaN(date)) {
            return date.toLocaleString();
          }
        } catch (e) {
          // If parsing fails, use the original value
        }
      }
    }
    
    // Handle enum/select fields
    if (optionMappings[columnPath]) {
      const mapping = optionMappings[columnPath];
      const option = mapping[value.toString()];
      if (option) return option;
    }
    
    // Default to string representation
    return value.toString();
  } catch (e) {
    Logger.log(`Error in formatValue for column ${columnPath}: ${e.message}`);
    return value ? value.toString() : '';
  }
}

/**
 * Checks if a field is a date field based on its key
 * @param {string} fieldKey - The field key to check
 * @return {boolean} True if field is a date field, false otherwise
 */
function isDateField(fieldKey) {
  if (!fieldKey) return false;
  
  const dateParts = ['date', '_at', '_time', 'deadline', 'birthday'];
  
  // Check if the field key contains any date-related parts
  for (const part of dateParts) {
    if (fieldKey.includes(part)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Checks if a field is a multi-option field
 * @param {string} fieldKey - The field key
 * @param {string} entityType - The entity type
 * @return {boolean} True if field is multi-option, false otherwise
 */
function isMultiOptionField(fieldKey, entityType) {
  try {
    // Get field definitions (cached if possible)
    let fieldDefinitions = [];
    const cacheKey = `fields_${entityType}`;

    if (!fieldDefinitionsCache[cacheKey]) {
      switch (entityType) {
        case ENTITY_TYPES.PERSONS:
          fieldDefinitions = getPersonFields();
          break;
        case ENTITY_TYPES.DEALS:
          fieldDefinitions = getDealFields();
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          fieldDefinitions = getOrganizationFields();
          break;
        case ENTITY_TYPES.ACTIVITIES:
          fieldDefinitions = getActivityFields();
          break;
        case ENTITY_TYPES.LEADS:
          fieldDefinitions = getLeadFields();
          break;
        case ENTITY_TYPES.PRODUCTS:
          fieldDefinitions = getProductFields();
          break;
        default:
          return false;
      }
      
      // Cache the field definitions
      fieldDefinitionsCache[cacheKey] = fieldDefinitions;
    } else {
      fieldDefinitions = fieldDefinitionsCache[cacheKey];
    }
    
    // Find the field definition
    const fieldDef = fieldDefinitions.find(field => field.key === fieldKey);
    if (!fieldDef) return false;
    
    // Check field type
    const isMultiOption = fieldDef.field_type === 'set' || 
                         fieldDef.field_type === 'enum' && fieldDef.edit_flag === true;
                         
    return isMultiOption;
  } catch (e) {
    Logger.log(`Error in isMultiOptionField: ${e.message}`);
    return false;
  }
}

/**
 * Converts a column letter to an index (e.g., A = 1, AA = 27)
 * @param {string} columnLetter - Column letter
 * @return {number} Column index (1-based)
 */
function columnLetterToIndex(columnLetter) {
  let column = 0;
  const length = columnLetter.length;
  
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  
  return column;
}

/**
 * Gets the option ID by its label for a specific field
 * @param {string} fieldKey - The field key
 * @param {string} optionLabel - The option label
 * @param {string} entityType - The entity type
 * @return {string} Option ID or null if not found
 */
function getOptionIdByLabel(fieldKey, optionLabel, entityType) {
  try {
    if (!fieldKey || !optionLabel || !entityType) return null;
    
    // Get field mappings
    const fieldMappings = getFieldOptionMappingsForEntity(entityType);
    if (!fieldMappings || !fieldMappings[fieldKey]) return null;
    
    // Get the mapping for this field
    const mapping = fieldMappings[fieldKey];
    
    // Look for the option ID by label
    for (const id in mapping) {
      if (mapping[id].toLowerCase() === optionLabel.toLowerCase()) {
        return id;
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`Error in getOptionIdByLabel: ${e.message}`);
    return null;
  }
}

/**
 * Detects if columns in the sheet have shifted and updates tracking accordingly
 */
function detectColumnShifts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    const scriptProperties = PropertiesService.getScriptProperties();

    // Get current and previous positions
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const currentColLetter = scriptProperties.getProperty(trackingColumnKey) || '';
    const previousPosStr = scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`) || '-1';
    const previousPos = parseInt(previousPosStr, 10);

    // Find all "Sync Status" headers in the sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColumns = [];

    // Find ALL instances of "Sync Status" headers
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        syncStatusColumns.push(i);
      }
    }

    // If we have multiple "Sync Status" columns, clean up all but the rightmost one
    if (syncStatusColumns.length > 1) {
      Logger.log(`Found ${syncStatusColumns.length} Sync Status columns`);
      // Keep only the rightmost column
      const rightmostIndex = Math.max(...syncStatusColumns);

      // Clean up all other columns
      for (const colIndex of syncStatusColumns) {
        if (colIndex !== rightmostIndex) {
          const colLetter = columnToLetter(colIndex);
          Logger.log(`Cleaning up duplicate Sync Status column at ${colLetter}`);
          cleanupColumnFormatting(sheet, colLetter);
        }
      }

      // Update the tracking to the rightmost column
      const rightmostColLetter = columnToLetter(rightmostIndex);
      scriptProperties.setProperty(trackingColumnKey, rightmostColLetter);
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, rightmostIndex.toString());
      return; // Exit after handling duplicates
    }

    let actualSyncStatusIndex = syncStatusColumns.length > 0 ? syncStatusColumns[0] : -1;

    if (actualSyncStatusIndex >= 0) {
      const actualColLetter = columnToLetter(actualSyncStatusIndex);

      // If there's a mismatch, columns might have shifted
      if (currentColLetter && actualColLetter !== currentColLetter) {
        Logger.log(`Column shift detected: was ${currentColLetter}, now ${actualColLetter}`);

        // If the actual position is less than the recorded position, columns were removed
        if (actualSyncStatusIndex < previousPos) {
          Logger.log(`Columns were likely removed (${previousPos} → ${actualSyncStatusIndex})`);

          // Clean ALL columns to be safe
          for (let i = 0; i < sheet.getLastColumn(); i++) {
            if (i !== actualSyncStatusIndex) { // Skip current Sync Status column
              cleanupColumnFormatting(sheet, columnToLetter(i));
            }
          }
        }

        // Clean up all potential previous locations
        scanAndCleanupAllSyncColumns(sheet, actualColLetter);

        // Update the tracking column property
        scriptProperties.setProperty(trackingColumnKey, actualColLetter);
        scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, actualSyncStatusIndex.toString());
      }
    }
  } catch (error) {
    Logger.log(`Error in detectColumnShifts: ${error.message}`);
  }
}

/**
 * Cleans up orphaned conditional formatting rules
 * @param {Sheet} sheet - The sheet to clean up
 * @param {number} currentColumnIndex - The index of the current Sync Status column
 */
function cleanupOrphanedConditionalFormatting(sheet, currentColumnIndex) {
  try {
    const rules = sheet.getConditionalFormatRules();
    const newRules = [];
    let removedRules = 0;

    for (const rule of rules) {
      const ranges = rule.getRanges();
      let keepRule = true;

      // Check if this rule applies to columns other than our current one
      // and has formatting that matches our Sync Status patterns
      for (const range of ranges) {
        const column = range.getColumn();

        // Skip our current column
        if (column === (currentColumnIndex + 1)) {
          continue;
        }

        // Check if this rule's formatting matches our Sync Status patterns
        const bgColor = rule.getBold() || rule.getBackground();
        if (bgColor) {
          const background = rule.getBackground();
          // If background matches our Sync Status colors, this is likely an orphaned rule
          if (background === '#FCE8E6' || background === '#E6F4EA' || background === '#F8F9FA') {
            keepRule = false;
            Logger.log(`Found orphaned conditional formatting at column ${columnToLetter(column - 1)}`);
            break;
          }
        }
      }

      if (keepRule) {
        newRules.push(rule);
      } else {
        removedRules++;
      }
    }

    if (removedRules > 0) {
      sheet.setConditionalFormatRules(newRules);
      Logger.log(`Removed ${removedRules} orphaned conditional formatting rules`);
    }
  } catch (error) {
    Logger.log(`Error cleaning up orphaned conditional formatting: ${error.message}`);
  }
}

/**
 * Cleans up formatting in a column
 * @param {Sheet} sheet - The sheet containing the column
 * @param {string} columnLetter - The letter of the column to clean up
 */
function cleanupColumnFormatting(sheet, columnLetter) {
  try {
    Logger.log(`Cleaning up formatting in column ${columnLetter}`);
    const columnIndex = columnLetterToIndex(columnLetter);
    
    // Clear the column header formatting if it's "Sync Status"
    const header = sheet.getRange(`${columnLetter}1`).getValue();
    if (header === "Sync Status") {
      sheet.getRange(`${columnLetter}1`).setValue(""); // Clear the header
    }
    
    // Clear cell formatting in this column
    const lastRow = Math.max(sheet.getLastRow(), 2);
    if (lastRow > 1) {
      const range = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1);
      range.clearContent();
      range.clearFormat();
    }
    
    // Clear validation rules in the column
    const dataValidations = sheet.getRange(1, columnIndex + 1, lastRow, 1).getDataValidations();
    const newValidations = [];
    
    for (let i = 0; i < dataValidations.length; i++) {
      newValidations.push([null]);
    }
    
    if (newValidations.length > 0) {
      sheet.getRange(1, columnIndex + 1, newValidations.length, 1).setDataValidations(newValidations);
    }
    
    // Clean up any conditional formatting rules for this column
    cleanupOrphanedConditionalFormatting(sheet, -1); // Pass -1 to clean up all
  } catch (error) {
    Logger.log(`Error cleaning up column formatting: ${error.message}`);
  }
}

/**
 * Scans and cleans up all sync status columns except the current one
 * @param {Sheet} sheet - The sheet to scan
 * @param {string} currentColumnLetter - The letter of the current sync status column
 */
function scanAndCleanupAllSyncColumns(sheet, currentColumnLetter) {
  try {
    Logger.log(`Scanning for Sync Status columns to clean up. Current: ${currentColumnLetter}`);
    const lastColumn = sheet.getLastColumn();
    
    // Loop through all columns
    for (let i = 0; i < lastColumn; i++) {
      const colLetter = columnToLetter(i);
      
      // Skip the current sync status column
      if (colLetter === currentColumnLetter) {
        continue;
      }
      
      // Check if this column has "Sync Status" as header
      const header = sheet.getRange(`${colLetter}1`).getValue();
      if (header === "Sync Status") {
        Logger.log(`Found Sync Status column at ${colLetter} to clean up`);
        cleanupColumnFormatting(sheet, colLetter);
      }
      
      // Check second row for keywords indicating it might be a sync status column
      if (sheet.getLastRow() >= 2) {
        const secondRowVal = sheet.getRange(`${colLetter}2`).getValue();
        if (typeof secondRowVal === 'string' && 
            (secondRowVal.includes('Synced') || 
             secondRowVal.includes('Modified') || 
             secondRowVal.includes('Error'))) {
          Logger.log(`Found potential Sync Status column at ${colLetter} based on content`);
          cleanupColumnFormatting(sheet, colLetter);
        }
      }
    }
  } catch (error) {
    Logger.log(`Error in scanAndCleanupAllSyncColumns: ${error.message}`);
  }
}

/**
 * Converts a date string to standard YYYY-MM-DD format for Pipedrive API
 * @param {string|Date} value - The date value to convert
 * @return {string} The formatted date string or original value if not a date
 */
function convertToStandardDateFormat(value) {
  // Skip if not a string or already in the correct format (YYYY-MM-DD)
  if (typeof value !== 'string' && !(value instanceof Date)) {
    return value;
  }

  if (typeof value === 'string') {
    // Skip if it's already in YYYY-MM-DD format
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return value;
    }

    // Skip if it's definitely not a date
    if (!/\d/.test(value)) {
      return value;
    }
  }

  try {
    // Try to parse the date using Date.parse
    const dateObj = new Date(value);

    // Check if it's a valid date
    if (isNaN(dateObj.getTime())) {
      return value;
    }

    // Format as YYYY-MM-DD
    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  } catch (e) {
    // If anything goes wrong, return the original value
    return value;
  }
}

/**
 * Converts a column index to a column letter (e.g., 1 -> A, 27 -> AA)
 * @param {number} columnIndex - Column index (1-based)
 * @return {string} Column letter
 */
function columnToLetter(columnIndex) {
  let temp;
  let letter = '';
  let col = columnIndex;
  
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  
  return letter;
} 