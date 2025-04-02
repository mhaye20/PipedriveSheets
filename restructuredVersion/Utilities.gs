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
  
  // Convert to string if not already a string
  let formatted = String(name);
  
  // Replace underscores with spaces
  formatted = formatted.replace(/_/g, ' ');
  
  // Capitalize first letter of each word
  formatted = formatted.replace(/\b\w/g, l => l.toUpperCase());
  
  return formatted;
}

/**
 * Gets a value from an object using a dot-notation path
 */
function getValueByPath(obj, path) {
  // If path is already an object with a key property, use that
  if (typeof path === 'object' && path.key) {
    path = path.key;
  }

  // Special handling for email and phone fields 
  // These are arrays with label and value properties in Pipedrive
  if (typeof path === 'string' && (path.startsWith('email.') || path.startsWith('phone.'))) {
    const parts = path.split('.');
    const fieldType = parts[0]; // 'email' or 'phone'
    const labelType = parts[1].toLowerCase(); // 'work', 'home', etc.
    
    // Get the array of email/phone objects
    const fieldArray = obj[fieldType];
    
    // If it's an array, try to find the specific type
    if (Array.isArray(fieldArray)) {
      // Try to find object with matching label
      const match = fieldArray.find(item => 
        item && item.label && item.label.toLowerCase() === labelType
      );
      
      if (match && match.value) {
        return match.value;
      }
      
      // If no match found for the specific type (like 'work'), and the type is 'primary',
      // try to find an item marked as primary
      if (labelType === 'primary') {
        const primary = fieldArray.find(item => item && item.primary);
        if (primary && primary.value) {
          return primary.value;
        }
      }
      
      // If no match found, return primary or first
      const primary = fieldArray.find(item => item && item.primary);
      if (primary && primary.value) {
        return primary.value;
      }
      
      // If no primary, return first value if available
      if (fieldArray.length > 0 && fieldArray[0] && fieldArray[0].value) {
        return fieldArray[0].value;
      }
      
      // No suitable value found
      return '';
    }
    
    // Not an array - just return the property if it exists
    return obj[fieldType] || '';
  }

  // Special handling for custom_fields in API v2
  if (path.startsWith('custom_fields.') && obj.custom_fields) {
    const parts = path.split('.');
    const fieldKey = parts[1];
    const nestedField = parts.length > 2 ? parts.slice(2).join('.') : null;

    // If the custom field exists
    if (obj.custom_fields[fieldKey] !== undefined) {
      const fieldValue = obj.custom_fields[fieldKey];

      // If we need to extract a nested property from the custom field
      if (nestedField) {
        return getValueByPath(fieldValue, nestedField);
      }

      // Otherwise return the field value itself
      return fieldValue;
    }

    return undefined;
  }

  // Handle simple non-nested case
  if (!path.includes('.')) {
    return obj[path];
  }

  // Handle nested paths
  const parts = path.split('.');
  let current = obj;

  for (const part of parts) {
    if (current === null || current === undefined) {
      return undefined;
    }

    // Handle array indexing
    if (!isNaN(part) && Array.isArray(current)) {
      const index = parseInt(part);
      current = current[index];
    } else {
      current = current[part];
    }
  }

  return current;
}

/**
 * Formats a value based on its type and column configuration
 * @param {*} value - The value to format
 * @param {Object|string} column - The column object or path
 * @param {Object} optionMappings - Option mappings for enum/select fields
 * @return {string} Formatted value
 */
function formatValue(value, columnPath, optionMappings = {}) {
  if (value === null || value === undefined) {
    return '';
  }

  // If columnPath is an object with a key property, use that key
  if (typeof columnPath === 'object' && columnPath.key) {
    columnPath = columnPath.key;
  }

  // Special handling for email and phone arrays
  if (typeof columnPath === 'string') {
    // Handle specific email/phone fields with labels like email.work, phone.mobile
    if (columnPath.startsWith('email.') || columnPath.startsWith('phone.')) {
      const parts = columnPath.split('.');
      const fieldType = parts[0]; // 'email' or 'phone'
      const labelType = parts[1].toLowerCase(); // 'work', 'home', etc.
      
      // Get the array of email/phone objects
      const fieldArray = value;
      
      // If it's an array, try to find the specific type
      if (Array.isArray(fieldArray)) {
        // Try to find object with matching label
        const match = fieldArray.find(item => 
          item && item.label && item.label.toLowerCase() === labelType
        );
        
        if (match && match.value) {
          return match.value;
        }
      }
      
      // Not found or not an array - use general handling below
    }
    // Handle base email/phone fields (get primary)
    else if (columnPath === 'email' || columnPath === 'phone') {
      // If the value is an array of contact objects
      if (Array.isArray(value)) {
        if (value.length === 0) {
          return '';
        }
        
        // First try to find primary value
        const primary = value.find(item => item && item.primary);
        if (primary && primary.value) {
          return primary.value;
        }
        
        // If no primary, return first value if it exists
        if (value[0] && value[0].value) {
          return value[0].value;
        }
        
        // Otherwise return empty string
        return '';
      }
      
      // If it's a simple value already
      if (typeof value === 'string') {
        return value;
      }
    }
  }

  // Handle custom fields from API v2
  // In API v2, custom fields are in a nested object
  if (columnPath.startsWith('custom_fields.')) {
    // The field key is the part after "custom_fields."
    const fieldKey = columnPath.split('.')[1];

    // If it's a multiple option field in new format (array of numbers)
    if (Array.isArray(value)) {
      // Check if we have option mappings for this field
      if (optionMappings[fieldKey]) {
        const labels = value.map(id => {
          // Return the label if we have it, otherwise just return the ID
          return optionMappings[fieldKey][id] || id;
        });
        return labels.join(', ');
      }
      return value.join(', ');
    }

    // Handle currency fields and other object-based custom fields
    if (typeof value === 'object' && value !== null) {
      if (value.value !== undefined && value.currency !== undefined) {
        return `${value.value} ${value.currency}`;
      }
      if (value.value !== undefined && value.until !== undefined) {
        return `${value.value} - ${value.until}`;
      }
      // For address fields
      if (value.value !== undefined && value.formatted_address !== undefined) {
        return value.formatted_address;
      }
      return JSON.stringify(value);
    }

    // For single option fields (now just a number)
    if (typeof value === 'number' && optionMappings[fieldKey]) {
      return optionMappings[fieldKey][value] || value;
    }
  }

  // Handle comma-separated IDs for multiple options fields (API v1 format - for backward compatibility)
  if (typeof value === 'string' && /^[0-9]+(,[0-9]+)*$/.test(value)) {
    // This looks like a comma-separated list of IDs, likely a multiple option field

    // Extract the field key from the column path (remove any array or nested indices)
    let fieldKey;

    // Special handling for custom_fields paths
    if (columnPath.startsWith('custom_fields.')) {
      // For custom fields, the field key is after "custom_fields."
      fieldKey = columnPath.split('.')[1];
    } else {
      // For regular fields, use the first part of the path
      fieldKey = columnPath.split('.')[0];
    }

    // Check if we have option mappings for this field
    if (optionMappings[fieldKey]) {
      const ids = value.split(',');
      const labels = ids.map(id => {
        // Return the label if we have it, otherwise just return the ID
        return optionMappings[fieldKey][id] || id;
      });
      return labels.join(', ');
    }
  }

  // Regular object handling
  if (typeof value === 'object') {
    // Check if it's an array of email/phone objects
    if (Array.isArray(value) && value.length > 0 && 
        value[0] && typeof value[0] === 'object' && 
        value[0].value !== undefined) {
      
      // If the objects have label and value properties (email/phone array)
      if (value[0].hasOwnProperty('value') && value[0].hasOwnProperty('label')) {
        // First try to find primary contact
        const primary = value.find(item => item && item.primary);
        if (primary && primary.value) {
          return primary.value;
        }
        
        // Otherwise return first value
        return value[0].value;
      }
      
      // For other arrays of objects with label property, extract and join labels
      if (value[0].label !== undefined) {
        return value.map(option => option.label).join(', ');
      }
    }
    // Check if it's a single option object
    else if (value.label !== undefined) {
      return value.label;
    }
    // Handle person/org objects
    else if (value.name !== undefined) {
      return value.name;
    }
    // Handle currency objects
    else if (value.currency !== undefined && value.value !== undefined) {
      return `${value.value} ${value.currency}`;
    }
    // For other objects, convert to JSON string
    return JSON.stringify(value);
  } else if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }

  return value.toString();
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

/**
 * Global include function for HTML templates
 * Allows templates to include other templates or script files
 * @param {string} filename - The name of the file to include
 * @return {string} The content of the file
 */
function include(filename) {
  // Handle special cases for TriggerManager UI components
  if (filename === 'TriggerManagerUI_Styles') {
    return TriggerManagerUI.getStyles();
  } else if (filename === 'TriggerManagerUI_Scripts') {
    return TriggerManagerUI.getScripts();
  }
  
  // Handle special cases for TwoWaySyncSettings UI components
  if (filename === 'TwoWaySyncSettingsUI_Styles') {
    return TwoWaySyncSettingsUI.getStyles();
  } else if (filename === 'TwoWaySyncSettingsUI_Scripts') {
    return TwoWaySyncSettingsUI.getScripts();
  }
  
  // For standard includes, return the content of the file
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
} 