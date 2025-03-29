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