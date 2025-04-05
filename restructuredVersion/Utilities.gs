/**
 * Utilities
 * 
 * This module contains utility functions used throughout the application:
 * - Data formatting and transformation
 * - Value extraction and manipulation
 * - Helper functions for common operations
 */

// Define a cache for field definitions to avoid repeated API calls
var fieldDefinitionsCache = {};

/**
 * Formats a column name for display in the UI
 * @param {string} name - Raw column name
 * @return {string} Formatted column name
 */
function formatColumnName(name) {
  if (!name) return '';
  
  // Convert to string if not already a string
  let key = String(name);
  
  // First check for standard entity object properties (person_id.name, org_id.address, etc.)
  if (key.includes('.')) {
    const parts = key.split('.');
    if (parts.length === 2) {
      const rootObject = parts[0];
      const property = parts[1];
      
      // Common entity root objects
      if (rootObject === 'person_id') {
        return `Contact - ${formatBasicName(property)}`;
      } else if (rootObject === 'org_id') {
        return `Organization - ${formatBasicName(property)}`;
      } else if (rootObject === 'user_id') {
        return `User - ${formatBasicName(property)}`;
      } else if (rootObject === 'creator_user_id') {
        return `Creator - ${formatBasicName(property)}`;
      } else if (rootObject === 'owner_id') {
        return `Owner - ${formatBasicName(property)}`;
      }
      
      // Array types like email.0.value or phone.0.label
      if (parts[1].match(/^\d+$/)) {
        // It's an array index
        if (parts.length >= 3) {
          // Something like email.0.value - use the third part
          const itemType = parts[0]; // email, phone, etc.
          const property = parts[2]; // value, label, etc.
          return `${formatBasicName(itemType)} ${formatBasicName(property)} (${parts[1]})`;
        }
      }
      
      // For custom field objects that have properties
      if (/[a-f0-9]{20,}/i.test(rootObject)) {
        const customFieldName = getCustomFieldNameFromCache(rootObject);
        return `${customFieldName} - ${formatBasicName(property)}`;
      }
      
      // Generic nested property
      return `${formatBasicName(rootObject)} - ${formatBasicName(property)}`;
    }
    
    // More than 2 parts (like email.0.value)
    if (parts.length > 2) {
      const collection = parts[0]; // email, phone, etc.
      const index = parts[1]; // usually a number
      const property = parts[2]; // value, label, etc.
      
      // For lists of emails, phones, etc.
      if (collection === 'email' || collection === 'phone') {
        return `${formatBasicName(collection)} ${formatBasicName(property)} #${Number(index) + 1}`;
      }
      
      // For other nested properties
      return `${formatBasicName(collection)} - ${formatBasicName(property)} (${index})`;
    }
  }
  
  // Check for custom field with hash ID and component
  if (/[a-f0-9]{20,}_[a-z_]+/i.test(key)) {
    // This is a custom field component (address part, etc.)
    const hashMatch = key.match(/([a-f0-9]{20,})_([a-z_]+)/i);
    if (hashMatch) {
      const hashId = hashMatch[1];
      const component = hashMatch[2];
      
      // Get base field name from cache
      let baseName = getCustomFieldNameFromCache(hashId);
      
      // Handle address components
      if (component === 'subpremise') {
        return `${baseName} - Suite/Apt`;
      } else if (component === 'street_number') {
        return `${baseName} - Street Number`;
      } else if (component === 'route') {
        return `${baseName} - Street Name`;
      } else if (component === 'locality') {
        return `${baseName} - City`;
      } else if (component === 'sublocality') {
        return `${baseName} - District`;
      } else if (component === 'admin_area_level_1') {
        return `${baseName} - State/Province`;
      } else if (component === 'admin_area_level_2') {
        return `${baseName} - County`;
      } else if (component === 'country') {
        return `${baseName} - Country`;
      } else if (component === 'postal_code') {
        return `${baseName} - ZIP/Postal Code`;
      } else if (component === 'formatted_address') {
        return `${baseName} - Complete Address`;
      }
      
      // Handle time and date components
      if (component === 'until') {
        return `${baseName} - End Time/Date`;
      } else if (component === 'timezone_id') {
        return `${baseName} - Timezone`;
      }
      
      // Handle currency and other components
      if (component === 'currency') {
        return `${baseName} - Currency`;
      }
      
      // Generic component
      return `${baseName} - ${formatBasicName(component)}`;
    }
  }
  
  // Main custom field (just the hash ID)
  if (/^[a-f0-9]{20,}$/i.test(key)) {
    return getCustomFieldNameFromCache(key);
  }
  
  // Custom field with hash embedded in it
  if (/[a-f0-9]{20,}/i.test(key)) {
    const hashMatch = key.match(/([a-f0-9]{20,})/i);
    if (hashMatch) {
      return getCustomFieldNameFromCache(hashMatch[1]);
    }
  }
  
  // Standard Pipedrive fields
  if (key === 'id') return 'Pipedrive ID';
  if (key === 'name') return 'Name';
  if (key === 'title') return 'Deal Title';
  if (key === 'value') return 'Deal Value';
  if (key === 'currency') return 'Currency';
  if (key === 'status') return 'Status';
  if (key === 'stage_id') return 'Pipeline Stage';
  if (key === 'pipeline_id') return 'Pipeline';
  if (key === 'expected_close_date') return 'Expected Close Date';
  if (key === 'won_time') return 'Won Date';
  if (key === 'lost_time') return 'Lost Date';
  if (key === 'close_time') return 'Closed Date';
  if (key === 'add_time') return 'Created Date';
  if (key === 'update_time') return 'Last Updated';
  if (key === 'active') return 'Active Status';
  if (key === 'deleted') return 'Deleted Status';
  if (key === 'visible_to') return 'Visibility';
  if (key === 'owner_name') return 'Owner Name';
  if (key === 'cc_email') return 'CC Email';
  if (key === 'org_name') return 'Organization Name';
  if (key === 'person_name') return 'Contact Name';
  if (key === 'next_activity_date') return 'Next Activity Date';
  if (key === 'next_activity_time') return 'Next Activity Time';
  if (key === 'next_activity_id') return 'Next Activity ID';
  if (key === 'last_activity_id') return 'Last Activity ID';
  if (key === 'last_activity_date') return 'Last Activity Date';
  
  // Default formatting
  return formatBasicName(key);
}

/**
 * Gets a custom field name from the cache
 * @param {string} hashId - The hash ID of the custom field
 * @return {string} The formatted name of the custom field
 */
function getCustomFieldNameFromCache(hashId) {
  if (typeof fieldDefinitionsCache.customFields !== 'undefined' && 
      fieldDefinitionsCache.customFields[hashId]) {
    return fieldDefinitionsCache.customFields[hashId];
  }
  
  // If not in cache, use a generic name
  return "Custom Field";
}

/**
 * Basic name formatting for column names
 * @param {string} name - Name to format
 * @return {string} Formatted name
 */
function formatBasicName(name) {
  // Replace underscores with spaces
  let formatted = name.replace(/_/g, ' ');
  
  // Replace dots with spaces (common in nested fields)
  formatted = formatted.replace(/\./g, ' ');
  
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

  // Special handling for address fields
  if (path.startsWith('address.')) {
    const parts = path.split('.');
    const component = parts[1];
    
    Logger.log(`Getting address component: ${component} from object`);
    
    // First check if we have a special flattened format like org['address.subpremise']
    if (obj[path] !== undefined) {
      Logger.log(`Found address component as direct flattened field: ${path} = ${obj[path]}`);
      return obj[path];
    }
    
    // Check if the address exists as an object
    if (obj.address && typeof obj.address === 'object') {
      if (obj.address[component] !== undefined) {
        Logger.log(`Found address component in address object: ${component} = ${obj.address[component]}`);
        return obj.address[component];
      } else {
        Logger.log(`Address component ${component} not found in address object: ${JSON.stringify(obj.address)}`);
      }
    } else {
      Logger.log(`Address object not found or not an object: ${JSON.stringify(obj.address)}`);
    }
    
    // If we get here, we didn't find the address component
    return undefined;
  }
  // Special handling for custom field address components
  else if (path.includes('_subpremise') || path.includes('_locality') || 
         path.includes('_sublocality') ||
         path.includes('_formatted_address') || path.includes('_street_number') ||
         path.includes('_route') || path.includes('_admin_area') || path.includes('_country') ||
         path.includes('_postal_code')) {
  
    // Custom field address components are stored at the top level of the item
    // They use the custom field ID + _ + component name pattern
    if (obj[path] !== undefined) {
      // Get component type for better logging
      const componentType = path.split('_').slice(1).join('_');
      Logger.log(`Found custom address component as direct field: ${path} = ${obj[path]} (${componentType})`);
      return obj[path];
    }
    
    // Special logging for locality vs sublocality to help debug
    if (path.includes('_locality') || path.includes('_sublocality')) {
      const componentType = path.split('_').slice(1).join('_');
      Logger.log(`Address component ${componentType} not found directly at ${path}`);
      
      // Try to find if the component exists under a different path
      if (componentType === 'locality' && path.replace('_locality', '_sublocality') in obj) {
        Logger.log(`Note: Found sublocality where locality was requested`);
      }
      else if (componentType === 'sublocality' && path.replace('_sublocality', '_locality') in obj) {
        Logger.log(`Note: Found locality where sublocality was requested`);
      }
    }
    
    return undefined;
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

  // Special handling for address fields
  if (columnPath.startsWith('address.')) {
    // The address component is the part after "address."
    const addressComponent = columnPath.split('.')[1];
    Logger.log(`Formatting address component: ${addressComponent} with value: ${value}`);
    
    // If the value is an object and has that component, extract it
    if (typeof value === 'object' && value !== null && value[addressComponent] !== undefined) {
      Logger.log(`Address component ${addressComponent} found in object: ${value[addressComponent]}`);
      return value[addressComponent].toString();
    }
    
    // Special handling for specific components to make them more readable
    if (addressComponent === 'subpremise' && value) {
      return `Apt/Suite: ${value}`;
    }
    
    // For direct string values (already extracted)
    if (value !== null && value !== undefined) {
      return value.toString();
    }
    
    // If we couldn't extract anything, return empty string
    return '';
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
 * Allows templates to include other templates or script files
 * @param {string} filename - The name of the file to include
 * @return {string} The content of the file
 */
function include(filename) {
  // Handle special cases for TriggerManager UI components
  if (filename === 'TriggerManagerUI_Styles') {
    return HtmlService.createHtmlOutputFromFile('TriggerManagerUI_Styles').getContent();
  } else if (filename === 'TriggerManagerUI_Scripts') {
    return HtmlService.createHtmlOutputFromFile('TriggerManagerUI_Scripts').getContent();
  }
  
  // Handle special cases for TwoWaySyncSettings UI components
  if (filename === 'TwoWaySyncSettingsUI_Styles') {
    return TwoWaySyncSettingsUI.getStyles();
  } else if (filename === 'TwoWaySyncSettingsUI_Scripts') {
    return TwoWaySyncSettingsUI.getScripts();
  }
  
  // Handle special case for ColumnSelectorUI
  if (filename === 'ColumnSelectorUI') {
    return ColumnSelectorUI.getScripts();
  }
  
  // For standard includes, return the content of the file
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initializes the custom field definitions cache with field names from Pipedrive
 * @param {string} entityType - The entity type (deals, persons, etc.)
 * @return {Object} - The field definitions cache
 */
function initializeCustomFieldsCache(entityType) {
  try {
    if (!fieldDefinitionsCache.customFields) {
      fieldDefinitionsCache.customFields = {};
    }
    
    // Get field definitions based on entity type
    let fields = [];
    
    switch (entityType) {
      case 'deals':
      case ENTITY_TYPES.DEALS:
        fields = getDealFields();
        break;
      case 'persons':
      case ENTITY_TYPES.PERSONS:
        fields = getPersonFields();
        break;
      case 'organizations':
      case ENTITY_TYPES.ORGANIZATIONS:
        fields = getOrganizationFields();
        break;
      case 'activities':
      case ENTITY_TYPES.ACTIVITIES:
        fields = getActivityFields();
        break;
      case 'leads':
      case ENTITY_TYPES.LEADS:
        fields = getLeadFields();
        break;
      case 'products':
      case ENTITY_TYPES.PRODUCTS:
        fields = getProductFields();
        break;
      default:
        // Try to get fields for all entity types if not specified
        fields = [
          ...getDealFields(),
          ...getPersonFields(),
          ...getOrganizationFields(),
          ...getActivityFields(),
          ...getLeadFields(),
          ...getProductFields()
        ];
    }
    
    Logger.log(`Processing ${fields.length} fields for entity type ${entityType} to build field cache`);
    
    // Process all fields including regular and custom fields
    for (const field of fields) {
      if (field.key && field.name) {
        // Add entity type prefix for standard Pipedrive fields to avoid confusion
        let displayName = field.name;
        
        // Add prefix for standard fields to better categorize them
        if (!field.edit_flag && !(/[a-f0-9]{20,}/i.test(field.key))) {
          const entityPrefix = getEntityPrefix(entityType);
          if (!displayName.startsWith(entityPrefix)) {
            displayName = `${entityPrefix} ${displayName}`;
          }
        }
        
        // Store both custom fields (hash IDs) and regular fields
        fieldDefinitionsCache.customFields[field.key] = displayName;
        
        // Also handle field keys with hash IDs embedded in them
        if (/[a-f0-9]{20,}/i.test(field.key)) {
          // Extract the hash part
          const hashMatch = field.key.match(/([a-f0-9]{20,})/i);
          if (hashMatch && hashMatch[1]) {
            const hashId = hashMatch[1];
            fieldDefinitionsCache.customFields[hashId] = displayName;
            
            // Also add entries for address components
            if (field.field_type === 'address') {
              const addressComponents = [
                '_formatted_address',
                '_subpremise',
                '_street_number',
                '_route',
                '_locality',
                '_sublocality',
                '_admin_area_level_1',
                '_admin_area_level_2',
                '_country',
                '_postal_code'
              ];
              
              // Add component-specific names
              addressComponents.forEach(component => {
                const componentKey = `${hashId}${component}`;
                let componentName = '';
                
                switch (component) {
                  case '_formatted_address':
                    componentName = `${displayName} - Complete Address`;
                    break;
                  case '_subpremise':
                    componentName = `${displayName} - Suite/Apt`;
                    break;
                  case '_street_number':
                    componentName = `${displayName} - Street Number`;
                    break;
                  case '_route':
                    componentName = `${displayName} - Street Name`;
                    break;
                  case '_locality':
                    componentName = `${displayName} - City`;
                    break;
                  case '_sublocality':
                    componentName = `${displayName} - District/Borough`;
                    break;
                  case '_admin_area_level_1':
                    componentName = `${displayName} - State/Province`;
                    break;
                  case '_admin_area_level_2':
                    componentName = `${displayName} - County`;
                    break; 
                  case '_country':
                    componentName = `${displayName} - Country`;
                    break;
                  case '_postal_code':
                    componentName = `${displayName} - ZIP/Postal Code`;
                    break;
                  default:
                    componentName = `${displayName} - ${component.substring(1)}`;
                }
                
                fieldDefinitionsCache.customFields[componentKey] = componentName;
              });
            }
          }
        }
      }
    }
    
    // Also process embedded fields like org_id, person_id, etc.
    const embeddedFields = {
      'org_id': 'Organization',
      'person_id': 'Contact',
      'user_id': 'User',
      'creator_user_id': 'Creator',
      'owner_id': 'Owner'
    };
    
    for (const [key, name] of Object.entries(embeddedFields)) {
      fieldDefinitionsCache.customFields[key] = name;
      
      // Add nested fields
      const nestedFields = {
        'name': `${name} Name`,
        'email': `${name} Email`,
        'phone': `${name} Phone`,
        'address': `${name} Address`,
        'active_flag': `${name} Active Status`
      };
      
      for (const [nestedKey, nestedName] of Object.entries(nestedFields)) {
        fieldDefinitionsCache.customFields[`${key}.${nestedKey}`] = nestedName;
      }
    }
    
    Logger.log(`Initialized custom fields cache with ${Object.keys(fieldDefinitionsCache.customFields).length} fields`);
    
    // Add some diagnostic logging
    const sampleEntries = Object.entries(fieldDefinitionsCache.customFields).slice(0, 5);
    Logger.log(`Sample field mappings: ${JSON.stringify(sampleEntries)}`);
    
    return fieldDefinitionsCache.customFields;
  } catch (e) {
    Logger.log(`Error initializing custom fields cache: ${e.message}`);
    Logger.log(`Error stack: ${e.stack}`);
    return {};
  }
}

/**
 * Gets the entity prefix for better field categorization
 * @param {string} entityType - The entity type
 * @return {string} The entity prefix
 */
function getEntityPrefix(entityType) {
  switch (entityType) {
    case 'deals':
    case ENTITY_TYPES.DEALS:
      return 'Deal';
    case 'persons':
    case ENTITY_TYPES.PERSONS:
      return 'Contact';
    case 'organizations':
    case ENTITY_TYPES.ORGANIZATIONS:
      return 'Organization';
    case 'activities':
    case ENTITY_TYPES.ACTIVITIES:
      return 'Activity';
    case 'leads':
    case ENTITY_TYPES.LEADS:
      return 'Lead';
    case 'products':
    case ENTITY_TYPES.PRODUCTS:
      return 'Product';
    default:
      return 'Pipedrive';
  }
}

/**
 * Gets a formatted column name for use in the UI
 * This is a wrapper function for formatColumnName that can be called from the UI
 * @param {string} key - The column key to format
 * @return {string} The formatted column name
 */
function getFormattedColumnName(key) {
  return formatColumnName(key);
} 