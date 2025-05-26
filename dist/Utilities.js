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
 * @param {string} name - Raw column name (usually the Pipedrive field key)
 * @return {string} Formatted column name
 */
function formatColumnName(key) {
  if (!key) return '';
  key = String(key);

  // 1. Handle Nested Fields (e.g., 'person_id.name')
  if (key.includes('.')) {
    const parts = key.split('.');
    // Basic case: parent.property
    if (parts.length === 2) {
      const parentKey = parts[0];
      const property = parts[1];
      const parentName = formatColumnName(parentKey); // Recursively format parent
      const propertyName = formatBasicName(property);
      // Avoid redundant names like "Organization - Organization Name"
      if (parentName && propertyName.toLowerCase().includes(parentName.toLowerCase())) {
        return propertyName;
      } 
      // Handle Owner special case
      else if (parentKey === 'owner_id' && property === 'name') {
        return 'Owner Name'; // Keep it simple
      }
      // General case
      return `${parentName} - ${propertyName}`;
    }
    // Deeper nesting (e.g., email.0.value) - less common, basic format
    else if (parts.length > 2) {
       const collection = parts[0];
       const index = parts[1];
       const property = parts[2];
       return `${formatBasicName(collection)} ${formatBasicName(property)} #${Number(index) + 1}`;
    }
  }

  // 2. Handle Custom Field Components (e.g., 'HASHID_locality')
  if (/[a-f0-9]{20,}_[a-z_]+/i.test(key)) {
    const hashMatch = key.match(/([a-f0-9]{20,})_([a-z_]+)/i);
    if (hashMatch) {
      const hashId = hashMatch[1];
      const component = hashMatch[2];
      
      // Get the RAW base name from cache
      let baseName = getRawFieldNameFromCache(hashId) || 'Custom Field';
      
      // Define component suffixes
      const componentSuffixes = {
        'subpremise': 'Suite/Apt',
        'street_number': 'Street Number',
        'route': 'Street Name',
        'locality': 'City',
        'sublocality': 'District',
        'admin_area_level_1': 'State/Province',
        'admin_area_level_2': 'County',
        'country': 'Country',
        'postal_code': 'ZIP/Postal Code',
        'formatted_address': 'Complete Address',
        'until': 'End Time/Date', // Date/Time range
        'timezone_id': 'Timezone', // Time
        'currency': 'Currency' // Monetary
      };

      // Return formatted name: "Base Name - Component Suffix"
      if (componentSuffixes[component]) {
         return `${baseName} - ${componentSuffixes[component]}`;
      }
      // Fallback for unknown components
      return `${baseName} - ${formatBasicName(component)}`;
    }
  }

  // 3. Handle Base Custom Fields (HASHID only)
  if (/^[a-f0-9]{20,}$/i.test(key)) {
    return getRawFieldNameFromCache(key) || 'Custom Field';
  }

  // 4. Handle Standard Pipedrive Fields (using cache if available, then specific overrides)
  const cachedName = getRawFieldNameFromCache(key);
  if (cachedName) {
    // Apply specific overrides for clarity even if cached
    if (key === 'title' && cachedName === 'Title') return 'Deal Title';
    if (key === 'person_id') return 'Contact'; // Prefer Contact over Person
    if (key === 'org_id') return 'Organization';
    if (key === 'owner_id') return 'Owner';
    if (key === 'user_id') return 'User';
    if (key === 'creator_user_id') return 'Creator';
    if (key === 'stage_id') return 'Pipeline Stage';
    if (key === 'pipeline_id') return 'Pipeline';
    return cachedName; // Use the name directly from Pipedrive API
  }
  
  // Specific overrides if not found in cache (or cache is empty)
  const specificOverrides = {
    'id': 'Pipedrive ID',
    'name': 'Name',
    'title': 'Deal Title',
    'value': 'Deal Value',
    'currency': 'Currency',
    'status': 'Status',
    'stage_id': 'Pipeline Stage',
    'pipeline_id': 'Pipeline',
    'expected_close_date': 'Expected Close Date',
    'won_time': 'Won Date',
    'lost_time': 'Lost Date',
    'close_time': 'Closed Date',
    'add_time': 'Created Date',
    'update_time': 'Last Updated',
    'active': 'Active Status',
    'deleted': 'Deleted Status',
    'visible_to': 'Visibility',
    'owner_id': 'Owner',
    'owner_name': 'Owner Name', // Covered by nested logic, but good fallback
    'org_id': 'Organization',
    'org_name': 'Organization Name', // Covered by nested logic
    'person_id': 'Contact', // Prefer Contact over Person
    'person_name': 'Contact Name', // Covered by nested logic
    'user_id': 'User',
    'creator_user_id': 'Creator',
    'cc_email': 'CC Email',
    'next_activity_date': 'Next Activity Date',
    'next_activity_time': 'Next Activity Time',
    'next_activity_id': 'Next Activity ID',
    'last_activity_id': 'Last Activity ID',
    'last_activity_date': 'Last Activity Date'
    // Add more standard fields as needed
  };
  if (specificOverrides[key]) {
    return specificOverrides[key];
  }

  // 5. Fallback: Basic Formatting
  return formatBasicName(key);
}

/**
 * Gets a custom field name from the cache
 * @param {string} hashId - The hash ID of the custom field
 * @return {string} The formatted name of the custom field
 */
function getCustomFieldNameFromCache(key) {
  if (typeof fieldDefinitionsCache.customFields !== 'undefined' && 
      fieldDefinitionsCache.customFields[key]) {
    return fieldDefinitionsCache.customFields[key];
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
 * Helper function to pad numbers with leading zeros for time formatting
 * @param {number} num - The number to pad
 * @return {string} Padded number string
 */
function padZero(num) {
  return String(num).padStart(2, '0');
}

/**
 * Formats a time value into HH:MM string format for Pipedrive API
 * @param {*} value - The time value (Date object, string like "HH:MM", "HH:MM:SS", "H:MM AM/PM")
 * @return {string|null} Formatted time string "HH:MM" or null if invalid
 */
function formatTimeField(timeValue) {
  if (!timeValue) return null;
  
  try {
    let hour = null, minute = null;
    
    // If already a properly formatted object with hour and minute
    if (typeof timeValue === 'object' && timeValue !== null &&
        timeValue.hour !== undefined && timeValue.minute !== undefined) {
      hour = parseInt(timeValue.hour, 10);
      minute = parseInt(timeValue.minute, 10);
      Logger.log(`[TIME DEBUG] Already an object with hour/minute: ${JSON.stringify(timeValue)}`);
    }
    // If it's a Date object
    else if (timeValue instanceof Date) {
      hour = timeValue.getHours();
      minute = timeValue.getMinutes();
      Logger.log(`[TIME DEBUG] Extracted from Date object: ${hour}:${minute}`);
    }
    // If it's a string in format HH:MM or H:MM
    else if (typeof timeValue === 'string') {
      const match = timeValue.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
      if (match) {
        hour = parseInt(match[1], 10);
        minute = parseInt(match[2], 10);
        Logger.log(`[TIME DEBUG] Parsed from HH:MM format: ${hour}:${minute}`);
      }
      // Try to parse as ISO date
      else if (timeValue.includes('T')) {
        try {
          const date = new Date(timeValue);
          if (!isNaN(date.getTime())) {
            hour = date.getHours();
            minute = date.getMinutes();
            Logger.log(`[TIME DEBUG] Extracted from ISO date: ${hour}:${minute}, original value: ${timeValue}`);
          } else {
            Logger.log(`[TIME DEBUG] Invalid ISO date: ${timeValue}`);
          }
        } catch (err) {
          Logger.log(`[TIME DEBUG] Error parsing ISO date: ${timeValue}, error: ${err.message}`);
        }
      }
      // Try parsing "1:30 PM" format
      else if (timeValue.match(/\d{1,2}:\d{2}\s*(am|pm)/i)) {
        const parts = timeValue.match(/(\d{1,2}):(\d{2})\s*(am|pm)/i);
        if (parts) {
          hour = parseInt(parts[1], 10);
          const ampm = parts[3].toLowerCase();
          
          if (ampm === 'pm' && hour < 12) {
            hour += 12;
          } else if (ampm === 'am' && hour === 12) {
            hour = 0;
          }
          
          minute = parseInt(parts[2], 10);
          Logger.log(`[TIME DEBUG] Parsed from AM/PM format: ${hour}:${minute}`);
        }
      }
    }
    // If it's a number, assume it's milliseconds since epoch
    else if (typeof timeValue === 'number') {
      const date = new Date(timeValue);
      if (!isNaN(date.getTime())) {
        hour = date.getHours();
        minute = date.getMinutes();
        Logger.log(`[TIME DEBUG] Extracted from timestamp: ${hour}:${minute}`);
      }
    }
    
    // Validate hour and minute
    if (hour !== null && minute !== null && 
        hour >= 0 && hour <= 23 && 
        minute >= 0 && minute <= 59) {
      // Return as object with hour and minute properties as Pipedrive API expects
      const result = { hour: hour, minute: minute };
      Logger.log(`[TIME DEBUG] Returning time object: ${JSON.stringify(result)}`);
      return result;
    }
    
    Logger.log(`[TIME DEBUG] Failed to format time value: ${JSON.stringify(timeValue)}`);
    return null;
  } catch (e) {
    Logger.log(`Error in formatTimeField: ${e.message}`);
    return null;
  }
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
      // For other objects, return a more readable format
      if (value.value !== undefined) {
        return String(value.value);
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

  // Handle label_ids field specifically (for leads)
  if (columnPath === 'label_ids' && Array.isArray(value)) {
    if (value.length === 0) {
      return '';
    }
    // Check if we have label mappings
    if (optionMappings['label_ids']) {
      const labels = value.map(id => {
        // Return the label name if we have it, otherwise just return the ID
        return optionMappings['label_ids'][id] || id;
      });
      return labels.join(', ');
    }
    // If no mappings, just join the IDs
    return value.join(', ');
  }

  // Handle prices field specifically
  if (columnPath === 'prices' && Array.isArray(value)) {
    if (value.length === 0) {
      return '';
    }
    // Format prices array to show price and currency
    const priceStrings = value.map(priceObj => {
      if (priceObj && priceObj.price !== undefined && priceObj.currency) {
        // Show cost if it's different from price
        if (priceObj.cost && priceObj.cost !== 0 && priceObj.cost !== priceObj.price) {
          return `${priceObj.price} (cost: ${priceObj.cost})`;
        }
        return `${priceObj.price}`;
      }
      return JSON.stringify(priceObj);
    });
    // Join multiple prices with semicolons
    return priceStrings.join('; ');
  }

  // Handle participants field specifically (for activities)
  if (columnPath === 'participants') {
    // If it's a string (from sheet), return as is
    if (typeof value === 'string') {
      return value;
    }
    
    // If it's an array (from Pipedrive)
    if (Array.isArray(value)) {
      if (value.length === 0) {
        return '';
      }
      
      // Just show person IDs for now - we'll map them to names in a separate process
      const participantIds = value.map(participant => {
        if (participant && participant.person_id) {
          return participant.person_id;
        }
        return '';
      }).filter(id => id); // Remove empty values
      
      // Return as comma-separated IDs
      return participantIds.join(',');
    }
    
    // For other types, convert to string
    return String(value);
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
 * @return {Object} - The field definitions cache (stores raw field names)
 */
function initializeCustomFieldsCache(entityType) {
  try {
    // Initialize cache object if it doesn't exist
    if (typeof fieldDefinitionsCache.rawNames === 'undefined') {
      fieldDefinitionsCache.rawNames = {};
      Logger.log('Initialized new rawNames cache.');
    }
    
    // Determine which fields to fetch based on entityType
    let fieldsToFetch = [];
    if (entityType) {
      switch (entityType.toLowerCase()) {
        case 'deals': fieldsToFetch.push(() => getDealFields(true)); break;
        case 'persons': fieldsToFetch.push(() => getPersonFields(true)); break;
        case 'organizations': fieldsToFetch.push(() => getOrganizationFields(true)); break;
        case 'activities': fieldsToFetch.push(() => getActivityFields(true)); break;
        case 'leads': fieldsToFetch.push(() => getLeadFields(true)); break;
        case 'products': fieldsToFetch.push(() => getProductFields(true)); break;
        default: Logger.log(`Unknown specific entity type for cache: ${entityType}`); break;
      }
    } else {
      // If no entityType, fetch all (useful for initial load or general formatting)
      Logger.log('No specific entity type, fetching fields for all types (forcing refresh).');
      fieldsToFetch = [
        () => getDealFields(true),
        () => getPersonFields(true),
        () => getOrganizationFields(true),
        () => getActivityFields(true),
        () => getLeadFields(true),
        () => getProductFields(true)
      ];
    }
    
    let totalFieldsProcessed = 0;
    fieldsToFetch.forEach(fetchFn => {
      const fields = fetchFn();
      totalFieldsProcessed += fields.length;
      // Process fetched fields
      for (const field of fields) {
        if (field.key && field.name) {
          // Store the RAW field name from Pipedrive API
          fieldDefinitionsCache.rawNames[field.key] = field.name;
          
          // Also store the base name for custom fields under their HASH ID key
          if (/[a-f0-9]{20,}/i.test(field.key)) {
            const hashMatch = field.key.match(/([a-f0-9]{20,})/i);
            if (hashMatch && hashMatch[1]) {
              const hashId = hashMatch[1];
              // Store the base name if we haven't already stored something for this hashId
              // (This prevents component keys like HASHID_locality overwriting the base HASHID entry)
              if (!fieldDefinitionsCache.rawNames[hashId]) {
                 fieldDefinitionsCache.rawNames[hashId] = field.name;
              }
            }
          }
        }
      }
    });
    
    Logger.log(`Processed ${totalFieldsProcessed} fields. Raw names cache size: ${Object.keys(fieldDefinitionsCache.rawNames).length}`);
    
    return fieldDefinitionsCache.rawNames; // Return the cache object
  } catch (e) {
    Logger.log(`Error in initializeCustomFieldsCache: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    // Ensure cache object exists even on error
    if (typeof fieldDefinitionsCache.rawNames === 'undefined') {
       fieldDefinitionsCache.rawNames = {};
    }
    return fieldDefinitionsCache.rawNames;
  }
}

/**
 * Gets a raw field name from the cache
 * @param {string} key - The field key
 * @return {string | null} The raw field name or null if not found
 */
function getRawFieldNameFromCache(key) {
  // Handle potential case differences if keys are sometimes inconsistent
  if (!fieldDefinitionsCache.rawNames) return null;
  return fieldDefinitionsCache.rawNames[key] || fieldDefinitionsCache.rawNames[key.toLowerCase()] || null;
}

/**
 * Gets the formatted column name for a given key
 * @param {string} key - The column key to format
 * @return {string} The formatted column name
 */
function getFormattedColumnName(key) {
  // Formatting logic now happens in formatColumnName, cache should be initialized beforehand.
  return formatColumnName(key);
}

/**
 * Gets a user-friendly category name for an entity type.
 * @param {string} entityType - The entity type (e.g., ENTITY_TYPES.DEALS).
 * @return {string} The category name (e.g., "Deal Fields").
 */
function getEntityCategoryName(entityType) {
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      return 'Deal Fields';
    case ENTITY_TYPES.PERSONS:
      return 'Contact Fields';
    case ENTITY_TYPES.ORGANIZATIONS:
      return 'Organization Fields';
    case ENTITY_TYPES.ACTIVITIES:
      return 'Activity Fields';
    case ENTITY_TYPES.LEADS:
      return 'Lead Fields';
    case ENTITY_TYPES.PRODUCTS:
      return 'Product Fields';
    default:
      return 'Standard Fields';
  }
}
