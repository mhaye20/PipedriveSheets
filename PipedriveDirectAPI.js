/**
 * Direct API wrapper for Pipedrive
 * 
 * This module provides direct API calls to Pipedrive that don't rely on the npm client's URL handling.
 * Use these functions when the standard client has issues with URL construction or parameter handling.
 */

/**
 * Processes date and time fields in the data before sending to Pipedrive API
 * This is especially important for time range fields (_until pairs)
 * @param {Object} data - The data object to process
 * @param {Object} fieldDefinitions - Optional field definitions from Pipedrive
 * @return {Object} - The processed data object
 */
function processDateTimeFields(payload, rowData, fieldDefinitions, headerToFieldKeyMap) {
  try {
    if (rowData) {
    }
    const fieldKeys = Object.keys(payload);

    // Track which fields were processed as date/time
    const processedDateTimeFields = [];

    // Create reverse mapping from fieldKeys to headers
    const fieldKeyToHeader = {};
    if (headerToFieldKeyMap) {
      // Build reverse mapping
      for (const header in headerToFieldKeyMap) {
        if (headerToFieldKeyMap.hasOwnProperty(header)) {
          fieldKeyToHeader[headerToFieldKeyMap[header]] = header;
        }
      }
    }
    
    // First identify date/time range pairs directly from field keys (not headers)
    const timeRangePairs = {};
    
    // Look for end time fields in rowData and add to tracking
    const rowDataUntilFields = [];
    if (rowData) {
      for (const headerKey in rowData) {
        // Check if this header maps to a field key
        const fieldKey = headerToFieldKeyMap ? headerToFieldKeyMap[headerKey] : null;
        if (fieldKey) {
          if (fieldKey.endsWith('_until')) {
            rowDataUntilFields.push(fieldKey);
            // Track the base field key too
            const baseFieldKey = fieldKey.replace(/_until$/, "");
            
            // CRITICAL: Add the _until field to the payload immediately
            if (!payload[fieldKey] && rowData[headerKey]) {
              payload[fieldKey] = rowData[headerKey];
            }
          }
        }
        
        // Also check if the header itself indicates an end time field
        if (headerKey.toLowerCase().includes('end time') && rowData[headerKey]) {
          // Try to find the corresponding field key
          const correspondingFieldKey = headerToFieldKeyMap ? headerToFieldKeyMap[headerKey] : null;
          if (correspondingFieldKey && !payload[correspondingFieldKey]) {
            payload[correspondingFieldKey] = rowData[headerKey];
          }
        }
      }
    }
    
    // Look in field definitions if available
    if (fieldDefinitions) {
      for (const fieldKey in fieldDefinitions) {
        if (fieldKey.endsWith("_until")) {
          const baseFieldKey = fieldKey.replace(/_until$/, "");
          if (fieldDefinitions[baseFieldKey]) {
            timeRangePairs[baseFieldKey] = fieldKey;
          } else {
          }
        }
      }
    } else {
    }
    
    // Also look for time range pairs in the payload itself - more aggressive detection
    for (const key in payload) {
      if (key.endsWith("_until")) {
        const baseKey = key.replace(/_until$/, "");
        // Accept even if base key doesn't exist yet - it might come from rowData or custom_fields
        timeRangePairs[baseKey] = key;
      }
    }

    // Also check in custom_fields object for time range pairs
    if (payload.custom_fields) {
      const customFieldKeys = Object.keys(payload.custom_fields);
      
      // First handle time range pairs - more aggressive detection
      for (const key of customFieldKeys) {
        if (key.endsWith("_until")) {
          const baseKey = key.replace(/_until$/, "");
          // Don't require the base key to exist
          timeRangePairs[baseKey] = key;
          
          // Ensure the time range end field exists at the root level
          payload[key] = payload.custom_fields[key];
        }
      }
      
      // Also check for any start fields that might have _until counterparts
      for (const key of customFieldKeys) {
        if (!key.endsWith("_until")) {
          const untilKey = `${key}_until`;
          // If we have this field in our data or in fieldDefinitions, add it to time range pairs
          if (payload.custom_fields[untilKey] || (fieldDefinitions && fieldDefinitions[untilKey])) {
            timeRangePairs[key] = untilKey;
          }
        }
      }

      // Check if we have any time range fields in rowData that need to be added to the payload
      for (const untilFieldKey of rowDataUntilFields) {
        const baseKey = untilFieldKey.replace(/_until$/, "");
        const headerKey = fieldKeyToHeader[untilFieldKey];
        
        if (headerKey && rowData[headerKey]) {
          if (!timeRangePairs[baseKey]) {
            timeRangePairs[baseKey] = untilFieldKey;
          }
        }
      }

      // Promote all custom fields to root level WITHOUT deleting custom_fields
      // Pipedrive API requires these values to be present in both places
      for (const key of customFieldKeys) {
        payload[key] = payload.custom_fields[key];
      }
      
      // DO NOT delete custom_fields - it's needed for proper API handling
      // delete payload.custom_fields; 

      // Continue with time range pair processing...
      for (const baseKey in timeRangePairs) {
        const untilKey = timeRangePairs[baseKey];
        
        // Get the start time value from any available source
        let startValue = null;
        if (payload[baseKey] !== undefined) {
          startValue = payload[baseKey];
        } else if (payload.custom_fields && payload.custom_fields[baseKey] !== undefined) {
          startValue = payload.custom_fields[baseKey];
        } else if (rowData && fieldKeyToHeader && fieldKeyToHeader[baseKey] && rowData[fieldKeyToHeader[baseKey]]) {
          startValue = rowData[fieldKeyToHeader[baseKey]];
        }
        
        // Get the end time value from any available source
        let endValue = null;
        if (payload[untilKey] !== undefined) {
          endValue = payload[untilKey];
        } else if (payload.custom_fields && payload.custom_fields[untilKey] !== undefined) {
          endValue = payload.custom_fields[untilKey];
        } else if (rowData && fieldKeyToHeader && fieldKeyToHeader[untilKey] && rowData[fieldKeyToHeader[untilKey]]) {
          endValue = rowData[fieldKeyToHeader[untilKey]];
        }
        
        // Format the values if they exist
        // Special case: "1899-12-30" is Excel/Sheets' way of storing time-only values, so treat it as time
        const isExcelTime = (startValue && String(startValue).includes('1899-12-30')) ||
                           (endValue && String(endValue).includes('1899-12-30'));
        
        // Check if this is a date range or time range based on the values
        const isDateRange = !isExcelTime && (
          (startValue && (String(startValue).includes('T') || String(startValue).match(/^\d{4}-\d{2}-\d{2}/))) ||
          (endValue && (String(endValue).includes('T') || String(endValue).match(/^\d{4}-\d{2}-\d{2}/)))
        );
        
        let formattedStartValue, formattedEndValue;
        
        if (isDateRange) {
          // Format as dates for date range fields
          formattedStartValue = startValue ? formatDateValue(startValue) : null;
          formattedEndValue = endValue ? formatDateValue(endValue) : null;
        } else {
          // Format as times for time range fields  
          formattedStartValue = startValue ? formatTimeValue(startValue) : null;
          formattedEndValue = endValue ? formatTimeValue(endValue) : null;
        }
        
        
        // Ensure we have both parts of the time range
        // If only one is present, use it for both (required by Pipedrive)
        if (formattedStartValue && !formattedEndValue) {
          formattedEndValue = formattedStartValue;
        } else if (!formattedStartValue && formattedEndValue) {
          formattedStartValue = formattedEndValue;
        }
        
        // Only proceed if we have at least one value
        if (formattedStartValue || formattedEndValue) {
          // Ensure custom_fields exists
          if (!payload.custom_fields) {
            payload.custom_fields = {};
          }
          
          // Set values at both root and custom_fields level
          if (formattedStartValue) {
            payload[baseKey] = formattedStartValue;
            payload.custom_fields[baseKey] = formattedStartValue;
          }
          
          if (formattedEndValue) {
            payload[untilKey] = formattedEndValue;
            payload.custom_fields[untilKey] = formattedEndValue;
          }
        } else {
        }
      }
    }

    // Log all time range pairs that were identified
    if (Object.keys(timeRangePairs).length > 0) {
      payload.__hasTimeRangeFields = true;
      
      // For each time range pair that we identified, ensure both parts are added to custom_fields
      // This is a fallback in case the normal processing missed them
      for (const baseKey in timeRangePairs) {
        const untilKey = timeRangePairs[baseKey];
        
        // Make sure we have custom_fields
        if (!payload.custom_fields) payload.custom_fields = {};
        
        // If we have start value but not end value in custom_fields
        if (payload.custom_fields[baseKey] && !payload.custom_fields[untilKey]) {
          // Check if the end time exists in root
          if (payload[untilKey]) {
            // Copy from root to custom_fields
            payload.custom_fields[untilKey] = payload[untilKey];
          } else {
            // Use start time as end time
            payload.custom_fields[untilKey] = payload.custom_fields[baseKey];
            payload[untilKey] = payload.custom_fields[baseKey];
          }
        }
        
        // If we have end value but not start value in custom_fields
        if (!payload.custom_fields[baseKey] && payload.custom_fields[untilKey]) {
          // Check if the start time exists in root
          if (payload[baseKey]) {
            // Copy from root to custom_fields
            payload.custom_fields[baseKey] = payload[baseKey];
          } else {
            // Use end time as start time
            payload.custom_fields[baseKey] = payload.custom_fields[untilKey];
            payload[baseKey] = payload.custom_fields[untilKey];
          }
        }
      }
    }

    // Now ensure both parts of any time range are included in the payload
    for (const baseKey in timeRangePairs) {
      const untilKey = timeRangePairs[baseKey];

      // Find the headers that map to these field keys
      const baseHeader = fieldKeyToHeader[baseKey];
      const untilHeader = fieldKeyToHeader[untilKey];


      // Check for values in both payload and rowData
      let startValue = payload[baseKey];
      let endValue = payload[untilKey];

      // If values not in payload, try to get from rowData
      if (!startValue && baseHeader && rowData[baseHeader]) {
        startValue = rowData[baseHeader];
      }

      if (!endValue && untilHeader && rowData[untilHeader]) {
        endValue = rowData[untilHeader];
      }

      // Update both values in payload and custom_fields
      if (startValue || endValue) {
        if (!payload.custom_fields) payload.custom_fields = {};

        if (startValue) {
          // Format time values properly
          const formattedStartValue = formatTimeValue(startValue);
          payload[baseKey] = formattedStartValue;
          payload.custom_fields[baseKey] = formattedStartValue;
        }

        if (endValue) {
          // Format time values properly
          const formattedEndValue = formatTimeValue(endValue);
          payload[untilKey] = formattedEndValue;
          payload.custom_fields[untilKey] = formattedEndValue;
        }

        // Always set both values, even if one is missing - this ensures proper time range handling
        if (startValue && !endValue) {
          // If we have start but no end, use the same value for end (required by Pipedrive API)
          const formattedValue = formatTimeValue(startValue);
          payload[untilKey] = formattedValue;
          payload.custom_fields[untilKey] = formattedValue;
        }
        // Similarly, if we have an end but no start, log a warning but do not auto-copy
        // as this might cause unexpected behavior
        if (!startValue && endValue) {
        }
        
        // Flag payload as having time range fields
        payload.__hasTimeRangeFields = true;
      }
    }

    return payload;
  } catch (error) {
    return payload;
  }
}

/**
 * Updates a Pipedrive deal using a single API call that includes all fields
 * @param {number|string} dealId - Deal ID to update
 * @param {Object} payload - Data to update in the deal
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @param {Object} fieldDefinitions - Field definitions from Pipedrive
 * @returns {Object} API response
 */
function updateDealDirect(dealId, payload, accessToken, basePath, fieldDefinitions) {
  try {
    // Ensure deal ID is a number
    dealId = Number(dealId);
    
    // Enhanced logging for debugging time range issues
    if (payload.custom_fields) {
    }
    
    // Look for time range _until fields and ensure they are at root level
    const timeRangePairs = {};
    
    // Extract time range pairs from both the payload and custom fields
    // More aggressive detection algorithm that finds all potential time range fields
    
    // First scan for all fields ending with _until in both root and custom_fields
    let untilFields = [];
    for (const key in payload) {
      if (key.endsWith("_until")) {
        untilFields.push(key);
      }
    }
    
    if (payload.custom_fields) {
      for (const key in payload.custom_fields) {
        if (key.endsWith("_until") && !untilFields.includes(key)) {
          untilFields.push(key);
        }
      }
    }
    
    // Log found _until fields
    
    // Process each _until field to find its matching start field
    for (const untilKey of untilFields) {
      const baseKey = untilKey.replace(/_until$/, "");
      
      // Check all possible places for values
      const startValueRoot = payload[baseKey];
      const startValueCustom = payload.custom_fields && payload.custom_fields[baseKey];
      const endValueRoot = payload[untilKey];
      const endValueCustom = payload.custom_fields && payload.custom_fields[untilKey];
      
      // Always create a time range pair regardless of whether we find values
      // This ensures we don't miss any potential time range fields
      timeRangePairs[baseKey] = {
        startKey: baseKey,
        startValue: startValueRoot !== undefined ? startValueRoot : startValueCustom,
        endKey: untilKey,
        endValue: endValueRoot !== undefined ? endValueRoot : endValueCustom
      };
      
    }
    
    // Look for other potential time fields that don't follow the _until convention
    // (This helps with older Pipedrive fields that might use different patterns)
    for (const key in payload) {
      // Check for fields that look like time fields but don't have a _until pair yet
      if (!key.endsWith("_until") && 
          !key.endsWith("_start") && 
          !Object.keys(timeRangePairs).includes(key) &&
          (key.includes("time") || key.includes("hour"))) {
          
        // Look for possible matching end time field with common patterns
        const possibleEndKeys = [
          `${key}_until`, 
          `${key}_end`, 
          `${key}_to`
        ];
        
        for (const endKey of possibleEndKeys) {
          if (payload[endKey] !== undefined || (payload.custom_fields && payload.custom_fields[endKey])) {
            // Found a match with a non-standard pattern
            timeRangePairs[key] = {
              startKey: key,
              startValue: payload[key] !== undefined ? payload[key] : (payload.custom_fields && payload.custom_fields[key]),
              endKey: endKey,
              endValue: payload[endKey] !== undefined ? payload[endKey] : (payload.custom_fields && payload.custom_fields[endKey])
            };
            break;
          }
        }
      }
    }
    
    // Create a flag to indicate we're preserving time range pairs
    payload.__preserveTimeRangePairs = true;
    payload.__timeRangePairs = timeRangePairs;
    
    // Remove empty custom_fields to avoid API errors
    if (payload.custom_fields && Object.keys(payload.custom_fields).length === 0) {
      delete payload.custom_fields;
    }
    
    // Create final payload with explicit time range handling
    const processedPayload = JSON.parse(JSON.stringify(payload));
    
    // Ensure time range pairs are preserved at ROOT LEVEL - this is critical
    Object.keys(timeRangePairs).forEach(baseKey => {
      const pair = timeRangePairs[baseKey];
      
      // Format both values consistently
      // Special case: "1899-12-30" is Excel/Sheets' way of storing time-only values, so treat it as time
      const isExcelTime = (pair.startValue && String(pair.startValue).includes('1899-12-30')) ||
                         (pair.endValue && String(pair.endValue).includes('1899-12-30'));
      
      // Check if this is a date range or time range based on the field ID pattern or value
      const isDateRange = !isExcelTime && (
        (pair.startValue && (String(pair.startValue).includes('T') || String(pair.startValue).match(/^\d{4}-\d{2}-\d{2}/))) ||
        (pair.endValue && (String(pair.endValue).includes('T') || String(pair.endValue).match(/^\d{4}-\d{2}-\d{2}/)))
      );
      
      let formattedStartValue, formattedEndValue;
      
      if (isDateRange) {
        // Format as dates for date range fields
        formattedStartValue = pair.startValue !== undefined ? formatDateValue(pair.startValue) : null;
        formattedEndValue = pair.endValue !== undefined ? formatDateValue(pair.endValue) : null;
      } else {
        // Format as times for time range fields
        formattedStartValue = pair.startValue !== undefined ? formatTimeValue(pair.startValue) : null;
        formattedEndValue = pair.endValue !== undefined ? formatTimeValue(pair.endValue) : null;
      }
      
      
      // Auto-fill missing end time with start time if needed
      let effectiveEndValue = formattedEndValue;
      if (formattedStartValue && !formattedEndValue) {
        effectiveEndValue = formattedStartValue;
      }
      
      // Auto-fill missing start time with end time if needed
      let effectiveStartValue = formattedStartValue;
      if (!formattedStartValue && formattedEndValue) {
        effectiveStartValue = formattedEndValue;
      }
      
      // ALWAYS set both values, even if one or both were null initially
      // This ensures the time range pair is properly handled by Pipedrive
      // But only use default if both values are truly missing
      if (!effectiveStartValue && !effectiveEndValue) {
        // Both are missing, use a default
        const defaultValue = "00:00:00";
        processedPayload[pair.startKey] = defaultValue;
        processedPayload[pair.endKey] = defaultValue;
      } else {
        // At least one value exists, use it
        processedPayload[pair.startKey] = effectiveStartValue || effectiveEndValue;
        processedPayload[pair.endKey] = effectiveEndValue || effectiveStartValue;
      }
      
      
      // Ensure custom_fields exists
      if (!processedPayload.custom_fields) {
        processedPayload.custom_fields = {};
      }
      
      // Set both values in custom_fields (use same values as root)
      processedPayload.custom_fields[pair.startKey] = processedPayload[pair.startKey];
      processedPayload.custom_fields[pair.endKey] = processedPayload[pair.endKey];
      
    });

    // Create URL for the request
    const dealUrl = `${basePath}/deals/${dealId}`;
    
    // Log the complete final payload to verify time range fields are included

    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(processedPayload)
    };
    
    // Make the API request
    
    // Additional check for address fields
    if (processedPayload.custom_fields && 
        Object.keys(processedPayload.custom_fields).some(key => 
          typeof processedPayload.custom_fields[key] === 'object' && 
          processedPayload.custom_fields[key] !== null)) {
    }
    
    // Fetch data from Pipedrive API
    const response = UrlFetchApp.fetch(dealUrl, options);
    const responseCode = response.getResponseCode();
    
    // Parse the response
    const responseData = JSON.parse(response.getContentText());
    
    return responseData;
  } catch (error) {
    throw error;
  }
}

/**
 * Format a date/time value based on its type
 * @param {*} value - The value to format
 * @param {string} fieldType - The field type (date, time, daterange)
 * @return {string} - The formatted value
 */
function formatDateTimeValue(value, fieldType) {
  if (value === null || value === undefined) {
    return value;
  }

  if (fieldType === 'time' || (typeof value === 'object' && value instanceof Date && value.getFullYear() === 1899)) {
    return formatTimeValue(value);
  } else if (fieldType === 'date' || (typeof value === 'string' && value.includes('-'))) {
    return formatDateValue(value);
  } else if (value instanceof Date) {
    // If it's a Date object but not specifically a time field, use ISO format
    return value.toISOString();
  }

  // If we can't determine the format, return as is
  return value;
}

/**
 * Format a time value to the HH:MM:SS format expected by Pipedrive
 * @param {*} value - The time value
 * @return {string} - The formatted time string
 */
function formatTimeValue(value) {
  try {
    if (!value && value !== 0 && value !== "0") return null;

    // CRITICAL: Handle objects that might have a 'value' property
    if (typeof value === 'object' && value !== null && !(value instanceof Date)) {
      if (value.value !== undefined) {
        value = value.value;
      } else {
        return null;
      }
    }

    
    // Quick check if it's already a properly formatted time string
    if (typeof value === 'string' && value.match(/^\d{2}:\d{2}:\d{2}$/)) {
      return value;
    }
    
    let timeObj;
    if (value instanceof Date) {
      // For Date objects, check if it's an Excel time-only date (1899-12-30)
      if (value.getFullYear() === 1899 && value.getMonth() === 11 && value.getDate() === 30) {
        // This is a time-only value, extract time components directly
        const hours = value.getHours();
        const minutes = value.getMinutes();
        const seconds = value.getSeconds();
        const formatted = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
        return formatted;
      }
      // For other Date objects, extract time normally
      timeObj = value;
    } else if (typeof value === 'string') {
      // Try to parse the string as a time
      if (value.match(/^\d{1,2}:\d{2}(:\d{2})?$/)) {
        // Already in proper format, ensure it has seconds
        const parts = value.split(':');
        if (parts.length === 2) {
          const formatted = `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:00`;
          return formatted;
        } else if (parts.length === 3) {
          // Already has seconds, just ensure proper padding
          const formatted = `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:${parts[2].padStart(2, '0')}`;
          return formatted;
        }
        return value;
      }

      // Special handling for "1899-12-30" date format (Excel/Sheets time-only format)
      if (value.includes("1899-12-30")) {
        // This format indicates a time-only value stored as a date
        // Extract time part directly from the ISO string to avoid timezone issues
        const timePart = value.split("T")[1];
        if (timePart) {
          // Extract hours, minutes, seconds from the UTC time string
          const timeMatch = timePart.match(/(\d{2}):(\d{2}):(\d{2})/);
          if (timeMatch) {
            const formatted = `${timeMatch[1]}:${timeMatch[2]}:${timeMatch[3]}`;
            return formatted;
          }
        }
        // Fallback to Date parsing if string extraction fails
        const datePart = new Date(value);
        if (!isNaN(datePart.getTime())) {
          timeObj = datePart;
        }
      }
      // Try parsing as a date+time
      else if (value.includes("T")) {
        // Handle ISO format strings
        const datePart = new Date(value);
        if (!isNaN(datePart.getTime())) {
          timeObj = datePart;
        } else {
          // Try extracting time part directly
          const timePart = value.split("T")[1];
          if (timePart && timePart.includes(":")) {
            const timeComponents = timePart.split(":");
            if (timeComponents.length >= 2) {
              const formatted = `${timeComponents[0].padStart(2, '0')}:${timeComponents[1].padStart(2, '0')}:00`;
              return formatted;
            }
          }
        }
      } else {
        // AM/PM format detection
        const amPmMatch = value.match(/(\d{1,2}):(\d{2})\s*(am|pm)/i);
        if (amPmMatch) {
          let hours = parseInt(amPmMatch[1], 10);
          const minutes = amPmMatch[2];
          const ampm = amPmMatch[3].toLowerCase();
          
          if (ampm === 'pm' && hours < 12) hours += 12;
          if (ampm === 'am' && hours === 12) hours = 0;
          
          const formatted = `${String(hours).padStart(2, '0')}:${minutes}:00`;
          return formatted;
        }
        
        // Try as regular time string with simpler regex (more permissive)
        const simpleTimeMatch = value.match(/(\d{1,2})[:\.](\d{2})/);
        if (simpleTimeMatch) {
          const hours = parseInt(simpleTimeMatch[1], 10);
          const minutes = parseInt(simpleTimeMatch[2], 10);
          if (hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
            const formatted = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:00`;
            return formatted;
          }
        }
        
        // Try as regular time string
        timeObj = new Date(value);
      }
    } else if (typeof value === 'number') {
      // Handle numeric time values (could be Excel time values)
      // If it's a small decimal (Excel time format), convert to hours/minutes
      if (value < 1) {
        const totalHours = value * 24;
        const hours = Math.floor(totalHours);
        const minutes = Math.floor((totalHours - hours) * 60);
        const seconds = Math.floor(((totalHours - hours) * 60 - minutes) * 60);
        
        const formatted = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
        return formatted;
      }
      
      // Otherwise try to create a date from timestamp
      timeObj = new Date(value);
    } else {
      return String(value);
    }

    if (!timeObj || isNaN(timeObj.getTime())) {
      // If we couldn't parse it as a date, try to extract time directly
      if (typeof value === 'string') {
        // Look for HH:MM pattern
        const timeMatch = value.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) {
          const hours = timeMatch[1].padStart(2, '0');
          const minutes = timeMatch[2].padStart(2, '0');
          const formatted = `${hours}:${minutes}:00`;
          return formatted;
        }
      }
      return String(value);
    }

    // Format as HH:MM:SS
    const formatted = `${String(timeObj.getHours()).padStart(2, '0')}:${String(timeObj.getMinutes()).padStart(2, '0')}:${String(timeObj.getSeconds()).padStart(2, '0')}`;
    return formatted;
  } catch (error) {
    return String(value);
  }
}

/**
 * Format a date value to the YYYY-MM-DD format expected by Pipedrive
 * @param {*} value - The date value
 * @return {string} - The formatted date string
 */
function formatDateValue(value) {
  try {
    if (!value) return null;
    

    let dateObj;
    if (value instanceof Date) {
      dateObj = value;
    } else if (typeof value === 'string') {
      // Check if already in YYYY-MM-DD format
      if (value.match(/^\d{4}-\d{2}-\d{2}$/)) {
        return value;
      }
      
      // Handle ISO date strings with time component
      if (value.includes('T')) {
        // Extract just the date part
        const datePart = value.split('T')[0];
        if (datePart.match(/^\d{4}-\d{2}-\d{2}$/)) {
          return datePart;
        }
      }
      
      // Handle time-only values that might be mistaken for dates
      if (value.match(/^\d{1,2}:\d{2}:\d{2}$/)) {
        // Return null or today's date, depending on requirements
        return null;
      }

      // Try parsing as a date
      dateObj = new Date(value);
    } else {
      return String(value);
    }

    if (isNaN(dateObj.getTime())) {
      return String(value);
    }

    // Format as YYYY-MM-DD
    const formatted = `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
    return formatted;
  } catch (error) {
    return String(value);
  }
}

/**
 * Updates a Pipedrive person using direct UrlFetchApp.fetch
 * @param {number|string} personId - Person ID to update
 * @param {Object} payload - Data to update in the person
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @param {Object} fieldDefinitions - Field definitions from Pipedrive
 * @returns {Object} API response
 */
function updatePersonDirect(personId, payload, accessToken, basePath, fieldDefinitions) {
  try {
    // Ensure person ID is a number
    personId = Number(personId);

    // Enhanced logging for debugging
    if (payload.custom_fields) {
    }
    
    // Process date/time fields similar to other entities
    let processedPayload = processDateTimeFields(payload, null, fieldDefinitions, null);
    
    // IMPORTANT: For persons in API v1, custom fields must be at root level
    // If we have custom_fields object, we need to flatten it to root level
    if (processedPayload.custom_fields) {
      const customFields = processedPayload.custom_fields;
      
      // Create new payload with custom fields at root level
      const flattenedPayload = {...processedPayload};
      delete flattenedPayload.custom_fields; // Remove the custom_fields object
      
      // Add each custom field to root level
      for (const key in customFields) {
        flattenedPayload[key] = customFields[key];
      }
      
      processedPayload = flattenedPayload;
    }

    // Create URL for the request
    const personUrl = `${basePath}/persons/${personId}`;
    
    // Log the complete final payload

    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(processedPayload)
    };

    // Make the request directly with UrlFetchApp
    const response = UrlFetchApp.fetch(personUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    // Parse response as JSON
    let responseData;
    try {
      responseData = JSON.parse(responseText);
    } catch (parseError) {
      responseData = {
        success: false,
        error: 'Error parsing response',
        error_info: parseError.message,
        raw_response: responseText.substring(0, 1000)
      };
    }

    return responseData;
  } catch (error) {
    return {
      success: false,
      error: error.message,
      error_info: 'Exception in updatePersonDirect method'
    };
  }
}

/**
 * Updates a Pipedrive organization using direct UrlFetchApp.fetch
 * @param {number|string} organizationId - Organization ID to update
 * @param {Object} payload - Data to update in the organization
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @param {Object} fieldDefinitions - Field definitions from Pipedrive
 * @returns {Object} API response
 */
function updateOrganizationDirect(organizationId, payload, accessToken, basePath, fieldDefinitions) {
  try {
    // Ensure organization ID is a number
    organizationId = Number(organizationId);
    
    // Enhanced logging for debugging
    if (payload.custom_fields) {
    }
    
    // For organizations, we need to handle the payload differently
    // Organizations in v1 API expect custom fields at root level
    const finalPayload = {};
    
    // First, copy any standard fields (non-custom fields)
    for (const key in payload) {
      if (key !== 'custom_fields' && !key.startsWith('__')) {
        finalPayload[key] = payload[key];
      }
    }
    
    // If there are custom fields, add them at root level
    if (payload.custom_fields) {
      
      // Process time range fields first
      const timeRangePairs = {};
      
      // Identify time range pairs
      for (const key in payload.custom_fields) {
        if (key.endsWith('_until')) {
          const baseKey = key.replace(/_until$/, '');
          timeRangePairs[baseKey] = key;
        }
      }
      
      // Add all custom fields to root level
      for (const key in payload.custom_fields) {
        const value = payload.custom_fields[key];
        
        // Check if this is part of a time range
        const isTimeRangeEnd = key.endsWith('_until');
        const isTimeRangeStart = timeRangePairs[key] !== undefined;
        
        if (isTimeRangeStart || isTimeRangeEnd) {
          // Format time/date values
          if (value) {
            // Check if it's a date or time based on the value
            if (String(value).match(/^\d{4}-\d{2}-\d{2}$/) || 
                (String(value).includes('T') && !String(value).includes('1899-12-30'))) {
              // It's a date
              finalPayload[key] = formatDateValue(value);
            } else {
              // It's a time
              finalPayload[key] = formatTimeValue(value);
            }
          } else {
            finalPayload[key] = value;
          }
        } else {
          // Regular custom field
          finalPayload[key] = value;
        }
      }
      
      // Ensure time range pairs are complete
      for (const baseKey in timeRangePairs) {
        const untilKey = timeRangePairs[baseKey];
        
        // If we have one but not the other, copy the value
        if (finalPayload[baseKey] && !finalPayload[untilKey]) {
          finalPayload[untilKey] = finalPayload[baseKey];
        } else if (!finalPayload[baseKey] && finalPayload[untilKey]) {
          finalPayload[baseKey] = finalPayload[untilKey];
        }
      }
    }
    
    // Remove any internal flags that processDateTimeFields might have added
    delete finalPayload.__hasTimeRangeFields;
    delete finalPayload.__preserveTimeRangePairs;
    delete finalPayload.__timeRangePairs;
    
    let processedPayload = finalPayload;
    
    // Create URL for the request
    const organizationUrl = `${basePath}/organizations/${organizationId}`;
    
    // Log the complete final payload

    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(processedPayload)
    };
    
    // Make the API request
    
    // Fetch data from Pipedrive API
    const response = UrlFetchApp.fetch(organizationUrl, options);
    const responseCode = response.getResponseCode();
    
    // Parse the response
    const responseData = JSON.parse(response.getContentText());
    
    return responseData;
  } catch (error) {
    return {
      success: false,
      error: error.message,
      error_info: 'Exception in updateOrganizationDirect method'
    };
  }
}

/**
 * Updates a Pipedrive product using direct UrlFetchApp.fetch
 * @param {number|string} productId - Product ID to update
 * @param {Object} payload - Data to update in the product
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @param {Object} fieldDefinitions - Field definitions from Pipedrive
 * @returns {Object} API response
 */
function updateProductDirect(productId, payload, accessToken, basePath, fieldDefinitions) {
  try {
    // Ensure product ID is a number
    productId = Number(productId);
    
    // Enhanced logging for debugging
    if (payload.custom_fields) {
    }
    
    // For products, we need to handle the payload differently
    // Products in v1 API expect custom fields at root level
    const finalPayload = {};
    
    // First, copy any standard fields (non-custom fields) with type conversion
    for (const key in payload) {
      if (key !== 'custom_fields' && !key.startsWith('__')) {
        // Handle product-specific field conversions
        if (key === 'unit') {
          // unit must be a string or null
          if (payload[key] !== null && payload[key] !== undefined && payload[key] !== '') {
            finalPayload[key] = String(payload[key]);
          } else {
            finalPayload[key] = null;
          }
        } else if (key === 'category') {
          // category must be a number or null (it's a category ID)
          if (payload[key] !== null && payload[key] !== undefined && payload[key] !== '') {
            // If it's a string that looks like a number, convert it
            const numValue = Number(payload[key]);
            if (!isNaN(numValue)) {
              finalPayload[key] = numValue;
            } else {
              // If it's not a number, skip it (it's probably a category name, not ID)
            }
          } else {
            finalPayload[key] = null;
          }
        } else if (key === 'owner_id') {
          // owner_id must be a number (user ID)
          if (payload[key] !== null && payload[key] !== undefined && payload[key] !== '') {
            const numValue = Number(payload[key]);
            if (!isNaN(numValue)) {
              finalPayload[key] = numValue;
            } else {
              // If it's a string like "Mike", skip it - we can't convert names to IDs here
            }
          }
        } else if (key === 'prices') {
          // prices must be an array of price objects
          if (Array.isArray(payload[key])) {
            finalPayload[key] = payload[key];
          } else if (payload[key] !== null && payload[key] !== undefined) {
            // Convert single price to array format
            // Assuming the value is the price amount
            finalPayload[key] = [{
              price: Number(payload[key]) || 0,
              currency: 'USD' // Default currency, should be configurable
            }];
          }
        } else {
          finalPayload[key] = payload[key];
        }
      }
    }
    
    // If there are custom fields, add them at root level
    if (payload.custom_fields) {
      
      // Process time range fields first
      const timeRangePairs = {};
      
      // Identify time range pairs
      for (const key in payload.custom_fields) {
        if (key.endsWith('_until')) {
          const baseKey = key.replace(/_until$/, '');
          timeRangePairs[baseKey] = key;
        }
      }
      
      // Add all custom fields to root level
      for (const key in payload.custom_fields) {
        const value = payload.custom_fields[key];
        
        // Check if this is part of a time range
        const isTimeRangeEnd = key.endsWith('_until');
        const isTimeRangeStart = timeRangePairs[key] !== undefined;
        
        if (isTimeRangeStart || isTimeRangeEnd) {
          // Format time/date values
          if (value) {
            // Check if it's a date or time based on the value
            if (String(value).match(/^\d{4}-\d{2}-\d{2}$/) || 
                (String(value).includes('T') && !String(value).includes('1899-12-30'))) {
              // It's a date
              finalPayload[key] = formatDateValue(value);
            } else {
              // It's a time
              finalPayload[key] = formatTimeValue(value);
            }
          } else {
            finalPayload[key] = value;
          }
        } else {
          // Regular custom field
          finalPayload[key] = value;
        }
      }
      
      // Ensure time range pairs are complete
      for (const baseKey in timeRangePairs) {
        const untilKey = timeRangePairs[baseKey];
        
        // If we have one but not the other, copy the value
        if (finalPayload[baseKey] && !finalPayload[untilKey]) {
          finalPayload[untilKey] = finalPayload[baseKey];
        } else if (!finalPayload[baseKey] && finalPayload[untilKey]) {
          finalPayload[baseKey] = finalPayload[untilKey];
        }
      }
    }
    
    // Remove any internal flags that processDateTimeFields might have added
    delete finalPayload.__hasTimeRangeFields;
    delete finalPayload.__preserveTimeRangePairs;
    delete finalPayload.__timeRangePairs;
    
    let processedPayload = finalPayload;
    
    // Create URL for the request
    const productUrl = `${basePath}/products/${productId}`;
    
    // Log the complete final payload

    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(processedPayload)
    };
    
    // Make the API request
    
    // Fetch data from Pipedrive API
    const response = UrlFetchApp.fetch(productUrl, options);
    const responseCode = response.getResponseCode();
    
    // Parse the response
    const responseData = JSON.parse(response.getContentText());
    
    return responseData;
  } catch (error) {
    return {
      success: false,
      error: error.message,
      error_info: 'Exception in updateProductDirect method'
    };
  }
}

// Add other direct API methods as needed for different entity types
