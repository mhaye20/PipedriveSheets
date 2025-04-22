/**
 * Direct API wrapper for Pipedrive
 * 
 * This module provides direct API calls to Pipedrive that don't rely on the npm client's URL handling.
 * Use these functions when the standard client has issues with URL construction or parameter handling.
 */

/**
 * Updates a Pipedrive deal using direct UrlFetchApp.fetch
 * @param {number|string} dealId - Deal ID to update
 * @param {Object} payload - Data to update in the deal
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @returns {Object} API response
 */
function updateDealDirect(dealId, payload, accessToken, basePath) {
  try {
    // Ensure deal ID is a number
    dealId = Number(dealId);
    
    // Create URL for the request
    const dealUrl = `${basePath}/deals/${dealId}`;
    Logger.log(`Direct API: Using URL: ${dealUrl}`);
    
    // Create a copy of the payload to prevent modifying the original
    const finalPayload = JSON.parse(JSON.stringify(payload));
    
    // Check for date fields in the payload root - for direct API calls,
    // date fields should use YYYY-MM-DD format
    for (const key in finalPayload) {
      const value = finalPayload[key];
      // Check if the value is an ISO date string with time part
      if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
        // If this looks like a date field, convert to YYYY-MM-DD format
        // This conversion might need to be conditional based on field definitions,
        // but for direct API we'll use a simpler approach
        if (key.includes('_date') || key.includes('date_') || key === 'expected_close_date') {
          Logger.log(`Direct API: Converting date format for field ${key}`);
          finalPayload[key] = value.split('T')[0];
        }
      }
    }
    
    // Also check in custom_fields object
    if (finalPayload.custom_fields) {
      for (const key in finalPayload.custom_fields) {
        const value = finalPayload.custom_fields[key];
        if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
          // Apply the same logic for custom fields
          if (key.includes('_date') || key.includes('date_')) {
            Logger.log(`Direct API: Converting date format for custom field ${key}`);
            finalPayload.custom_fields[key] = value.split('T')[0];
          }
        }
      }
    }
    
    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(finalPayload)
    };
    
    // Log the request for debugging
    Logger.log(`Direct API: Making PUT request to: ${dealUrl}`);
    Logger.log(`Direct API: With payload: ${JSON.stringify(finalPayload).substring(0, 500)}...`);
    
    // Make the request directly with UrlFetchApp
    const response = UrlFetchApp.fetch(dealUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Direct API: Response code: ${responseCode}`);
    
    // Parse response as JSON
    let responseData;
    try {
      responseData = JSON.parse(responseText);
      Logger.log(`Direct API: Response data: ${JSON.stringify(responseData).substring(0, 500)}...`);
    } catch (parseError) {
      Logger.log(`Direct API: Error parsing response: ${parseError.message}`);
      Logger.log(`Direct API: Raw response: ${responseText.substring(0, 500)}...`);
      responseData = { 
        success: false, 
        error: 'Error parsing response',
        error_info: parseError.message,
        raw_response: responseText.substring(0, 1000)
      };
    }
    
    return responseData;
  } catch (error) {
    Logger.log(`Direct API: Error in updateDealDirect: ${error.message}`);
    if (error.stack) {
      Logger.log(`Direct API: Error stack: ${error.stack}`);
    }
    
    return {
      success: false,
      error: error.message,
      error_info: 'Exception in updateDealDirect method'
    };
  }
}

/**
 * Updates a Pipedrive person using direct UrlFetchApp.fetch
 * @param {number|string} personId - Person ID to update
 * @param {Object} payload - Data to update in the person
 * @param {string} accessToken - OAuth access token
 * @param {string} basePath - API base path (e.g., https://mycompany.pipedrive.com/v1)
 * @returns {Object} API response
 */
function updatePersonDirect(personId, payload, accessToken, basePath) {
  try {
    // Ensure person ID is a number
    personId = Number(personId);
    
    // Create URL for the request
    const personUrl = `${basePath}/persons/${personId}`;
    Logger.log(`Direct API: Using URL: ${personUrl}`);
    
    // Create fetch options
    const options = {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
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

// Add other direct API methods as needed for different entity types 