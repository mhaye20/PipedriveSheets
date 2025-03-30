/**
 * Pipedrive API Communication
 * 
 * This module handles all direct communication with the Pipedrive API:
 * - API requests and response handling
 * - Data retrieval for different entity types
 * - Field definitions and metadata
 */

/**
 * Gets the Pipedrive API URL
 * @return {string} The complete API URL with subdomain
 */
function getPipedriveApiUrl() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  return PIPEDRIVE_API_URL_PREFIX + subdomain + PIPEDRIVE_API_URL_SUFFIX;
}

/**
 * Generic function to make authenticated API requests to Pipedrive
 */
function makeAuthenticatedRequest(url, options = {}) {
  // Ensure we have a valid token
  if (!refreshAccessTokenIfNeeded()) {
    throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  
  // Add authorization header
  if (!options.headers) {
    options.headers = {};
  }
  options.headers['Authorization'] = 'Bearer ' + accessToken;
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log('API request error: ' + error.message);
    throw error;
  }
}

/**
 * Gets deals using a specific filter
 */
function getDealsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.DEALS, filterId, limit);
}

/**
 * Gets persons using a specific filter
 */
function getPersonsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PERSONS, filterId, limit);
}

/**
 * Gets organizations using a specific filter
 */
function getOrganizationsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ORGANIZATIONS, filterId, limit);
}

/**
 * Gets activities using a specific filter
 */
function getActivitiesWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ACTIVITIES, filterId, limit);
}

/**
 * Gets leads using a specific filter
 */
function getLeadsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.LEADS, filterId, limit);
}

/**
 * Gets data from Pipedrive using a specific filter based on entity type
 */
function getFilteredDataFromPipedrive(entityType, filterId, limit = 100) {
  try {
    // If we have a filter ID, use it, otherwise get all data
    const filterParam = filterId ? `?filter_id=${filterId}` : '';
    const url = `${getPipedriveApiUrl()}/${entityType}${filterParam}`;
    
    let allItems = [];
    let hasMore = true;
    let start = 0;
    const pageLimit = 100; // Pipedrive API limit per page is 100
    
    // Handle pagination
    while (hasMore) {
      // Add pagination parameters
      let paginatedUrl = url;
      if (filterParam) {
        paginatedUrl += `&start=${start}&limit=${pageLimit}`;
      } else {
        paginatedUrl += `?start=${start}&limit=${pageLimit}`;
      }
      
      // Make the authenticated request
      const responseData = makeAuthenticatedRequest(paginatedUrl);
      
      if (responseData.success) {
        const items = responseData.data;
        if (items && items.length > 0) {
          allItems = allItems.concat(items);
          
          // If we have a specified limit (not 0), check if we've reached it
          if (limit > 0 && allItems.length >= limit) {
            allItems = allItems.slice(0, limit);
            hasMore = false;
          }
          // Check if there are more results
          else if (responseData.additional_data && 
              responseData.additional_data.pagination && 
              responseData.additional_data.pagination.more_items_in_collection) {
            hasMore = true;
            start += pageLimit;
            
            // Status update for large datasets
            if (allItems.length % 500 === 0) {
              SpreadsheetApp.getActiveSpreadsheet().toast(`Retrieved ${allItems.length} ${entityType} so far...`);
            }
          } else {
            hasMore = false;
          }
        } else {
          hasMore = false;
        }
      } else {
        Logger.log(`Failed to retrieve ${entityType}: ${responseData.error}`);
        hasMore = false;
      }
    }
    
    // Log completion status for large datasets
    if (allItems.length > 100) {
      Logger.log(`Retrieved ${allItems.length} ${entityType} from Pipedrive filter`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Retrieved ${allItems.length} ${entityType} from Pipedrive filter. Preparing data for the sheet...`);
    }
    
    return allItems;
  } catch (error) {
    Logger.log(`Error retrieving ${entityType}: ${error.message}`);
    return [];
  }
}

/**
 * Gets products with a specific filter
 * @param {string} filterId - Filter ID (optional)
 * @param {number} limit - Maximum number of items to retrieve (0 for all)
 * @return {Array} Array of product objects
 */
function getProductsWithFilter(filterId, limit = 0) {
  return getFilteredEntityData('products', filterId, limit);
}

/**
 * Generic function to get entity data with filtering
 * @param {string} entityType - Type of entity (deals, persons, etc.)
 * @param {string} filterId - Filter ID (optional)
 * @param {number} limit - Maximum number of items to retrieve (0 for all)
 * @return {Array} Array of entity objects
 */
function getFilteredEntityData(entityType, filterId, limit = 0) {
  try {
    // Build the endpoint with filter if provided
    let endpoint = entityType;
    if (filterId) {
      endpoint += `?filter_id=${filterId}`;
    }
    
    // Get all items with pagination
    const items = getAllItemsWithPagination(endpoint, limit);
    
    return items;
  } catch (e) {
    Logger.log(`Error getting ${entityType} data: ${e.message}`);
    throw new Error(`Failed to get ${entityType} data: ${e.message}`);
  }
}

/**
 * Gets all items from a paginated endpoint
 * @param {string} endpoint - The API endpoint
 * @param {number} limit - Maximum number of items to retrieve (0 for all)
 * @return {Array} All items combined from all pages
 */
function getAllItemsWithPagination(endpoint, limit = 0) {
  try {
    const allItems = [];
    let hasMore = true;
    let start = 0;
    const pageSize = 100; // Pipedrive API page size
    
    // Add toast message for user feedback
    SpreadsheetApp.getActiveSpreadsheet().toast('Starting data retrieval from Pipedrive...', 'Pipedrive Sync', 5);
    
    while (hasMore) {
      // Add pagination parameters
      let paginatedEndpoint = endpoint;
      if (endpoint.includes('?')) {
        paginatedEndpoint += `&start=${start}&limit=${pageSize}`;
      } else {
        paginatedEndpoint += `?start=${start}&limit=${pageSize}`;
      }
      
      // Make the request
      const response = makePipedriveRequest(paginatedEndpoint);
      
      if (response.success && response.data && response.data.length > 0) {
        // Add items to our collection
        allItems.push(...response.data);
        
        // Show progress for large data sets
        if (allItems.length % 500 === 0) {
          SpreadsheetApp.getActiveSpreadsheet().toast(`Retrieved ${allItems.length} items...`, 'Pipedrive Sync', 3);
        }
        
        // Check if we've hit the requested limit
        if (limit > 0 && allItems.length >= limit) {
          return allItems.slice(0, limit);
        }
        
        // Check if there are more pages
        if (response.additional_data && 
            response.additional_data.pagination && 
            response.additional_data.pagination.more_items_in_collection) {
          start += pageSize;
        } else {
          hasMore = false;
        }
      } else {
        // No more data or error
        hasMore = false;
      }
    }
    
    return allItems;
  } catch (e) {
    Logger.log(`Error in getAllItemsWithPagination: ${e.message}`);
    throw e;
  }
}

/**
 * Gets field definitions for deals
 * @return {Array} Array of field definition objects
 */
function getDealFields() {
  return getEntityFields('dealFields');
}

/**
 * Gets field definitions for persons
 * @return {Array} Array of field definition objects
 */
function getPersonFields() {
  return getEntityFields('personFields');
}

/**
 * Gets field definitions for organizations
 * @return {Array} Array of field definition objects
 */
function getOrganizationFields() {
  return getEntityFields('organizationFields');
}

/**
 * Gets field definitions for activities
 * @return {Array} Array of field definition objects
 */
function getActivityFields() {
  return getEntityFields('activityFields');
}

/**
 * Gets field definitions for leads
 * @return {Array} Array of field definition objects
 */
function getLeadFields() {
  return getEntityFields('leadFields');
}

/**
 * Gets field definitions for products
 * @return {Array} Array of field definition objects
 */
function getProductFields() {
  return getEntityFields('productFields');
}

/**
 * Generic function to get field definitions for an entity type
 * @param {string} fieldEndpoint - API endpoint for field definitions
 * @return {Array} Array of field definition objects
 */
function getEntityFields(fieldEndpoint) {
  try {
    const response = makePipedriveRequest(fieldEndpoint);
    
    if (response.success && response.data) {
      return response.data;
    }
    
    return [];
  } catch (e) {
    Logger.log(`Error getting field definitions: ${e.message}`);
    throw new Error(`Failed to get fields: ${e.message}`);
  }
}

/**
 * Gets all available filters from Pipedrive
 * @return {Array} Array of filter objects
 */
function getPipedriveFilters() {
  try {
    // Ensure we have a valid token
    if (!refreshAccessTokenIfNeeded()) {
      throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    if (!accessToken) {
      throw new Error('Pipedrive authentication not configured. Please connect to Pipedrive first.');
    }
    
    const baseUrl = getPipedriveApiUrl();
    const url = `${baseUrl}/filters`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      },
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      // Enhance filters with their type in human-readable form
      const filters = responseData.data || [];
      filters.forEach(filter => {
        filter.typeFormatted = formatFilterType(filter.type);
        
        // Add a normalized type property to handle inconsistencies
        filter.normalizedType = normalizeFilterType(filter.type);
        
        // Log filter details for debugging
        Logger.log(`Filter: ${filter.name}, Type: ${filter.type}, Normalized: ${filter.normalizedType}`);
      });
      
      return filters;
    } else {
      throw new Error(`Failed to get filters: ${responseData.error || 'Unknown error'}`);
    }
  } catch (e) {
    Logger.log(`Error in getPipedriveFilters: ${e.message}`);
    throw e;
  }
}

/**
 * Formats filter type to more human-readable form
 * @param {string} type - The filter type from Pipedrive
 * @return {string} Formatted filter type
 */
function formatFilterType(type) {
  switch (type) {
    case 'deals':
      return 'Deals';
    case 'person':
    case 'persons':
      return 'Contacts';
    case 'org':
    case 'organization':
    case 'organizations':
      return 'Organizations';
    case 'product':
    case 'products':
      return 'Products';
    case 'activity':
    case 'activities':
      return 'Activities';
    case 'lead':
    case 'leads':
      return 'Leads';
    default:
      return type.charAt(0).toUpperCase() + type.slice(1);
  }
}

/**
 * Normalizes filter type to match our entity type constants
 * @param {string} type - The filter type from Pipedrive
 * @return {string} Normalized filter type matching ENTITY_TYPES
 */
function normalizeFilterType(type) {
  switch (type) {
    case 'deals':
      return 'deals';
    case 'person': 
    case 'people':
      return 'persons';
    case 'org':
    case 'organization':
      return 'organizations';
    case 'product':
      return 'products';
    case 'activity':
      return 'activities';
    case 'lead':
      return 'leads';
    default:
      return type;
  }
}

/**
 * Gets field option mappings for a specific entity type
 * @param {string} entityType - The entity type
 * @return {Object} Mapping of field keys to option mappings
 */
function getFieldOptionMappingsForEntity(entityType) {
  try {
    // Get field definitions
    let fields = [];
    
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fields = getDealFields();
        break;
      case ENTITY_TYPES.PERSONS:
        fields = getPersonFields();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fields = getOrganizationFields();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        fields = getActivityFields();
        break;
      case ENTITY_TYPES.LEADS:
        fields = getLeadFields();
        break;
      case ENTITY_TYPES.PRODUCTS:
        fields = getProductFields();
        break;
      default:
        return {};
    }
    
    // Build the mappings
    const mappings = {};
    
    fields.forEach(field => {
      // Only process fields with options
      if (field.options && field.options.length > 0) {
        mappings[field.key] = {};
        
        // Map each option ID to its label
        field.options.forEach(option => {
          if (option.id !== undefined && option.label !== undefined) {
            mappings[field.key][option.id] = option.label;
          }
        });
      }
    });
    
    return mappings;
  } catch (e) {
    Logger.log(`Error getting field option mappings: ${e.message}`);
    return {};
  }
}

/**
 * Makes an authenticated request to the Pipedrive API
 * @param {string} endpoint - The API endpoint (without the base URL)
 * @param {Object} options - Request options
 * @return {Object} The parsed JSON response
 */
function makePipedriveRequest(endpoint, options = {}) {
  // Ensure we have a valid token
  if (!refreshAccessTokenIfNeeded()) {
    throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  
  // Build the complete URL
  const baseUrl = getPipedriveApiUrl();
  const url = baseUrl + '/' + endpoint;
  
  // Set up request options
  const requestOptions = Object.assign({
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + accessToken
    },
    muteHttpExceptions: true
  }, options);
  
  // Add authorization header if not already present
  if (!requestOptions.headers) {
    requestOptions.headers = {};
  }
  if (!requestOptions.headers['Authorization']) {
    requestOptions.headers['Authorization'] = 'Bearer ' + accessToken;
  }
  
  // Make the request
  try {
    const response = UrlFetchApp.fetch(url, requestOptions);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    // Check if the request was successful
    if (statusCode >= 200 && statusCode < 300 && responseData.success) {
      return responseData;
    } else {
      // If we got a 401, try to refresh the token and retry
      if (statusCode === 401) {
        Logger.log('Received 401 Unauthorized, attempting to refresh token and retry...');
        
        // Force token refresh
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
        if (refreshAccessTokenIfNeeded()) {
          // Retry the request with the new token
          return makePipedriveRequest(endpoint, options);
        }
      }
      
      // Handle error response
      const errorMessage = responseData.error || `API request failed with status code ${statusCode}`;
      Logger.log(`Pipedrive API error (${endpoint}): ${errorMessage}`);
      throw new Error(errorMessage);
    }
  } catch (error) {
    Logger.log(`Error in makePipedriveRequest: ${error.message}`);
    throw error;
  }
}