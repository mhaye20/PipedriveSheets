/**
 * Pipedrive API Communication
 * 
 * This module handles all direct communication with the Pipedrive API:
 * - API requests and response handling
 * - Data retrieval for different entity types
 * - Field definitions and metadata
 */

var PipedriveAPI = PipedriveAPI || {};

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
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  
  if (!accessToken) {
    throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
  }
  
  // Set up request options with proper headers
  const requestOptions = {
    method: options.method || 'get',
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Accept': 'application/json',
      ...(options.headers || {})
    },
    muteHttpExceptions: true,
    ...options
  };
  
  try {
    Logger.log(`Making authenticated request to: ${url}`);
    const response = UrlFetchApp.fetch(url, requestOptions);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Log response details for debugging
    Logger.log(`Response status code: ${statusCode}`);
    
    // Handle different status codes
    if (statusCode === 401) {
      Logger.log('Received 401 Unauthorized, attempting to refresh token...');
      // Try to refresh the token
      if (refreshAccessTokenIfNeeded()) {
        // Get the new token
        const newToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
        if (!newToken) {
          throw new Error('Failed to refresh access token.');
        }
        
        // Update the authorization header with the new token
        requestOptions.headers['Authorization'] = 'Bearer ' + newToken;
        
        // Retry the request with the new token
        const retryResponse = UrlFetchApp.fetch(url, requestOptions);
        const retryStatusCode = retryResponse.getResponseCode();
        const retryResponseText = retryResponse.getContentText();
        
        if (retryStatusCode === 401) {
          // If we still get 401 after refresh, clear tokens and force re-auth
          scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
          scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
          scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
          throw new Error('Authentication failed. Please reconnect to Pipedrive.');
        }
        
        // Try to parse the retry response
        try {
          const retryData = JSON.parse(retryResponseText);
          if (retryStatusCode >= 200 && retryStatusCode < 300 && retryData.success) {
            return retryData;
          } else {
            throw new Error(retryData.error || `API request failed with status ${retryStatusCode}`);
          }
        } catch (parseError) {
          throw new Error(`Invalid response from Pipedrive API: ${retryResponseText.substring(0, 100)}...`);
        }
      } else {
        // Token refresh failed, clear tokens and force re-auth
        scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
        throw new Error('Authentication failed. Please reconnect to Pipedrive.');
      }
    }
    
    // Try to parse the response as JSON
    try {
      const responseData = JSON.parse(responseText);
      
      // Check if the request was successful
      if (statusCode >= 200 && statusCode < 300 && responseData.success) {
        return responseData;
      } else {
        // Handle error response
        const errorMessage = responseData.error || `API request failed with status ${statusCode}`;
        Logger.log(`Pipedrive API error: ${errorMessage}`);
        throw new Error(errorMessage);
      }
    } catch (parseError) {
      Logger.log(`Error parsing response as JSON: ${parseError.message}`);
      Logger.log(`Response text: ${responseText}`);
      
      // If we got HTML instead of JSON, it's likely an authentication issue
      if (responseText.includes('<!DOCTYPE html>')) {
        throw new Error('Authentication error. Please reconnect to Pipedrive.');
      }
      
      throw new Error(`Invalid response from Pipedrive API: ${responseText.substring(0, 100)}...`);
    }
  } catch (error) {
    Logger.log(`Error in makeAuthenticatedRequest: ${error.message}`);
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
 * Gets products using a specific filter
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
}

/**
 * Searches for persons by name and returns a mapping of names to person IDs
 * @param {Array<string>} names - Array of person names to search for
 * @return {Object} Mapping of person names to person IDs
 */
function searchPersonsByName(names) {
  try {
    const nameToIdMap = {};
    
    // Process each name
    for (const name of names) {
      if (!name || typeof name !== 'string') continue;
      
      const searchTerm = name.trim();
      if (!searchTerm) continue;
      
      // Search for the person using Pipedrive search API
      const searchUrl = `${getPipedriveApiUrl()}/persons/search?term=${encodeURIComponent(searchTerm)}&limit=10`;
      
      try {
        const response = makeAuthenticatedRequest(searchUrl);
        
        if (response.success && response.data && response.data.items) {
          // Look for exact match first
          let found = false;
          for (const item of response.data.items) {
            if (item.item && item.item.name && item.item.name.toLowerCase() === searchTerm.toLowerCase()) {
              nameToIdMap[name] = item.item.id;
              found = true;
              Logger.log(`Found exact match for "${name}": Person ID ${item.item.id}`);
              break;
            }
          }
          
          // If no exact match, use the first result
          if (!found && response.data.items.length > 0 && response.data.items[0].item) {
            nameToIdMap[name] = response.data.items[0].item.id;
            Logger.log(`Using first match for "${name}": Person ID ${response.data.items[0].item.id} (${response.data.items[0].item.name})`);
          } else if (!found) {
            Logger.log(`No person found for name: ${name}`);
          }
        }
      } catch (searchError) {
        Logger.log(`Error searching for person "${name}": ${searchError.message}`);
      }
    }
    
    return nameToIdMap;
  } catch (error) {
    Logger.log(`Error in searchPersonsByName: ${error.message}`);
    return {};
  }
}

/**
 * Builds comprehensive entity mappings for ID to name conversions
 * @param {Array} items - Array of items from Pipedrive
 * @param {string} entityType - The entity type being processed
 * @return {Object} Object containing all entity mappings
 */
function buildEntityMappings(items, entityType) {
  try {
    Logger.log(`Building entity mappings for ${entityType}...`);
    
    const mappings = {
      personIdToName: {},
      orgIdToName: {},
      dealIdToTitle: {},
      userIdToName: {},
      stageIdToName: {},
      pipelineIdToName: {}
    };
    
    // Collect IDs and names from the current data
    const personIds = new Set();
    const orgIds = new Set();
    const dealIds = new Set();
    const userIds = new Set();
    const stageIds = new Set();
    const pipelineIds = new Set();
    
    for (const item of items) {
      // Person mappings
      if (item.person_id && item.person_name) {
        mappings.personIdToName[item.person_id] = item.person_name;
      }
      if (item.person_id) personIds.add(item.person_id);
      
      // Organization mappings
      if (item.org_id && item.org_name) {
        mappings.orgIdToName[item.org_id] = item.org_name;
      }
      if (item.org_id) orgIds.add(item.org_id);
      
      // Deal mappings
      if (item.deal_id && item.deal_title) {
        mappings.dealIdToTitle[item.deal_id] = item.deal_title;
      }
      if (item.deal_id) dealIds.add(item.deal_id);
      
      // User mappings (owner, creator, etc.)
      if (item.owner_id && item.owner_name) {
        mappings.userIdToName[item.owner_id] = item.owner_name;
      }
      if (item.creator_user_id && item.creator_name) {
        mappings.userIdToName[item.creator_user_id] = item.creator_name;
      }
      if (item.user_id && item.user_name) {
        mappings.userIdToName[item.user_id] = item.user_name;
      }
      
      // Collect user IDs
      if (item.owner_id) userIds.add(item.owner_id);
      if (item.creator_user_id) userIds.add(item.creator_user_id);
      if (item.user_id) userIds.add(item.user_id);
      
      // Stage and pipeline mappings
      if (item.stage_id && item.stage_name) {
        mappings.stageIdToName[item.stage_id] = item.stage_name;
      }
      if (item.pipeline_id && item.pipeline_name) {
        mappings.pipelineIdToName[item.pipeline_id] = item.pipeline_name;
      }
      if (item.stage_id) stageIds.add(item.stage_id);
      if (item.pipeline_id) pipelineIds.add(item.pipeline_id);
      
      // Handle participants for activities
      if (entityType === 'activities' && item.participants && Array.isArray(item.participants)) {
        for (const participant of item.participants) {
          if (participant.person_id) {
            personIds.add(participant.person_id);
          }
        }
      }
    }
    
    Logger.log(`Pre-mapped: ${Object.keys(mappings.personIdToName).length} persons, ${Object.keys(mappings.orgIdToName).length} orgs, ${Object.keys(mappings.dealIdToTitle).length} deals, ${Object.keys(mappings.userIdToName).length} users`);
    
    // Fetch missing entity names in batches
    const batchFetchPromises = [];
    
    // Fetch missing person names
    const missingPersonIds = Array.from(personIds).filter(id => !mappings.personIdToName[id]);
    if (missingPersonIds.length > 0) {
      batchFetchPromises.push(batchFetchEntities('persons', missingPersonIds, mappings.personIdToName, 'name'));
    }
    
    // Fetch missing organization names
    const missingOrgIds = Array.from(orgIds).filter(id => !mappings.orgIdToName[id]);
    if (missingOrgIds.length > 0) {
      batchFetchPromises.push(batchFetchEntities('organizations', missingOrgIds, mappings.orgIdToName, 'name'));
    }
    
    // Fetch missing deal titles
    const missingDealIds = Array.from(dealIds).filter(id => !mappings.dealIdToTitle[id]);
    if (missingDealIds.length > 0) {
      batchFetchPromises.push(batchFetchEntities('deals', missingDealIds, mappings.dealIdToTitle, 'title'));
    }
    
    // Fetch missing user names
    const missingUserIds = Array.from(userIds).filter(id => !mappings.userIdToName[id]);
    if (missingUserIds.length > 0) {
      batchFetchPromises.push(batchFetchEntities('users', missingUserIds, mappings.userIdToName, 'name'));
    }
    
    Logger.log(`Final mappings: ${Object.keys(mappings.personIdToName).length} persons, ${Object.keys(mappings.orgIdToName).length} orgs, ${Object.keys(mappings.dealIdToTitle).length} deals, ${Object.keys(mappings.userIdToName).length} users`);
    
    return mappings;
  } catch (error) {
    Logger.log(`Error building entity mappings: ${error.message}`);
    return {
      personIdToName: {},
      orgIdToName: {},
      dealIdToTitle: {},
      userIdToName: {},
      stageIdToName: {},
      pipelineIdToName: {}
    };
  }
}

/**
 * Batch fetch entities to build ID to name mappings
 * @param {string} entityType - Type of entity (persons, organizations, deals, users)
 * @param {Array} ids - Array of IDs to fetch
 * @param {Object} mapping - Mapping object to populate
 * @param {string} nameField - Field name containing the display name
 */
function batchFetchEntities(entityType, ids, mapping, nameField) {
  try {
    Logger.log(`Fetching ${ids.length} ${entityType} names...`);
    
    const batchSize = 100;
    for (let i = 0; i < ids.length; i += batchSize) {
      const batch = ids.slice(i, i + batchSize);
      try {
        const url = `${getPipedriveApiUrl()}/${entityType}?ids=${batch.join(',')}&limit=${batchSize}`;
        const response = makeAuthenticatedRequest(url);
        
        if (response.success && response.data) {
          for (const entity of response.data) {
            if (entity.id && entity[nameField]) {
              mapping[entity.id] = entity[nameField];
            }
          }
        }
      } catch (e) {
        Logger.log(`Error fetching ${entityType} names for batch: ${e.message}`);
      }
    }
  } catch (error) {
    Logger.log(`Error in batchFetchEntities: ${error.message}`);
  }
}

/**
 * Searches for entities by name and returns a mapping of names to IDs
 * @param {string} entityType - Type of entity (persons, organizations, deals)
 * @param {Array<string>} names - Array of entity names to search for
 * @return {Object} Mapping of entity names to IDs
 */
function searchEntitiesByName(entityType, names) {
  try {
    const nameToIdMap = {};
    
    // Process each name
    for (const name of names) {
      if (!name || typeof name !== 'string') continue;
      
      const searchTerm = name.trim();
      if (!searchTerm) continue;
      
      // Search for the entity using Pipedrive search API
      const searchUrl = `${getPipedriveApiUrl()}/${entityType}/search?term=${encodeURIComponent(searchTerm)}&limit=10`;
      
      try {
        const response = makeAuthenticatedRequest(searchUrl);
        
        if (response.success && response.data && response.data.items) {
          // Look for exact match first
          let found = false;
          for (const item of response.data.items) {
            const itemName = item.item.name || item.item.title;
            if (item.item && itemName && itemName.toLowerCase() === searchTerm.toLowerCase()) {
              nameToIdMap[name] = item.item.id;
              found = true;
              Logger.log(`Found exact match for "${name}": ${entityType} ID ${item.item.id}`);
              break;
            }
          }
          
          // If no exact match, use the first result
          if (!found && response.data.items.length > 0 && response.data.items[0].item) {
            nameToIdMap[name] = response.data.items[0].item.id;
            const itemName = response.data.items[0].item.name || response.data.items[0].item.title;
            Logger.log(`Using first match for "${name}": ${entityType} ID ${response.data.items[0].item.id} (${itemName})`);
          } else if (!found) {
            Logger.log(`No ${entityType} found for name: ${name}`);
          }
        }
      } catch (searchError) {
        Logger.log(`Error searching for ${entityType} "${name}": ${searchError.message}`);
      }
    }
    
    return nameToIdMap;
  } catch (error) {
    Logger.log(`Error in searchEntitiesByName: ${error.message}`);
    return {};
  }
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
 * Gets field definitions for deals
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getDealFields(forceRefresh = false) {
  return getEntityFields('dealFields', forceRefresh);
}

/**
 * Gets field definitions for persons
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getPersonFields(forceRefresh = false) {
  return getEntityFields('personFields', forceRefresh);
}

/**
 * Gets field definitions for organizations
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getOrganizationFields(forceRefresh = false) {
  return getEntityFields('organizationFields', forceRefresh);
}

/**
 * Gets field definitions for activities
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getActivityFields(forceRefresh = false) {
  return getEntityFields('activityFields', forceRefresh);
}

/**
 * Gets field definitions for leads
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getLeadFields(forceRefresh = false) {
  // Leads don't have their own custom fields - they inherit from deals
  // So we need to get deal fields and return them for leads
  Logger.log('Getting lead fields - using deal fields as leads inherit from deals');
  return getEntityFields('dealFields', forceRefresh);
}

/**
 * Gets all lead labels
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of lead label objects
 */
function getLeadLabels(forceRefresh = false) {
  const cacheKey = 'LEAD_LABELS_CACHE';
  const cache = CacheService.getScriptCache();
  
  // Try to get from cache first
  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      try {
        const labels = JSON.parse(cachedData);
        Logger.log(`Retrieved ${labels.length} lead labels from cache`);
        return labels;
      } catch (e) {
        Logger.log('Failed to parse cached lead labels');
      }
    }
  }
  
  try {
    Logger.log('Fetching lead labels from Pipedrive API');
    const response = makePipedriveRequest('leadLabels');
    
    if (response && response.data) {
      // Cache for 1 hour
      cache.put(cacheKey, JSON.stringify(response.data), 3600);
      Logger.log(`Retrieved ${response.data.length} lead labels from API`);
      return response.data;
    } else {
      Logger.log('No lead labels data in response');
      return [];
    }
  } catch (e) {
    Logger.log(`Error fetching lead labels: ${e.message}`);
    return [];
  }
}

/**
 * Gets field definitions for products
 * @param {boolean} forceRefresh - Whether to force a refresh from the API
 * @return {Array} Array of field definition objects
 */
function getProductFields(forceRefresh = false) {
  return getEntityFields('productFields', forceRefresh);
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
 * Generic function to get field definitions for an entity type
 * @param {string} fieldEndpoint - API endpoint for field definitions
 * @param {boolean} forceRefresh - Whether to force a refresh from API instead of using cache
 * @return {Array} Array of field definition objects
 */
function getEntityFields(fieldEndpoint, forceRefresh = false) {
  try {
    // Check if we have cached results and should use them
    const cacheKey = `CACHE_${fieldEndpoint}`;
    const scriptProperties = PropertiesService.getScriptProperties();
    const cachedDataJson = scriptProperties.getProperty(cacheKey);
    
    // Use cache if available and not forcing refresh
    if (!forceRefresh && cachedDataJson) {
      try {
        const cachedData = JSON.parse(cachedDataJson);
        // Check if cache is still valid (less than 1 hour old)
        const cacheTimeKey = `CACHE_TIME_${fieldEndpoint}`;
        const cacheTimeStr = scriptProperties.getProperty(cacheTimeKey);
        
        if (cacheTimeStr) {
          const cacheTime = parseInt(cacheTimeStr, 10);
          const currentTime = new Date().getTime();
          const oneHour = 60 * 60 * 1000; // 1 hour in milliseconds
          
          if (currentTime - cacheTime < oneHour && cachedData && cachedData.length > 0) {
            Logger.log(`Using cached ${fieldEndpoint} data (${cachedData.length} fields)`);
            return cachedData;
          }
        }
      } catch (cacheError) {
        Logger.log(`Error parsing cached ${fieldEndpoint} data: ${cacheError.message}`);
        // Cache error is non-fatal, continue to fetch new data
      }
    }
    
    // Make the request to get the fields
    Logger.log(`Fetching ${fieldEndpoint} from Pipedrive API`);
    const response = makePipedriveRequest(fieldEndpoint);
    
    if (response.success && response.data) {
      // If the data exists but is empty, that's a problem
      if (!response.data.length) {
        Logger.log(`Warning: ${fieldEndpoint} returned empty data array. This may indicate an API issue.`);
      }
      
      // Cache the response for future use
      scriptProperties.setProperty(cacheKey, JSON.stringify(response.data));
      scriptProperties.setProperty(`CACHE_TIME_${fieldEndpoint}`, new Date().getTime().toString());
      
      return response.data;
    }
    
    Logger.log(`API response for ${fieldEndpoint} was unsuccessful or missing data`);
    
    // If we have cached data, use it as a fallback
    if (cachedDataJson) {
      try {
        const cachedData = JSON.parse(cachedDataJson);
        Logger.log(`Using cached ${fieldEndpoint} data as fallback`);
        return cachedData;
      } catch (e) {
        Logger.log(`Error parsing cached ${fieldEndpoint} fallback data: ${e.message}`);
      }
    }
    
    // Return empty array as last resort
    return [];
  } catch (e) {
    Logger.log(`Error getting field definitions for ${fieldEndpoint}: ${e.message}`);
    
    // Check if we have cached data to use as fallback
    try {
      const cacheKey = `CACHE_${fieldEndpoint}`;
      const scriptProperties = PropertiesService.getScriptProperties();
      const cachedDataJson = scriptProperties.getProperty(cacheKey);
      
      if (cachedDataJson) {
        const cachedData = JSON.parse(cachedDataJson);
        Logger.log(`Using cached ${fieldEndpoint} data as fallback after error`);
        return cachedData;
      }
    } catch (cacheError) {
      Logger.log(`Error accessing cache after API failure: ${cacheError.message}`);
    }
    
    // If all else fails, return empty array
    return [];
  }
}

/**
 * Gets all available filters from Pipedrive
 * @return {Array} Array of filter objects
 */
function getPipedriveFilters() {
  try {
    Logger.log('Getting filters from Pipedrive');
    
    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    if (!accessToken) {
      throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
    }
    
    // Use v1 API endpoint for filters
    const url = `https://${subdomain}.pipedrive.com/v1/filters`;
    
    // Make authenticated request
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Filter API response code: ${statusCode}`);
    
    // Handle different status codes
    if (statusCode === 401) {
      Logger.log('Received 401 Unauthorized, attempting to refresh token...');
      // Try to refresh the token
      if (refreshAccessTokenIfNeeded()) {
        // Get the new token
        const newToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
        if (!newToken) {
          throw new Error('Failed to refresh access token.');
        }
        
        // Retry the request with the new token
        const retryResponse = UrlFetchApp.fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + newToken,
            'Accept': 'application/json'
          },
          muteHttpExceptions: true
        });
        
        const retryStatusCode = retryResponse.getResponseCode();
        const retryResponseText = retryResponse.getContentText();
        
        if (retryStatusCode === 401) {
          throw new Error('Authentication failed. Please reconnect to Pipedrive.');
        }
        
        // Parse retry response
        const retryData = JSON.parse(retryResponseText);
        if (retryStatusCode >= 200 && retryStatusCode < 300 && retryData.success) {
          // Enhance filters with their type in human-readable form
          const filters = retryData.data || [];
          Logger.log(`Retrieved ${filters.length} filters from Pipedrive`);
          
          filters.forEach(filter => {
            filter.typeFormatted = formatFilterType(filter.type);
            filter.normalizedType = normalizeFilterType(filter.type);
            Logger.log(`Filter: ${filter.name}, Type: ${filter.type}, Normalized: ${filter.normalizedType}`);
          });
          
          return filters;
        }
      }
      
      throw new Error('Authentication failed. Please reconnect to Pipedrive.');
    }
    
    // Parse the response
    try {
      const responseData = JSON.parse(responseText);
      
      if (!responseData.success) {
        throw new Error(responseData.error || 'Unknown error');
      }
      
      // Enhance filters with their type in human-readable form
      const filters = responseData.data || [];
      Logger.log(`Retrieved ${filters.length} filters from Pipedrive`);
      
      filters.forEach(filter => {
        filter.typeFormatted = formatFilterType(filter.type);
        filter.normalizedType = normalizeFilterType(filter.type);
        Logger.log(`Filter: ${filter.name}, Type: ${filter.type}, Normalized: ${filter.normalizedType}`);
      });
      
      return filters;
    } catch (parseError) {
      Logger.log(`Error parsing response as JSON: ${parseError.message}`);
      Logger.log(`Response text: ${responseText}`);
      
      if (responseText.includes('<!DOCTYPE html>')) {
        throw new Error('Authentication error. Please reconnect to Pipedrive.');
      }
      
      throw new Error(`Invalid response from Pipedrive API: ${responseText.substring(0, 100)}...`);
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
  if (!type) return '';
  
  // Convert to lowercase for consistent comparison
  const lowerType = type.toLowerCase();
  
  switch (lowerType) {
    case 'deals':
      return ENTITY_TYPES.DEALS;
    case 'person':
    case 'people':
    case 'persons':
      return ENTITY_TYPES.PERSONS;
    case 'org':
    case 'organization':
    case 'organizations':
      return ENTITY_TYPES.ORGANIZATIONS;
    case 'product':
    case 'products':
      return ENTITY_TYPES.PRODUCTS;
    case 'activity':
    case 'activities':
      return ENTITY_TYPES.ACTIVITIES;
    case 'lead':
    case 'leads':
      return ENTITY_TYPES.LEADS;
    default:
      Logger.log(`Unknown filter type: ${type}`);
      return lowerType;
  }
}

/**
 * Gets filters for a specific entity type
 * @param {string} entityType - The entity type to filter for (e.g., 'deals', 'persons')
 * @return {Array} Array of filter objects matching the entity type
 */
function getFiltersForEntityType(entityType) {
  try {
    Logger.log(`Getting filters for entity type: ${entityType}`);
    
    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    if (!accessToken) {
      throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
    }
    
    // Use v1 API endpoint for filters
    const url = `https://${subdomain}.pipedrive.com/v1/filters`;
    
    // Make authenticated request
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Filter API response code: ${statusCode}`);
    
    // Handle different status codes
    if (statusCode === 401) {
      Logger.log('Received 401 Unauthorized, attempting to refresh token...');
      // Try to refresh the token
      if (refreshAccessTokenIfNeeded()) {
        // Get the new token
        const newToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
        if (!newToken) {
          throw new Error('Failed to refresh access token.');
        }
        
        // Retry the request with the new token
        const retryResponse = UrlFetchApp.fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + newToken,
            'Accept': 'application/json'
          },
          muteHttpExceptions: true
        });
        
        const retryStatusCode = retryResponse.getResponseCode();
        const retryResponseText = retryResponse.getContentText();
        
        if (retryStatusCode === 401) {
          throw new Error('Authentication failed. Please reconnect to Pipedrive.');
        }
        
        // Parse retry response
        const retryData = JSON.parse(retryResponseText);
        if (retryStatusCode >= 200 && retryStatusCode < 300 && retryData.success) {
          const filters = retryData.data || [];
          Logger.log(`Retrieved ${filters.length} total filters`);
          
          // Filter based on normalized type matching the requested entity type
          const matchingFilters = filters.filter(filter => {
            const normalizedType = normalizeFilterType(filter.type);
            const isMatch = normalizedType === entityType;
            Logger.log(`Filter: ${filter.name}, Type: ${filter.type}, Normalized: ${normalizedType}, Matches ${entityType}: ${isMatch}`);
            return isMatch;
          });
          
          Logger.log(`Found ${matchingFilters.length} matching filters for ${entityType}`);
          return matchingFilters;
        }
      }
      
      throw new Error('Authentication failed. Please reconnect to Pipedrive.');
    }
    
    // Parse the response
    try {
      const responseData = JSON.parse(responseText);
      
      if (!responseData.success) {
        throw new Error(responseData.error || 'Unknown error');
      }
      
      const filters = responseData.data || [];
      Logger.log(`Retrieved ${filters.length} total filters`);
      
      // Filter based on normalized type matching the requested entity type
      const matchingFilters = filters.filter(filter => {
        const normalizedType = normalizeFilterType(filter.type);
        const isMatch = normalizedType === entityType;
        Logger.log(`Filter: ${filter.name}, Type: ${filter.type}, Normalized: ${normalizedType}, Matches ${entityType}: ${isMatch}`);
        return isMatch;
      });
      
      Logger.log(`Found ${matchingFilters.length} matching filters for ${entityType}`);
      return matchingFilters;
    } catch (parseError) {
      Logger.log(`Error parsing response as JSON: ${parseError.message}`);
      Logger.log(`Response text: ${responseText}`);
      
      if (responseText.includes('<!DOCTYPE html>')) {
        throw new Error('Authentication error. Please reconnect to Pipedrive.');
      }
      
      throw new Error(`Invalid response from Pipedrive API: ${responseText.substring(0, 100)}...`);
    }
  } catch (e) {
    Logger.log(`Error in getFiltersForEntityType: ${e.message}`);
    throw new Error(`Failed to get filters for ${entityType}: ${e.message}`);
  }
}

/**
 * Gets a map of field definitions for an entity type, keyed by field key (hash)
 * @param {string} entityType - The entity type (e.g., 'deals', 'persons')
 * @param {boolean} forceRefresh - Whether to force refresh from API
 * @return {Object} Map of field definitions { fieldKey: fieldDefinition }
 */
function getFieldDefinitionsMap(entityType, forceRefresh = false) {
  let fieldsArray = [];
  let endpoint = '';

  switch (entityType) {
    case ENTITY_TYPES.DEALS: endpoint = 'dealFields'; break;
    case ENTITY_TYPES.PERSONS: endpoint = 'personFields'; break;
    case ENTITY_TYPES.ORGANIZATIONS: endpoint = 'organizationFields'; break;
    case ENTITY_TYPES.ACTIVITIES: endpoint = 'activityFields'; break;
    case ENTITY_TYPES.LEADS: endpoint = 'dealFields'; break; // Leads inherit custom fields from deals
    case ENTITY_TYPES.PRODUCTS: endpoint = 'productFields'; break;
    default:
      Logger.log(`Unknown entity type for field definitions: ${entityType}`);
      return {};
  }
  
  try {
    fieldsArray = getEntityFields(endpoint, forceRefresh);
    const fieldMap = {};
    if (fieldsArray && fieldsArray.length > 0) {
      fieldsArray.forEach(field => {
        if (field && field.key) {
          fieldMap[field.key] = field;
        }
      });
    }
    Logger.log(`Created field definition map with ${Object.keys(fieldMap).length} entries for ${entityType}`);
    return fieldMap;
  } catch (e) {
    Logger.log(`Error creating field definition map for ${entityType}: ${e.message}`);
    return {}; // Return empty map on error
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
    
    // Special handling for leads - add label_ids mapping
    if (entityType === ENTITY_TYPES.LEADS) {
      try {
        const leadLabels = getLeadLabels();
        if (leadLabels && leadLabels.length > 0) {
          mappings['label_ids'] = {};
          leadLabels.forEach(label => {
            if (label.id && label.name) {
              mappings['label_ids'][label.id] = label.name;
            }
          });
          Logger.log(`Added ${leadLabels.length} lead label mappings`);
        }
      } catch (e) {
        Logger.log(`Error getting lead labels for mapping: ${e.message}`);
      }
    }
    
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

/**
 * Creates a mapping between custom field keys and their human-readable names
 * @param {string} entityType The entity type (deals, persons, etc.)
 * @return {Object} Mapping of custom field keys to field names
 */
function getCustomFieldMappings(entityType) {
  Logger.log(`Getting custom field mappings for ${entityType}`);

  let customFields = getCustomFieldsForEntity(entityType);
  const map = {};

  Logger.log(`Retrieved ${customFields.length} custom fields for ${entityType}`);

  customFields.forEach(field => {
    map[field.key] = field.name;
    // Log each mapping for debugging
    Logger.log(`Custom field mapping: ${field.key} => ${field.name}`);
  });
  Logger.log(`Total custom field mappings: ${Object.keys(map).length}`);

  return map;
}

/**
 * Gets custom fields for a specific entity type
 * @param {string} entityType The entity type (deals, persons, etc.)
 * @return {Array} Array of custom fields
 */
function getCustomFieldsForEntity(entityType) {
  try {
    // Ensure we have a valid token
    if (!refreshAccessTokenIfNeeded()) {
      Logger.log('Not authenticated with Pipedrive');
      return [];
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    if (!accessToken) {
      Logger.log('No OAuth access token found');
      return [];
    }
    
    let endpoint = '';
    
    // Different endpoints for different entity types
    switch(entityType) {
      case 'deals':
        endpoint = 'dealFields';
        break;
      case 'persons':
        endpoint = 'personFields';
        break;
      case 'organizations':
        endpoint = 'organizationFields';
        break;
      case 'activities':
        endpoint = 'activityFields';
        break;
      case 'leads':
        endpoint = 'dealFields'; // Leads inherit custom fields from deals
        break;
      case 'products':
        endpoint = 'productFields';
        break;
      default:
        Logger.log(`Unknown entity type: ${entityType}`);
        return [];
    }
    
    // Build the complete URL
    const baseUrl = getPipedriveApiUrl();
    const url = `${baseUrl}/${endpoint}`;
    
    // Make the authenticated request
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + accessToken
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`Error fetching custom fields: ${responseCode}`);
      return [];
    }
    
    const result = JSON.parse(response.getContentText());
    
    if (!result.success) {
      Logger.log(`API returned error: ${result.error || 'Unknown error'}`);
      return [];
    }
    
    // Extract only custom fields
    const customFields = result.data.filter(field => field.edit_flag);
    Logger.log(`Found ${customFields.length} custom fields for ${entityType}`);
    
    return customFields;
  } catch (e) {
    Logger.log(`Error in getCustomFieldsForEntity: ${e.message}`);
    return [];
  }
}

// Export API functions to be available through the PipedriveAPI namespace
Object.assign(PipedriveAPI, {
  getDealsWithFilter,
  getPersonsWithFilter,
  getOrganizationsWithFilter,
  getActivitiesWithFilter,
  getLeadsWithFilter,
  getProductsWithFilter,
  getDealFields,
  getPersonFields,
  getOrganizationFields,
  getActivityFields,
  getLeadFields,
  getProductFields,
  getFieldOptionMappingsForEntity,
  getFiltersForEntityType,
  getCustomFieldMappings,
  getCustomFieldsForEntity,
  getFieldDefinitionsMap // Export the new function
});
