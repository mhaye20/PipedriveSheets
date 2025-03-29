/**
 * Pipedrive API Integration
 * 
 * This module handles all direct interactions with the Pipedrive API:
 * - Fetching data (deals, persons, organizations, etc.)
 * - Pushing updates back to Pipedrive
 * - Retrieving filter and field definitions
 */

/**
 * Constructs the base Pipedrive API URL based on the configured subdomain
 * @return {string} The base Pipedrive API URL
 */
function getPipedriveApiUrl() {
  const subdomain = PropertiesService.getDocumentProperties().getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  return PIPEDRIVE_API_URL_PREFIX + subdomain + PIPEDRIVE_API_URL_SUFFIX;
}

/**
 * Gets all available deal fields from Pipedrive
 * @return {Array} Array of deal field definitions
 */
function getDealFields() {
  return getPipedriverFields(ENTITY_TYPES.DEALS);
}

/**
 * Gets all available person fields from Pipedrive
 * @return {Array} Array of person field definitions
 */
function getPersonFields() {
  return getPipedriverFields(ENTITY_TYPES.PERSONS);
}

/**
 * Gets all available organization fields from Pipedrive
 * @return {Array} Array of organization field definitions
 */
function getOrganizationFields() {
  return getPipedriverFields(ENTITY_TYPES.ORGANIZATIONS);
}

/**
 * Gets all available activity fields from Pipedrive
 * @return {Array} Array of activity field definitions
 */
function getActivityFields() {
  return getPipedriverFields(ENTITY_TYPES.ACTIVITIES);
}

/**
 * Gets all available lead fields from Pipedrive
 * @return {Array} Array of lead field definitions
 */
function getLeadFields() {
  return getPipedriverFields(ENTITY_TYPES.LEADS);
}

/**
 * Gets all available product fields from Pipedrive
 * @return {Array} Array of product field definitions
 */
function getProductFields() {
  return getPipedriverFields(ENTITY_TYPES.PRODUCTS);
}

/**
 * Gets field definitions for a specific entity type
 * @param {string} entityType - The entity type to get fields for
 * @return {Array} Array of field definitions
 */
function getPipedriverFields(entityType) {
  try {
    // Get from cache if available
    const cacheKey = `fields_${entityType}`;
    if (fieldDefinitionsCache[cacheKey]) {
      return fieldDefinitionsCache[cacheKey];
    }
    
    const apiKey = PropertiesService.getDocumentProperties().getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    if (!apiKey) {
      throw new Error('Pipedrive API key not configured. Please set up the API key in the settings.');
    }
    
    const baseUrl = getPipedriveApiUrl();
    let url = `${baseUrl}/${entityType}/fields?api_token=${apiKey}`;
    
    // Special case for activities which has a different endpoint structure
    if (entityType === ENTITY_TYPES.ACTIVITIES) {
      url = `${baseUrl}/activityFields?api_token=${apiKey}`;
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      const fields = responseData.data || [];
      // Cache the fields
      fieldDefinitionsCache[cacheKey] = fields;
      return fields;
    } else {
      throw new Error(`Failed to get ${entityType} fields: ${responseData.error || 'Unknown error'}`);
    }
  } catch (e) {
    Logger.log(`Error in getPipedriverFields for ${entityType}: ${e.message}`);
    throw e;
  }
}

/**
 * Gets all filters from Pipedrive for the current user
 * @return {Array} Array of filter definitions
 */
function getPipedriveFilters() {
  try {
    const apiKey = PropertiesService.getDocumentProperties().getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    if (!apiKey) {
      throw new Error('Pipedrive API key not configured. Please set up the API key in the settings.');
    }
    
    const baseUrl = getPipedriveApiUrl();
    const url = `${baseUrl}/filters?api_token=${apiKey}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      // Enhance filters with their type in human-readable form
      const filters = responseData.data || [];
      filters.forEach(filter => {
        filter.typeFormatted = formatFilterType(filter.type);
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
    case 'persons':
      return 'Contacts';
    case 'organizations':
      return 'Organizations';
    case 'products':
      return 'Products';
    case 'activities':
      return 'Activities';
    case 'leads':
      return 'Leads';
    default:
      return type.charAt(0).toUpperCase() + type.slice(1);
  }
}

/**
 * Gets Pipedrive data using a specific filter
 * @param {string} entityType - Type of entity to fetch
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of records from Pipedrive
 */
function getFilteredDataFromPipedrive(entityType, filterId, limit = 100) {
  try {
    const apiKey = PropertiesService.getDocumentProperties().getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    if (!apiKey) {
      throw new Error('Pipedrive API key not configured. Please set up the API key in the settings.');
    }
    
    // Validate filter ID
    if (!filterId) {
      throw new Error(`No filter ID provided for ${entityType}. Please select a filter in the settings.`);
    }
    
    const baseUrl = getPipedriveApiUrl();
    const url = `${baseUrl}/${entityType}?api_token=${apiKey}&filter_id=${filterId}&limit=${limit}`;
    
    Logger.log(`Fetching data from ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      if (responseData.data && responseData.data.length > 0) {
        Logger.log(`Retrieved ${responseData.data.length} ${entityType} from Pipedrive`);
        return responseData.data;
      } else {
        Logger.log(`No ${entityType} found with the selected filter`);
        return [];
      }
    } else {
      throw new Error(`Failed to get ${entityType}: ${responseData.error || 'Unknown error'}`);
    }
  } catch (e) {
    Logger.log(`Error in getFilteredDataFromPipedrive for ${entityType}: ${e.message}`);
    throw e;
  }
}

/**
 * Gets deals using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of deals from Pipedrive
 */
function getDealsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.DEALS, filterId, limit);
}

/**
 * Gets persons using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of persons from Pipedrive
 */
function getPersonsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PERSONS, filterId, limit);
}

/**
 * Gets organizations using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of organizations from Pipedrive
 */
function getOrganizationsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ORGANIZATIONS, filterId, limit);
}

/**
 * Gets activities using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of activities from Pipedrive
 */
function getActivitiesWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ACTIVITIES, filterId, limit);
}

/**
 * Gets leads using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of leads from Pipedrive
 */
function getLeadsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.LEADS, filterId, limit);
}

/**
 * Gets products using a specific filter
 * @param {string} filterId - ID of the filter to use
 * @param {number} limit - Maximum number of records to fetch
 * @return {Array} Array of products from Pipedrive
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
} 