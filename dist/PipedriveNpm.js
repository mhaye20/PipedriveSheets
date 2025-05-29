/**
 * PipedriveNpm.js
 * Helper module for working with the Pipedrive npm package in Google™ Apps Script
 */

/**
 * Initialize Pipedrive API client with all necessary fixes for Google™ Apps Script
 * @param {string} apiToken - The OAuth token for Pipedrive
 * @param {string} basePath - The API base path (optional)
 * @returns {Object} API client with proper configuration
 */
function initializePipedriveClient(apiToken, basePath = null) {
  try {
    // Load npm package
    const pipedriveLib = getNpmPackage();
    
    // If basePath wasn't provided, try to get it from script properties
    if (!basePath) {
      const scriptProperties = PropertiesService.getScriptProperties();
      const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN');
      if (subdomain) {
        basePath = `https://${subdomain}.pipedrive.com/v1`;
      } else {
        basePath = 'https://api.pipedrive.com/v1';
      }
    }
    
    // Create configuration
    const { Configuration } = pipedriveLib.v1;
    const config = new Configuration({
      apiKey: apiToken,
      basePath: basePath
    });
    
    
    // Create a map to hold API clients
    const apiClients = {};
    
    // Setup polyfills from AppLib if available
    if (getNpmPackageHelpers() && getNpmPackageHelpers().setupPolyfills) {
      getNpmPackageHelpers().setupPolyfills();
    } else {
      setupPolyfills();
    }
    
    // Create API clients with fixed methods
    apiClients.deals = createFixedApiClient('DealsApi', pipedriveLib, config);
    apiClients.persons = createFixedApiClient('PersonsApi', pipedriveLib, config);
    apiClients.organizations = createFixedApiClient('OrganizationsApi', pipedriveLib, config);
    apiClients.activities = createFixedApiClient('ActivitiesApi', pipedriveLib, config);
    apiClients.leads = createFixedApiClient('LeadsApi', pipedriveLib, config);
    apiClients.products = createFixedApiClient('ProductsApi', pipedriveLib, config);
    
    return apiClients;
  } catch (error) {
    throw error;
  }
}

/**
 * Creates an API client with fixed methods for Google™ Apps Script
 * @param {string} apiName - The name of the API client class
 * @param {Object} pipedriveLib - The Pipedrive library object
 * @param {Object} config - The API configuration
 * @returns {Object} Fixed API client
 */
function createFixedApiClient(apiName, pipedriveLib, config) {
  try {
    // Create the API client
    const ApiClass = pipedriveLib.v1[apiName];
    if (!ApiClass) {
      throw new Error(`API class ${apiName} not found in Pipedrive library`);
    }
    
    const apiClient = new ApiClass(config);
    
    // Apply custom adapter to fix URL and payload issues
    if (apiClient.axios && apiClient.axios.defaults) {
      const customAdapter = createGASAdapter();
      apiClient.axios.defaults.adapter = customAdapter;
    }
    
    // Fix update method based on API type
    switch (apiName) {
      case 'DealsApi':
        fixUpdateMethod(apiClient, 'updateDeal', 'deals');
        break;
      case 'PersonsApi':
        fixUpdateMethod(apiClient, 'updatePerson', 'persons');
        break;
      case 'OrganizationsApi':
        fixUpdateMethod(apiClient, 'updateOrganization', 'organizations');
        break;
      case 'ActivitiesApi':
        fixUpdateMethod(apiClient, 'updateActivity', 'activities');
        break;
      case 'LeadsApi':
        fixUpdateMethod(apiClient, 'updateLead', 'leads');
        break;
      case 'ProductsApi':
        fixUpdateMethod(apiClient, 'updateProduct', 'products');
        break;
    }
    
    // Add additional helper methods
    apiClient.entityType = apiName.replace('Api', '').toLowerCase();
    
    return apiClient;
  } catch (error) {
    throw error;
  }
}

/**
 * Fix the update method of an API client to work with Google™ Apps Script
 * @param {Object} apiClient - The API client to fix
 * @param {string} methodName - The name of the update method
 * @param {string} entityPath - The API path for the entity type
 */
function fixUpdateMethod(apiClient, methodName, entityPath) {
  if (!apiClient[methodName]) {
    return;
  }
  
  const originalMethod = apiClient[methodName];
  
  apiClient[methodName] = async function(params) {
    try {
      // Validate parameters
      if (!params) {
        throw new Error(`No parameters provided to ${methodName}`);
      }
      
      // Make sure ID is a number
      if (params.id && typeof params.id === 'string') {
        params.id = Number(params.id);
      }
      
      if (!params.body) {
        throw new Error(`No body provided in ${methodName} parameters`);
      }
      
      
      // Clean up the body payload
      const cleanBody = sanitizePayload(params.body);
      
      // Create a clone of the parameters to avoid reference issues
      const updatedParams = {
        id: params.id,
        body: cleanBody
      };
      
      // Log the payload for debugging
      
      // Create explicit request options with fixed URL
      const requestOptions = {
        url: `${apiClient.basePath}/${entityPath}/${params.id}`
      };
      
      // Call the original method with our fixed parameters and options
      return await originalMethod.call(apiClient, updatedParams, requestOptions);
    } catch (error) {
      if (error.stack) {
      }
      throw error;
    }
  };
  
}

/**
 * Sanitize payload for Pipedrive API
 * @param {Object} payload - The payload to sanitize
 * @returns {Object} Sanitized payload
 */
function sanitizePayload(payload) {
  if (!payload || typeof payload !== 'object') {
    return payload;
  }
  
  // Create a deep clone to avoid modifying the original
  const cleaned = JSON.parse(JSON.stringify(payload));
  
  // Ensure custom_fields is properly formatted
  if (cleaned.custom_fields && typeof cleaned.custom_fields === 'object') {
    for (const key in cleaned.custom_fields) {
      const value = cleaned.custom_fields[key];
      
      // Handle null/undefined values
      if (value === null || value === undefined) {
        delete cleaned.custom_fields[key];
        continue;
      }
      
      // Handle address fields (objects with 'value' property)
      if (typeof value === 'object' && value !== null) {
        // Ensure all values in address objects are strings
        for (const prop in value) {
          if (value[prop] !== null && value[prop] !== undefined) {
            value[prop] = String(value[prop]);
          }
        }
      }
    }
  }
  
  return cleaned;
}

/**
 * Create a custom Google™ Apps Script adapter for axios
 * @returns {Function} Adapter function
 */
function createGASAdapter() {
  return function gasAdapter(config) {
    return new Promise((resolve, reject) => {
      try {
        
        // Convert axios config to UrlFetchApp options
        const options = {
          method: config.method.toUpperCase(),
          muteHttpExceptions: true,
          contentType: 'application/json',
          headers: config.headers || {}
        };
        
        // Fix authorization header
        if (config.headers && config.headers['Authorization']) {
          // Auth header already set, keep it
        } else if (config.headers && config.headers['api_token']) {
          // Using API token as query param (fallback)
        }
        
        // Fix URL format - handle '[object Object]' issue
        let finalUrl = config.url;
        if (finalUrl && finalUrl.includes('[object Object]')) {
          finalUrl = finalUrl.split('[object Object]')[0];
        }
        
        // Add query parameters if any
        if (config.params && typeof config.params === 'object') {
          const queryParams = [];
          for (const key in config.params) {
            if (config.params.hasOwnProperty(key) && config.params[key] != null) {
              queryParams.push(`${encodeURIComponent(key)}=${encodeURIComponent(config.params[key])}`);
            }
          }
          
          if (queryParams.length > 0) {
            finalUrl += (finalUrl.includes('?') ? '&' : '?') + queryParams.join('&');
          }
        }
        
        // Handle payload for non-GET requests
        if (config.method !== 'get' && config.data) {
          let payloadData = config.data;
          
          // Extract body from {id, body} format
          if (typeof payloadData === 'object' && payloadData.id !== undefined && payloadData.body !== undefined) {
            payloadData = payloadData.body;
          }
          
          // Always stringify objects for payload
          options.payload = typeof payloadData === 'object' 
            ? JSON.stringify(payloadData) 
            : String(payloadData);
            
        }
        
        
        // Make the request
        const response = UrlFetchApp.fetch(finalUrl, options);
        
        // Parse response
        let responseData;
        try {
          responseData = JSON.parse(response.getContentText());
        } catch (e) {
          responseData = response.getContentText();
        }
        
        // Create axios-like response
        const axiosResponse = {
          data: responseData,
          status: response.getResponseCode(),
          statusText: '',
          headers: response.getAllHeaders(),
          config: config
        };
        
        
        // Resolve the promise
        resolve(axiosResponse);
      } catch (error) {
        reject(error);
      }
    });
  };
}

/**
 * Setup URL and URLSearchParams polyfills for Google™ Apps Script
 */
function setupPolyfills() {
  try {
    // Access npm helpers for polyfills
    const npmHelpers = getNpmPackageHelpers();
    
    // Install polyfills globally
    if (npmHelpers && npmHelpers.getURLPolyfill) {
      global.URL = npmHelpers.getURLPolyfill();
    }
    
    if (npmHelpers && npmHelpers.getURLSearchParamsPolyfill) {
      global.URLSearchParams = npmHelpers.getURLSearchParamsPolyfill();
    }
  } catch (error) {
  }
}

/**
 * Get the Pipedrive npm package
 * @returns {Object} Pipedrive npm package
 */
function getNpmPackage() {
  try {
    // First try to get from npm helpers
    const npmHelpers = getNpmPackageHelpers();
    if (npmHelpers && npmHelpers.getPipedriveLib) {
      const lib = npmHelpers.getPipedriveLib();
      if (lib) {
        return lib;
      }
    }
    
    // Fallback to direct import if available
    if (typeof pipedrive !== 'undefined') {
      return pipedrive;
    }
    
    throw new Error('Pipedrive npm package not found');
  } catch (error) {
    throw error;
  }
}

/**
 * Get npm package helpers from index.js
 * @returns {Object} npm package helpers
 */
function getNpmPackageHelpers() {
  try {
    // AppLib is the global variable created by webpack
    if (typeof AppLib !== 'undefined') {
      return AppLib;
    }
    if (typeof NPMPackageHelpers !== 'undefined') {
      return NPMPackageHelpers;
    }
    return null;
  } catch (error) {
    return null;
  }
}

// Export functions to global scope for access in other files
const PipedriveNpm = {
  initializePipedriveClient,
  createFixedApiClient,
  fixUpdateMethod,
  sanitizePayload,
  createGASAdapter
}; 