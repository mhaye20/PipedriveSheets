/**
 * API functions that Google Apps Script will call directly.
 * These functions can use the bundled code through the global AppLib variable.
 */

/**
 * Configure Pipedrive client with custom Axios adapter for Google Apps Script
 * This resolves the "no suitable adapter" error
 */
function setupPipedriveClient() {
  // Get the Pipedrive library
  const pipedriveLib = AppLib.getPipedriveLib();
  
  // If we have access to the axios instance used by Pipedrive, configure it
  if (pipedriveLib.v1 && pipedriveLib.v1.axios) {
    // Get our custom adapter
    const gasAdapter = AppLib.getGASAxiosAdapter();
    
    // Configure axios to use our custom adapter
    pipedriveLib.v1.axios.defaults.adapter = gasAdapter;
    Logger.log('Configured Pipedrive client with custom Google Apps Script adapter');
    return true;
  }
  
  // Try to access axios through a different path
  try {
    // Set up global axios if it exists
    if (typeof axios !== 'undefined') {
      if (AppLib.applyGASAdapterToAxios(axios)) {
        Logger.log('Configured global axios with custom Google Apps Script adapter');
        return true;
      }
    }
    
    // Look for axios in the v1 API classes
    if (pipedriveLib.v1) {
      // Loop through all API clients to find and patch their axios instances
      let configured = false;
      
      // Try to find a DealsApi instance and configure its axios
      if (pipedriveLib.v1.DealsApi) {
        const dealsApi = new pipedriveLib.v1.DealsApi();
        if (dealsApi.axios && AppLib.applyGASAdapterToAxios(dealsApi.axios)) {
          configured = true;
          Logger.log('Configured DealsApi axios with custom adapter');
        }
      }
      
      // Get a sample API to extract axios configuration
      const apiClasses = Object.keys(pipedriveLib.v1).filter(key => key.endsWith('Api'));
      for (let i = 0; i < apiClasses.length && !configured; i++) {
        const ApiClass = pipedriveLib.v1[apiClasses[i]];
        if (typeof ApiClass === 'function') {
          try {
            const apiInstance = new ApiClass();
            if (apiInstance.axios && AppLib.applyGASAdapterToAxios(apiInstance.axios)) {
              configured = true;
              Logger.log(`Configured ${apiClasses[i]} axios with custom adapter`);
              break;
            }
          } catch (e) {
            // Skip if we can't instantiate this API class
          }
        }
      }
      
      // Check if axios is in the Configuration class
      if (!configured && pipedriveLib.v1.Configuration) {
        try {
          const config = new pipedriveLib.v1.Configuration();
          if (config.axios && AppLib.applyGASAdapterToAxios(config.axios)) {
            configured = true;
            Logger.log('Configured Configuration axios with custom adapter');
          }
        } catch (e) {
          // Skip if we can't instantiate Configuration
        }
      }
      
      if (configured) {
        return true;
      }
    }
    
    Logger.log('Could not configure Pipedrive client with custom adapter');
    return false;
  } catch (e) {
    Logger.log('Error configuring Pipedrive client: ' + e.toString());
    return false;
  }
}

/**
 * Polyfill URL and URLSearchParams classes in the global scope for use with npm packages
 * This is needed for the Pipedrive npm package which uses these classes internally
 */
function installPolyfills() {
  // In Google Apps Script, the global object is 'this' at the top level
  // Fallback to 'this' or {} if global is not defined
  const globalObj = typeof globalThis !== 'undefined' ? globalThis : 
                    typeof global !== 'undefined' ? global : 
                    typeof window !== 'undefined' ? window : 
                    typeof self !== 'undefined' ? self : this || {};
  
  // Install polyfills globally
  let installed = false;
  
  // URL polyfill
  if (typeof URL === 'undefined') {
    globalObj.URL = AppLib.getURLPolyfill();
    Logger.log('URL polyfill installed globally');
    installed = true;
  }
  
  // URLSearchParams polyfill
  if (typeof URLSearchParams === 'undefined') {
    globalObj.URLSearchParams = AppLib.getURLSearchParamsPolyfill();
    Logger.log('URLSearchParams polyfill installed globally');
    installed = true;
  }
  
  return installed;
}

/**
 * Get the bundled Pipedrive npm package library
 * This function can be called from any Google Apps Script file
 * @returns {Object} The Pipedrive npm package
 */
function getPipedriveNpmLibrary() {
  // Make sure polyfills are installed
  installPolyfills();
  
  // Set up the Pipedrive client with our custom adapter
  setupPipedriveClient();
  
  return AppLib.getPipedriveLib();
}

/**
 * Initialize Pipedrive client with API token
 * @param {string} apiToken - Pipedrive API token
 * @returns {Object} Initialized Pipedrive client
 */
function initPipedriveClient(apiToken) {
  // Make sure polyfills are installed
  installPolyfills();
  
  // Set up the Pipedrive client with our custom adapter
  setupPipedriveClient();
  
  return AppLib.initializePipedrive(apiToken);
}

/**
 * Get all npm packages bundled with webpack
 * @returns {Object} All npm packages
 */
function getNpm() {
  return AppLib.getNpmPackages();
}

/**
 * Required for Google Apps Script to run as a web app
 */
function doGet(e) {
  return HtmlService.createHtmlOutput('App is running!');
}

/**
 * Configure a specific API client instance with the custom adapter
 * This should be called for the specific API client you're about to use
 * @param {Object} apiClient - The API client instance to configure (e.g., dealsApi)
 * @returns {Boolean} - Whether the configuration was successful
 */
function configureApiClientAdapter(apiClient) {
  try {
    if (!apiClient) {
      Logger.log('No API client provided to configure');
      return false;
    }
    
    // Get our custom adapter
    const gasAdapter = AppLib.getGASAxiosAdapter();
    
    // Check if the client has an axios property
    if (apiClient.axios) {
      apiClient.axios.defaults.adapter = gasAdapter;
      Logger.log('Configured specific API client with custom adapter');
      return true;
    }
    
    // Try to access axios through the configuration
    if (apiClient.configuration && apiClient.configuration.axios) {
      apiClient.configuration.axios.defaults.adapter = gasAdapter;
      Logger.log('Configured API client configuration with custom adapter');
      return true;
    }
    
    // Check if the client has basePath and axios properties (seen in logs)
    if (apiClient.basePath && typeof apiClient.axios !== 'undefined') {
      apiClient.axios.defaults.adapter = gasAdapter;
      Logger.log('Configured API client axios with custom adapter');
      return true;
    }
    
    Logger.log('Could not configure specific API client with custom adapter');
    return false;
  } catch (e) {
    Logger.log('Error configuring API client: ' + e.toString());
    return false;
  }
}

/**
 * Inspect the axios structure of an API client for debugging
 * @param {Object} apiClient - The API client to inspect
 * @returns {Boolean} - Whether the client has an axios instance
 */
function inspectApiClientAxios(apiClient) {
  if (!apiClient) {
    Logger.log('No API client provided to inspect');
    return false;
  }
  
  Logger.log('Inspecting API client for axios properties');
  Logger.log('Client type: ' + typeof apiClient);
  
  if (apiClient.basePath) {
    Logger.log('Client basePath: ' + apiClient.basePath);
  }
  
  let hasAxios = false;
  
  // Check direct axios property
  if (apiClient.axios) {
    Logger.log('Client has direct axios property');
    hasAxios = true;
    
    // Check if axios has an adapter
    if (apiClient.axios.defaults && apiClient.axios.defaults.adapter) {
      Logger.log('Axios has adapter: ' + typeof apiClient.axios.defaults.adapter);
    } else {
      Logger.log('Axios missing defaults.adapter property');
    }
  }
  
  // Check configuration
  if (apiClient.configuration) {
    Logger.log('Client has configuration property');
    
    if (apiClient.configuration.axios) {
      Logger.log('Configuration has axios property');
      hasAxios = true;
      
      // Check if config axios has an adapter
      if (apiClient.configuration.axios.defaults && apiClient.configuration.axios.defaults.adapter) {
        Logger.log('Configuration axios has adapter: ' + typeof apiClient.configuration.axios.defaults.adapter);
      } else {
        Logger.log('Configuration axios missing defaults.adapter property');
      }
    }
  }
  
  return hasAxios;
}

// These functions are globally available to all Google Apps Script files
// In Google Apps Script, functions are exported by simply defining them
// No need to explicitly export them with 'global'

// The following functions will be automatically available in Google Apps Script:
// - setupPipedriveClient
// - installPolyfills  
// - getPipedriveNpmLibrary
// - initPipedriveClient
// - getNpm
// - doGet
// - configureApiClientAdapter
// - inspectApiClientAxios 