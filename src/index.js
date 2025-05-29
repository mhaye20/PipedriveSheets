// Import the polyfill for older JavaScript features
import '@babel/polyfill';

// Simpler URL polyfill for Google™ Apps Script (no DOM dependencies)
class URLPolyfill {
  constructor(url, base) {
    // Simple implementation that handles the basics
    let fullUrl = url;
    
    if (base && !url.match(/^https?:\/\//)) {
      // Combine base and URL if url is relative
      fullUrl = base.replace(/\/+$/, '') + '/' + url.replace(/^\/+/, '');
    }
    
    this.href = fullUrl;
    
    // Extract protocol (e.g., "https:")
    const protocolMatch = fullUrl.match(/^(https?:)\/\//);
    this.protocol = protocolMatch ? protocolMatch[1] : '';
    
    // Remove protocol for further parsing
    const withoutProtocol = fullUrl.replace(/^(https?:)?\/\//, '');
    
    // Extract hostname, port, and pathname
    const parts = withoutProtocol.split('/');
    const hostPart = parts[0] || '';
    
    // Handle hostname and port
    const hostParts = hostPart.split(':');
    this.hostname = hostParts[0] || '';
    this.port = hostParts.length > 1 ? hostParts[1] : '';
    
    // Extract pathname
    this.pathname = '/' + (parts.slice(1).join('/').split('?')[0] || '');
    
    // Extract search and hash
    const queryHashMatch = fullUrl.match(/\?([^#]*)(?:#(.*))?$/);
    this.search = queryHashMatch ? '?' + queryHashMatch[1] : '';
    this.hash = queryHashMatch && queryHashMatch[2] ? '#' + queryHashMatch[2] : '';
    
    // Combine hostname and port for host
    this.host = this.port ? this.hostname + ':' + this.port : this.hostname;
    
    // Set origin
    this.origin = this.protocol + '//' + this.host;
    
    // Add basic searchParams implementation
    this.searchParams = {
      get: (param) => {
        const search = this.search.substring(1); // remove leading '?'
        const pairs = search.split('&');
        for (let i = 0; i < pairs.length; i++) {
          const pair = pairs[i].split('=');
          if (decodeURIComponent(pair[0]) === param) {
            return pair.length > 1 ? decodeURIComponent(pair[1]) : '';
          }
        }
        return null;
      }
    };
  }
  
  toString() {
    return this.href;
  }
}

// Set URL polyfill on global object for Google™ Apps Script
if (typeof URL === 'undefined') {
  // Try multiple global object references
  if (typeof globalThis !== 'undefined') {
    globalThis.URL = URLPolyfill;
  } else if (typeof global !== 'undefined') {
    global.URL = URLPolyfill;
  } else if (typeof window !== 'undefined') {
    window.URL = URLPolyfill;
  } else if (typeof this !== 'undefined') {
    this.URL = URLPolyfill;
  }
  // Also set it directly on the global scope
  URL = URLPolyfill;
}

// Create a global URLSearchParams polyfill
class URLSearchParamsPolyfill {
  constructor(init) {
    this.params = [];
    
    if (typeof init === 'string') {
      // Remove leading '?' if present
      const query = init.charAt(0) === '?' ? init.substring(1) : init;
      
      // Parse the query string
      const pairs = query.split('&');
      for (let i = 0; i < pairs.length; i++) {
        const pair = pairs[i].split('=');
        const key = decodeURIComponent(pair[0] || '');
        const value = pair.length > 1 ? decodeURIComponent(pair[1] || '') : '';
        if (key) {
          this.params.push([key, value]);
        }
      }
    }
  }
  
  get(name) {
    for (let i = 0; i < this.params.length; i++) {
      if (this.params[i][0] === name) {
        return this.params[i][1];
      }
    }
    return null;
  }
  
  has(name) {
    for (let i = 0; i < this.params.length; i++) {
      if (this.params[i][0] === name) {
        return true;
      }
    }
    return false;
  }
  
  set(name, value) {
    // Remove existing entries with same name
    this.params = this.params.filter(param => param[0] !== name);
    // Add new entry
    this.params.push([name, value]);
  }
  
  append(name, value) {
    this.params.push([name, value]);
  }
  
  delete(name) {
    this.params = this.params.filter(param => param[0] !== name);
  }
  
  toString() {
    return this.params
      .map(param => encodeURIComponent(param[0]) + '=' + encodeURIComponent(param[1]))
      .join('&');
  }
  
  forEach(callback, thisArg) {
    this.params.forEach((param) => {
      callback.call(thisArg, param[1], param[0], this);
    });
  }
  
  keys() {
    return this.params.map(param => param[0]);
  }
  
  values() {
    return this.params.map(param => param[1]);
  }
  
  entries() {
    return this.params.slice();
  }
}

// Set URLSearchParams polyfill on global object for Google™ Apps Script
if (typeof URLSearchParams === 'undefined') {
  // Try multiple global object references
  if (typeof globalThis !== 'undefined') {
    globalThis.URLSearchParams = URLSearchParamsPolyfill;
  } else if (typeof global !== 'undefined') {
    global.URLSearchParams = URLSearchParamsPolyfill;
  } else if (typeof window !== 'undefined') {
    window.URLSearchParams = URLSearchParamsPolyfill;
  } else if (typeof this !== 'undefined') {
    this.URLSearchParams = URLSearchParamsPolyfill;
  }
  // Also set it directly on the global scope
  URLSearchParams = URLSearchParamsPolyfill;
}

// Axios adapter for Google™ Apps Script
function createGASAxiosAdapter() {
  // Return a function that implements the Axios adapter interface
  return function gasAdapter(config) {
    return new Promise((resolve, reject) => {
      try {
        // Debug the incoming config
        Logger.log('GAS adapter received config: ' + JSON.stringify({
          method: config.method,
          url: config.url,
          hasParams: !!config.params,
          hasData: !!config.data
        }));
        
        // Convert Axios config to UrlFetchApp options
        const options = {
          method: config.method.toUpperCase(),
          muteHttpExceptions: true,
          contentType: 'application/json',
          headers: config.headers || {}
        };

        // Ensure proper Authorization header for OAuth
        if (config.headers && config.headers['Authorization']) {
          // Auth header already set, keep it
          Logger.log('Using existing Authorization header');
        } else if (config.headers && config.headers['api_token']) {
          // Using API token as query param (fallback)
          Logger.log('Using API token from headers');
        } else {
          // If we have access token in the URL, extract and use it as a header instead
          const url = new URLPolyfill(config.url);
          const accessToken = url.searchParams.get('api_token');
          
          if (accessToken) {
            options.headers['Authorization'] = `Bearer ${accessToken}`;
            Logger.log('Extracted and set Authorization header from URL params');
            
            // Remove from URL to avoid duplication
            url.searchParams.delete('api_token');
            config.url = url.toString();
          }
        }

        // Ensure Content-Type header is properly set
        if (!options.headers['Content-Type'] && !options.headers['content-type']) {
          options.headers['Content-Type'] = 'application/json';
        }

        // Fix for URL parameters
        let finalUrl = config.url;
        
        // Save a reference to the original data before manipulating URLs
        const originalData = config.data;
        
        // Handle Pipedrive update API calls specifically
        if (config.url && config.url.includes('[object Object]')) {
          Logger.log('Path parameter substitution needed in URL: ' + config.url);
          
          // For Pipedrive updates, the pattern is typically /endpoint/{id}[object Object]
          // Extract just the base path with the ID
          let basePath = config.url.split('[object Object]')[0];
          
          // Remove any trailing character that might not belong to the path
          if (basePath.endsWith('/')) {
            basePath = basePath.slice(0, -1);
          }
          
          Logger.log('Using base path: ' + basePath);
          finalUrl = basePath;
        }
        
        // Handle params object by properly appending to URL
        if (config.params && typeof config.params === 'object') {
          const queryParams = [];
          for (const key in config.params) {
            if (config.params.hasOwnProperty(key)) {
              const value = config.params[key];
              if (value !== null && value !== undefined) {
                queryParams.push(`${encodeURIComponent(key)}=${encodeURIComponent(value)}`);
              }
            }
          }
          
          // Append params to URL
          if (queryParams.length > 0) {
            finalUrl += (finalUrl.includes('?') ? '&' : '?') + queryParams.join('&');
          }
        }

        // Add payload for non-GET requests
        if ((originalData || config.data) && (config.method === 'put' || config.method === 'post' || config.method === 'patch')) {
            // Use the original data saved before URL manipulations
            let payloadData = originalData || config.data;

            // For PUT requests to /deals/{id}, special handling for Pipedrive SDK
            if (finalUrl.includes('/deals/') && config.method === 'put') {
                Logger.log('Processing Pipedrive deals update request');
                
                // Check if the data has an ID and Body, typical for Pipedrive update calls
                if (payloadData && payloadData.id !== undefined && payloadData.body !== undefined) {
                    Logger.log('Payload has Pipedrive SDK update format {id, body}. Extracting body.');
                    payloadData = payloadData.body;
                } else {
                    Logger.log('Using payload data directly for deals update: ' + JSON.stringify(payloadData).substring(0, 200));
                }
            } else if (payloadData && payloadData.id !== undefined && payloadData.body !== undefined) {
                Logger.log('Payload seems to be Pipedrive SDK update format {id, body}. Extracting body.');
                payloadData = payloadData.body;
            } else {
                Logger.log('Using payload data directly: ' + JSON.stringify(payloadData).substring(0, 200));
            }

            // Ensure payload is a string
            if (typeof payloadData === 'object') {
                options.payload = JSON.stringify(payloadData);
            } else if (payloadData !== null && payloadData !== undefined) {
                options.payload = String(payloadData); // Ensure it's a string even if primitive
            } else {
                options.payload = ''; // Default to empty string if null/undefined
            }
            Logger.log(`Prepared payload: ${options.payload ? options.payload.substring(0, 500) + '...' : '[EMPTY]'}`);
        } else {
            options.payload = ''; // Ensure payload is empty string if not set
            Logger.log('No payload needed for this request method or no data provided.');
        }

        // Log options right before fetch
        try {
          // Stringify options carefully, handling potential circular refs if any
          const optionsString = JSON.stringify(options, (key, value) => {
            if (typeof value === 'function') {
              return 'function'; // Avoid issues with functions in headers/config
            }
            return value;
          }, 2); // Pretty print with 2 spaces
          Logger.log(`Final UrlFetchApp options: ${optionsString.substring(0, 1000)}...`);
        } catch (stringifyError) {
          Logger.log(`Could not stringify UrlFetchApp options: ${stringifyError.message}`);
          Logger.log(`Options Method: ${options.method}, Headers: ${JSON.stringify(options.headers)}, Payload Exists: ${!!options.payload}`);
        }
        
        // Handle case where body objects might incorrectly be used in URL
        if (finalUrl.includes('[object Object]')) {
          Logger.log('Warning: URL contains [object Object], fixing URL: ' + finalUrl);
          // Extract the base URL by removing everything after the object notation
          finalUrl = finalUrl.split('[object Object]')[0];
          Logger.log('Fixed URL: ' + finalUrl);
        }
        
        Logger.log('Making fetch request to: ' + finalUrl);

        // Make the HTTP request using UrlFetchApp
        const response = UrlFetchApp.fetch(finalUrl, options);
        
        // Convert UrlFetchApp response to Axios response format
        const responseData = response.getContentText();
        const responseHeaders = {};
        
        // Handle the headers
        const headers = response.getAllHeaders();
        for (const key in headers) {
          if (headers.hasOwnProperty(key)) {
            responseHeaders[key.toLowerCase()] = headers[key];
          }
        }
        
        // Create the response object
        const axiosResponse = {
          data: responseData,
          status: response.getResponseCode(),
          statusText: '',
          headers: responseHeaders,
          config: config,
          request: {
            responseURL: finalUrl
          }
        };
        
        // Try to parse JSON response if content type indicates JSON
        const contentType = response.getHeaders()['Content-Type'] || '';
        if (contentType.includes('application/json')) {
          try {
            axiosResponse.data = JSON.parse(responseData);
            // Log response body for debugging
            Logger.log(`API Response data (parsed): ${JSON.stringify(axiosResponse.data)}`);
          } catch (e) {
            // If parsing fails, keep the original response text
            Logger.log(`Failed to parse JSON response: ${responseData}`);
          }
        } else {
          Logger.log(`Response is not JSON. Content-Type: ${contentType}, Data: ${responseData.substring(0, 500)}`);
        }
        
        Logger.log('Response status: ' + axiosResponse.status);
        
        // Resolve the promise with the response
        resolve(axiosResponse);
      } catch (error) {
        // Improved error logging
        Logger.log('GAS adapter error: ' + error.toString());
        if (error.stack) {
          Logger.log('Error stack: ' + error.stack);
        }
        
        // Reject with error
        reject(error);
      }
    });
  };
}

// Import npm packages you want to use
import * as pipedrive from 'pipedrive';

// Monkey-patch the Pipedrive SDK to use our polyfills
// This is necessary because the SDK uses 'new URL()' and 'new URLSearchParams()' directly
try {
  // Patch the URL constructor in the global scope to use our polyfill
  if (typeof URL === 'undefined' || URL === URLPolyfill) {
    // Already using polyfill
  } else {
    // Save original if it exists
    const OriginalURL = URL;
    // Replace with a wrapper that checks for our polyfill
    URL = function(url, base) {
      try {
        return new OriginalURL(url, base);
      } catch (e) {
        return new URLPolyfill(url, base);
      }
    };
  }
  
  // Patch URLSearchParams constructor
  if (typeof URLSearchParams === 'undefined' || URLSearchParams === URLSearchParamsPolyfill) {
    // Already using polyfill
  } else {
    // Save original if it exists
    const OriginalURLSearchParams = URLSearchParams;
    // Replace with a wrapper that ensures our methods exist
    URLSearchParams = function(init) {
      const instance = new (OriginalURLSearchParams || URLSearchParamsPolyfill)(init);
      
      // Ensure all required methods exist
      if (!instance.has) {
        // Copy our polyfill methods
        const polyfillInstance = new URLSearchParamsPolyfill(init);
        instance.has = polyfillInstance.has.bind(polyfillInstance);
        instance.set = polyfillInstance.set.bind(polyfillInstance);
        instance.append = polyfillInstance.append.bind(polyfillInstance);
        instance.delete = polyfillInstance.delete.bind(polyfillInstance);
        instance.toString = polyfillInstance.toString.bind(polyfillInstance);
        instance.forEach = polyfillInstance.forEach.bind(polyfillInstance);
        instance.keys = polyfillInstance.keys.bind(polyfillInstance);
        instance.values = polyfillInstance.values.bind(polyfillInstance);
        instance.entries = polyfillInstance.entries.bind(polyfillInstance);
      }
      
      return instance;
    };
  }
  
  Logger.log('Patched URL and URLSearchParams constructors for Pipedrive SDK compatibility');
} catch (patchError) {
  Logger.log('Error patching constructors: ' + patchError.message);
}

/**
 * Get URL polyfill for Google™ Apps Script
 * @returns {Class} URL polyfill class
 */
function getURLPolyfill() {
  return URLPolyfill;
}

/**
 * Get URLSearchParams polyfill for Google™ Apps Script
 * @returns {Class} URLSearchParams polyfill class
 */
function getURLSearchParamsPolyfill() {
  return URLSearchParamsPolyfill;
}

/**
 * Get Axios adapter for Google™ Apps Script
 * @returns {Function} The adapter function
 */
function getGASAxiosAdapter() {
  return createGASAxiosAdapter();
}

/**
 * Utility function to apply the custom GAS adapter directly to any Axios instance
 * @param {Object} axiosInstance - The Axios instance to configure
 * @returns {Boolean} Whether the adapter was successfully applied
 */
function applyGASAdapterToAxios(axiosInstance) {
  try {
    if (!axiosInstance || typeof axiosInstance !== 'object') {
      return false;
    }
    
    const adapter = createGASAxiosAdapter();
    
    // Try to apply to defaults.adapter
    if (axiosInstance.defaults) {
      axiosInstance.defaults.adapter = adapter;
      return true;
    }
    
    // Try direct assignment if no defaults property
    axiosInstance.adapter = adapter;
    return true;
  } catch (e) {
    // Silent failure, will be handled by caller
    return false;
  }
}

/**
 * Accessor function to get the Pipedrive npm package
 * This can be called from any Google™ Apps Script file
 */
function getPipedriveLib() {
  return pipedrive;
}

/**
 * Initialize Pipedrive with an API token
 * @param {string} apiToken - Your Pipedrive API token
 * @returns {Object} - Initialized Pipedrive client
 */
function initializePipedrive(apiToken) {
  try {
    // Create the Pipedrive client
    const client = new pipedrive.Pipedrive(apiToken);
    
    // Intercept and modify APIs to work with Google™ Apps Script
    const customAdapter = createGASAxiosAdapter();
    
    // First apply adapter to the main client
    if (client.axios && client.axios.defaults) {
      client.axios.defaults.adapter = customAdapter;
      Logger.log('Applied GAS adapter to main Pipedrive client');
    }
    
    // Apply adapter to all API clients that client exposes
    if (client.v1) {
      // List of API clients to patch with our custom adapter
      const apiClientNames = [
        'DealsApi', 'PersonsApi', 'OrganizationsApi', 'ActivitiesApi',
        'LeadsApi', 'ProductsApi', 'FilesApi', 'NotesApi', 'UsersApi'
      ];
      
      for (const apiName of apiClientNames) {
        if (client.v1[apiName]) {
          try {
            const apiClient = new client.v1[apiName]();
            
            // Apply custom adapter to axios instance
            if (apiClient.axios && apiClient.axios.defaults) {
              apiClient.axios.defaults.adapter = customAdapter;
              Logger.log(`Applied GAS adapter to ${apiName}`);
            }
            
            // Monkey patch update methods to fix URL issues
            if (apiName === 'DealsApi' && apiClient.updateDeal) {
              const originalUpdateDeal = apiClient.updateDeal;
              apiClient.updateDeal = async function(params) {
                if (params && params.id && params.body) {
                  // Ensure correct URL format by forcing proper URL construction
                  Logger.log(`Enhanced updateDeal called with ID ${params.id}`);
                  return originalUpdateDeal.call(this, params);
                } else {
                  return originalUpdateDeal.apply(this, arguments);
                }
              };
              Logger.log('Enhanced updateDeal method for DealsApi');
            }
            
            client.v1[apiName] = apiClient;
          } catch (apiError) {
            Logger.log(`Error configuring ${apiName}: ${apiError.message}`);
          }
        }
      }
    }
    
    return client;
  } catch (e) {
    Logger.log(`Error initializing Pipedrive: ${e.message}`);
    throw e;
  }
}

/**
 * Get all available npm packages
 * @returns {Object} - Object containing all npm packages
 */
function getNpmPackages() {
  return {
    pipedrive
  };
}

/**
 * Creates a properly configured DealsApi instance for Google™ Apps Script
 * Use this for updating deals to avoid URL and payload issues
 * @param {string} apiToken - Your Pipedrive API token
 * @param {string} basePath - Optional base path for the API
 * @returns {Object} - Configured DealsApi instance
 */
function getUpdatedDealsApi(apiToken, basePath = null) {
  try {
    // Import directly from v1 namespace to avoid initialization issues
    const { DealsApi, Configuration } = pipedrive.v1;
    
    // Create configuration with the provided token
    const config = new Configuration({
      apiKey: apiToken
    });
    
    // If basePath is provided, set it in the configuration
    if (basePath) {
      config.basePath = basePath;
    }
    
    // Create the API client with the configuration
    const dealsApi = new DealsApi(config);
    
    // Apply our custom adapter
    if (dealsApi.axios && dealsApi.axios.defaults) {
      dealsApi.axios.defaults.adapter = createGASAxiosAdapter();
      Logger.log('Applied GAS adapter to DealsApi instance');
    } else {
      Logger.log('Could not apply adapter to DealsApi instance - axios unavailable');
    }
    
    // Monkey patch the updateDeal method to handle URL and payload correctly
    const originalUpdateDeal = dealsApi.updateDeal;
    dealsApi.updateDeal = async function(params) {
      try {
        // Ensure params is properly structured
        if (!params) {
          throw new Error('No parameters provided to updateDeal');
        }
        
        // Make sure ID is a number
        let dealId = params.id;
        if (typeof dealId === 'string') {
          dealId = Number(dealId);
        }
        
        // Ensure body is properly formatted
        if (!params.body) {
          throw new Error('No body provided in updateDeal parameters');
        }
        
        // Create a clean parameters object with just id and body to avoid URL issues
        const cleanParams = {
          id: dealId,
          body: params.body
        };
        
        Logger.log(`Enhanced updateDeal called with ID ${cleanParams.id} and ${Object.keys(cleanParams.body).length} fields`);
        Logger.log(`Update payload: ${JSON.stringify(cleanParams.body).substring(0, 500)}`);
        
        // Call original method with cleaned parameters
        return await originalUpdateDeal.call(this, cleanParams);
      } catch (error) {
        Logger.log(`Error in enhanced updateDeal: ${error.message}`);
        throw error;
      }
    };
    
    return dealsApi;
  } catch (e) {
    Logger.log(`Error creating DealsApi: ${e.message}`);
    throw e;
  }
}

/**
 * Set up polyfills for Google Apps Script environment
 * Call this before using any npm packages that depend on URL/URLSearchParams
 */
function setupPolyfills() {
  // Set URL polyfill
  if (typeof URL === 'undefined' || !URL) {
    if (typeof globalThis !== 'undefined') {
      globalThis.URL = URLPolyfill;
    }
    if (typeof global !== 'undefined') {
      global.URL = URLPolyfill;
    }
    if (typeof window !== 'undefined') {
      window.URL = URLPolyfill;
    }
    // Force set on global scope
    try {
      URL = URLPolyfill;
    } catch (e) {
      // Ignore if we can't set it directly
    }
  }
  
  // Set URLSearchParams polyfill  
  if (typeof URLSearchParams === 'undefined' || !URLSearchParams) {
    if (typeof globalThis !== 'undefined') {
      globalThis.URLSearchParams = URLSearchParamsPolyfill;
    }
    if (typeof global !== 'undefined') {
      global.URLSearchParams = URLSearchParamsPolyfill;
    }
    if (typeof window !== 'undefined') {
      window.URLSearchParams = URLSearchParamsPolyfill;
    }
    // Force set on global scope
    try {
      URLSearchParams = URLSearchParamsPolyfill;
    } catch (e) {
      // Ignore if we can't set it directly
    }
  } else {
    // Even if URLSearchParams exists, it might not have all methods
    // Wrap it to ensure our polyfill methods are available
    const OriginalURLSearchParams = URLSearchParams;
    try {
      // Test if it has the methods we need
      const test = new OriginalURLSearchParams();
      if (!test.has || !test.set || !test.append) {
        // Replace with our polyfill
        URLSearchParams = URLSearchParamsPolyfill;
        if (typeof globalThis !== 'undefined') {
          globalThis.URLSearchParams = URLSearchParamsPolyfill;
        }
        if (typeof global !== 'undefined') {
          global.URLSearchParams = URLSearchParamsPolyfill;
        }
        if (typeof window !== 'undefined') {
          window.URLSearchParams = URLSearchParamsPolyfill;
        }
        Logger.log('Replaced incomplete URLSearchParams with polyfill');
      }
    } catch (e) {
      // If we can't instantiate it, use our polyfill
      URLSearchParams = URLSearchParamsPolyfill;
      Logger.log('URLSearchParams test failed, using polyfill: ' + e.message);
    }
  }
  
  Logger.log('Polyfills set up - URL: ' + (typeof URL !== 'undefined') + ', URLSearchParams: ' + (typeof URLSearchParams !== 'undefined'));
}

// Call setupPolyfills immediately
setupPolyfills();

// Export all functions that need to be available in Google Apps Script
export {
  getPipedriveLib,
  initializePipedrive,
  getNpmPackages,
  getURLPolyfill,
  getURLSearchParamsPolyfill,
  getGASAxiosAdapter,
  applyGASAdapterToAxios,
  getUpdatedDealsApi,
  setupPolyfills
}; 