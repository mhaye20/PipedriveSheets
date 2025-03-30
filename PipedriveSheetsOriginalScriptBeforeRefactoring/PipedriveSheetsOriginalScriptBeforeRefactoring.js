/**
 * Pipedrive to Google Sheets Integration
 * 
 * This script connects to Pipedrive API and fetches data based on filters
 * to populate a Google Sheet with the requested fields.
 */

/**
 * Define the OAuth scopes the script needs
 * @OnlyCurrentDoc
 */

// Configuration constants (can be overridden in the script properties)
const API_KEY = ''; // Default API key
const FILTER_ID = ''; // Default filter ID
const DEFAULT_PIPEDRIVE_SUBDOMAIN = 'api';
const DEFAULT_SHEET_NAME = 'PDexport'; // Default sheet name
const PIPEDRIVE_API_URL_PREFIX = 'https://';
const PIPEDRIVE_API_URL_SUFFIX = '.pipedrive.com/v1';
const ENTITY_TYPES = {
  DEALS: 'deals',
  PERSONS: 'persons',
  ORGANIZATIONS: 'organizations',
  ACTIVITIES: 'activities',
  LEADS: 'leads'
};

// OAuth Constants - YOU NEED TO REGISTER YOUR APP WITH PIPEDRIVE TO GET THESE
// Go to https://developers.pipedrive.com/docs/api/v1/oauth2/auth and create an app
const PIPEDRIVE_CLIENT_ID = 'f48c99e028029bab'; // Client ID from Pipedrive
const PIPEDRIVE_CLIENT_SECRET = '2d245de02052108d8c22d8f7ea8004bc00e7aac7'; // Client Secret from Pipedrive
// Important: If you're having authentication issues, create a new deployment and use that URL here
const REDIRECT_URI = 'https://script.google.com/macros/s/AKfycbwIIm6uKIGWNfvcPdZdNM4Qpj81HIGyfNQTz1KxtJeDusX4K1cbnhan8ufqJpL2Mu5YZw/exec?page=oauthCallback';

/**
 * Creates the menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Pipedrive')
    .addItem('Connect to Pipedrive', 'showAuthorizationDialog')
    .addItem('Sync Data', 'syncFromPipedrive')
    .addItem('Select Columns', 'showColumnSelector')
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Shows a dialog to authorize with Pipedrive
 */
function showAuthorizationDialog() {
  // Check if already authorized
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  
  if (accessToken) {
    // Check if token is still valid by making a test request
    try {
      const testUrl = `${getPipedriveApiUrl()}/users/me?api_token=${accessToken}`;
      const response = UrlFetchApp.fetch(testUrl);
      const data = JSON.parse(response.getContentText());
      
      if (data.success) {
        const ui = SpreadsheetApp.getUi();
        const result = ui.alert(
          'Already Connected',
          'You are already connected to Pipedrive as ' + data.data.name + '. Do you want to reconnect?',
          ui.ButtonSet.YES_NO
        );
        
        if (result === ui.Button.NO) {
          return;
        }
      }
    } catch (e) {
      // Token is probably invalid, continue with auth
      Logger.log('Error checking token: ' + e.message);
    }
  }
  
  // Create the OAuth2 authorization URL
  const authUrl = `https://oauth.pipedrive.com/oauth/authorize?client_id=${PIPEDRIVE_CLIENT_ID}&redirect_uri=${REDIRECT_URI}&state=${generateRandomState()}&scope=contacts:read deals:read`;
  
  // Display the authorization dialog
  const template = HtmlService.createTemplate(
    '<html>'
    + '<head>'
    + '<style>'
    + 'button { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }'
    + 'button:hover { background-color: #45a049; }'
    + '.container { text-align: center; padding: 20px; }'
    + 'h2 { color: #333; }'
    + '</style>'
    + '</head>'
    + '<body>'
    + '<div class="container">'
    + '<h2>Connect to Pipedrive</h2>'
    + '<p>Click the button below to authorize this app to access your Pipedrive data.</p>'
    + '<button onclick="openAuthWindow()">Sign in with Pipedrive</button>'
    + '</div>'
    + '<script>'
    + 'function openAuthWindow() {'
    + '  window.open("<?= authUrl ?>", "_blank", "width=600,height=700");'
    + '  google.script.host.close();'
    + '}'
    + '</script>'
    + '</body>'
    + '</html>'
  );
  
  template.authUrl = authUrl;
  
  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setHeight(300)
    .setTitle('Connect to Pipedrive');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Connect to Pipedrive');
}

/**
 * Generates a random state string for OAuth security
 */
function generateRandomState() {
  return Math.random().toString(36).substring(2, 15);
}

/**
 * Handles the OAuth callback
 */
function doGet(e) {
  if (e.parameter.page === 'oauthCallback') {
    // This function is called when Pipedrive redirects back to the app
    const code = e.parameter.code;
    
    if (code) {
      try {
        // Exchange the authorization code for an access token
        const tokenResponse = UrlFetchApp.fetch('https://oauth.pipedrive.com/oauth/token', {
          method: 'post',
          payload: {
            grant_type: 'authorization_code',
            code: code,
            redirect_uri: REDIRECT_URI,
            client_id: PIPEDRIVE_CLIENT_ID,
            client_secret: PIPEDRIVE_CLIENT_SECRET
          }
        });
        
        const tokenData = JSON.parse(tokenResponse.getContentText());
        
        // Save the tokens in script properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
        scriptProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
        scriptProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
        
        // Get user info to determine the subdomain
        const userResponse = UrlFetchApp.fetch('https://api.pipedrive.com/v1/users/me?api_token=' + tokenData.access_token);
        const userData = JSON.parse(userResponse.getContentText());
        
        if (userData.success) {
          // Extract the company domain from the user data
          const companyDomain = userData.data.company_domain;
          scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', companyDomain);
        }
        
        // Display success page
        return HtmlService.createHtmlOutput(
          '<html>'
          + '<head>'
          + '<style>'
          + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
          + '.success { color: #4CAF50; font-size: 24px; }'
          + '</style>'
          + '</head>'
          + '<body>'
          + '<h1 class="success">✓ Successfully Connected!</h1>'
          + '<p>You have successfully connected your Pipedrive account. You can close this window and return to Google Sheets.</p>'
          + '</body>'
          + '</html>'
        )
        .setTitle('Connected to Pipedrive');
      } catch (error) {
        // Handle authorization error
        Logger.log('Token exchange error: ' + error.message);
        return HtmlService.createHtmlOutput(
          '<html>'
          + '<head>'
          + '<style>'
          + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
          + '.error { color: #f44336; font-size: 24px; }'
          + '</style>'
          + '</head>'
          + '<body>'
          + '<h1 class="error">Error Connecting</h1>'
          + '<p>There was an error connecting to Pipedrive: ' + error.message + '</p>'
          + '<p>Please try again or contact support.</p>'
          + '</body>'
          + '</html>'
        )
        .setTitle('Connection Error');
      }
    } else {
      // Authorization was denied or errored
      return HtmlService.createHtmlOutput(
        '<html>'
        + '<head>'
        + '<style>'
        + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
        + '.error { color: #f44336; font-size: 24px; }'
        + '</style>'
        + '</head>'
        + '<body>'
        + '<h1 class="error">Authorization Failed</h1>'
        + '<p>The authorization process was cancelled or failed.</p>'
        + '</body>'
        + '</html>'
      )
      .setTitle('Authorization Error');
    }
  }
  
  // Default page if no specific action
  return HtmlService.createHtmlOutput('Pipedrive Integration App');
}

/**
 * Refreshes the access token if needed
 */
function refreshAccessTokenIfNeeded() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  const refreshToken = scriptProperties.getProperty('PIPEDRIVE_REFRESH_TOKEN');
  const expiresAt = scriptProperties.getProperty('PIPEDRIVE_TOKEN_EXPIRES');
  
  // If no token or refresh token, we can't refresh
  if (!accessToken || !refreshToken) {
    return false;
  }
  
  // Check if token is expired or about to expire (5 minutes buffer)
  const now = new Date().getTime();
  if (!expiresAt || now > (parseInt(expiresAt) - (5 * 60 * 1000))) {
    try {
      // Refresh the token
      const tokenResponse = UrlFetchApp.fetch('https://oauth.pipedrive.com/oauth/token', {
        method: 'post',
        payload: {
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          client_id: PIPEDRIVE_CLIENT_ID,
          client_secret: PIPEDRIVE_CLIENT_SECRET
        }
      });
      
      const tokenData = JSON.parse(tokenResponse.getContentText());
      
      // Save the new tokens
      scriptProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
      if (tokenData.refresh_token) {
        scriptProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
      }
      scriptProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
      
      return true;
    } catch (error) {
      Logger.log('Token refresh error: ' + error.message);
      return false;
    }
  }
  
  // Token is still valid
  return true;
}

/**
 * Gets the Pipedrive API URL with the subdomain from settings
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
 * Main entry point for syncing data from Pipedrive
 */
function syncFromPipedrive() {
  // Show a loading message
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  
  spreadsheet.toast('Starting Pipedrive sync...', 'Sync Status', -1);
  
  // Get the configured entity type for this specific sheet
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Set the active sheet as the current sheet for this operation
  scriptProperties.setProperty('SHEET_NAME', activeSheetName);
  
  try {
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        syncDealsFromFilter();
        break;
      case ENTITY_TYPES.PERSONS:
        syncPersonsFromFilter();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        syncOrganizationsFromFilter();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        syncActivitiesFromFilter();
        break;
      case ENTITY_TYPES.LEADS:
        syncLeadsFromFilter();
        break;
      default:
        spreadsheet.toast('Unknown entity type. Please check settings.', 'Sync Error', 10);
        break;
    }
  } catch (error) {
    // If there's an error, show it
    spreadsheet.toast('Error: ' + error.message, 'Sync Error', 10);
    Logger.log('Sync error: ' + error.message);
  }
}

/**
 * Shows settings dialog where users can configure filter ID and entity type
 */
function showSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();
  
  // Get current settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Check if we're authenticated
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  if (!accessToken) {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Not Connected',
      'You need to connect to Pipedrive before configuring settings. Connect now?',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      showAuthorizationDialog();
    }
    return;
  }
  
  const savedSubdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  
  // Get sheet-specific settings using the sheet name
  const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  
  const savedFilterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const savedEntityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Create HTML content for the settings dialog
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 5px; margin-top: 5px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 12px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      .note { font-size: 12px; color: #666; margin-top: 5px; }
      .domain-container { display: flex; align-items: center; }
      .domain-input { flex: 1; margin-right: 5px; }
      .domain-suffix { padding: 5px; background: #f0f0f0; border: 1px solid #ccc; }
      .loading { display: none; margin-right: 10px; }
      .loader { 
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid rgba(255,255,255,.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 1s ease-in-out infinite;
        vertical-align: middle;
      }
      .sheet-info {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        font-size: 14px;
        border-left: 4px solid #4285f4;
      }
      .connected-status {
        background-color: #e6f4ea;
        border-left: 4px solid #34a853;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
      }
      .connected-status i {
        color: #34a853;
        margin-right: 10px;
        font-size: 24px;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
    <h3>Pipedrive Integration Settings</h3>
    
    <div class="connected-status">
      <span style="color: #34a853; font-size: 24px; margin-right: 10px;">✓</span>
      <div>
        <strong>Connected to Pipedrive</strong><br>
        <span style="font-size: 12px;">Company: ${savedSubdomain}.pipedrive.com</span>
      </div>
      <button style="margin-left: auto; background-color: #f1f3f4; color: #202124;" onclick="reconnect()">Reconnect</button>
    </div>
    
    <div class="sheet-info">
      Configuring settings for sheet: <strong>"${activeSheetName}"</strong>
    </div>
    
    <form id="settingsForm">
      <label for="entityType">Entity Type:</label>
      <select id="entityType">
        <option value="deals" ${savedEntityType === 'deals' ? 'selected' : ''}>Deals</option>
        <option value="persons" ${savedEntityType === 'persons' ? 'selected' : ''}>Persons</option>
        <option value="organizations" ${savedEntityType === 'organizations' ? 'selected' : ''}>Organizations</option>
        <option value="activities" ${savedEntityType === 'activities' ? 'selected' : ''}>Activities</option>
        <option value="leads" ${savedEntityType === 'leads' ? 'selected' : ''}>Leads</option>
      </select>
      <p class="note">Select which entity type you want to sync from Pipedrive</p>
      
      <label for="filterId">Filter ID:</label>
      <input type="text" id="filterId" value="${savedFilterId}" />
      <p class="note">The filter ID can be found in the URL when viewing your filter in Pipedrive</p>
      
      <input type="hidden" id="sheetName" value="${activeSheetName}" />
      <p class="note">The data will be exported to the current sheet: "${activeSheetName}"</p>
      
      <div class="button-container">
        <span class="loading" id="saveLoading"><span class="loader"></span> Saving...</span>
        <button type="button" id="saveBtn" onclick="saveSettings()">Save Settings</button>
      </div>
    </form>
    
    <script>
      function saveSettings() {
        const entityType = document.getElementById('entityType').value;
        const filterId = document.getElementById('filterId').value;
        const sheetName = document.getElementById('sheetName').value;
        
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'inline-block';
        document.getElementById('saveBtn').disabled = true;
        
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            alert('Settings saved successfully!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            alert('Error saving settings: ' + error.message);
          })
          .saveSettings('', entityType, filterId, '', sheetName);
      }
      
      function reconnect() {
        google.script.host.close();
        google.script.run.showAuthorizationDialog();
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(520)
  .setTitle(`Pipedrive Settings for "${activeSheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
}

/**
 * Shows a column selector UI for users to choose which columns to display
 */
function showColumnSelector() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get API key and filter ID from properties or use defaults
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    
    // Get sheet-specific settings
    const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    
    const filterId = scriptProperties.getProperty(sheetFilterIdKey) || FILTER_ID;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    // Use the active sheet for the operation
    scriptProperties.setProperty('SHEET_NAME', activeSheetName);
    
    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      SpreadsheetApp.getUi().alert('Please configure your Pipedrive API Key in Settings first.');
      showSettings();
      return;
    }
    
    // Check if we can connect to Pipedrive and get a sample item
    SpreadsheetApp.getActiveSpreadsheet().toast(`Connecting to Pipedrive to retrieve ${entityType} data...`);
    
    // Get sample data based on entity type (1 item only)
    let sampleData = [];
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        sampleData = getDealsWithFilter(filterId, 1);
        break;
      case ENTITY_TYPES.PERSONS:
        sampleData = getPersonsWithFilter(filterId, 1);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        sampleData = getOrganizationsWithFilter(filterId, 1);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        sampleData = getActivitiesWithFilter(filterId, 1);
        break;
      case ENTITY_TYPES.LEADS:
        sampleData = getLeadsWithFilter(filterId, 1);
        break;
    }
    
    if (!sampleData || sampleData.length === 0) {
      SpreadsheetApp.getUi().alert(`Could not retrieve any ${entityType} from Pipedrive. Please check your API key and filter ID.`);
      return;
    }
    
    // Log the raw sample data for debugging
    Logger.log(`PIPEDRIVE DEBUG - Sample ${entityType} raw data:`);
    Logger.log(JSON.stringify(sampleData[0], null, 2));
    
    // Get all available columns from the sample data
    const sampleItem = sampleData[0];
    const availableColumns = [];
    
    // Get field definitions to show friendly names
    let fieldDefinitions = [];
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fieldDefinitions = getDealFields();
        break;
      case ENTITY_TYPES.PERSONS:
        fieldDefinitions = getPersonFields();
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
    }
    
    const fieldMap = {};
    fieldDefinitions.forEach(field => {
      fieldMap[field.key] = field.name;
    });
    
    // Function to recursively extract fields from complex objects
    function extractFields(obj, parentPath = '', parentName = '') {
      // Skip if null or not an object
      if (obj === null || typeof obj !== 'object') {
        return;
      }
      
      // Handle arrays by looking at the first item
      if (Array.isArray(obj)) {
        // If it's an array of objects with a common structure, extract from first item
        if (obj.length > 0 && typeof obj[0] === 'object' && obj[0] !== null) {
          // Special handling for multiple options fields
          if (obj[0].hasOwnProperty('label') && obj[0].hasOwnProperty('id')) {
            // Add a field for the entire multiple options array
            availableColumns.push({
              key: parentPath,
              name: parentName + ' (Multiple Options)',
              isNested: true,
              parentKey: parentPath.split('.').slice(0, -1).join('.')
            });
            return;
          }
          
          // For arrays of structured objects, like emails or phones
          if (obj[0].hasOwnProperty('value') && obj[0].hasOwnProperty('primary')) {
            let displayName = 'Primary ' + (parentName || 'Item');
            if (obj[0].hasOwnProperty('label')) {
              displayName = 'Primary ' + parentName + ' (' + obj[0].label + ')';
            }
            
            availableColumns.push({
              key: parentPath + '.0.value',
              name: displayName,
              isNested: true,
              parentKey: parentPath
            });
          } else {
            extractFields(obj[0], parentPath + '.0', parentName + ' (First Item)');
          }
        }
        return;
      }
      
      // Extract properties from this object
      for (const key in obj) {
        // Skip internal properties, functions, or empty objects
        if (key.startsWith('_') || typeof obj[key] === 'function') {
          continue;
        }
        
        const currentPath = parentPath ? parentPath + '.' + key : key;
        
        // For common Pipedrive objects, create shortcuts
        if (key === 'name' && parentPath && (obj.hasOwnProperty('email') || obj.hasOwnProperty('phone') || obj.hasOwnProperty('address'))) {
          // For person or organization name
          availableColumns.push({
            key: currentPath,
            name: (parentName || parentPath) + ' Name',
            isNested: true,
            parentKey: parentPath
          });
        } else if (typeof obj[key] === 'object' && obj[key] !== null) {
          // Recursively extract from nested objects
          let nestedParentName = parentName ? parentName + ' ' + formatColumnName(key) : formatColumnName(key);
          extractFields(obj[key], currentPath, nestedParentName);
        } else {
          // Simple property
          const displayName = parentName ? parentName + ' ' + formatColumnName(key) : formatColumnName(key);
          availableColumns.push({
            key: currentPath,
            name: displayName,
            isNested: parentPath ? true : false,
            parentKey: parentPath
          });
        }
      }
    }
    
    // Build top-level column data first
    for (const key in sampleItem) {
      // Skip internal properties or functions
      if (key.startsWith('_') || typeof sampleItem[key] === 'function') {
        continue;
      }
      
      const displayName = fieldMap[key] || formatColumnName(key);
      
      // Add the top-level column
      availableColumns.push({
        key: key,
        name: displayName,
        isNested: false
      });
      
      // If it's a complex object, extract nested fields
      if (typeof sampleItem[key] === 'object' && sampleItem[key] !== null) {
        extractFields(sampleItem[key], key, displayName);
      }
    }
    
    // Log all extracted columns for debugging
    Logger.log(`PIPEDRIVE DEBUG - All available columns for ${entityType}:`);
    Logger.log(JSON.stringify(availableColumns, null, 2));
    
    // Get previously saved column preferences for this specific sheet and entity type
    const columnSettingsKey = `COLUMNS_${activeSheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let selectedColumns = [];
    
    if (savedColumnsJson) {
      try {
        selectedColumns = JSON.parse(savedColumnsJson);
        // Filter out any columns that no longer exist
        selectedColumns = selectedColumns.filter(col => 
          availableColumns.some(availCol => availCol.key === col.key)
        );
      } catch (e) {
        Logger.log('Error parsing saved columns: ' + e.message);
        selectedColumns = [];
      }
    }
    
    // Show the column selector UI
    showColumnSelectorUI(availableColumns, selectedColumns, entityType, activeSheetName);
  } catch (error) {
    Logger.log('Error in column selector: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
  }
}

/**
 * Shows the column selector UI
 */
function showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName) {
  const htmlContent = `
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
      .container { display: flex; height: 400px; }
      .column { width: 50%; padding: 10px; box-sizing: border-box; }
      .scrollable { height: 340px; overflow-y: auto; border: 1px solid #ccc; padding: 5px; }
      .header { font-weight: bold; margin-bottom: 10px; }
      .item { padding: 5px; margin: 2px 0; cursor: pointer; border-radius: 3px; }
      .item:hover { background-color: #f0f0f0; }
      .selected { background-color: #e8f0fe; }
      .footer { margin-top: 10px; display: flex; justify-content: space-between; }
      .search { margin-bottom: 10px; width: 100%; padding: 5px; }
      button { padding: 8px 12px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      button.secondary { background-color: #f0f0f0; color: #333; }
      .action-btns { display: flex; gap: 5px; align-items: center; }
      .category { font-weight: bold; margin-top: 5px; padding: 5px; background-color: #f0f0f0; }
      .nested { margin-left: 15px; }
      .info { font-size: 12px; color: #666; margin-bottom: 5px; }
      .loading { display: none; margin-right: 10px; }
      .loader { 
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid rgba(255,255,255,.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 1s ease-in-out infinite;
        vertical-align: middle;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
      .drag-handle {
        display: inline-block;
        width: 10px;
        height: 16px;
        background-image: radial-gradient(circle, #999 1px, transparent 1px);
        background-size: 3px 3px;
        background-position: 0 center;
        background-repeat: repeat;
        margin-right: 8px;
        cursor: grab;
        opacity: 0.5;
      }
      .selected:hover .drag-handle {
        opacity: 1;
      }
      .dragging {
        opacity: 0.4;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
      }
      .over {
        border-top: 2px solid #4285f4;
      }
      .sheet-info {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        font-size: 14px;
        border-left: 4px solid #4285f4;
      }
    </style>
    
    <div class="header">
      <div class="sheet-info">
        Configuring columns for <strong>${entityType}</strong> in sheet <strong>"${sheetName}"</strong>
      </div>
      <p class="info">Select columns from the left panel and add them to the right panel. Drag items in the right panel to reorder them.</p>
    </div>
    
    <div class="container">
      <div class="column">
        <input type="text" id="searchBox" class="search" placeholder="Search for columns...">
        <div class="header">Available Columns</div>
        <div id="availableList" class="scrollable">
          <!-- Available columns will be populated here by JavaScript -->
        </div>
      </div>
      
      <div class="column">
        <div class="header">Selected Columns</div>
        <div id="selectedList" class="scrollable">
          <!-- Selected columns will be populated here by JavaScript -->
        </div>
      </div>
    </div>
    
    <div class="footer">
      <div class="action-btns">
        <button class="secondary" id="debug">View Debug Info</button>
      </div>
      <div class="action-btns">
        <span class="loading" id="saveLoading"><span class="loader"></span> Saving...</span>
        <button class="secondary" id="cancel">Cancel</button>
        <button id="save">Save & Close</button>
      </div>
    </div>

    <script>
      // Initialize data
      let availableColumns = ${JSON.stringify(availableColumns)};
      let selectedColumns = ${JSON.stringify(selectedColumns)};
      const entityType = "${entityType}";
      const sheetName = "${sheetName}";
      
      // DOM elements
      const availableList = document.getElementById('availableList');
      const selectedList = document.getElementById('selectedList');
      const searchBox = document.getElementById('searchBox');
      
      // Render the lists
      function renderAvailableList(searchTerm = '') {
        availableList.innerHTML = '';
        
        // Group columns by parent key or top-level
        const topLevel = [];
        const nested = {};
        
        availableColumns.forEach(col => {
          if (!selectedColumns.some(selected => selected.key === col.key)) {
            if (col.name.toLowerCase().includes(searchTerm.toLowerCase())) {
              if (!col.isNested) {
                topLevel.push(col);
              } else {
                const parentKey = col.parentKey || 'unknown';
                if (!nested[parentKey]) {
                  nested[parentKey] = [];
                }
                nested[parentKey].push(col);
              }
            }
          }
        });
        
        // Add top-level columns first
        if (topLevel.length > 0) {
          const topLevelHeader = document.createElement('div');
          topLevelHeader.className = 'category';
          topLevelHeader.textContent = 'Main Fields';
          availableList.appendChild(topLevelHeader);
          
          topLevel.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.textContent = col.name;
            item.dataset.key = col.key;
            item.onclick = () => addColumn(col);
            availableList.appendChild(item);
          });
        }
        
        // Then add nested columns by parent
        for (const parentKey in nested) {
          if (nested[parentKey].length > 0) {
            const parentName = availableColumns.find(col => col.key === parentKey)?.name || parentKey;
            
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'category';
            categoryHeader.textContent = parentName;
            availableList.appendChild(categoryHeader);
            
            nested[parentKey].forEach(col => {
              const item = document.createElement('div');
              item.className = 'item nested';
              item.textContent = col.name;
              item.dataset.key = col.key;
              item.onclick = () => addColumn(col);
              availableList.appendChild(item);
            });
          }
        }
      }
      
      function renderSelectedList() {
        selectedList.innerHTML = '';
        selectedColumns.forEach((col, index) => {
          const item = document.createElement('div');
          item.className = 'item selected';
          item.dataset.key = col.key;
          item.dataset.index = index;
          item.draggable = true;
          
          // Add drag handle
          const dragHandle = document.createElement('span');
          dragHandle.className = 'drag-handle';
          dragHandle.innerHTML = '&nbsp;&nbsp;&nbsp;';
          item.appendChild(dragHandle);
          
          // Add column name
          const nameSpan = document.createElement('span');
          nameSpan.textContent = col.name;
          item.appendChild(nameSpan);
          
          item.ondragstart = handleDragStart;
          item.ondragover = handleDragOver;
          item.ondrop = handleDrop;
          item.ondragend = handleDragEnd;
          
          const removeBtn = document.createElement('span');
          removeBtn.textContent = ' ✕';
          removeBtn.style.color = 'red';
          removeBtn.style.float = 'right';
          removeBtn.style.cursor = 'pointer';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeColumn(col);
          };
          
          item.appendChild(removeBtn);
          selectedList.appendChild(item);
        });
      }
      
      // Drag and drop functionality
      let draggedItem = null;
      
      function handleDragStart(e) {
        draggedItem = this;
        this.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', this.dataset.index);
        
        // Add a small delay to make the visual change noticeable
        setTimeout(() => {
          this.style.opacity = '0.4';
        }, 0);
      }
      
      function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this.classList.add('over');
      }
      
      function handleDrop(e) {
        e.preventDefault();
        this.classList.remove('over');
        
        const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
        const toIndex = parseInt(this.dataset.index);
        
        if (fromIndex !== toIndex) {
          const item = selectedColumns[fromIndex];
          selectedColumns.splice(fromIndex, 1);
          selectedColumns.splice(toIndex, 0, item);
          renderSelectedList();
        }
      }
      
      function handleDragEnd() {
        this.classList.remove('dragging');
        document.querySelectorAll('.item').forEach(item => {
          item.classList.remove('over');
        });
      }
      
      // Column management
      function addColumn(column) {
        selectedColumns.push(column);
        renderAvailableList(searchBox.value);
        renderSelectedList();
      }
      
      function removeColumn(column) {
        selectedColumns = selectedColumns.filter(col => col.key !== column.key);
        renderAvailableList(searchBox.value);
        renderSelectedList();
      }
      
      // Event listeners
      document.getElementById('save').onclick = () => {
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'inline-block';
        document.getElementById('save').disabled = true;
        document.getElementById('cancel').disabled = true;
        
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('saveLoading').style.display = 'none';
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('save').disabled = false;
            document.getElementById('cancel').disabled = false;
            alert('Error saving column preferences: ' + error.message);
          })
          .saveColumnPreferences(selectedColumns, entityType, sheetName);
      };
      
      document.getElementById('cancel').onclick = () => {
        google.script.host.close();
      };
      
      document.getElementById('debug').onclick = () => {
        google.script.run.logDebugInfo();
        alert('Debug information has been logged to the Apps Script execution log. You can view it from View > Logs in the Apps Script editor.');
      };
      
      searchBox.oninput = () => {
        renderAvailableList(searchBox.value);
      };
      
      // Initial render
      renderAvailableList();
      renderSelectedList();
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(800)
    .setHeight(550)
    .setTitle(`Select Columns for ${entityType} in "${sheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Select Columns for ${entityType} in "${sheetName}"`);
}

/**
 * Saves column preferences to script properties
 */
function saveColumnPreferences(columns, entityType, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Store columns based on both entity type and sheet name for sheet-specific preferences
  const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
  scriptProperties.setProperty(columnSettingsKey, JSON.stringify(columns));
}

/**
 * Saves settings to script properties
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Save global settings (only if provided)
  if (apiKey) scriptProperties.setProperty('PIPEDRIVE_API_KEY', apiKey);
  if (subdomain) scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
  
  // Save sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  
  scriptProperties.setProperty(sheetFilterIdKey, filterId);
  scriptProperties.setProperty(sheetEntityTypeKey, entityType);
  scriptProperties.setProperty('SHEET_NAME', sheetName);
}

/**
 * Main function to sync deals from a Pipedrive filter to the Google Sheet
 */
function syncDealsFromFilter() {
  syncPipedriveDataToSheet(ENTITY_TYPES.DEALS);
}

/**
 * Main function to sync persons from a Pipedrive filter to the Google Sheet
 */
function syncPersonsFromFilter() {
  syncPipedriveDataToSheet(ENTITY_TYPES.PERSONS);
}

/**
 * Main function to sync organizations from a Pipedrive filter to the Google Sheet
 */
function syncOrganizationsFromFilter() {
  syncPipedriveDataToSheet(ENTITY_TYPES.ORGANIZATIONS);
}

/**
 * Main function to sync activities from a Pipedrive filter to the Google Sheet
 */
function syncActivitiesFromFilter() {
  syncPipedriveDataToSheet(ENTITY_TYPES.ACTIVITIES);
}

/**
 * Main function to sync leads from a Pipedrive filter to the Google Sheet
 */
function syncLeadsFromFilter() {
  syncPipedriveDataToSheet(ENTITY_TYPES.LEADS);
}

/**
 * Generic function to sync any Pipedrive entity to a Google Sheet
 */
function syncPipedriveDataToSheet(entityType) {
  try {
    // Get API key and filter ID from properties or use defaults
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    const sheetName = scriptProperties.getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Get sheet-specific filter ID
    const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
    const filterId = scriptProperties.getProperty(sheetFilterIdKey) || FILTER_ID;
    
    // Check if API key is configured
    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Pipedrive API Key Required', 
              'Please configure your Pipedrive API Key in Settings first.', 
              ui.ButtonSet.OK);
      showSettings();
      return;
    }
    
    // Get saved column preferences for this specific sheet and entity type
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let selectedColumns = [];
    
    if (savedColumnsJson) {
      try {
        selectedColumns = JSON.parse(savedColumnsJson);
      } catch (e) {
        Logger.log('Error parsing saved columns: ' + e.message);
        selectedColumns = [];
      }
    }
    
    // If no columns are selected, prompt to configure
    if (selectedColumns.length === 0) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert('No Columns Selected', 
                              `You need to select which columns to display for ${entityType} in "${sheetName}". Would you like to configure them now?`, 
                              ui.ButtonSet.YES_NO);
      
      if (response === ui.Button.YES) {
        showColumnSelector();
      }
      
      return;
    }
    
    // Create column mapping for the export
    const columnsToUse = selectedColumns.map(col => col.key);
    const headerRow = selectedColumns.map(col => col.name);
    
    // Get option mappings for multiple option fields
    const optionMappings = getFieldOptionMappingsForEntity(entityType);
    Logger.log(`Option mappings for ${entityType} fields: ${JSON.stringify(optionMappings)}`);
    
    // Display status message
    SpreadsheetApp.getActiveSpreadsheet().toast(`Syncing ${entityType} from Pipedrive to "${sheetName}"... This may take a while for large datasets.`);
    
    // Get all data for the export based on entity type
    let items = [];
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        items = getDealsWithFilter(filterId, 0);
        break;
      case ENTITY_TYPES.PERSONS:
        items = getPersonsWithFilter(filterId, 0);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        items = getOrganizationsWithFilter(filterId, 0);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        items = getActivitiesWithFilter(filterId, 0);
        break;
      case ENTITY_TYPES.LEADS:
        items = getLeadsWithFilter(filterId, 0);
        break;
    }
    
    if (!items || items.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`No ${entityType} found with the specified filter.`);
      return;
    }
    
    // Log the full result for debugging
    Logger.log(`PIPEDRIVE DEBUG - Sync triggered. Number of ${entityType} retrieved: ${items.length}`);
    Logger.log(`PIPEDRIVE DEBUG - First ${entityType} data (sample):`);
    Logger.log(JSON.stringify(items[0], null, 2));
    Logger.log(`PIPEDRIVE DEBUG - Selected columns for ${sheetName}:`);
    Logger.log(JSON.stringify(selectedColumns, null, 2));
    
    // Prepare and write data to sheet with the selected columns
    writeDataToSheet(items, { 
      columns: columnsToUse, 
      headerRow: headerRow, 
      optionMappings: optionMappings,
      entityType: entityType,
      sheetName: sheetName
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`${entityType} successfully synced from Pipedrive to "${sheetName}"! (${items.length} items total)`);
  } catch (error) {
    Logger.log('Error: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
  }
}

/**
 * Writes data to the Google Sheet with columns matching the filter
 */
function writeDataToSheet(items, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get sheet name from options or properties
  const sheetName = options.sheetName || 
                    PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 
                    DEFAULT_SHEET_NAME;
  
  let sheet = ss.getSheetByName(sheetName);
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    // Clear existing data
    sheet.clear();
  }
  
  // Use the columns and headers from the columns info
  const columns = options.columns;
  const headerRow = options.headerRow;
  const optionMappings = options.optionMappings || {};
  const entityType = options.entityType || ENTITY_TYPES.DEALS;
  
  // Add header row (directly in row 1)
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
  
  // Add data rows
  const data = [];
  items.forEach(item => {
    const row = [];
    columns.forEach(column => {
      // Get the value using path notation
      const value = getValueByPath(item, column);
      
      // Format the value, passing the column path to determine field type
      row.push(formatValue(value, column, optionMappings));
    });
    data.push(row);
  });
  
  // Write all rows at once for better performance
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headerRow.length).setValues(data);
  }
  
  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, headerRow.length);
  
  // Add timestamp of last sync
  const timestampRow = sheet.getLastRow() + 2;
  sheet.getRange(timestampRow, 1).setValue('Last synced:');
  sheet.getRange(timestampRow, 2).setValue(new Date()).setNumberFormat('yyyy-MM-dd HH:mm:ss');
}

/**
 * Logs debug information about the Pipedrive data
 */
function logDebugInfo() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetName = scriptProperties.getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
  
  // Get sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  
  const filterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Show which column selections are available for the current entity type and sheet
  const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
  const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
  
  if (savedColumnsJson) {
    Logger.log(`\n===== COLUMN SETTINGS FOR ${sheetName} - ${entityType} =====`);
    try {
      const selectedColumns = JSON.parse(savedColumnsJson);
      Logger.log(`Number of selected columns: ${selectedColumns.length}`);
      Logger.log(JSON.stringify(selectedColumns, null, 2));
    } catch (e) {
      Logger.log(`Error parsing column settings: ${e.message}`);
    }
  } else {
    Logger.log(`\n===== NO COLUMN SETTINGS FOUND FOR ${sheetName} - ${entityType} =====`);
  }
  
  // Get a sample item to see what data is available
  let sampleData = [];
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      sampleData = getDealsWithFilter(filterId, 1);
      break;
    case ENTITY_TYPES.PERSONS:
      sampleData = getPersonsWithFilter(filterId, 1);
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      sampleData = getOrganizationsWithFilter(filterId, 1);
      break;
    case ENTITY_TYPES.ACTIVITIES:
      sampleData = getActivitiesWithFilter(filterId, 1);
      break;
    case ENTITY_TYPES.LEADS:
      sampleData = getLeadsWithFilter(filterId, 1);
      break;
  }
  
  if (sampleData && sampleData.length > 0) {
    const sampleItem = sampleData[0];
    
    // Log filter ID and entity type
    Logger.log('===== DEBUG INFORMATION =====');
    Logger.log(`Entity Type: ${entityType}`);
    Logger.log(`Filter ID: ${filterId}`);
    Logger.log(`Sheet Name: ${sheetName}`);
    
    // Log complete raw deal data for inspection
    Logger.log(`\n===== COMPLETE RAW ${entityType.toUpperCase()} DATA =====`);
    Logger.log(JSON.stringify(sampleItem, null, 2));
    
    // Extract all fields including nested ones
    Logger.log('\n===== ALL AVAILABLE FIELDS =====');
    const allFields = {};
    
    // Recursive function to extract all fields with their paths
    function extractAllFields(obj, path = '') {
      if (!obj || typeof obj !== 'object') return;
      
      if (Array.isArray(obj)) {
        // For arrays, log the length and extract fields from first item if exists
        Logger.log(`${path} (Array with ${obj.length} items)`);
        if (obj.length > 0 && typeof obj[0] === 'object') {
          extractAllFields(obj[0], `${path}[0]`);
        }
      } else {
        // For objects, extract each property
        for (const key in obj) {
          const value = obj[key];
          const newPath = path ? `${path}.${key}` : key;
          
          if (value === null) {
            allFields[newPath] = 'null';
            continue;
          }
          
          const type = typeof value;
          
          if (type === 'object') {
            if (Array.isArray(value)) {
              allFields[newPath] = `array[${value.length}]`;
              Logger.log(`${newPath}: array[${value.length}]`);
              
              // Special case for custom fields with options
              if (key === 'options' && value.length > 0 && value[0] && value[0].label) {
                Logger.log(`  - Multiple options field with values: ${value.map(opt => opt.label).join(', ')}`);
              }
              
              // For small arrays with objects, recursively extract from the first item
              if (value.length > 0 && typeof value[0] === 'object') {
                extractAllFields(value[0], `${newPath}[0]`);
              }
            } else {
              allFields[newPath] = 'object';
              Logger.log(`${newPath}: object`);
              extractAllFields(value, newPath);
            }
          } else {
            allFields[newPath] = type;
            
            // Log a preview of the value unless it's a string longer than 50 chars
            const preview = type === 'string' && value.length > 50 
              ? value.substring(0, 50) + '...' 
              : value;
              
            Logger.log(`${newPath}: ${type} = ${preview}`);
          }
        }
      }
    }
    
    // Start extraction from the top level
    extractAllFields(sampleItem);
    
    // Specifically focus on custom fields section if it exists
    if (sampleItem.custom_fields) {
      Logger.log('\n===== CUSTOM FIELDS DETAIL =====');
      for (const key in sampleItem.custom_fields) {
        const field = sampleItem.custom_fields[key];
        const fieldType = typeof field;
        
        if (fieldType === 'object' && Array.isArray(field)) {
          Logger.log(`${key}: array[${field.length}]`);
          // Check if this is a multiple options field
          if (field.length > 0 && field[0] && field[0].label) {
            Logger.log(`  - Multiple options with values: ${field.map(opt => opt.label).join(', ')}`);
          }
        } else {
          const preview = fieldType === 'string' && field.length > 50 
            ? field.substring(0, 50) + '...' 
            : field;
          Logger.log(`${key}: ${fieldType} = ${preview}`);
        }
      }
    }
    
    // Count unique fields
    const fieldPaths = Object.keys(allFields).sort();
    Logger.log(`\nTotal unique fields found: ${fieldPaths.length}`);
    
    // Log all field paths in alphabetical order for easy lookup
    Logger.log('\n===== ALPHABETICAL LIST OF ALL FIELD PATHS =====');
    fieldPaths.forEach(path => {
      Logger.log(`${path}: ${allFields[path]}`);
    });
    
  } else {
    Logger.log(`No ${entityType} found with this filter. Please check the filter ID.`);
  }
}

/**
 * Gets the field definitions for a specific entity type
 */
function getPipedriverFields(entityType) {
  try {
    // Convert plural entity type to singular for fields endpoints
    let fieldsEndpoint;
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fieldsEndpoint = 'dealFields';
        break;
      case ENTITY_TYPES.PERSONS:
        fieldsEndpoint = 'personFields';
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fieldsEndpoint = 'organizationFields';
        break;
      case ENTITY_TYPES.ACTIVITIES:
        fieldsEndpoint = 'activityFields';
        break;
      case ENTITY_TYPES.LEADS:
        fieldsEndpoint = 'leadFields';
        break;
      default:
        Logger.log(`Unknown entity type: ${entityType}`);
        return [];
    }
    
    const url = `${getPipedriveApiUrl()}/${fieldsEndpoint}`;
    const responseData = makeAuthenticatedRequest(url);
    
    if (responseData.success) {
      return responseData.data;
    } else {
      Logger.log(`Failed to retrieve ${entityType} fields: ${responseData.error}`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error retrieving ${entityType} fields: ${error.message}`);
    return [];
  }
}

/**
 * Builds a mapping of field option IDs to their labels for a specific entity type
 */
function getFieldOptionMappingsForEntity(entityType) {
  // Get both standard and custom fields
  const standardFields = getPipedriverFields(entityType);
  const customFields = getCustomFieldsForEntity(entityType);
  
  // Combine all fields for processing
  const allFields = [...standardFields, ...customFields];
  
  const optionMappings = {};
  
  allFields.forEach(field => {
    // Only process fields with options (dropdown, multiple options, etc.)
    if (field.options && field.options.length > 0) {
      // Create a mapping for this field
      optionMappings[field.key] = {};
      
      // Map each option ID to its label
      field.options.forEach(option => {
        optionMappings[field.key][option.id] = option.label;
      });
      
      // Log for debugging
      Logger.log(`Field ${field.name} (${field.key}) has ${field.options.length} options`);
    }
  });
  
  return optionMappings;
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
 * Gets deal fields
 */
function getDealFields() {
  return getPipedriverFields(ENTITY_TYPES.DEALS);
}

/**
 * Gets person fields
 */
function getPersonFields() {
  return getPipedriverFields(ENTITY_TYPES.PERSONS);
}

/**
 * Gets organization fields
 */
function getOrganizationFields() {
  return getPipedriverFields(ENTITY_TYPES.ORGANIZATIONS);
}

/**
 * Gets activity fields
 */
function getActivityFields() {
  return getPipedriverFields(ENTITY_TYPES.ACTIVITIES);
}

/**
 * Gets lead fields
 */
function getLeadFields() {
  return getPipedriverFields(ENTITY_TYPES.LEADS);
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
 * Format the value for display in the spreadsheet
 */
function formatValue(value, columnPath, optionMappings = {}) {
  if (value === null || value === undefined) {
    return '';
  }
  
  // Handle comma-separated IDs for multiple options fields (custom fields)
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
    // Check if it's an array of option objects (multiple options field)
    if (Array.isArray(value) && value.length > 0 && value[0] && typeof value[0] === 'object' && value[0].label) {
      // It's a multiple options field with label property, extract and join labels
      return value.map(option => option.label).join(', ');
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
 * Gets a value from an object by path notation
 * Handles nested properties using dot notation (e.g., "creator_user.name")
 * Supports array indexing with numeric indices
 */
function getValueByPath(obj, path) {
  // If path is already an object with a key property, use that
  if (typeof path === 'object' && path.key) {
    path = path.key;
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
 * Format column names for better readability
 */
function formatColumnName(name) {
  // Convert snake_case to Title Case
  return name.split('_')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

/**
 * Gets custom field definitions for a specific entity type
 */
function getCustomFieldsForEntity(entityType) {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
  
  // Convert plural entity type to singular for fields endpoints
  let fieldsEndpoint;
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      fieldsEndpoint = 'dealFields';
      break;
    case ENTITY_TYPES.PERSONS:
      fieldsEndpoint = 'personFields';
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      fieldsEndpoint = 'organizationFields';
      break;
    case ENTITY_TYPES.ACTIVITIES:
      fieldsEndpoint = 'activityFields';
      break;
    case ENTITY_TYPES.LEADS:
      fieldsEndpoint = 'leadFields';
      break;
    default:
      Logger.log(`Unknown entity type: ${entityType}`);
      return [];
  }
  
  const url = `${getPipedriveApiUrl()}/${fieldsEndpoint}?api_token=${apiKey}&filter_type=custom_field`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      Logger.log(`Retrieved ${responseData.data.length} custom fields for ${entityType}`);
      return responseData.data;
    } else {
      Logger.log(`Failed to retrieve custom fields for ${entityType}: ${responseData.error}`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error retrieving custom fields for ${entityType}: ${error.message}`);
    return [];
  }
}

