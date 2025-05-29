/**
 * OAuth Handler
 * 
 * This module handles OAuth authentication with Pipedrive:
 * - OAuth flow for token retrieval
 * - Token storage and refresh
 * - Authorization validation
 */

// Redirect URI for OAuth flow - Important: Use the web app URL for OAuth
// Important: If you're having authentication issues, create a new deployment and use that URL here
const REDIRECT_URI = 'https://script.google.com/macros/s/AKfycbwFP346JXePDSleJZwhLbF4SbgnNqY81H5PRWsHWSml5LnzLPNPgmBuRVqfRY_fJYQ1bw/exec?page=oauthCallback';

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
      const testUrl = `${getPipedriveApiUrl()}/users/me`;
      const response = makeAuthenticatedRequest(testUrl);
      
      if (response && response.success) {
        const ui = SpreadsheetApp.getUi();
        const result = ui.alert(
          'Already Connected',
          'You are already connected to Pipedrive as ' + response.data.name + '. Do you want to reconnect?',
          ui.ButtonSet.YES_NO
        );
        
        if (result === ui.Button.NO) {
          return;
        }
        
        // User wants to reconnect, clear existing tokens
        scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
      }
    } catch (e) {
      // Token is invalid, continue with auth
      // Clear any existing tokens
      scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
      scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
      scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
    }
  }
  
  // Create the OAuth2 authorization URL
  const state = generateRandomState();
  const scopes = [
    'base',
    'deals:read', 'deals:write',
    'persons:read', 'persons:write',
    'organizations:read', 'organizations:write',
    'activities:read', 'activities:write',
    'leads:read', 'leads:write',
    'products:read', 'products:write'
  ].join(' ');
  
  const authUrl = `https://oauth.pipedrive.com/oauth/authorize?client_id=${PIPEDRIVE_CLIENT_ID}&redirect_uri=${REDIRECT_URI}&state=${state}&scope=${encodeURIComponent(scopes)}`;
  
  
  // Save state for validation in callback
  scriptProperties.setProperty('OAUTH_STATE', state);
  
  // Display the authorization dialog
  const template = HtmlService.createTemplate(
    '<html>'
    + '<head>'
    + '<style>'
    + 'body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }'
    + 'button { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; transition: background-color 0.3s; }'
    + 'button:hover { background-color: #45a049; }'
    + '.container { text-align: center; padding: 20px; }'
    + 'h2 { color: #333; margin-bottom: 16px; }'
    + 'p { color: #666; margin-bottom: 24px; }'
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
    const state = e.parameter.state;
    
    // Validate state to prevent CSRF attacks
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedState = scriptProperties.getProperty('OAUTH_STATE');
    scriptProperties.deleteProperty('OAUTH_STATE'); // Clear state after use
    
    if (!state || state !== savedState) {
      return HtmlService.createHtmlOutput(
        '<html>'
        + '<head>'
        + '<style>'
        + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
        + '.error { color: #f44336; font-size: 24px; }'
        + '</style>'
        + '</head>'
        + '<body>'
        + '<h1 class="error">Security Error</h1>'
        + '<p>Invalid state parameter. This could be a security issue or the authorization process was interrupted.</p>'
        + '<p>Please try again.</p>'
        + '</body>'
        + '</html>'
      )
      .setTitle('Security Error');
    }
    
    if (code) {
      try {
        // Create Authorization header with Base64 encoded client_id:client_secret
        const authHeader = "Basic " + Utilities.base64Encode(PIPEDRIVE_CLIENT_ID + ":" + PIPEDRIVE_CLIENT_SECRET);
        
        // Exchange the authorization code for an access token
        const tokenResponse = UrlFetchApp.fetch('https://oauth.pipedrive.com/oauth/token', {
          method: 'post',
          headers: {
            'Authorization': authHeader,
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json'
          },
          payload: {
            grant_type: 'authorization_code',
            code: code,
            redirect_uri: REDIRECT_URI
          },
          muteHttpExceptions: true
        });
        
        const responseCode = tokenResponse.getResponseCode();
        const responseText = tokenResponse.getContentText();
        
        if (responseCode !== 200) {
          throw new Error(`Token exchange failed with status ${responseCode}: ${responseText}`);
        }
        
        let tokenData;
        try {
          tokenData = JSON.parse(responseText);
        } catch (parseError) {
          throw new Error('Invalid response from Pipedrive OAuth server');
        }
        
        if (!tokenData.access_token) {
          throw new Error(`No access token returned: ${responseText}`);
        }
        
        // Clear any existing tokens first
        scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
        scriptProperties.deleteProperty('PIPEDRIVE_SUBDOMAIN');
        
        // Save the new tokens
        scriptProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
        scriptProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
        scriptProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
        
        // Save the API domain if provided
        if (tokenData.api_domain) {
          // Extract just the subdomain part
          const apiDomainMatch = tokenData.api_domain.match(/https:\/\/([^.]+)/);
          if (apiDomainMatch && apiDomainMatch[1]) {
            const subdomain = apiDomainMatch[1];
            scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
          }
        }
        
        // Get user info to verify connection and get company domain if needed
        try {
          const userResponse = makeAuthenticatedRequest(`${getPipedriveApiUrl()}/users/me`);
          if (userResponse && userResponse.success && userResponse.data) {
            const userData = userResponse.data;
            if (!scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') && userData.company_domain) {
              scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', userData.company_domain);
            }
          }
        } catch (userError) {
          // Non-fatal error, continue with success page
        }
        
        // Display success page
        return HtmlService.createHtmlOutput(
          '<html>'
          + '<head>'
          + '<style>'
          + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
          + '.success { color: #4CAF50; font-size: 24px; }'
          + 'button { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; margin-top: 20px; }'
          + 'button:hover { background-color: #45a049; }'
          + '</style>'
          + '</head>'
          + '<body>'
          + '<h1 class="success">✓ Successfully Connected!</h1>'
          + '<p>You have successfully connected your Pipedrive account. You can close this window and return to Google Sheets™.</p>'
          + '<button onclick="window.close()">Close Window</button>'
          + '</body>'
          + '</html>'
        )
        .setTitle('Connected to Pipedrive');
      } catch (error) {
        // Handle authorization error
        return HtmlService.createHtmlOutput(
          '<html>'
          + '<head>'
          + '<style>'
          + 'body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }'
          + '.error { color: #f44336; font-size: 24px; }'
          + 'pre { background: #f5f5f5; padding: 10px; border-radius: 4px; text-align: left; overflow-x: auto; }'
          + '</style>'
          + '</head>'
          + '<body>'
          + '<h1 class="error">Error Connecting</h1>'
          + '<p>There was an error connecting to Pipedrive:</p>'
          + '<pre>' + error.message + '</pre>'
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
        + '<p>Error details: ' + (e.parameter.error || 'No error details provided') + '</p>'
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
 * @return {boolean} True if token is valid or was refreshed successfully, false otherwise
 */
function refreshAccessTokenIfNeeded() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  const refreshToken = scriptProperties.getProperty('PIPEDRIVE_REFRESH_TOKEN');
  const expiresAt = scriptProperties.getProperty('PIPEDRIVE_TOKEN_EXPIRES');
  
  // If no token or refresh token, we can't refresh
  if (!accessToken || !refreshToken) {
    // Clear any existing tokens to force re-authentication
    scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
    scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
    scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
    return false;
  }
  
  // Check if token is expired or about to expire (5 minutes buffer)
  const now = new Date().getTime();
  if (!expiresAt || now > (parseInt(expiresAt) - (5 * 60 * 1000))) {
    try {
      // Create Authorization header with Base64 encoded client_id:client_secret
      const authHeader = "Basic " + Utilities.base64Encode(PIPEDRIVE_CLIENT_ID + ":" + PIPEDRIVE_CLIENT_SECRET);
      
      // Refresh the token
      const tokenResponse = UrlFetchApp.fetch('https://oauth.pipedrive.com/oauth/token', {
        method: 'post',
        headers: {
          'Authorization': authHeader,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
          grant_type: 'refresh_token',
          refresh_token: refreshToken
        },
        muteHttpExceptions: true
      });
      
      const responseCode = tokenResponse.getResponseCode();
      const responseText = tokenResponse.getContentText();
      
      if (responseCode !== 200) {
        // Clear tokens to force re-authentication
        scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
        return false;
      }
      
      const tokenData = JSON.parse(responseText);
      
      if (!tokenData.access_token) {
        // Clear tokens to force re-authentication
        scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
        scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
        return false;
      }
      
      // Save the new tokens
      scriptProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
      if (tokenData.refresh_token) {
        scriptProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
      }
      scriptProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
      
      // Save the API domain if provided
      if (tokenData.api_domain) {
        const apiDomainMatch = tokenData.api_domain.match(/https:\/\/([^.]+)/);
        if (apiDomainMatch && apiDomainMatch[1]) {
          const subdomain = apiDomainMatch[1];
          scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
        }
      }
      
      return true;
    } catch (error) {
      // Clear tokens to force re-authentication
      scriptProperties.deleteProperty('PIPEDRIVE_ACCESS_TOKEN');
      scriptProperties.deleteProperty('PIPEDRIVE_REFRESH_TOKEN');
      scriptProperties.deleteProperty('PIPEDRIVE_TOKEN_EXPIRES');
      return false;
    }
  }
  
  // Token is still valid
  return true;
}
