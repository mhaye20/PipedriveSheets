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
      // Use proper OAuth authentication with Bearer token instead of api_token parameter
      const testUrl = `${getPipedriveApiUrl()}/users/me`;
      const response = UrlFetchApp.fetch(testUrl, {
        headers: {
          'Authorization': 'Bearer ' + accessToken
        },
        muteHttpExceptions: true
      });
      
      const statusCode = response.getResponseCode();
      const data = JSON.parse(response.getContentText());
      
      if (statusCode === 200 && data.success) {
        const ui = SpreadsheetApp.getUi();
        const result = ui.alert(
          'Already Connected',
          'You are already connected to Pipedrive as ' + data.data.name + '. Do you want to reconnect?',
          ui.ButtonSet.YES_NO
        );
        
        if (result === ui.Button.NO) {
          return;
        }
      } else {
        // Token is invalid, continue with auth
        Logger.log('Token validation failed: ' + data.error + ' (status code: ' + statusCode + ')');
      }
    } catch (e) {
      // Token is probably invalid, continue with auth
      Logger.log('Error checking token: ' + e.message);
    }
  }
  
  // Create the OAuth2 authorization URL
  const authUrl = `https://oauth.pipedrive.com/oauth/authorize?client_id=${PIPEDRIVE_CLIENT_ID}&redirect_uri=${REDIRECT_URI}&state=${generateRandomState()}&scope=base deals:read deals:write persons:read persons:write organizations:read organizations:write activities:read activities:write leads:read leads:write products:read products:write`;
  
  Logger.log(`Generated authorization URL with redirect URI: ${REDIRECT_URI}`);
  Logger.log(`Authorization URL: ${authUrl}`);
  
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
    Logger.log(`OAuth callback received with code: ${code ? 'present' : 'missing'}`);
    
    if (code) {
      try {
        // Create Authorization header with Base64 encoded client_id:client_secret
        const authHeader = "Basic " + Utilities.base64Encode(PIPEDRIVE_CLIENT_ID + ":" + PIPEDRIVE_CLIENT_SECRET);
        Logger.log(`Attempting to exchange authorization code for tokens using client ID: ${PIPEDRIVE_CLIENT_ID}`);
        
        // Exchange the authorization code for an access token
        const tokenResponse = UrlFetchApp.fetch('https://oauth.pipedrive.com/oauth/token', {
          method: 'post',
          headers: {
            'Authorization': authHeader,
            'Content-Type': 'application/x-www-form-urlencoded'
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
        Logger.log(`Token exchange response code: ${responseCode}`);
        Logger.log(`Token exchange response: ${responseText}`);
        
        if (responseCode !== 200) {
          throw new Error(`Token exchange failed with status ${responseCode}: ${responseText}`);
        }
        
        const tokenData = JSON.parse(responseText);
        
        if (!tokenData.access_token) {
          throw new Error(`No access token returned: ${responseText}`);
        }
        
        // Save the tokens in script properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
        scriptProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
        scriptProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
        
        // Save the API domain if provided
        if (tokenData.api_domain) {
          Logger.log(`API domain received: ${tokenData.api_domain}`);
          // Extract just the subdomain part
          const apiDomainMatch = tokenData.api_domain.match(/https:\/\/([^.]+)/);
          if (apiDomainMatch && apiDomainMatch[1]) {
            const subdomain = apiDomainMatch[1];
            Logger.log(`Setting subdomain to: ${subdomain}`);
            scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
          }
        } else {
          Logger.log(`No API domain in token response, querying user info`);
          // Get user info to determine the subdomain if not provided in token response
          const userResponse = UrlFetchApp.fetch('https://api.pipedrive.com/v1/users/me', {
            headers: {
              'Authorization': 'Bearer ' + tokenData.access_token
            },
            muteHttpExceptions: true
          });
          
          const userStatusCode = userResponse.getResponseCode();
          const userResponseText = userResponse.getContentText();
          Logger.log(`User info response code: ${userStatusCode}`);
          
          if (userStatusCode === 200) {
            const userData = JSON.parse(userResponseText);
            
            if (userData.success) {
              // Extract the company domain from the user data
              const companyDomain = userData.data.company_domain;
              Logger.log(`Setting subdomain to: ${companyDomain}`);
              scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', companyDomain);
            } else {
              Logger.log(`User info request wasn't successful: ${userResponseText}`);
            }
          } else {
            Logger.log(`Failed to get user info: ${userResponseText}`);
          }
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
    Logger.log('Cannot refresh token: missing access token or refresh token');
    return false;
  }
  
  // Check if token is expired or about to expire (5 minutes buffer)
  const now = new Date().getTime();
  if (!expiresAt || now > (parseInt(expiresAt) - (5 * 60 * 1000))) {
    Logger.log('Token expired or about to expire, refreshing...');
    try {
      // Create Authorization header with Base64 encoded client_id:client_secret
      const authHeader = "Basic " + Utilities.base64Encode(PIPEDRIVE_CLIENT_ID + ":" + PIPEDRIVE_CLIENT_SECRET);
      Logger.log(`Refreshing token using client ID: ${PIPEDRIVE_CLIENT_ID}`);
      
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
      Logger.log(`Token refresh response code: ${responseCode}`);
      
      if (responseCode !== 200) {
        Logger.log(`Token refresh failed: ${responseText}`);
        return false;
      }
      
      const tokenData = JSON.parse(responseText);
      
      if (!tokenData.access_token) {
        Logger.log(`No access token in refresh response: ${responseText}`);
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
          Logger.log(`Setting subdomain from refresh to: ${subdomain}`);
          scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
        }
      }
      
      Logger.log('Token refreshed successfully');
      return true;
    } catch (error) {
      Logger.log('Token refresh error: ' + error.message);
      return false;
    }
  }
  
  // Token is still valid
  return true;
}
