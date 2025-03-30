/**
 * OAuth Handler
 * 
 * This module handles OAuth authentication with Pipedrive:
 * - OAuth flow for token retrieval
 * - Token storage and refresh
 * - Authorization validation
 */

/**
 * Starts the OAuth flow with Pipedrive
 */
function startOAuth() {
  try {
    // Define OAuth parameters
    const oauthConfig = {
      clientId: PIPEDRIVE_CLIENT_ID,
      clientSecret: PIPEDRIVE_CLIENT_SECRET,
      authUrl: 'https://oauth.pipedrive.com/oauth/authorize',
      tokenUrl: 'https://oauth.pipedrive.com/oauth/token',
      scope: 'contacts:read deals:read',
      callbackFunction: 'processOAuthResponse'
    };
    
    // Create the authorization URL
    const authorizationUrl = oauthConfig.authUrl + 
      '?client_id=' + oauthConfig.clientId +
      '&redirect_uri=' + encodeURIComponent(ScriptApp.getService().getUrl()) +
      '&scope=' + encodeURIComponent(oauthConfig.scope) +
      '&response_type=code' +
      '&state=' + encodeURIComponent(Utilities.getUuid());
    
    // Store state in user properties for later verification
    PropertiesService.getUserProperties().setProperty('OAUTH_STATE', state);
    
    // Display the authorization dialog
    const template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize with Pipedrive</a>'
    );
    template.authorizationUrl = authorizationUrl;
    
    const htmlOutput = template.evaluate()
      .setWidth(600)
      .setHeight(400)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Authorize with Pipedrive');
  } catch (e) {
    Logger.log('Error in startOAuth: ' + e.message);
    SpreadsheetApp.getUi().alert('Authorization Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Processes the OAuth response from Pipedrive
 * @param {Object} response - The OAuth response
 * @return {boolean} Success status
 */
function processOAuthResponse(response) {
  try {
    // Validate state parameter
    const storedState = PropertiesService.getUserProperties().getProperty('OAUTH_STATE');
    if (response.state !== storedState) {
      throw new Error('Invalid state parameter. Request may have been tampered with.');
    }
    
    // Exchange authorization code for tokens
    const tokenResponse = UrlFetchApp.fetch(
      'https://oauth.pipedrive.com/oauth/token',
      {
        method: 'post',
        payload: {
          client_id: PIPEDRIVE_CLIENT_ID,
          client_secret: PIPEDRIVE_CLIENT_SECRET,
          grant_type: 'authorization_code',
          code: response.code,
          redirect_uri: ScriptApp.getService().getUrl()
        },
        muteHttpExceptions: true
      }
    );
    
    const tokenData = JSON.parse(tokenResponse.getContentText());
    
    if (tokenData.error) {
      throw new Error('Error getting access token: ' + tokenData.error_description);
    }
    
    // Store the tokens
    PropertiesService.getUserProperties().setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
    PropertiesService.getUserProperties().setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token);
    PropertiesService.getUserProperties().setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
    
    return true;
  } catch (e) {
    Logger.log('Error in processOAuthResponse: ' + e.message);
    return false;
  }
}

/**
 * Refreshes the Pipedrive access token if expired
 * @return {string} Valid access token or null if unable to refresh
 */
function getValidAccessToken() {
  try {
    // Get current tokens
    const userProperties = PropertiesService.getUserProperties();
    const accessToken = userProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    const refreshToken = userProperties.getProperty('PIPEDRIVE_REFRESH_TOKEN');
    const expiresAt = userProperties.getProperty('PIPEDRIVE_TOKEN_EXPIRES');
    
    // If we don't have tokens yet, return null
    if (!accessToken || !refreshToken) {
      return null;
    }
    
    // Check if token is expired
    const now = new Date().getTime();
    if (now < expiresAt - 60000) { // 1 minute buffer
      // Token is still valid
      return accessToken;
    }
    
    // Token is expired, attempt to refresh
    const response = UrlFetchApp.fetch(
      'https://oauth.pipedrive.com/oauth/token',
      {
        method: 'post',
        payload: {
          client_id: PIPEDRIVE_CLIENT_ID,
          client_secret: PIPEDRIVE_CLIENT_SECRET,
          grant_type: 'refresh_token',
          refresh_token: refreshToken
        },
        muteHttpExceptions: true
      }
    );
    
    const tokenData = JSON.parse(response.getContentText());
    
    if (tokenData.error) {
      throw new Error('Error refreshing token: ' + tokenData.error_description);
    }
    
    // Store the new tokens
    userProperties.setProperty('PIPEDRIVE_ACCESS_TOKEN', tokenData.access_token);
    userProperties.setProperty('PIPEDRIVE_REFRESH_TOKEN', tokenData.refresh_token || refreshToken);
    userProperties.setProperty('PIPEDRIVE_TOKEN_EXPIRES', new Date().getTime() + (tokenData.expires_in * 1000));
    
    return tokenData.access_token;
  } catch (e) {
    Logger.log('Error in getValidAccessToken: ' + e.message);
    return null;
  }
}
