If you would like to use an OAuth access token for making API calls, then make sure the API key described in the previous section is not set or is set to an empty string. If both API token and OAuth access token are set, then the API token takes precedence.

To set up authentication in the API client, you need the following information. You can receive the necessary client tokens through a Sandbox account (get it here) and generate the tokens (detailed steps here).

Parameter	Description
clientId	OAuth 2 Client ID
clientSecret	OAuth 2 Client Secret
redirectUri	OAuth 2 Redirection endpoint or Callback Uri
Next, initialize the API client as follows:

import { OAuth2Configuration, Configuration } from 'pipedrive/v1';

// Configuration parameters and credentials
const oauth2 = new OAuth2Configuration({
  clientId: "clientId", // OAuth 2 Client ID
  clientSecret: "clientSecret",  // OAuth 2 Client Secret
  redirectUri: 'redirectUri' // OAuth 2 Redirection endpoint or Callback Uri
});

const apiConfig = new Configuration({
    accessToken: oauth2.getAccessToken,
    basePath: oauth2.basePath,
});
You must now authorize the client.

Authorizing your client
Your application must obtain user authorization before it can execute an endpoint call. The SDK uses OAuth 2.0 authorization to obtain a user's consent to perform an API request on the user's behalf. Details about how the OAuth2.0 flow works in Pipedrive, how long tokens are valid, and more, can be found here.

1. Obtaining user consent
To obtain user's consent, you must redirect the user to the authorization page. The authorizationUrl returns the URL to the authorization page.

// open up the authUrl in the browser
const authUrl = oauth2.authorizationUrl;
2. Handle the OAuth server response
Once the user responds to the consent request, the OAuth 2.0 server responds to your application's access request by using the URL specified in the request.

If the user approves the request, the authorization code will be sent as the code query string:

https://example.com/oauth/callback?code=XXXXXXXXXXXXXXXXXXXXXXXXX
If the user does not approve the request, the response contains an error query string:

https://example.com/oauth/callback?error=access_denied
3. Authorize the client using the code
After the server receives the code, it can exchange this for an access token. The access token is an object containing information for authorizing the client and refreshing the token itself. In the API client all the access token fields are held separately in the OAuth2Configuration class. Additionally access token expiration time as an OAuth2Configuration.expiresAt field is calculated. It is measured in the number of milliseconds elapsed since January 1, 1970 00:00:00 UTC.

const token = await oauth2.authorize(code);
The Node.js SDK supports only promises. So, the authorize call returns a promise.

Refreshing token
Access tokens may expire after sometime, if it necessary you can do it manually.

const newToken = await oauth2.tokenRefresh();
If the access token expires, the SDK will attempt to automatically refresh it before the next endpoint call which requires authentication.