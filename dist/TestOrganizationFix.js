/**
 * Quick test for organization updates
 * Run this from Google Apps Script to test the fix
 */
function testOrganizationUpdateFix() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
  const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN");
  
  if (!accessToken || !subdomain) {
    Logger.log("ERROR: Missing credentials");
    return;
  }
  
  const basePath = `https://${subdomain}.pipedrive.com/v1`;
  const orgId = 1; // Your test organization ID
  
  // Test 1: Simple update with custom fields at root
  Logger.log("\n=== Test 1: Custom fields at root ===");
  const test1Payload = {
    "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
    "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
    "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29",
    "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
  };
  
  try {
    const response1 = makeDirectApiCall(orgId, test1Payload, accessToken, basePath);
    Logger.log(`Response: ${JSON.stringify(response1)}`);
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
  }
  
  // Test 2: Using updateOrganizationDirect with custom_fields wrapper
  Logger.log("\n=== Test 2: Using updateOrganizationDirect ===");
  const test2Payload = {
    custom_fields: {
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
    }
  };
  
  try {
    const response2 = updateOrganizationDirect(orgId, test2Payload, accessToken, basePath, null);
    Logger.log(`Response: ${JSON.stringify(response2)}`);
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
  }
}

function makeDirectApiCall(orgId, payload, accessToken, basePath) {
  const url = `${basePath}/organizations/${orgId}`;
  
  const options = {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };
  
  Logger.log(`URL: ${url}`);
  Logger.log(`Payload: ${JSON.stringify(payload)}`);
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseData = JSON.parse(response.getContentText());
  
  Logger.log(`Response code: ${responseCode}`);
  
  return responseData;
}