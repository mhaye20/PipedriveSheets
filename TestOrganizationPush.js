/**
 * Test the specific organization push scenario from the user
 */
function testOrganizationPushScenario() {
  const config = getTestConfig();
  if (!config) {
    Logger.log("ERROR: Missing configuration");
    return;
  }
  
  // The exact payload from the user's error log
  const testPayload = {
    custom_fields: {
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
    }
  };
  
  Logger.log("Testing organization update with exact user payload...");
  Logger.log(`Payload: ${JSON.stringify(testPayload)}`);
  
  const result = updateOrganizationDirect(
    1, // Organization ID
    testPayload,
    config.accessToken,
    config.basePath,
    null
  );
  
  Logger.log(`\nResult: ${JSON.stringify(result)}`);
  
  if (result.success) {
    Logger.log("\n✅ SUCCESS! Organization update is now working correctly.");
    Logger.log("\nThe fix successfully:");
    Logger.log("1. Converts custom_fields wrapper to root level fields");
    Logger.log("2. Properly formats date and time fields");
    Logger.log("3. Handles time range fields with _until suffix");
    
    // Show what fields were updated
    if (result.data) {
      Logger.log("\nUpdated fields in response:");
      Logger.log(`- Date range: ${result.data["6fefa2fd421fe09f9f92571cfebca4caff8274b7"]} to ${result.data["6fefa2fd421fe09f9f92571cfebca4caff8274b7_until"]}`);
      Logger.log(`- Time range: ${result.data["4ae9767e7c7b836e8ef1049b6a2728b4d8207abb"]} to ${result.data["4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until"]}`);
    }
  } else {
    Logger.log("\n❌ FAILED! There's still an issue with organization updates.");
    Logger.log(`Error: ${result.error}`);
    Logger.log(`Error info: ${result.error_info}`);
  }
  
  return result;
}

/**
 * Test direct update bypassing the sync service
 */
function testDirectOrganizationUpdate() {
  const config = getTestConfig();
  if (!config) return;
  
  // Test with fields at root level (how Pipedrive v1 expects them)
  const directPayload = {
    name: "Test Direct Update " + new Date().getTime(),
    "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
    "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
    "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29",
    "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
  };
  
  const url = `${config.basePath}/organizations/1`;
  
  const options = {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${config.accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(directPayload)
  };
  
  Logger.log("Testing direct API call with fields at root level...");
  Logger.log(`URL: ${url}`);
  Logger.log(`Payload: ${JSON.stringify(directPayload)}`);
  
  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());
  
  Logger.log(`Response: ${JSON.stringify(responseData)}`);
  
  return responseData;
}

/**
 * Get test configuration
 */
function getTestConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
  const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN");
  
  if (!accessToken || !subdomain) {
    Logger.log("ERROR: Missing Pipedrive credentials");
    return null;
  }
  
  return {
    accessToken,
    subdomain,
    basePath: `https://${subdomain}.pipedrive.com/v1`
  };
}