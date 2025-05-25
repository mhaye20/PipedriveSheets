/**
 * Test functions for Pipedrive field updates
 * This module provides test functions to verify different field types work correctly
 */

/**
 * Test updating various field types for each entity type
 * Run this function from Google Apps Script editor to test field updates
 */
function testPipedriveFieldUpdates() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
  const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN");
  
  if (!accessToken || !subdomain) {
    Logger.log("ERROR: Missing Pipedrive credentials. Please ensure PIPEDRIVE_ACCESS_TOKEN and PIPEDRIVE_SUBDOMAIN are set.");
    return;
  }
  
  const basePath = `https://${subdomain}.pipedrive.com/v1`;
  
  // Test payloads for different field types
  const testCases = {
    organizations: {
      entityId: 1, // Replace with a test organization ID
      testPayloads: [
        {
          name: "Test 1: Simple text field update",
          payload: {
            name: "Test Organization Updated"
          }
        },
        {
          name: "Test 2: Custom field at root level",
          payload: {
            "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14"
          }
        },
        {
          name: "Test 3: Time range fields at root level",
          payload: {
            "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
            "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
          }
        },
        {
          name: "Test 4: Mixed standard and custom fields",
          payload: {
            name: "Test Organization Mixed",
            "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14"
          }
        }
      ]
    },
    persons: {
      entityId: 1, // Replace with a test person ID
      testPayloads: [
        {
          name: "Test 1: Simple name update",
          payload: {
            name: "Test Person Updated"
          }
        },
        {
          name: "Test 2: Custom fields",
          payload: {
            name: "Test Person Custom",
            // Add your person custom field IDs here
          }
        }
      ]
    },
    deals: {
      entityId: 1, // Replace with a test deal ID
      testPayloads: [
        {
          name: "Test 1: Simple title update",
          payload: {
            title: "Test Deal Updated"
          }
        },
        {
          name: "Test 2: Custom fields with wrapper",
          payload: {
            title: "Test Deal Custom",
            custom_fields: {
              // Add your deal custom field IDs here
            }
          }
        }
      ]
    }
  };
  
  // Run tests for each entity type
  for (const entityType in testCases) {
    Logger.log(`\n========== Testing ${entityType} ==========`);
    const tests = testCases[entityType];
    
    for (const test of tests.testPayloads) {
      Logger.log(`\nRunning: ${test.name}`);
      Logger.log(`Payload: ${JSON.stringify(test.payload)}`);
      
      try {
        let response;
        
        switch (entityType) {
          case "organizations":
            response = testOrganizationUpdate(tests.entityId, test.payload, accessToken, basePath);
            break;
          case "persons":
            response = testPersonUpdate(tests.entityId, test.payload, accessToken, basePath);
            break;
          case "deals":
            response = testDealUpdate(tests.entityId, test.payload, accessToken, basePath);
            break;
        }
        
        if (response.success) {
          Logger.log(`✓ SUCCESS: ${test.name}`);
          Logger.log(`Response: ${JSON.stringify(response.data).substring(0, 200)}...`);
        } else {
          Logger.log(`✗ FAILED: ${test.name}`);
          Logger.log(`Error: ${response.error}`);
          Logger.log(`Error info: ${response.error_info}`);
        }
      } catch (error) {
        Logger.log(`✗ EXCEPTION: ${test.name}`);
        Logger.log(`Error: ${error.message}`);
      }
    }
  }
}

/**
 * Test organization update with different payload structures
 */
function testOrganizationUpdate(orgId, payload, accessToken, basePath) {
  const orgUrl = `${basePath}/organizations/${orgId}`;
  
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
  
  const response = UrlFetchApp.fetch(orgUrl, options);
  const responseData = JSON.parse(response.getContentText());
  
  return responseData;
}

/**
 * Test person update
 */
function testPersonUpdate(personId, payload, accessToken, basePath) {
  const personUrl = `${basePath}/persons/${personId}`;
  
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
  
  const response = UrlFetchApp.fetch(personUrl, options);
  const responseData = JSON.parse(response.getContentText());
  
  return responseData;
}

/**
 * Test deal update
 */
function testDealUpdate(dealId, payload, accessToken, basePath) {
  const dealUrl = `${basePath}/deals/${dealId}`;
  
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
  
  const response = UrlFetchApp.fetch(dealUrl, options);
  const responseData = JSON.parse(response.getContentText());
  
  return responseData;
}

/**
 * Get current field values for an organization to understand the structure
 */
function inspectOrganizationStructure() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
  const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN");
  const orgId = 1; // Replace with your test organization ID
  
  const orgUrl = `https://${subdomain}.pipedrive.com/v1/organizations/${orgId}`;
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(orgUrl, options);
  const responseData = JSON.parse(response.getContentText());
  
  Logger.log("Organization structure:");
  Logger.log(JSON.stringify(responseData, null, 2));
  
  return responseData;
}