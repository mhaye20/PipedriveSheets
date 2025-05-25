/**
 * Comprehensive field testing for all Pipedrive entity types
 * This tests different field types to ensure proper API integration
 */

/**
 * Main test runner - tests all entity types with various field types
 */
function runComprehensiveFieldTests() {
  const results = {
    organizations: testOrganizationFields(),
    persons: testPersonFields(),
    deals: testDealFields(),
    activities: testActivityFields(),
    leads: testLeadFields(),
    products: testProductFields()
  };
  
  // Generate summary report
  Logger.log("\n========== TEST SUMMARY ==========");
  for (const entityType in results) {
    const entityResults = results[entityType];
    const passed = entityResults.filter(r => r.success).length;
    const total = entityResults.length;
    Logger.log(`${entityType}: ${passed}/${total} tests passed`);
  }
  
  return results;
}

/**
 * Test organization field updates
 */
function testOrganizationFields() {
  Logger.log("\n========== TESTING ORGANIZATIONS ==========");
  const results = [];
  
  // Get test configuration
  const config = getTestConfig();
  if (!config) return [];
  
  const testOrgId = 1; // Replace with your test organization ID
  
  // Test 1: Standard field update
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Standard field (name)',
    { name: "Test Organization " + new Date().getTime() },
    config
  ));
  
  // Test 2: Single custom field - SKIP if no valid field ID
  // Uncomment and add a valid custom field ID to test
  /*
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Single custom text field',
    { 
      custom_fields: {
        // Add your custom field ID here
        "valid_field_id_here": "Test Value " + new Date().getTime()
      }
    },
    config
  ));
  */
  
  // Test 3: Date field
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Date field',
    {
      custom_fields: {
        "6fefa2fd421fe09f9f92571cfebca4caff8274b7": new Date().toISOString().split('T')[0]
      }
    },
    config
  ));
  
  // Test 4: Time field
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Time field',
    {
      custom_fields: {
        "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "14:30:00"
      }
    },
    config
  ));
  
  // Test 5: Date range fields
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Date range fields',
    {
      custom_fields: {
        "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
        "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29"
      }
    },
    config
  ));
  
  // Test 6: Time range fields
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Time range fields',
    {
      custom_fields: {
        "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
        "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
      }
    },
    config
  ));
  
  // Test 7: Mixed fields
  results.push(testFieldUpdate(
    'organizations',
    testOrgId,
    'Mixed standard and custom fields',
    {
      name: "Test Org Mixed " + new Date().getTime(),
      custom_fields: {
        "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
        "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "14:30:00"
      }
    },
    config
  ));
  
  return results;
}

/**
 * Test person field updates
 */
function testPersonFields() {
  Logger.log("\n========== TESTING PERSONS ==========");
  const results = [];
  
  const config = getTestConfig();
  if (!config) return [];
  
  const testPersonId = 1; // Replace with your test person ID
  
  // Test 1: Standard field
  results.push(testFieldUpdate(
    'persons',
    testPersonId,
    'Standard field (name)',
    { name: "Test Person " + new Date().getTime() },
    config
  ));
  
  // Test 2: Email field
  results.push(testFieldUpdate(
    'persons',
    testPersonId,
    'Email field',
    { 
      email: [{
        value: "test" + new Date().getTime() + "@example.com",
        primary: true,
        label: "work"
      }]
    },
    config
  ));
  
  // Test 3: Phone field
  results.push(testFieldUpdate(
    'persons',
    testPersonId,
    'Phone field',
    {
      phone: [{
        value: "+1234567890",
        primary: true,
        label: "work"
      }]
    },
    config
  ));
  
  // Add more person-specific tests...
  
  return results;
}

/**
 * Test deal field updates
 */
function testDealFields() {
  Logger.log("\n========== TESTING DEALS ==========");
  const results = [];
  
  const config = getTestConfig();
  if (!config) return [];
  
  const testDealId = 1; // Replace with your test deal ID
  
  // Test 1: Standard field
  results.push(testFieldUpdate(
    'deals',
    testDealId,
    'Standard field (title)',
    { title: "Test Deal " + new Date().getTime() },
    config
  ));
  
  // Test 2: Value field
  results.push(testFieldUpdate(
    'deals',
    testDealId,
    'Value field',
    { value: Math.floor(Math.random() * 10000) },
    config
  ));
  
  // Test 3: Custom fields with wrapper - SKIP if empty
  // Uncomment and add valid custom field IDs to test
  /*
  results.push(testFieldUpdate(
    'deals',
    testDealId,
    'Custom fields in wrapper',
    {
      custom_fields: {
        // Add your deal custom field IDs
        "field_id": "value"
      }
    },
    config
  ));
  */
  
  return results;
}

/**
 * Test activity field updates
 */
function testActivityFields() {
  Logger.log("\n========== TESTING ACTIVITIES ==========");
  const results = [];
  
  const config = getTestConfig();
  if (!config) return [];
  
  const testActivityId = 1; // Replace with your test activity ID
  
  // Test 1: Subject update
  results.push(testFieldUpdate(
    'activities',
    testActivityId,
    'Subject field',
    { subject: "Test Activity " + new Date().getTime() },
    config
  ));
  
  // Test 2: Due date
  results.push(testFieldUpdate(
    'activities',
    testActivityId,
    'Due date',
    { due_date: new Date().toISOString().split('T')[0] },
    config
  ));
  
  return results;
}

/**
 * Test lead field updates
 */
function testLeadFields() {
  Logger.log("\n========== TESTING LEADS ==========");
  const results = [];
  
  const config = getTestConfig();
  if (!config) return [];
  
  const testLeadId = "abc-123"; // Replace with your test lead ID (string)
  
  // Test 1: Title update
  results.push(testFieldUpdate(
    'leads',
    testLeadId,
    'Title field',
    { title: "Test Lead " + new Date().getTime() },
    config
  ));
  
  return results;
}

/**
 * Test product field updates
 */
function testProductFields() {
  Logger.log("\n========== TESTING PRODUCTS ==========");
  const results = [];
  
  const config = getTestConfig();
  if (!config) return [];
  
  const testProductId = 1; // Replace with your test product ID
  
  // Test 1: Name update
  results.push(testFieldUpdate(
    'products',
    testProductId,
    'Name field',
    { name: "Test Product " + new Date().getTime() },
    config
  ));
  
  // Test 2: Price update
  results.push(testFieldUpdate(
    'products',
    testProductId,
    'Prices field',
    {
      prices: [{
        price: Math.floor(Math.random() * 1000),
        currency: "USD",
        cost: 0
      }]
    },
    config
  ));
  
  return results;
}

/**
 * Helper function to test a single field update
 */
function testFieldUpdate(entityType, entityId, testName, payload, config) {
  Logger.log(`\nTest: ${testName}`);
  Logger.log(`Payload: ${JSON.stringify(payload)}`);
  
  try {
    let response;
    
    // Use appropriate update function based on entity type
    switch (entityType) {
      case 'organizations':
        response = updateOrganizationDirect(
          entityId, 
          payload, 
          config.accessToken, 
          config.basePath,
          null
        );
        break;
        
      case 'persons':
        response = updatePersonDirect(
          entityId,
          payload,
          config.accessToken,
          config.basePath
        );
        break;
        
      case 'deals':
        response = updateDealDirect(
          entityId,
          payload,
          config.accessToken,
          config.basePath,
          null
        );
        break;
        
      case 'activities':
      case 'leads':
      case 'products':
        // For these, use direct API call
        response = makeDirectApiCall(
          entityType,
          entityId,
          payload,
          config.accessToken,
          config.basePath
        );
        break;
        
      default:
        throw new Error(`Unknown entity type: ${entityType}`);
    }
    
    if (response.success) {
      Logger.log(`✓ PASSED: ${testName}`);
      return { testName, success: true };
    } else {
      Logger.log(`✗ FAILED: ${testName}`);
      Logger.log(`Error: ${response.error}`);
      return { testName, success: false, error: response.error };
    }
  } catch (error) {
    Logger.log(`✗ EXCEPTION: ${testName}`);
    Logger.log(`Error: ${error.message}`);
    return { testName, success: false, error: error.message };
  }
}

/**
 * Make direct API call for entity types without specific functions
 */
function makeDirectApiCall(entityType, entityId, payload, accessToken, basePath) {
  // Special handling for leads which use PATCH instead of PUT
  const method = entityType === 'leads' ? 'PATCH' : 'PUT';
  
  // Construct the URL correctly
  const url = `${basePath}/${entityType}/${entityId}`;
  
  Logger.log(`Direct API call to: ${url}`);
  Logger.log(`Method: ${method}`);
  Logger.log(`Payload: ${JSON.stringify(payload)}`);
  
  const options = {
    method: method,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());
  
  return responseData;
}

/**
 * Get test configuration from script properties
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

/**
 * Quick test for the specific organization error
 */
function quickTestOrganizationFix() {
  const config = getTestConfig();
  if (!config) return;
  
  const payload = {
    custom_fields: {
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7": "2025-05-14",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb": "02:24:00",
      "6fefa2fd421fe09f9f92571cfebca4caff8274b7_until": "2025-05-29",
      "4ae9767e7c7b836e8ef1049b6a2728b4d8207abb_until": "04:24:00"
    }
  };
  
  Logger.log("Testing organization update with time range fields...");
  const response = updateOrganizationDirect(1, payload, config.accessToken, config.basePath, null);
  
  Logger.log(`Response: ${JSON.stringify(response)}`);
  
  if (response.success) {
    Logger.log("✓ SUCCESS: Organization update works correctly!");
  } else {
    Logger.log("✗ FAILED: Organization update still has issues");
    Logger.log(`Error: ${response.error}`);
  }
}