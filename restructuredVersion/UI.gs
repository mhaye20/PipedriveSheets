/**
 * User Interface
 * 
 * This module handles all UI-related functions:
 * - Showing dialogs and sidebars
 * - Building UI components
 * - Managing user interactions
 */

// Create UI namespace if it doesn't exist
var UI = UI || {};

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
  
  // Create a template from the HTML file
  const template = HtmlService.createTemplateFromFile('SettingsDialog');
  
  // Pass data to the template
  template.activeSheetName = activeSheetName;
  template.savedSubdomain = savedSubdomain;
  template.savedFilterId = savedFilterId;
  template.savedEntityType = savedEntityType;
  
  // Create the HTML output from the template
  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setHeight(520)
    .setTitle(`Pipedrive Settings for "${activeSheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
}

/**
 * Shows the column selector dialog
 */
function showColumnSelector() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = activeSheet.getName();
    
    // Get Pipedrive authentication details
    const scriptProperties = PropertiesService.getScriptProperties();
    // Check for OAuth access token instead of API key
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    // Get sheet-specific settings
    const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
    const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
    
    const filterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    Logger.log(`=== COLUMN SELECTOR DEBUG ===`);
    Logger.log(`Sheet: ${sheetName}, Entity: ${entityType}, Filter ID: ${filterId}`);
    Logger.log(`Using OAuth access token: ${Boolean(accessToken)}`);
    
    // Check if we're authenticated
    if (!accessToken) {
      const ui = SpreadsheetApp.getUi();
      const result = ui.alert(
        'Authentication Required',
        'Please connect to Pipedrive to access your data.',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        showAuthorizationDialog();
      }
      return;
    }
    
    // First get actual sample data based on filter ID
    let sampleData = [];
    Logger.log(`Getting sample data for ${entityType} with filter ID ${filterId}`);
    
    try {
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
        case ENTITY_TYPES.PRODUCTS:
          sampleData = getProductsWithFilter(filterId, 1);
          break;
      }
      
      if (sampleData && sampleData.length > 0) {
        Logger.log(`Successfully retrieved sample data for ${entityType}`);
        Logger.log(`Sample data first item has ${Object.keys(sampleData[0]).length} properties`);
        
        // Quick sample of what fields are available
        if (sampleData[0]) {
          const sampleKeys = Object.keys(sampleData[0]).slice(0, 10);
          Logger.log(`Sample fields (first 10): ${sampleKeys.join(', ')}`);
          
          // Check for custom fields
          if (sampleData[0].custom_fields) {
            const customKeys = Object.keys(sampleData[0].custom_fields).slice(0, 10);
            Logger.log(`Sample custom fields (first 10): ${customKeys.join(', ')}`);
          }
        }
        
        // Directly use the sample data for fields extraction
        if (sampleData && sampleData[0]) {
          // Process fields to create UI-friendly structure
          const availableColumns = extractFields(sampleData);
          Logger.log(`Extracted ${availableColumns ? availableColumns.length : 0} available columns from sample data`);
          
          // Make sure availableColumns is valid
          if (availableColumns && availableColumns.length > 0) {
            // Get currently selected columns
            const selectedColumns = UI.getTeamAwareColumnPreferences(entityType, sheetName) || [];
            
            Logger.log(`Found ${selectedColumns.length} selected columns`);
            
            // Add key property to match the path property for compatibility
            availableColumns.forEach(col => {
              if (!col.key) { // Only set if not already present
                col.key = col.path;
              }
            });
            
            // Show the column selector UI
            showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName);
            return;
          }
        }
      } else {
        Logger.log(`No sample data found for ${entityType} with filter ID ${filterId}`);
      }
    } catch (sampleError) {
      Logger.log(`Error getting sample data: ${sampleError.message}`);
      Logger.log(`Stack trace: ${sampleError.stack}`);
    }
    
    // If we couldn't get sample data, try using the field definitions directly
    if (!sampleData || sampleData.length === 0) {
      Logger.log(`No sample data found, falling back to field definitions`);
      
      // Get fields based on entity type
      let fields = [];
      
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          fields = getDealFields(true); // Force a refresh
          break;
        case ENTITY_TYPES.PERSONS:
          fields = getPersonFields(true);
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          fields = getOrganizationFields(true);
          break;
        case ENTITY_TYPES.ACTIVITIES:
          fields = getActivityFields(true);
          break;
        case ENTITY_TYPES.LEADS:
          fields = getLeadFields(true);
          break;
        case ENTITY_TYPES.PRODUCTS:
          fields = getProductFields(true);
          break;
      }
      
      if (fields && fields.length > 0) {
        Logger.log(`Got ${fields.length} fields for ${entityType}`);
        Logger.log(`First field: ${fields[0].name} (${fields[0].key})`);
        
        const availableColumns = fields.map(field => {
          return {
            key: field.key,
            path: field.key,
            name: field.name || field.label || field.key,
            type: field.field_type || 'string',
            isNested: false,
            options: field.options
          };
        });
        
        // Get currently selected columns
        const selectedColumns = UI.getTeamAwareColumnPreferences(entityType, sheetName) || [];
        
        Logger.log(`Found ${selectedColumns.length} selected columns from definitions`);
        
        // Show the selector
        showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName);
      } else {
        Logger.log(`No fields found for ${entityType}, checking field names`);
        
        // If all else fails, use a hardcoded list of common fields based on entity type
        const commonFields = getCommonFieldsForEntity(entityType);
        
        if (commonFields && commonFields.length > 0) {
          const availableColumns = commonFields.map(field => {
            return {
              key: field.key,
              path: field.key,
              name: field.name,
              type: field.type || 'string',
              isNested: false
            };
          });
          
          // Get currently selected columns
          const selectedColumns = UI.getTeamAwareColumnPreferences(entityType, sheetName) || [];
          
          // Show the selector
          showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName);
        } else {
          throw new Error(`Could not fetch fields for ${entityType}. Please check your settings and try again.`);
        }
      }
    }
  } catch (e) {
    Logger.log(`Error in showColumnSelector: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Extracts fields into a UI-friendly structure
 * @param {Array} fields - The fields to extract
 * @param {string} parentPath - Parent path for nested fields
 * @param {string} parentName - Parent name for nested fields
 * @return {Array} Processed fields
 */
function extractFields(fields, parentPath = '', parentName = '') {
  try {
    const result = [];
    
    // Add detailed debugging
    Logger.log(`=== EXTRACT FIELDS DEBUG ===`);
    Logger.log(`Extracting fields with parent path: ${parentPath}, parent name: ${parentName}`);
    Logger.log(`Fields type: ${typeof fields}, is array: ${Array.isArray(fields)}`);
    if (Array.isArray(fields)) {
      Logger.log(`Array length: ${fields.length}`);
      if (fields.length > 0) {
        Logger.log(`First item type: ${typeof fields[0]}`);
        if (typeof fields[0] === 'object') {
          Logger.log(`First item has properties: ${Object.keys(fields[0]).join(', ').substring(0, 100)}...`);
        }
      }
    }
    
    if (!fields) {
      Logger.log('No fields to extract');
      // Return default fields instead of empty array
      return [
        { path: 'id', key: 'id', name: 'ID', type: 'integer' },
        { path: 'name', key: 'name', name: 'Name', type: 'string' },
        { path: 'add_time', key: 'add_time', name: 'Created Date', type: 'date' }
      ];
    }
    
    // Handle the case when we get a sample of data items instead of field definitions
    if (Array.isArray(fields) && fields.length > 0 && fields[0] && typeof fields[0] === 'object') {
      // Check if this is sample data (has id but not key/name properties that field definitions would have)
      if (fields[0].id !== undefined && (!fields[0].key || !fields[0].name)) {
        // This looks like a sample data item, not field definitions
        Logger.log('Detected sample data items instead of field definitions');
        
        const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const sheetName = activeSheet.getName();
        const scriptProperties = PropertiesService.getScriptProperties();
        const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
        const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || 'deals';
        
        // Get custom field mappings
        let customFieldMap = {};
        try {
          // Use PipedriveAPI namespace to access the function
          customFieldMap = PipedriveAPI.getCustomFieldMappings(entityType);
          Logger.log(`Got custom field map with ${Object.keys(customFieldMap).length} entries`);
          if (Object.keys(customFieldMap).length > 0) {
            const sampleKeys = Object.keys(customFieldMap).slice(0, 5);
            Logger.log(`Sample custom field mappings (first 5): ${sampleKeys.map(k => `${k} => ${customFieldMap[k]}`).join(', ')}`);
          } else {
            // If we don't have custom field mappings, we'll still continue and use key names directly
            Logger.log(`No custom field mappings found for ${entityType}, will use field keys as names`);
          }
        } catch (e) {
          // Log error but continue processing - custom field mappings are helpful but not critical
          Logger.log(`Error getting custom field mappings: ${e.message}`);
          Logger.log(`Stack trace: ${e.stack}`);
          Logger.log(`Continuing without custom field mappings, will use field keys as names`);
        }
        
        const addedCustomFields = new Set(); // Track custom fields we've already processed
        const processedKeys = new Set(); // Track all processed keys to avoid duplicates
        
        // Process all items in the sample data to gather as many fields as possible
        for (const sampleItem of fields) {
          // Extract all top-level properties
          for (const key in sampleItem) {
            // Skip functions and internal properties and already processed keys
            if (typeof sampleItem[key] === 'function' || key.startsWith('_') || processedKeys.has(key)) continue;
            
            // Mark key as processed
            processedKeys.add(key);
            
            // Check if this is a hash-pattern key (custom field directly at root level)
            const hashPattern = /^[0-9a-f]{40}$/;
            if (hashPattern.test(key) && customFieldMap[key] && !addedCustomFields.has(key)) {
              addedCustomFields.add(key);
              Logger.log(`Found hash-pattern key: ${key} => ${customFieldMap[key]}`);
              
              // Add this custom field with its friendly name
              result.push({
                path: key,
                key: key,
                name: `${customFieldMap[key]}`,
                type: typeof sampleItem[key],
                isNested: false
              });
              
              // If it's a complex object, also add fields for its properties
              const customValue = sampleItem[key];
              if (customValue && typeof customValue === 'object') {
                for (const propKey in customValue) {
                  if (propKey !== 'value') { // Skip the main value property
                    result.push({
                      path: `${key}.${propKey}`,
                      key: `${key}.${propKey}`,
                      name: `${customFieldMap[key]} - ${propKey}`,
                      type: typeof customValue[propKey],
                      isNested: true,
                      parentKey: key
                    });
                  }
                }
              }
              continue;
            }
            
            // Regular top-level property - add to results
            result.push({
              path: key,
              key: key,
              name: key.charAt(0).toUpperCase() + key.slice(1).replace(/_/g, ' '),
              type: typeof sampleItem[key],
              isNested: false
            });
            
            // If this property is an object, extract its properties as nested fields
            if (sampleItem[key] && typeof sampleItem[key] === 'object' && !Array.isArray(sampleItem[key])) {
              for (const nestedKey in sampleItem[key]) {
                // Skip functions and internal properties
                if (typeof sampleItem[key][nestedKey] === 'function' || nestedKey.startsWith('_')) continue;
                
                const nestedPath = `${key}.${nestedKey}`;
                if (!result.some(col => col.key === nestedPath)) {
                  result.push({
                    path: nestedPath,
                    key: nestedPath,
                    name: `${key.charAt(0).toUpperCase() + key.slice(1).replace(/_/g, ' ')} > ${nestedKey}`,
                    type: typeof sampleItem[key][nestedKey],
                    isNested: true,
                    parentKey: key
                  });
                }
              }
            }
          }
          
          // Special handling for custom_fields
          if (sampleItem.custom_fields) {
            Logger.log(`Found custom_fields in sample data with ${Object.keys(sampleItem.custom_fields).length} entries`);
            
            // Extract custom fields
            for (const key in sampleItem.custom_fields) {
              if (addedCustomFields.has(`custom_fields.${key}`)) continue;
              addedCustomFields.add(`custom_fields.${key}`);
              
              const customValue = sampleItem.custom_fields[key];
              const friendlyName = customFieldMap[key] || key;
              const currentPath = `custom_fields.${key}`;
              
              Logger.log(`Processing custom field: ${key} => ${friendlyName}`);
              
              // Simple value custom fields
              if (typeof customValue !== 'object' || customValue === null) {
                result.push({
                  path: currentPath,
                  key: currentPath,
                  name: friendlyName,
                  type: typeof customValue,
                  isNested: true,
                  parentKey: 'custom_fields'
                });
                continue;
              }
              
              // Handle complex custom fields (object-based)
              if (typeof customValue === 'object') {
                // Log the custom field structure
                Logger.log(`Custom field ${key} is an object with properties: ${Object.keys(customValue).join(', ')}`);
                
                // Currency fields
                if (customValue.value !== undefined && customValue.currency !== undefined) {
                  result.push({
                    path: currentPath,
                    key: currentPath,
                    name: `${friendlyName} (Currency)`,
                    type: typeof customValue.value,
                    isNested: true,
                    parentKey: 'custom_fields'
                  });
                }
                // Date/Time range fields
                else if (customValue.value !== undefined && customValue.until !== undefined) {
                  result.push({
                    path: currentPath,
                    key: currentPath,
                    name: `${friendlyName} (Range)`,
                    type: typeof customValue.value,
                    isNested: true,
                    parentKey: 'custom_fields'
                  });
                }
                // Address fields
                else if (customValue.value !== undefined && customValue.formatted_address !== undefined) {
                  result.push({
                    path: currentPath,
                    key: currentPath,
                    name: `${friendlyName} (Address)`,
                    type: typeof customValue.value,
                    isNested: true,
                    parentKey: 'custom_fields'
                  });
                  
                  // Add formatted address as a separate column option
                  result.push({
                    path: `${currentPath}.formatted_address`,
                    key: `${currentPath}.formatted_address`,
                    name: `${friendlyName} (Formatted Address)`,
                    type: typeof customValue.formatted_address,
                    isNested: true,
                    parentKey: currentPath
                  });
                }
                // For all other object types
                else {
                  result.push({
                    path: currentPath,
                    key: currentPath,
                    name: `${friendlyName} (Complex)`,
                    type: 'object',
                    isNested: true,
                    parentKey: 'custom_fields'
                  });
                  
                  // Extract nested properties from complex custom field
                  for (const propKey in customValue) {
                    if (propKey !== 'value') { // Skip the main value property
                      result.push({
                        path: `${currentPath}.${propKey}`,
                        key: `${currentPath}.${propKey}`,
                        name: `${friendlyName} - ${propKey}`,
                        type: typeof customValue[propKey],
                        isNested: true,
                        parentKey: currentPath
                      });
                    }
                  }
                }
              }
            }
          }
        }
        
        Logger.log(`Extracted ${result.length} columns total from sample data`);
        return result;
      }
    }
    
    // If this is an array of field definitions like from Pipedrive API
    if (Array.isArray(fields)) {
      Logger.log(`Processing array of ${fields.length} fields`);
      
      fields.forEach(field => {
        // Skip unsupported field types
        if (field.field_type === 'picture' ||
            field.field_type === 'file' ||
            field.field_type === 'visibleto') {
          return;
        }
        
        // Ensure the field has a key
        if (!field.key && !field.id) {
          Logger.log(`Skipping field with no key or id: ${JSON.stringify(field).substring(0, 100)}`);
          return;
        }
        
        // Use key or id as fallback
        const fieldKey = field.key || `field_${field.id}`;
        
        // Construct path
        const path = parentPath ? `${parentPath}.${fieldKey}` : fieldKey;
        
        // Construct readable name
        const fieldName = field.name || field.label || 'Unnamed Field';
        const name = parentName ? `${parentName} > ${fieldName}` : fieldName;
        
        // Add the field to the result
        result.push({
          path: path,
          key: path,
          name: name,
          type: field.field_type || 'string',
          isNested: !!parentPath,
          parentKey: parentPath || null
        });
        
        // Handle nested objects
        if (field.options && Array.isArray(field.options)) {
          // This is a field with options (like dropdown)
          Logger.log(`Field ${fieldName} has ${field.options.length} options`);
        } else if (field.subfields && Array.isArray(field.subfields)) {
          // This is a field with subfields (like address)
          extractFields(field.subfields, path, name);
        }
      });
      
      return result;
    }
    
    // If fields is a single object, extract its properties
    if (fields && typeof fields === 'object' && !Array.isArray(fields)) {
      for (const key in fields) {
        if (typeof fields[key] === 'function' || key.startsWith('_')) continue;
        
        const path = parentPath ? `${parentPath}.${key}` : key;
        const fieldName = key.charAt(0).toUpperCase() + key.slice(1).replace(/_/g, ' ');
        const name = parentName ? `${parentName} > ${fieldName}` : fieldName;
        
        result.push({
          path: path,
          key: path,
          name: name,
          type: typeof fields[key],
          isNested: !!parentPath,
          parentKey: parentPath || null
        });
        
        // Recursively extract nested objects
        if (fields[key] && typeof fields[key] === 'object' && !Array.isArray(fields[key])) {
          const nestedFields = extractFields(fields[key], path, name);
          result.push(...nestedFields);
        }
      }
      
      return result;
    }
    
    // Return what we have
    return result;
  } catch (e) {
    Logger.log(`Error in extractFields: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    return [
      { path: 'id', key: 'id', name: 'ID', type: 'integer' },
      { path: 'name', key: 'name', name: 'Name', type: 'string' },
      { path: 'add_time', key: 'add_time', name: 'Created Date', type: 'date' }
    ];
  }
}

/**
 * Shows the column selector UI
 * @param {Array} availableColumns - Available columns
 * @param {Array} selectedColumns - Selected columns
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
function showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName) {
  try {
    // Log detailed information about the data
    Logger.log(`*** UI DATA PASSING DEBUG ***`);
    Logger.log(`Passing ${availableColumns.length} available columns to template`);
    if (availableColumns.length > 0) {
      Logger.log(`First available column: ${JSON.stringify(availableColumns[0])}`);
      Logger.log(`Column keys preview: ${availableColumns.slice(0, 5).map(c => c.key).join(', ')}`);
    }
    
    // Create a simple HTML UI instead of using the complex template
    let html = `
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h3 { margin-bottom: 15px; }
        .container { display: flex; gap: 20px; height: 400px; }
        .column { flex: 1; border: 1px solid #ccc; border-radius: 4px; padding: 10px; overflow: hidden; display: flex; flex-direction: column; }
        .column-heading { font-weight: bold; margin-bottom: 10px; }
        .item { padding: 5px; margin: 5px 0; background: #f5f5f5; border-radius: 3px; cursor: pointer; display: flex; align-items: center; }
        .item:hover { background: #e0e0e0; }
        .item-name { flex: 1; overflow: hidden; text-overflow: ellipsis; }
        .item-button { margin-left: 5px; color: #4285f4; cursor: pointer; user-select: none; }
        .scrollable { overflow-y: auto; flex-grow: 1; }
        .search { width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; }
        .buttons { margin-top: 20px; text-align: right; }
        button { padding: 8px 16px; margin-left: 10px; }
        .primary-btn { background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
        .primary-btn:hover { background-color: #3367d6; }
        .info { margin-bottom: 10px; padding: 10px; background: #f9f9f9; border-radius: 4px; color: #666; }
      </style>
      
      <h3>Column Selection for ${entityType} in "${sheetName}"</h3>
      
      <div class="info">
        We found ${availableColumns.length} available columns. Select which ones to display in your sheet.
      </div>
      
      <div class="container">
        <div class="column">
          <div class="column-heading">Available Columns (<span id="availableCount">${availableColumns.length}</span>)</div>
          <input type="text" class="search" id="searchBox" placeholder="Search columns...">
          <div class="scrollable" id="availableList">
    `;
    
    // Add all available columns
    for (const col of availableColumns) {
      html += `
        <div class="item" data-key="${col.key}" onclick="addColumn('${col.key}')">
          <span class="item-name">${col.name}</span>
          <span class="item-button">+</span>
        </div>
      `;
    }
    
    html += `
          </div>
        </div>
        
        <div class="column">
          <div class="column-heading">Selected Columns (<span id="selectedCount">${selectedColumns.length}</span>)</div>
          <div class="scrollable" id="selectedList">
    `;
    
    // Add selected columns
    for (const col of selectedColumns) {
      html += `
        <div class="item" data-key="${col.key}">
          <span class="item-name">${col.name}</span>
          <span class="item-button" onclick="removeColumn('${col.key}')">✕</span>
        </div>
      `;
    }
    
    if (selectedColumns.length === 0) {
      html += `<div style="padding: 10px; color: #666;">No columns selected yet. Click on columns on the left to add them.</div>`;
    }
    
    html += `
          </div>
        </div>
      </div>
      
      <div class="buttons">
        <button onclick="google.script.host.close()">Cancel</button>
        <button class="primary-btn" onclick="saveColumns()">Save & Close</button>
      </div>
      
      <script>
        // Column data
        const availableColumnsData = ${JSON.stringify(availableColumns)};
        let selectedColumnsData = ${JSON.stringify(selectedColumns)};
        
        // Element references
        const availableList = document.getElementById('availableList');
        const selectedList = document.getElementById('selectedList');
        const searchBox = document.getElementById('searchBox');
        const availableCount = document.getElementById('availableCount');
        const selectedCount = document.getElementById('selectedCount');
        
        // Add a column to selected
        function addColumn(key) {
          // Skip if already selected
          if (selectedColumnsData.some(col => col.key === key)) return;
          
          // Find column in available columns
          const column = availableColumnsData.find(col => col.key === key);
          if (column) {
            // Add to selected
            selectedColumnsData.push(column);
            updateLists();
          }
        }
        
        // Remove a column from selected
        function removeColumn(key) {
          // Remove from selected
          selectedColumnsData = selectedColumnsData.filter(col => col.key !== key);
          updateLists();
        }
        
        // Update both lists
        function updateLists() {
          // Update counts
          selectedCount.textContent = selectedColumnsData.length;
          
          // Update selected list
          selectedList.innerHTML = '';
          
          if (selectedColumnsData.length === 0) {
            selectedList.innerHTML = '<div style="padding: 10px; color: #666;">No columns selected yet. Click on columns on the left to add them.</div>';
          } else {
            selectedColumnsData.forEach(col => {
              const item = document.createElement('div');
              item.className = 'item';
              item.dataset.key = col.key;
              item.innerHTML = \`
                <span class="item-name">\${col.name}</span>
                <span class="item-button" onclick="removeColumn('\${col.key}')">✕</span>
              \`;
              selectedList.appendChild(item);
            });
          }
          
          // Filter available list
          filterAvailableList();
        }
        
        // Filter available columns based on search
        function filterAvailableList() {
          const searchTerm = searchBox.value.toLowerCase();
          const filteredColumns = availableColumnsData.filter(col => 
            !selectedColumnsData.some(selected => selected.key === col.key) &&
            (searchTerm === '' || col.name.toLowerCase().includes(searchTerm))
          );
          
          // Update UI
          availableCount.textContent = filteredColumns.length;
          availableList.innerHTML = '';
          
          filteredColumns.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.dataset.key = col.key;
            item.onclick = () => addColumn(col.key);
            item.innerHTML = \`
              <span class="item-name">\${col.name}</span>
              <span class="item-button">+</span>
            \`;
            availableList.appendChild(item);
          });
          
          if (filteredColumns.length === 0) {
            availableList.innerHTML = '<div style="padding: 10px; color: #666;">No matching columns found</div>';
          }
        }
        
        // Save the column selection
        function saveColumns() {
          if (selectedColumnsData.length === 0) {
            alert('Please select at least one column');
            return;
          }
          
          // Call back to the server to save
          google.script.run
            .withSuccessHandler(() => {
              alert('Column preferences saved successfully!');
              google.script.host.close();
            })
            .withFailureHandler(error => {
              alert('Error saving column preferences: ' + error.message);
            })
            .saveColumnPreferences(selectedColumnsData, "${entityType}", "${sheetName}");
        }
        
        // Set up event listeners
        searchBox.addEventListener('input', filterAvailableList);
        
        // Initialize the UI
        updateLists();
      </script>
    `;
    
    // Create HTML output
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(800)
      .setHeight(600)
      .setTitle(`Select Columns for ${entityType} in "${sheetName}"`);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Column Selection`);
    
  } catch (e) {
    Logger.log(`Error showing column selector UI: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to load column selector: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Saves column preferences after user selection
 * @param {Array} selectedColumns - Array of selected column objects
 * @param {string} entityType - The entity type (persons, deals, etc.)
 * @param {string} sheetName - The name of the sheet
 * @return {boolean} - Success status
 */
function saveColumnPreferences(selectedColumns, entityType, sheetName) {
  try {
    Logger.log(`Saving column preferences for ${entityType} in sheet "${sheetName}"`);
    Logger.log(`Selected ${selectedColumns.length} columns: ${selectedColumns.map(c => c.key).join(', ')}`);
    
    // Get the current user
    const user = Session.getActiveUser().getEmail();
    
    // Save to SyncService - use namespace
    return SyncService.saveColumnPreferences(selectedColumns, entityType, sheetName, user);
  } catch (e) {
    Logger.log(`Error saving column preferences: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    throw new Error(`Failed to save column preferences: ${e.message}`);
  }
}

/**
 * Shows the sync status dialog
 * @param {string} sheetName - Sheet name
 */
function showSyncStatus(sheetName) {
  try {
    // Get the current entity type for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('SyncStatus');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.entityType = entityType;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Syncing Data')
      .setWidth(400)
      .setHeight(300);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Syncing Data');
  } catch (e) {
    Logger.log(`Error in showSyncStatus: ${e.message}`);
    // Just log the error but don't throw, as this is a UI enhancement
  }
}

/**
 * Shows the team manager dialog
 * @param {boolean} joinOnly - Whether to show only the join team UI
 */
function showTeamManager(joinOnly = false) {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Unable to determine your email address. Please ensure you are signed in.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    
    // Get user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    
    // Determine if user is in a team
    const hasTeam = userTeam !== null;
    
    // If join-only mode and user isn't in a team, use the simpler join dialog
    if (joinOnly && !hasTeam) {
      return showTeamJoinRequest();
    }
    
    // Create a direct HTML string rather than using a complex template
    // This is more robust and less likely to have loading issues
    let html = `
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
        }
        h3 {
          margin-top: 0;
          margin-bottom: 15px;
          color: #4285F4;
        }
        .card {
          border: 1px solid #ddd;
          border-radius: 4px;
          padding: 15px;
          margin-bottom: 20px;
          background-color: #f9f9f9;
        }
        .section-title {
          font-weight: bold;
          margin-bottom: 10px;
        }
        input[type="text"] {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-bottom: 10px;
          box-sizing: border-box;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 8px 15px;
          border-radius: 4px;
          cursor: pointer;
          margin-right: 10px;
          margin-bottom: 10px;
        }
        button:hover {
          background-color: #3367D6;
        }
        button.secondary {
          background-color: #f1f1f1;
          color: #333;
        }
        button.secondary:hover {
          background-color: #e4e4e4;
        }
        button.danger {
          background-color: #EA4335;
        }
        button.danger:hover {
          background-color: #D62516;
        }
        .footer {
          margin-top: 20px;
          display: flex;
          justify-content: flex-end;
        }
        .tab-container {
          display: flex;
          margin-bottom: 20px;
          border-bottom: 1px solid #ddd;
        }
        .tab {
          padding: 10px 15px;
          cursor: pointer;
        }
        .tab.active {
          color: #4285F4;
          font-weight: bold;
          border-bottom: 2px solid #4285F4;
        }
        .tab-content {
          display: none;
        }
        .tab-content.active {
          display: block;
        }
        .member-list {
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-top: 10px;
          max-height: 200px;
          overflow-y: auto;
        }
        .member {
          padding: 8px 12px;
          border-bottom: 1px solid #eee;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .member:last-child {
          border-bottom: none;
        }
        .status-message {
          padding: 10px;
          margin: 10px 0;
          border-radius: 4px;
        }
        .status-success {
          background-color: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }
        .status-error {
          background-color: #f8d7da;
          color: #721c24;
          border: 1px solid #f5c6cb;
        }
      </style>
      
      <div class="container">
        <h3>${hasTeam ? 'Team Management' : 'Team Access'}</h3>
        
        <div id="status-container"></div>
        
        <div class="tab-container">
          ${hasTeam ? '' : '<div class="tab active" data-tab="join">Join Team</div>'}
          ${hasTeam ? '' : '<div class="tab" data-tab="create">Create Team</div>'}
          ${hasTeam ? '<div class="tab active" data-tab="manage">Manage Team</div>' : ''}
        </div>
    `;
    
    // Join Team Tab
    if (!hasTeam) {
      html += `
        <div id="join-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Join an Existing Team</div>
            <p>Enter the team ID provided by your team administrator:</p>
            <input type="text" id="team-id-input" placeholder="Team ID">
            <div class="footer">
              <button id="join-team-button">Join Team</button>
            </div>
          </div>
        </div>
        
        <div id="create-tab" class="tab-content">
          <div class="card">
            <div class="section-title">Create a New Team</div>
            <p>Create a new team to share Pipedrive configuration with your colleagues:</p>
            <input type="text" id="team-name-input" placeholder="Team Name">
            <div class="footer">
              <button id="create-team-button">Create Team</button>
            </div>
          </div>
        </div>
      `;
    }
    
    // Manage Team Tab
    if (hasTeam) {
      const teamName = userTeam.name || 'Unnamed Team';
      const teamId = userTeam.teamId || '';
      const userRole = userTeam.role || 'Member';
      const isAdmin = userRole === 'Admin';
      
      // Get team members
      const members = [];
      if (teamsData[teamId]) {
        if (teamsData[teamId].members) {
          // New format
          Object.keys(teamsData[teamId].members).forEach(email => {
            members.push({
              email: email,
              role: teamsData[teamId].members[email]
            });
          });
        } else if (teamsData[teamId].memberEmails) {
          // Legacy format
          teamsData[teamId].memberEmails.forEach(email => {
            const isAdmin = teamsData[teamId].adminEmails && 
                           teamsData[teamId].adminEmails.includes(email);
            members.push({
              email: email,
              role: isAdmin ? 'Admin' : 'Member'
            });
          });
        }
      }
      
      html += `
        <div id="manage-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Current Team</div>
            <div style="margin-bottom: 15px;">
              <div><strong>Team Name:</strong> ${teamName}</div>
              <div><strong>Team ID:</strong> <span id="team-id-display">${teamId}</span></div>
              <div><strong>Your Role:</strong> ${userRole}</div>
            </div>
            
            <button id="copy-team-id" class="secondary">Copy Team ID</button>
            
            <div class="section-title" style="margin-top: 20px;">Team Members</div>
            <div class="member-list">
      `;
      
      if (members.length > 0) {
        members.forEach(member => {
          html += `
            <div class="member">
              <div>${member.email}</div>
              <div>
                <span style="margin-right: 10px;">${member.role}</span>
                ${isAdmin && member.email !== userEmail ? 
                  `<button class="secondary remove-member" data-email="${member.email}">Remove</button>` : ''}
              </div>
            </div>
          `;
        });
      } else {
        html += '<div class="member">No members found</div>';
      }
      
      html += `
            </div>
      `;
      
      if (isAdmin) {
        html += `
            <div class="section-title" style="margin-top: 15px;">Add Team Member</div>
            <input type="text" id="new-member-email" placeholder="Email Address">
            <button id="add-member-button">Add Member</button>
        `;
      }
      
      html += `
            <div style="margin-top: 20px;">
              <button id="leave-team-button" class="danger">Leave Team</button>
              ${isAdmin ? '<button id="delete-team-button" class="danger">Delete Team</button>' : ''}
            </div>
          </div>
        </div>
      `;
    }
    
    // Footer
    html += `
        <div class="footer">
          <button id="close-button" class="secondary">Close</button>
        </div>
      </div>
      
      <script>
        // Document ready
        (function() {
          // Tab switching
          document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', function() {
              // Update active tab
              document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
              this.classList.add('active');
              
              // Show corresponding content
              const tabId = this.dataset.tab + '-tab';
              document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
              });
              document.getElementById(tabId).classList.add('active');
            });
          });
          
          // Status message helper
          window.showStatus = function(message, type) {
            const container = document.getElementById('status-container');
            const statusDiv = document.createElement('div');
            statusDiv.className = 'status-message status-' + type;
            statusDiv.textContent = message;
            container.innerHTML = '';
            container.appendChild(statusDiv);
            
            if (type === 'success') {
              setTimeout(() => statusDiv.remove(), 5000);
            }
          };
          
          // Close button
          document.getElementById('close-button').addEventListener('click', function() {
            google.script.host.close();
          });
    `;
    
    // Add join-specific handlers
    if (!hasTeam) {
      html += `
          // Join team handler
          document.getElementById('join-team-button').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-input').value.trim();
            if (!teamId) {
              showStatus('Please enter a team ID', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Joining team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Successfully joined the team!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('join-team-button').disabled = false;
                  showStatus(result.message || 'Error joining team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('join-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .joinTeam(teamId);
          });
          
          // Create team handler
          document.getElementById('create-team-button').addEventListener('click', function() {
            const teamName = document.getElementById('team-name-input').value.trim();
            if (!teamName) {
              showStatus('Please enter a team name', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Creating team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team created successfully!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('create-team-button').disabled = false;
                  showStatus(result.message || 'Error creating team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('create-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .createTeam(teamName);
          });
      `;
    }
    
    // Add team management handlers
    if (hasTeam) {
      html += `
          // Copy team ID handler
          document.getElementById('copy-team-id').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-display').textContent;
            
            // Create a temporary input element
            const tempInput = document.createElement('input');
            tempInput.value = teamId;
            document.body.appendChild(tempInput);
            tempInput.select();
            document.execCommand('copy');
            document.body.removeChild(tempInput);
            
            showStatus('Team ID copied to clipboard!', 'success');
          });
          
          // Leave team handler
          document.getElementById('leave-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to leave this team? You will lose access to shared configurations.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Leaving team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('You have left the team.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('leave-team-button').disabled = false;
                  showStatus(result.message || 'Error leaving team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('leave-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .leaveTeam();
          });
      `;
      
      // Add admin-specific handlers
      if (userTeam && userTeam.role === 'Admin') {
        html += `
          // Add member handler
          document.getElementById('add-member-button').addEventListener('click', function() {
            const email = document.getElementById('new-member-email').value.trim();
            if (!email) {
              showStatus('Please enter an email address', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Adding member...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Member added successfully.', 'success');
                  document.getElementById('new-member-email').value = '';
                  document.getElementById('add-member-button').disabled = false;
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('add-member-button').disabled = false;
                  showStatus(result.message || 'Error adding member', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('add-member-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .addTeamMember(email);
          });
          
          // Remove member handlers
          document.querySelectorAll('.remove-member').forEach(button => {
            button.addEventListener('click', function() {
              const email = this.dataset.email;
              if (!confirm('Are you sure you want to remove ' + email + ' from the team?')) {
                return;
              }
              
              this.disabled = true;
              showStatus('Removing member...', 'info');
              
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    showStatus('Member removed successfully.', 'success');
                    setTimeout(() => window.top.location.reload(), 2000);
                  } else {
                    document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                    showStatus(result.message || 'Error removing member', 'error');
                  }
                })
                .withFailureHandler(function(error) {
                  document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                  showStatus('Error: ' + error.message, 'error');
                })
                .removeTeamMember(email);
            });
          });
          
          // Delete team handler
          document.getElementById('delete-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to delete this team? This will remove access for all team members and cannot be undone.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Deleting team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team has been deleted.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('delete-team-button').disabled = false;
                  showStatus(result.message || 'Error deleting team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('delete-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .deleteTeam();
          });
        `;
      }
    }
    
    // Close script tag
    html += `
        })(); // End of self-executing function
      </script>
    `;
    
    // Create HTML service object
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setTitle(hasTeam ? 'Team Management' : 'Join Team')
      .setWidth(joinOnly ? 450 : 600)
      .setHeight(joinOnly ? 400 : 550);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Join Team');
  } catch (e) {
    Logger.log(`Error in showTeamManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open team manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reopens the team manager after an action
 */
function reopenTeamManager() {
  showTeamManager(false);
}

/**
 * Saves team-aware column preferences
 * @param {Array} columns - Array of column keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @returns {boolean} Success status
 */
UI.saveTeamAwareColumnPreferences = function(columns, entityType, sheetName) {
  try {
    if (!entityType || !sheetName) {
      throw new Error("Missing required parameters for saveTeamAwareColumnPreferences");
    }
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const userEmail = Session.getActiveUser().getEmail();
    
    // Store columns in user-specific property
    const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
    scriptProperties.setProperty(userColumnSettingsKey, JSON.stringify(columns));
    Logger.log(`Saved user-specific column preferences with key: ${userColumnSettingsKey}`);
    
    // Also store in global property for backward compatibility
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    scriptProperties.setProperty(columnSettingsKey, JSON.stringify(columns));
    
    // Check if two-way sync is enabled for this sheet
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // When columns are changed and two-way sync is enabled, handle tracking column
    if (twoWaySyncEnabled) {
      Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Checking if we need to adjust sync column.`);
      
      // When columns are changed, delete the tracking column property to force repositioning
      const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
      scriptProperties.deleteProperty(twoWaySyncTrackingColumnKey);
      
      // Add a flag to indicate that the Sync Status column should be repositioned at the end
      const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${sheetName}`;
      scriptProperties.setProperty(twoWaySyncColumnAtEndKey, 'true');
      
      Logger.log(`Removed tracking column property for sheet "${sheetName}" to ensure correct positioning on next sync.`);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTeamAwareColumnPreferences: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    throw e;
  }
}

/**
 * Gets team-aware column preferences
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @returns {Array} Array of column configurations
 */
UI.getTeamAwareColumnPreferences = function(entityType, sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const userEmail = Session.getActiveUser().getEmail();

    // Get the user's team and check sharing settings
    const userTeam = getUserTeam(userEmail);

    // First try to get user-specific column preferences
    const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
    let savedColumnsJson = scriptProperties.getProperty(userColumnSettingsKey);

    // If the user is part of a team and column sharing is enabled, check team preferences
    if ((!savedColumnsJson || savedColumnsJson === '[]') && userTeam && userTeam.shareColumns) {
      // Look for team members' configurations
      const memberEmails = userTeam.memberEmails || [];
      for (const teamMemberEmail of memberEmails) {
        if (teamMemberEmail === userEmail) continue; // Skip the current user

        const teamMemberColumnSettingsKey = `COLUMNS_${teamMemberEmail}_${sheetName}_${entityType}`;
        const teamMemberColumnsJson = scriptProperties.getProperty(teamMemberColumnSettingsKey);

        if (teamMemberColumnsJson && teamMemberColumnsJson !== '[]') {
          savedColumnsJson = teamMemberColumnsJson;
          Logger.log(`Using team member ${teamMemberEmail}'s column preferences for ${entityType}`);
          break;
        }
      }
    }

    // If still no team column preferences, fall back to the global setting
    if (!savedColumnsJson || savedColumnsJson === '[]') {
      const globalColumnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
      savedColumnsJson = scriptProperties.getProperty(globalColumnSettingsKey);
    }

    let selectedColumns = [];
    if (savedColumnsJson) {
      try {
        selectedColumns = JSON.parse(savedColumnsJson);
      } catch (e) {
        Logger.log(`Error parsing saved columns: ${e.message}`);
        selectedColumns = [];
      }
    }

    return selectedColumns;
  } catch (e) {
    Logger.log(`Error in getTeamAwareColumnPreferences: ${e.message}`);
    return [];
  }
}

/**
 * Gets team-aware Pipedrive filters
 * @return {Array} Array of filter objects
 */
function getTeamAwarePipedriveFilters() {
  try {
    // Get user filters
    const userFilters = getPipedriveFilters();
    
    // Check if user is in a team with shared filters
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    // If not in a team or team doesn't share filters, return user filters
    if (!userTeam || !userTeam.settings || !userTeam.settings.shareFilters) {
      return userFilters;
    }
    
    // Return team filters if available
    return userFilters;
  } catch (e) {
    Logger.log(`Error in getTeamAwarePipedriveFilters: ${e.message}`);
    return [];
  }
}

/**
 * Gets team-aware Pipedrive filters for a specific entity type
 * @param {string} entityType - The entity type to get filters for
 * @return {Array} Array of filter objects
 */
function getFiltersForEntityType(entityType) {
  try {
    // Directly call the function with the same name in the PipedriveAPI.gs file
    return PipedriveAPI.getFiltersForEntityType(entityType);
  } catch (e) {
    Logger.log(`Error in UI.getFiltersForEntityType: ${e.message}`);
    return [];
  }
}

/**
 * Saves settings for the add-on
 * @param {string} apiKey - Pipedrive API key
 * @param {string} entityType - Entity type
 * @param {string} filterId - Filter ID
 * @param {string} subdomain - Pipedrive subdomain
 * @param {string} sheetName - Sheet name
 * @param {boolean} enableTimestamp - Whether to enable timestamp
 * @return {boolean} True if successful, false otherwise
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp = false) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();

    // Save global settings (API key and subdomain are global)
    if (apiKey) scriptProperties.setProperty('PIPEDRIVE_API_KEY', apiKey);
    if (subdomain) scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
    if (sheetName) scriptProperties.setProperty('SHEET_NAME', sheetName);
    
    // Save sheet-specific settings
    if (entityType && sheetName) {
      const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
      scriptProperties.setProperty(sheetEntityTypeKey, entityType);
    }
    
    // Only update filter ID if provided (may be empty intentionally)
    if (filterId !== undefined && sheetName) {
      const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
      scriptProperties.setProperty(sheetFilterIdKey, filterId);
    }
    
    // Handle boolean property
    const timestampEnabledKey = `TIMESTAMP_ENABLED_${sheetName}`;
    scriptProperties.setProperty(timestampEnabledKey, enableTimestamp.toString());
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveSettings: ${e.message}`);
    throw e;
  }
}

function showTriggerManager() {
  try {
    // Get current triggers
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Get all triggers for this sheet
    const triggers = getTriggersForSheet(sheetName);
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TriggerManager');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.triggers = triggers;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Schedule Sync')
      .setWidth(600)
      .setHeight(500);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Sync');
  } catch (e) {
    Logger.log(`Error in showTriggerManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open trigger manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showTwoWaySyncSettings() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get current settings
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const enableTwoWaySync = scriptProps.getProperty(twoWaySyncEnabledKey) === 'true';
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const trackingColumn = scriptProps.getProperty(twoWaySyncTrackingColumnKey) || '';
    
    // Create a settings object with all properties the template might need
    const settings = {
      enableTwoWaySync: enableTwoWaySync,
      trackingColumn: trackingColumn,
      sheetName: activeSheetName
    };
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
    
    // Pass data to the template
    htmlTemplate.enableTwoWaySync = enableTwoWaySync;
    htmlTemplate.trackingColumn = trackingColumn;
    htmlTemplate.settings = settings; // Pass the entire settings object
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Two-Way Sync Settings')
      .setWidth(500)
      .setHeight(400);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
  } catch (e) {
    Logger.log(`Error in showTwoWaySyncSettings: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open two-way sync settings: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Save settings with sheet-specific keys
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    
    // Save the settings
    scriptProps.setProperty(twoWaySyncEnabledKey, enableTwoWaySync ? 'true' : 'false');
    
    // Debug log to verify what we're saving
    Logger.log(`Saving two-way sync setting: ${twoWaySyncEnabledKey} = ${enableTwoWaySync ? 'true' : 'false'}`);
    
    if (trackingColumn) {
      // Save the tracking column
      scriptProps.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
    } else {
      // If no column specified, clean up the property
      scriptProps.deleteProperty(twoWaySyncTrackingColumnKey);
    }
    
    // If two-way sync is enabled, set up the onEdit trigger and immediately add the Sync Status column
    if (enableTwoWaySync) {
      // Set up the onEdit trigger
      setupOnEditTrigger();
      
      // Add Sync Status column if it doesn't exist yet
      addSyncStatusColumn(activeSheet, trackingColumn);
    } else {
      removeOnEditTrigger();
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTwoWaySyncSettings: ${e.message}`);
    throw e;
  }
}

function addSyncStatusColumn(sheet, specificColumn = '') {
  try {
    // First, check if there's already a Sync Status column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColumnIndex = -1;
    
    // Search for existing Sync Status column
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'Sync Status') {
        syncStatusColumnIndex = i;
        break;
      }
    }
    
    // If we found an existing column, use it
    if (syncStatusColumnIndex >= 0) {
      const columnLetter = columnToLetter(syncStatusColumnIndex + 1); // +1 because it's 1-based
      Logger.log(`Found existing Sync Status column at ${columnLetter}`);
      return;
    }
    
    // If we specified a specific column, use that
    if (specificColumn) {
      const columnIndex = columnLetterToIndex(specificColumn);
      
      // Check if this column already has a header
      if (columnIndex <= sheet.getLastColumn()) {
        const existingHeader = sheet.getRange(1, columnIndex).getValue();
        if (existingHeader) {
          // If there's already a header, append to the end instead
          Logger.log(`Column ${specificColumn} already has header "${existingHeader}". Will append Sync Status to the end instead.`);
          specificColumn = '';
        }
      }
    }
    
    // Add the Sync Status column
    let targetColumnIndex;
    if (specificColumn) {
      // Use the specified column
      targetColumnIndex = columnLetterToIndex(specificColumn);
    } else {
      // Append to the end
      targetColumnIndex = sheet.getLastColumn() + 1;
    }
    
    // Set the header
    sheet.getRange(1, targetColumnIndex).setValue('Sync Status');
    
    // Format the header cell
    sheet.getRange(1, targetColumnIndex)
         .setFontWeight('bold')
         .setBackground('#E8F0FE');
    
    // Save the column location for later
    const sheetName = sheet.getName();
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const columnLetter = columnToLetter(targetColumnIndex);
    PropertiesService.getScriptProperties().setProperty(trackingColumnKey, columnLetter);
    
    Logger.log(`Added Sync Status column at column ${columnLetter}`);
    
    // Add conditional formatting for the Sync Status column
    const lastRow = Math.max(sheet.getLastRow(), 100); // Format at least 100 rows
    if (lastRow > 1) {
      const statusRange = sheet.getRange(2, targetColumnIndex, lastRow - 1, 1);
      const rules = sheet.getConditionalFormatRules();
      
      // Create rule for "Modified" cells
      const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FCE8E6') // Light red
        .setBold(true)
        .setRanges([statusRange])
        .build();
      
      // Create rule for "Synced" cells
      const syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA') // Light green
        .setRanges([statusRange])
        .build();
      
      // Create rule for "Error" cells
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('Error')
        .setBackground('#FCE8E6') // Light red
        .setBold(true)
        .setRanges([statusRange])
        .build();
      
      // Add the new rules to the existing rules
      rules.push(modifiedRule);
      rules.push(syncedRule);
      rules.push(errorRule);
      sheet.setConditionalFormatRules(rules);
    }
    
    // Add data validation for the Sync Status column
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Modified', 'Synced'], true)
      .build();
    
    if (lastRow > 1) {
      sheet.getRange(2, targetColumnIndex, lastRow - 1, 1).setDataValidation(validationRule);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error adding Sync Status column: ${e.message}`);
    return false;
  }
}

/**
 * Converts a column index to letter format (e.g., 1 = A, 27 = AA)
 * @param {number} column - The column index (1-based)
 * @return {string} The column letter
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Converts a column letter to index format (e.g., A = 1, AA = 27)
 * @param {string} letter - The column letter
 * @return {number} The column index (1-based)
 */
function columnLetterToIndex(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}