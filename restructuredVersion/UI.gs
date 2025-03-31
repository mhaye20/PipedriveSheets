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
    // Direct the function to the new implementation
    showColumnSelectorUI();
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
    
    // List of fields to treat as simple values instead of expanding nested properties
    // This makes the UI cleaner for common fields
    const simplifiedFields = ['owner_id', 'creator_id', 'person_id', 'org_id', 'user_id', 'deal_id'];
    
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
            Logger.log(`No custom field mappings found for ${entityType}, will use field keys as names`);
          }
        } catch (e) {
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
            
            // Format the field name in a user-friendly way
            const fieldName = key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
            
            // Regular top-level property - add to results
            result.push({
              path: key,
              key: key,
              name: fieldName,
              type: typeof sampleItem[key],
              isNested: false
            });
            
            // Check if this is a field we should simplify rather than expand nested properties
            if (simplifiedFields.includes(key)) {
              continue;
            }
            
            // If this property is an object, extract its properties as nested fields
            if (sampleItem[key] && typeof sampleItem[key] === 'object' && !Array.isArray(sampleItem[key])) {
              for (const nestedKey in sampleItem[key]) {
                // Skip functions, internal properties, and numeric indices
                if (typeof sampleItem[key][nestedKey] === 'function' || 
                    nestedKey.startsWith('_') ||
                    !isNaN(parseInt(nestedKey))) continue;
                
                const nestedPath = `${key}.${nestedKey}`;
                if (!result.some(col => col.key === nestedPath)) {
                  // Format the nested field name
                  const nestedFieldName = fieldName + ' ' + nestedKey.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
                  
                  result.push({
                    path: nestedPath,
                    key: nestedPath,
                    name: nestedFieldName,
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
        
        // Check if this is a field we should simplify rather than expand nested properties
        if (simplifiedFields.includes(key)) {
          continue;
        }
        
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
 * Shows the column selector UI for the current sheet and entity type
 */
function showColumnSelectorUI() {
  // Get the current sheet and entity type
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = activeSheet.getName();
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || 'deals';
  
  // Get all available fields (columns) for this entity type
  const availableColumns = getAvailableColumns(entityType);
  
  // Debug log the available columns
  Logger.log(`Available columns for ${entityType}: ${availableColumns.length}`);
  Logger.log(`First 5 available columns: ${JSON.stringify(availableColumns.slice(0, 5))}`);
  
  // Get currently selected columns
  const selectedColumns = SyncService.getTeamAwareColumnPreferences(entityType, sheetName) || [];
  
  // Ensure selected columns are normalized to objects with key and name properties
  const normalizedSelectedColumns = selectedColumns.map(col => {
    if (typeof col === 'string') {
      // Find name for this key in available columns
      const found = availableColumns.find(avail => avail.key === col);
      return {
        key: col,
        name: found ? found.name : col.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
      };
    }
    return col;
  });
  
  // Filter and group available columns
  // Separate into core vs nested fields
  const coreColumns = availableColumns.filter(col => !col.isNested);
  const nestedColumns = availableColumns.filter(col => col.isNested);
  
  // Group nested columns by parent (e.g., owner_id.name and owner_id.email grouped under owner_id)
  const nestedGroups = {};
  nestedColumns.forEach(col => {
    const parent = col.parentKey || 'other';
    if (!nestedGroups[parent]) {
      nestedGroups[parent] = [];
    }
    nestedGroups[parent].push(col);
  });
  
  // Create list items for available columns with groups
  let availableColumnsHtml = '';
  
  // First add core columns
  coreColumns.forEach(col => {
    // Check if this column is already selected
    const isSelected = normalizedSelectedColumns.some(selected => selected.key === col.key);
    if (!isSelected) {
      availableColumnsHtml += `<div class="column-item" data-key="${col.key}" data-name="${col.name}">
        ${col.name} <span class="column-type">(${col.type || 'field'})</span>
      </div>`;
    }
  });
  
  // Then add nested columns with group headers
  Object.keys(nestedGroups).sort().forEach(groupKey => {
    // Get a friendly name for the group
    let groupName = groupKey;
    // For custom fields group, use a friendly name
    if (groupKey === 'custom_fields') {
      groupName = 'Custom Fields';
    } else {
      // Look up the field in available columns to get a friendly name
      const parentField = availableColumns.find(col => col.key === groupKey);
      if (parentField) {
        groupName = parentField.name;
      } else {
        // Try to format the key name as a fallback
        groupName = groupKey.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
      }
    }
    
    // Check if any fields in this group are not already selected
    const unselectedFieldsInGroup = nestedGroups[groupKey].filter(col => 
      !normalizedSelectedColumns.some(selected => selected.key === col.key)
    );
    
    if (unselectedFieldsInGroup.length > 0) {
      // Add a group header
      availableColumnsHtml += `<div class="group-header">${groupName}</div>`;
      
      // Add the fields in this group
      unselectedFieldsInGroup.forEach(col => {
        // Display a simplified name without redundant parent info
        const displayName = col.name.includes(' > ') 
          ? col.name.split(' > ').slice(1).join(' > ')  // Remove parent prefix
          : col.name;
          
        availableColumnsHtml += `<div class="column-item nested-column" data-key="${col.key}" data-name="${col.name}">
          ${displayName} <span class="column-type">(${col.type || 'field'})</span>
        </div>`;
      });
    }
  });
  
  // Create list items for selected columns
  let selectedColumnsHtml = '';
  normalizedSelectedColumns.forEach(col => {
    selectedColumnsHtml += `<div class="column-item selected" data-key="${col.key}" data-name="${col.name}">
      ${col.name} <span class="column-type">(${col.type || 'field'})</span>
    </div>`;
  });
  
  // Calculate counts for display
  const availableCount = availableColumns.length - normalizedSelectedColumns.length;
  const selectedCount = normalizedSelectedColumns.length;
  
  // Build the HTML
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            color: #424242;
          }
          
          .container {
            display: flex;
            height: 100vh;
            max-height: 100vh;
            overflow: hidden;
          }
          
          .column {
            flex: 1;
            padding: 16px;
            display: flex;
            flex-direction: column;
            height: 100%;
            box-sizing: border-box;
            overflow: hidden;
          }
          
          .column-header {
            margin-bottom: 8px;
            padding-bottom: 8px;
            font-weight: bold;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .column-count {
            background-color: #f5f5f5;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            color: #616161;
          }
          
          .search-box {
            padding: 8px;
            margin-bottom: 8px;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            width: 100%;
            box-sizing: border-box;
          }
          
          .column-list {
            flex: 1;
            overflow-y: auto;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            background-color: #ffffff;
          }
          
          .column-item {
            padding: 8px 12px;
            border-bottom: 1px solid #f5f5f5;
            cursor: pointer;
            transition: background-color 0.2s;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .column-item:hover {
            background-color: #f5f5f5;
          }
          
          .column-item.selected {
            position: relative;
          }
          
          .nested-column {
            padding-left: 20px;
            font-size: 0.95em;
          }
          
          .group-header {
            padding: 4px 12px;
            font-weight: bold;
            background-color: #f0f0f0;
            border-bottom: 1px solid #e0e0e0;
            font-size: 0.9em;
            color: #616161;
          }
          
          .column-type {
            font-size: 0.8em;
            color: #9e9e9e;
          }
          
          .buttons {
            margin-top: 16px;
            display: flex;
            justify-content: flex-end;
            gap: 8px;
          }
          
          .btn {
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          
          .btn-primary {
            background-color: #4285f4;
            color: white;
          }
          
          .btn-secondary {
            background-color: #f5f5f5;
            color: #424242;
          }
          
          .btn:hover {
            opacity: 0.9;
          }
          
          .no-columns {
            padding: 16px;
            color: #9e9e9e;
            text-align: center;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="column">
            <div class="column-header">
              Available Columns <span class="column-count">${availableCount}</span>
            </div>
            <input type="text" class="search-box" placeholder="Search available columns..." id="search-available">
            <div class="column-list" id="available-columns">
              ${availableColumnsHtml || '<div class="no-columns">No columns available</div>'}
            </div>
          </div>
          <div class="column">
            <div class="column-header">
              Selected Columns <span class="column-count">${selectedCount}</span>
            </div>
            <input type="text" class="search-box" placeholder="Search selected columns..." id="search-selected">
            <div class="column-list" id="selected-columns">
              ${selectedColumnsHtml || '<div class="no-columns">No columns selected</div>'}
            </div>
            <div class="buttons">
              <button class="btn btn-secondary" id="btn-cancel">Cancel</button>
              <button class="btn btn-primary" id="btn-save">Save & Close</button>
            </div>
          </div>
        </div>
        
        <script>
          // Function to filter items based on search text
          function filterItems(searchText, containerId) {
            const container = document.getElementById(containerId);
            const items = container.getElementsByClassName('column-item');
            const headers = container.getElementsByClassName('group-header');
            const lowerText = searchText.toLowerCase();
            
            let visibleGroupMembers = {};
            
            // First check which items will be visible
            for (let i = 0; i < items.length; i++) {
              const itemText = items[i].innerText.toLowerCase();
              const headerBefore = items[i].previousElementSibling && 
                                   items[i].previousElementSibling.classList.contains('group-header') ? 
                                   items[i].previousElementSibling : null;
              
              if (itemText.includes(lowerText)) {
                items[i].style.display = '';
                
                // Track that this group has visible members
                if (headerBefore) {
                  const headerText = headerBefore.innerText;
                  visibleGroupMembers[headerText] = true;
                }
              } else {
                items[i].style.display = 'none';
              }
            }
            
            // Then show/hide headers based on whether any items in their group are visible
            for (let i = 0; i < headers.length; i++) {
              const headerText = headers[i].innerText;
              headers[i].style.display = visibleGroupMembers[headerText] ? '' : 'none';
            }
          }
          
          // Add event listeners for search boxes
          document.getElementById('search-available').addEventListener('input', function(e) {
            filterItems(e.target.value, 'available-columns');
          });
          
          document.getElementById('search-selected').addEventListener('input', function(e) {
            filterItems(e.target.value, 'selected-columns');
          });
          
          // Get elements
          const availableColumns = document.getElementById('available-columns');
          const selectedColumns = document.getElementById('selected-columns');
          
          // Add click event for available columns (to select them)
          availableColumns.addEventListener('click', function(e) {
            const columnItem = e.target.closest('.column-item');
            if (columnItem) {
              const key = columnItem.getAttribute('data-key');
              const name = columnItem.getAttribute('data-name');
              
              // Move this item to selected columns
              selectedColumns.appendChild(columnItem);
              columnItem.classList.add('selected');
              
              // Update counts
              updateCounts();
            }
          });
          
          // Add click event for selected columns (to unselect them)
          selectedColumns.addEventListener('click', function(e) {
            const columnItem = e.target.closest('.column-item');
            if (columnItem) {
              const key = columnItem.getAttribute('data-key');
              
              // Move back to available columns
              // First find where to insert it alphabetically
              let inserted = false;
              const availableItems = availableColumns.getElementsByClassName('column-item');
              
              // Check if this is a nested column (for special handling)
              const isNested = columnItem.classList.contains('nested-column');
              
              // For nested columns, we need to find or create the appropriate group section
              if (isNested) {
                // Extract the parent group from the key (usually the part before the first dot)
                const parentKey = key.split('.')[0];
                
                // Check if there's already a group header for this parent
                let groupHeader = null;
                let insertAfterHeader = availableColumns.firstChild;
                
                // Look for the group header with this parent's name
                const headers = availableColumns.getElementsByClassName('group-header');
                for (let i = 0; i < headers.length; i++) {
                  if (headers[i].innerText.toLowerCase().includes(parentKey.replace(/_/g, ' '))) {
                    groupHeader = headers[i];
                    break;
                  }
                }
                
                // If we found a group header, insert after it
                if (groupHeader) {
                  let sibling = groupHeader.nextSibling;
                  while (sibling && sibling.classList && sibling.classList.contains('nested-column')) {
                    const siblingKey = sibling.getAttribute('data-key');
                    if (key < siblingKey) {
                      availableColumns.insertBefore(columnItem, sibling);
                      inserted = true;
                      break;
                    }
                    sibling = sibling.nextSibling;
                  }
                  
                  // If we didn't find a place within this group, add at the end of the group
                  if (!inserted) {
                    if (sibling) {
                      availableColumns.insertBefore(columnItem, sibling);
                    } else {
                      availableColumns.appendChild(columnItem);
                    }
                    inserted = true;
                  }
                }
              }
              
              // For non-nested or if we couldn't find the appropriate group
              if (!inserted) {
                // Just insert alphabetically among the available columns
                for (let i = 0; i < availableItems.length; i++) {
                  const availableKey = availableItems[i].getAttribute('data-key');
                  if (key < availableKey) {
                    availableColumns.insertBefore(columnItem, availableItems[i]);
                    inserted = true;
                    break;
                  }
                }
                
                // If we couldn't find a place to insert, add at the end
                if (!inserted) {
                  availableColumns.appendChild(columnItem);
                }
              }
              
              // Update counts
              updateCounts();
            }
          });
          
          // Function to update the counts in the UI
          function updateCounts() {
            const availableCount = availableColumns.getElementsByClassName('column-item').length;
            const selectedCount = selectedColumns.getElementsByClassName('column-item').length;
            
            document.querySelector('.column:first-child .column-count').textContent = availableCount;
            document.querySelector('.column:last-child .column-count').textContent = selectedCount;
          }
          
          // Add event listener for save button
          document.getElementById('btn-save').addEventListener('click', function() {
            // Collect selected columns
            const selectedItems = selectedColumns.getElementsByClassName('column-item');
            const selectedData = [];
            
            for (let i = 0; i < selectedItems.length; i++) {
              selectedData.push({
                key: selectedItems[i].getAttribute('data-key'),
                name: selectedItems[i].getAttribute('data-name')
              });
            }
            
            // Return data to server
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .saveColumnPreferences(selectedData, "${entityType}", "${sheetName}");
          });
          
          // Add event listener for cancel button
          document.getElementById('btn-cancel').addEventListener('click', function() {
            // Close the dialog without saving
            google.script.host.close();
          });
        </script>
      </body>
    </html>
  `;
  
  // Calculate dialog dimensions
  const dialogWidth = 800;
  const dialogHeight = 600;
  
  // Show the HTML in a modal dialog
  const title = `Select Columns for ${entityType.charAt(0).toUpperCase() + entityType.slice(1)} in '${sheetName}'`;
  const ui = HtmlService.createHtmlOutput(html)
    .setWidth(dialogWidth)
    .setHeight(dialogHeight);
  
  SpreadsheetApp.getUi().showModalDialog(ui, title);
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
 * @param {Array} columns - Array of column objects or keys to save
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
    
    // Store full column objects in user-specific property
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

/**
 * Gets all available columns for a given entity type
 * @param {string} entityType - The entity type to get columns for (deals, persons, etc.)
 * @return {Array} Array of column objects
 */
function getAvailableColumns(entityType) {
  try {
    Logger.log(`Getting available columns for ${entityType}`);
    
    // Get sample data to extract fields from
    let sampleData = [];
    let fields = [];
    
    // First try to get sample data to determine available fields
    try {
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          sampleData = getDealsWithFilter('', 1);
          break;
        case ENTITY_TYPES.PERSONS:
          sampleData = getPersonsWithFilter('', 1);
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          sampleData = getOrganizationsWithFilter('', 1);
          break;
        case ENTITY_TYPES.ACTIVITIES:
          sampleData = getActivitiesWithFilter('', 1);
          break;
        case ENTITY_TYPES.LEADS:
          sampleData = getLeadsWithFilter('', 1);
          break;
        case ENTITY_TYPES.PRODUCTS:
          sampleData = getProductsWithFilter('', 1);
          break;
      }
      
      if (sampleData && sampleData.length > 0) {
        Logger.log(`Successfully retrieved sample data for ${entityType}`);
      }
    } catch (e) {
      Logger.log(`Error getting sample data: ${e.message}`);
    }
    
    // If we got sample data, extract fields from it
    if (sampleData && sampleData.length > 0) {
      return extractFields(sampleData);
    }
    
    // If no sample data, try to get field definitions
    try {
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          fields = getDealFields(true);
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
        Logger.log(`Got ${fields.length} field definitions for ${entityType}`);
        
        // Convert field definitions to column objects
        return fields.map(field => {
          return {
            key: field.key,
            path: field.key,
            name: field.name || field.label || field.key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase()),
            type: field.field_type || 'text',
            isNested: false
          };
        });
      }
    } catch (e) {
      Logger.log(`Error getting field definitions: ${e.message}`);
    }
    
    // If all else fails, return common fields for this entity type
    return getCommonFieldsForEntity(entityType);
  } catch (e) {
    Logger.log(`Error in getAvailableColumns: ${e.message}`);
    return getCommonFieldsForEntity(entityType);
  }
}

/**
 * Gets common fields for an entity type as a fallback
 * @param {string} entityType - The entity type
 * @return {Array} Array of column objects
 */
function getCommonFieldsForEntity(entityType) {
  // Define common fields by entity type
  const commonFields = {
    [ENTITY_TYPES.DEALS]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'title', name: 'Title', type: 'text' },
      { key: 'value', name: 'Value', type: 'number' },
      { key: 'currency', name: 'Currency', type: 'text' },
      { key: 'status', name: 'Status', type: 'text' },
      { key: 'stage_id', name: 'Stage', type: 'integer' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' },
      { key: 'owner_id', name: 'Owner', type: 'user' },
      { key: 'person_id', name: 'Person', type: 'person' },
      { key: 'org_id', name: 'Organization', type: 'organization' }
    ],
    [ENTITY_TYPES.PERSONS]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'name', name: 'Name', type: 'text' },
      { key: 'email', name: 'Email', type: 'text' },
      { key: 'phone', name: 'Phone', type: 'text' },
      { key: 'owner_id', name: 'Owner', type: 'user' },
      { key: 'org_id', name: 'Organization', type: 'organization' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' }
    ],
    [ENTITY_TYPES.ORGANIZATIONS]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'name', name: 'Name', type: 'text' },
      { key: 'address', name: 'Address', type: 'text' },
      { key: 'owner_id', name: 'Owner', type: 'user' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' }
    ],
    [ENTITY_TYPES.ACTIVITIES]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'subject', name: 'Subject', type: 'text' },
      { key: 'type', name: 'Type', type: 'text' },
      { key: 'due_date', name: 'Due Date', type: 'date' },
      { key: 'deal_id', name: 'Deal', type: 'deal' },
      { key: 'person_id', name: 'Person', type: 'person' },
      { key: 'org_id', name: 'Organization', type: 'organization' },
      { key: 'user_id', name: 'User', type: 'user' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' }
    ],
    [ENTITY_TYPES.LEADS]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'title', name: 'Title', type: 'text' },
      { key: 'owner_id', name: 'Owner', type: 'user' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' }
    ],
    [ENTITY_TYPES.PRODUCTS]: [
      { key: 'id', name: 'ID', type: 'integer' },
      { key: 'name', name: 'Name', type: 'text' },
      { key: 'code', name: 'Code', type: 'text' },
      { key: 'unit', name: 'Unit', type: 'text' },
      { key: 'price', name: 'Price', type: 'number' },
      { key: 'add_time', name: 'Created Date', type: 'date' },
      { key: 'update_time', name: 'Updated Date', type: 'date' }
    ]
  };
  
  // Return common fields or default fields if not found
  return commonFields[entityType] || [
    { key: 'id', name: 'ID', type: 'integer' },
    { key: 'name', name: 'Name', type: 'text' },
    { key: 'add_time', name: 'Created Date', type: 'date' },
    { key: 'update_time', name: 'Updated Date', type: 'date' }
  ];
}