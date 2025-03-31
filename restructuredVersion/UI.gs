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
    
    // Initialize sync status properties
    scriptProperties.setProperty('SYNC_PHASE_1_STATUS', 'active');
    scriptProperties.setProperty('SYNC_PHASE_1_DETAIL', 'Connecting to Pipedrive...');
    scriptProperties.setProperty('SYNC_PHASE_1_PROGRESS', '50');
    
    scriptProperties.setProperty('SYNC_PHASE_2_STATUS', 'pending');
    scriptProperties.setProperty('SYNC_PHASE_2_DETAIL', 'Waiting to start...');
    scriptProperties.setProperty('SYNC_PHASE_2_PROGRESS', '0');
    
    scriptProperties.setProperty('SYNC_PHASE_3_STATUS', 'pending');
    scriptProperties.setProperty('SYNC_PHASE_3_DETAIL', 'Waiting to start...');
    scriptProperties.setProperty('SYNC_PHASE_3_PROGRESS', '0');
    
    scriptProperties.setProperty('SYNC_CURRENT_PHASE', '1');
    scriptProperties.setProperty('SYNC_COMPLETED', 'false');
    scriptProperties.setProperty('SYNC_ERROR', '');
    
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
 * Includes an HTML file with the content of a script file.
 * @param {string} filename - The name of the file to include without the extension
 * @return {string} The content to be included
 */
function include(filename) {
  return HtmlService.createHtmlOutput(getFileContent_(filename)).getContent();
}

/**
 * Gets the content of a script file.
 * @param {string} filename - The name of the file without the extension
 * @return {string} The content of the file
 * @private
 */
function getFileContent_(filename) {
  if (filename === 'TeamManagerUI') {
    return `<style>${TeamManagerUI.getStyles()}</style><script>${TeamManagerUI.getScripts()}</script>`;
  }
  return '';
}

/**
 * Shows the team management UI.
 * @param {boolean} joinOnly - Whether to show only the join team section.
 */
function showTeamManager(joinOnly = false) {
  try {
    // Get the active user's email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to retrieve your email address. Please make sure you are logged in.');
    }

    // Get team data
    const teamAccess = new TeamAccess();
    const hasTeam = teamAccess.isUserInTeam(userEmail);
    let teamName = '';
    let teamId = '';
    let teamMembers = [];
    let userRole = '';

    if (hasTeam) {
      const teamData = teamAccess.getUserTeamData(userEmail);
      teamName = teamData.name;
      teamId = teamData.id;
      teamMembers = teamAccess.getTeamMembers(teamId);
      userRole = teamData.role;
    }

    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('TeamManager');
    
    // Set template variables
    template.userEmail = userEmail;
    template.hasTeam = hasTeam;
    template.teamName = teamName;
    template.teamId = teamId;
    template.teamMembers = teamMembers;
    template.userRole = userRole;
    template.initialTab = joinOnly ? 'join' : (hasTeam ? 'manage' : 'create');
    
    // Make include function available to the template
    template.include = include;
    
    // Evaluate the template
    const htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(hasTeam ? 600 : 400)
      .setTitle(hasTeam ? 'Team Management' : 'Team Access')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Team Access');
  } catch (error) {
    Logger.log(`Error in showTeamManager: ${error.message}`);
    showError('An error occurred while loading the team management interface: ' + error.message);
  }
}

/**
 * Shows an error message to the user in a dialog
 * @param {string} message - The error message to display
 */
function showError(message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', message, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('Error showing error dialog: ' + error.message);
  }
}

/**
 * Process a template string with PHP-like syntax using a data object
 * @param {string} template - Template string with PHP-like syntax
 * @param {Object} data - Data object containing values to substitute
 * @return {string} - Processed template
 */
function processTemplate(template, data) {
  let processedTemplate = template;
  
  // Handle conditionals
  processedTemplate = processIfElse(processedTemplate, data);
  
  // Handle foreach loops
  processedTemplate = processForEach(processedTemplate, data);
  
  // Process simple variable substitutions like <?= variable ?>
  processedTemplate = processedTemplate.replace(/\<\?=\s*([^?>]+?)\s*\?>/g, (match, variable) => {
    try {
      // Handle nested properties like member.email or complex expressions
      const value = evalInContext(variable, data) || '';
      return value;
    } catch (e) {
      Logger.log('Error processing variable: ' + e.message);
      return '';
    }
  });
  
  return processedTemplate;
}

/**
 * Process if/else statements in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
function processIfElse(template, data) {
  let result = template;
  
  // Handle if statements: <?php if (condition): ?>content<?php endif; ?>
  const IF_PATTERN = /\<\?php\s+if\s+\((.+?)\)\s*:\s*\?>([\s\S]*?)(?:\<\?php\s+else\s*:\s*\?>([\s\S]*?))?\<\?php\s+endif;\s*\?>/g;
  
  result = result.replace(IF_PATTERN, (match, condition, ifContent, elseContent = '') => {
    try {
      // Convert PHP-like condition to JavaScript
      let jsCondition = condition
        .replace(/\!\=/g, '!==')
        .replace(/\=\=/g, '===')
        .replace(/\!([\w\.]+)/g, '!$1');
      
      // Evaluate the condition in the context of the data object
      const conditionResult = evalInContext(jsCondition, data);
      
      return conditionResult ? ifContent : elseContent;
    } catch (e) {
      Logger.log('Error processing if condition: ' + e.message);
      return '';
    }
  });
  
  return result;
}

/**
 * Process foreach loops in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
function processForEach(template, data) {
  let result = template;
  
  // Handle foreach loops: <?php foreach ($items as $item): ?>content<?php endforeach; ?>
  const FOREACH_PATTERN = /\<\?php\s+foreach\s+\(\$(\w+)\s+as\s+\$(\w+)\)\s*:\s*\?>([\s\S]*?)\<\?php\s+endforeach;\s*\?>/g;
  
  result = result.replace(FOREACH_PATTERN, (match, collection, item, content) => {
    try {
      const items = data[collection];
      if (!Array.isArray(items) || items.length === 0) {
        return '';
      }
      
      return items.map(itemData => {
        // Create a context with the item for nested variable replacement
        const itemContext = Object.assign({}, data, { [item]: itemData });
        
        // Replace item variables within the loop content
        let itemContent = content;
        
        // Process variables like <?= item.property ?>
        itemContent = itemContent.replace(/\<\?=\s*(\$?)([\w\.]+)\s*\?>/g, (m, dollar, varName) => {
          try {
            // If it's a loop item variable like $member or $member.property
            if (varName.startsWith(item + '.')) {
              const propPath = varName.substring(item.length + 1);
              return evalPropertyPath(itemData, propPath) || '';
            } 
            // If it's just $member (the whole item)
            else if (varName === item) {
              return itemData || '';
            } 
            // Other variables from the parent context
            else {
              return evalInContext(varName, itemContext) || '';
            }
          } catch (e) {
            Logger.log('Error in foreach variable replacement: ' + e.message);
            return '';
          }
        });
        
        return itemContent;
      }).join('');
    } catch (e) {
      Logger.log('Error processing foreach: ' + e.message);
      return '';
    }
  });
  
  return result;
}

/**
 * Evaluate a JavaScript expression in the context of a data object
 * @param {string} expr - The expression to evaluate
 * @param {Object} context - The context object
 * @return {*} - The result of the evaluation
 */
function evalInContext(expr, context) {
  try {
    // Handle simple variable access first
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(expr)) {
      return context[expr];
    }
    
    // Handle nested properties with dot notation
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*(\.[a-zA-Z_$][a-zA-Z0-9_$]*)+$/.test(expr)) {
      return evalPropertyPath(context, expr);
    }
    
    // Handle comparisons and more complex expressions
    // Create a safe function to evaluate in the context
    const keys = Object.keys(context);
    const values = keys.map(key => context[key]);
    const evaluator = new Function(...keys, `return ${expr};`);
    return evaluator(...values);
  } catch (e) {
    Logger.log('Error evaluating expression: ' + e.message);
    return null;
  }
}

/**
 * Evaluate a property path on an object (e.g. "user.profile.name")
 * @param {Object} obj - The object to evaluate on
 * @param {string} path - The property path
 * @return {*} - The value at the path or undefined
 */
function evalPropertyPath(obj, path) {
  try {
    return path.split('.').reduce((o, p) => o && o[p], obj);
  } catch (e) {
    Logger.log('Error evaluating property path: ' + e.message);
    return undefined;
  }
}