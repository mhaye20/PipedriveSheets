/**
 * Column Selector UI Helper
 * 
 * This module provides helper functions for the column selector UI:
 * - Styles for the column selector dialog
 * - Scripts for the column selector dialog
 */

var ColumnSelectorUI = ColumnSelectorUI || {};

/**
 * Shows the column selector UI
 */
function showColumnSelectorUI() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get sheet-specific settings
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;

    const filterId = scriptProperties.getProperty(sheetFilterIdKey);
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    Logger.log(`Getting data for sheet "${activeSheetName}" with entityType: ${entityType}, filterId: ${filterId}`);

    // Check if we can connect to Pipedrive
    SpreadsheetApp.getActiveSpreadsheet().toast(`Connecting to Pipedrive to retrieve ${entityType} data...`);

    // First verify we have a valid access token
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    if (!accessToken) {
      throw new Error('Not connected to Pipedrive. Please connect your account first.');
    }

    // Get sample data based on entity type (1 item only)
    let sampleData = [];
    try {
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          sampleData = PipedriveAPI.getDealsWithFilter(filterId, 1);
          break;
        case ENTITY_TYPES.PERSONS:
          sampleData = PipedriveAPI.getPersonsWithFilter(filterId, 1);
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          sampleData = PipedriveAPI.getOrganizationsWithFilter(filterId, 1);
          break;
        case ENTITY_TYPES.ACTIVITIES:
          sampleData = PipedriveAPI.getActivitiesWithFilter(filterId, 1);
          break;
        case ENTITY_TYPES.LEADS:
          sampleData = PipedriveAPI.getLeadsWithFilter(filterId, 1);
          break;
        case ENTITY_TYPES.PRODUCTS:
          sampleData = PipedriveAPI.getProductsWithFilter(filterId, 1);
          break;
      }
      Logger.log(`Retrieved ${sampleData ? sampleData.length : 0} items for ${entityType}`);
    } catch (apiError) {
      Logger.log(`Error retrieving ${entityType} data: ${apiError.message}`);
      throw new Error(`Failed to retrieve ${entityType} data from Pipedrive: ${apiError.message}`);
    }

    if (!sampleData || sampleData.length === 0) {
      // Try to get available filters to provide better error message
      const filters = PipedriveAPI.getFiltersForEntityType(entityType);
      if (!filters || filters.length === 0) {
        throw new Error(`No filters found for ${entityType}. Please create a filter in Pipedrive first.`);
      } else if (!filterId) {
        throw new Error(`No filter selected for ${entityType}. Please configure filter settings first.`);
      }
      throw new Error(`No ${entityType} data found for the selected filter. Please check your filter settings.`);
    }

    const sampleItem = sampleData[0];

    // Get field mappings for this entity type
    const fieldMap = getCustomFieldMappings(entityType);

    // Extract fields from sample data
    function extractFields(obj, parentPath = '', parentName = '') {
      // Special handling for custom_fields in API v2
      if (parentPath === 'custom_fields') {
        Logger.log('Processing custom_fields object');
        
        for (const key in obj) {
          if (obj.hasOwnProperty(key) && !addedCustomFields.has(key)) {
            addedCustomFields.add(key);
            const customFieldValue = obj[key];
            const currentPath = `custom_fields.${key}`;
            const displayName = fieldMap[key] || formatColumnName(key);
            
            // Simple value custom fields
            if (typeof customFieldValue !== 'object' || customFieldValue === null) {
              availableColumns.push({
                key: currentPath,
                name: displayName,
                isNested: true,
                parentKey: 'custom_fields'
              });
              continue;
            }
            
            // Handle complex custom fields (object-based)
            if (typeof customFieldValue === 'object') {
              // Currency fields
              if (customFieldValue.value !== undefined && customFieldValue.currency !== undefined) {
                availableColumns.push({
                  key: currentPath,
                  name: `${displayName} (Currency)`,
                  isNested: true,
                  parentKey: 'custom_fields'
                });
              }
              // Date/Time range fields
              else if (customFieldValue.value !== undefined && customFieldValue.until !== undefined) {
                availableColumns.push({
                  key: currentPath,
                  name: `${displayName} (Range)`,
                  isNested: true,
                  parentKey: 'custom_fields'
                });
              }
              // Address fields
              else if (customFieldValue.value !== undefined && customFieldValue.formatted_address !== undefined) {
                availableColumns.push({
                  key: currentPath,
                  name: `${displayName} (Address)`,
                  isNested: true,
                  parentKey: 'custom_fields'
                });
                
                // Add formatted address as a separate column option
                availableColumns.push({
                  key: `${currentPath}.formatted_address`,
                  name: `${displayName} (Formatted Address)`,
                  isNested: true,
                  parentKey: currentPath
                });
              }
              // For all other object types
              else {
                availableColumns.push({
                  key: currentPath,
                  name: `${displayName} (Complex)`,
                  isNested: true,
                  parentKey: 'custom_fields'
                });
                
                // Extract nested fields from complex custom field
                extractFields(customFieldValue, currentPath, `Custom Field: ${key}`);
              }
            }
          }
        }
        return;
      }
      
      // Handle arrays
      if (Array.isArray(obj)) {
        Logger.log(`Processing array with ${obj.length} items, parent: ${parentPath}`);
        
        // For arrays of structured objects, like emails or phones
        if (obj.length > 0 && typeof obj[0] === 'object' && obj[0] !== null) {
          // Special handling for email/phone arrays
          if (obj[0].hasOwnProperty('value') && obj[0].hasOwnProperty('primary')) {
            let displayName = 'Primary ' + (parentName || 'Item');
            
            if (obj[0].hasOwnProperty('label')) {
              // For contact fields like email/phone
              if ((parentPath === 'email' || parentPath === 'phone') && obj[0].label) {
                // Add primary version first
                availableColumns.push({
                  key: `${parentPath}.0.value`,
                  name: `Primary ${formatColumnName(parentPath)}`,
                  isNested: true,
                  parentKey: parentPath
                });
                
                // Then add specific types
                const typeLabels = new Set();
                obj.forEach(item => {
                  if (item && item.label && !typeLabels.has(item.label.toLowerCase())) {
                    typeLabels.add(item.label.toLowerCase());
                    
                    // Create specific field like "Email Work" or "Phone Mobile"
                    const itemLabel = formatColumnName(item.label);
                    const columnKey = `${parentPath}.${item.label.toLowerCase()}`;
                    const columnName = `${formatColumnName(parentPath)} ${itemLabel}`;
                    
                    availableColumns.push({
                      key: columnKey,
                      name: columnName,
                      isNested: true,
                      parentKey: parentPath
                    });
                  }
                });
                
                return;
              }
            }
          } else {
            // For other object arrays, process the first item
            extractFields(obj[0], parentPath + '.0', parentName + ' (First Item)');
          }
        }
        return;
      }
      
      // Extract properties from this object
      for (const key in obj) {
        // Skip internal properties, functions, or empty objects
        if (key.startsWith('_') || typeof obj[key] === 'function') {
          continue;
        }
        
        const currentPath = parentPath ? parentPath + '.' + key : key;
        // Use field map for display name if available
        let displayName = parentName ? 
          parentName + ' ' + (fieldMap[key] || formatColumnName(key)) : 
          (fieldMap[key] || formatColumnName(key));
        
        // Special handling for name in person/org objects
        if (key === 'name' && parentPath && 
           (obj.hasOwnProperty('email') || obj.hasOwnProperty('phone') || obj.hasOwnProperty('address'))) {
          availableColumns.push({
            key: currentPath,
            name: (parentName || formatColumnName(parentPath)) + ' Name',
            isNested: true,
            parentKey: parentPath
          });
        } 
        else if (typeof obj[key] === 'object' && obj[key] !== null) {
          // Skip adding email/phone as separate objects - we'll handle them specially
          if (parentPath === '' && (key === 'email' || key === 'phone')) {
            extractFields(obj[key], key, displayName);
          } else {
            // Recursively extract from nested objects
            extractFields(obj[key], currentPath, displayName);
          }
        } 
        else {
          // Simple property
          availableColumns.push({
            key: currentPath,
            name: displayName,
            isNested: parentPath ? true : false,
            parentKey: parentPath
          });
        }
      }
    }

    // Sample data and field mapping are ready, extract available columns
    Logger.log(`Beginning field extraction from sample data`);
    let availableColumns = [];
    
    try {
      // Get field mappings for better display names
      const fieldMap = {};
      let fieldDefinitions = [];
      
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          fieldDefinitions = PipedriveAPI.getDealFields();
          break;
        case ENTITY_TYPES.PERSONS:
          fieldDefinitions = PipedriveAPI.getPersonFields();
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          fieldDefinitions = PipedriveAPI.getOrganizationFields();
          break;
        case ENTITY_TYPES.ACTIVITIES:
          fieldDefinitions = PipedriveAPI.getActivityFields();
          break;
        case ENTITY_TYPES.LEADS:
          fieldDefinitions = PipedriveAPI.getLeadFields();
          break;
        case ENTITY_TYPES.PRODUCTS:
          fieldDefinitions = PipedriveAPI.getProductFields();
          break;
      }
      
      // Create a mapping of field keys to field names
      fieldDefinitions.forEach(field => {
        fieldMap[field.key] = field.name;
      });
      
      Logger.log(`Got ${fieldDefinitions.length} field definitions and created field map`);
      
      // Create a record of already added custom fields to avoid duplicates
      const addedCustomFields = new Set();
      
      // First, build top-level column data
      Logger.log(`Building top-level columns first`);
      for (const key in sampleItem) {
        // Skip internal properties or functions
        if (key.startsWith('_') || typeof sampleItem[key] === 'function') {
          continue;
        }
        
        // Skip email and phone at top level - handled specially
        if (key === 'email' || key === 'phone') {
          continue;
        }
        
        const displayName = fieldMap[key] || formatColumnName(key);
        
        // Add the top-level column
        availableColumns.push({
          key: key,
          name: displayName,
          isNested: false
        });
      }
      
      // Special handling for top-level email/phone fields
      if (sampleItem.email) {
        availableColumns.push({
          key: 'email',
          name: 'Email',
          isNested: false
        });
        
        extractFields(sampleItem.email, 'email', 'Email');
      }
      
      if (sampleItem.phone) {
        availableColumns.push({
          key: 'phone',
          name: 'Phone',
          isNested: false
        });
        
        extractFields(sampleItem.phone, 'phone', 'Phone');
      }
      
      // Then extract nested fields for all complex objects
      for (const key in sampleItem) {
        if (key.startsWith('_') || typeof sampleItem[key] === 'function') {
          continue;
        }
        
        // Skip email/phone - already handled
        if (key === 'email' || key === 'phone') {
          continue;
        }
        
        // If it's a complex object, extract nested fields
        if (typeof sampleItem[key] === 'object' && sampleItem[key] !== null) {
          const displayName = fieldMap[key] || formatColumnName(key);
          extractFields(sampleItem[key], key, displayName);
        }
      }
      
      Logger.log(`Extracted ${availableColumns.length} available columns from sample data`);
      
      // Remove any duplicate fields
      const keySet = new Set();
      availableColumns = availableColumns.filter(col => {
        if (keySet.has(col.key)) {
          return false;
        }
        keySet.add(col.key);
        return true;
      });
      
      Logger.log(`After deduplication: ${availableColumns.length} columns`);
    } catch (extractError) {
      Logger.log(`Error during field extraction: ${extractError.message}`);
      Logger.log(`Error stack: ${extractError.stack}`);
      
      // Create some minimal columns as fallback
      availableColumns = [
        { key: 'id', name: 'ID', isNested: false, parentKey: null },
        { key: 'name', name: 'Name', isNested: false, parentKey: null }
      ];
      
      if (entityType === ENTITY_TYPES.PERSONS) {
        availableColumns.push(
          { key: 'email', name: 'Email', isNested: false, parentKey: null },
          { key: 'phone', name: 'Phone', isNested: false, parentKey: null }
        );
      }
      
      Logger.log(`Using ${availableColumns.length} fallback columns after extraction error`);
    }
    
    // Filter out problematic fields
    availableColumns = availableColumns.filter(col => {
      // Remove problematic fields like "im" and "lm"
      return !col.key.startsWith('im.') && 
             !col.key.startsWith('lm.') &&
             col.key !== 'im' && 
             col.key !== 'lm';
    });
    
    // Log the structure of email and phone fields if they exist
    if (sampleItem.email) {
      Logger.log(`Email field structure: ${JSON.stringify(sampleItem.email)}`);
    }
    if (sampleItem.phone) {
      Logger.log(`Phone field structure: ${JSON.stringify(sampleItem.phone)}`);
    }
    
    // Sort columns to prioritize important fields
    availableColumns.sort((a, b) => {
      // ID always comes first
      if (a.key === 'id') return -1;
      if (b.key === 'id') return 1;
      
      // Then name
      if (a.key === 'name') return -1;
      if (b.key === 'name') return 1;
      
      // Top-level fields before nested
      if (!a.isNested && b.isNested) return -1;
      if (a.isNested && !b.isNested) return 1;
      
      // Special fields like email and phone should come early
      if (a.key === 'email' || a.key === 'phone') return -1;
      if (b.key === 'email' || b.key === 'phone') return 1;
      
      // For nested fields, sort by parent first
      if (a.isNested && b.isNested && a.parentKey !== b.parentKey) {
        return a.parentKey.localeCompare(b.parentKey);
      }
      
      // Then by display name
      return a.name.localeCompare(b.name);
    });
    
    // Log the first 10 available columns after sorting to see their structure
    Logger.log(`First 10 available columns after sorting:`);
    availableColumns.slice(0, 10).forEach((col, index) => {
      Logger.log(`Column ${index + 1}: key=${col.key}, name=${col.name}, isNested=${col.isNested}, parentKey=${col.parentKey || 'null'}`);
    });
    
    // Now get the currently saved columns (if any)
    let savedColumns = getColumnPreferences(entityType, activeSheetName);
    
    // If no columns are saved yet, use defaults
    if (!savedColumns || savedColumns.length === 0) {
      Logger.log(`No saved columns found, using defaults for ${entityType}`);
      
      // Use entity-specific defaults if available, otherwise use common defaults
      const defaultKeys = DEFAULT_COLUMNS[entityType.toUpperCase()] || DEFAULT_COLUMNS.COMMON;
      
      // Find matching columns from available columns
      savedColumns = availableColumns.filter(col => defaultKeys.includes(col.key));
      
      // For persons, also add email and phone fields if available
      if (entityType === ENTITY_TYPES.PERSONS) {
        // Add all email fields
        const emailColumns = availableColumns.filter(col => col.key && (col.key === 'email' || col.key.startsWith('email.')));
        savedColumns = savedColumns.concat(emailColumns);
        
        // Add all phone fields
        const phoneColumns = availableColumns.filter(col => col.key && (col.key === 'phone' || col.key.startsWith('phone.')));
        savedColumns = savedColumns.concat(phoneColumns);
      }
      
      Logger.log(`Using ${savedColumns.length} default columns`);
    } else {
      Logger.log(`Found ${savedColumns.length} saved columns`);
    }
    
    // Create HTML for the UI
    const html = HtmlService.createHtmlOutput(renderColumnSelectorHtml(availableColumns, savedColumns, entityType, activeSheetName))
      .setTitle('Select Columns')
      .setWidth(800)
      .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Columns to Display');
  } catch (error) {
    Logger.log('Error in showColumnSelectorUI: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
    throw error;
  }
}

/**
 * Shows the column selector dialog
 */
ColumnSelectorUI.showColumnSelector = function() {
  try {
    // Implementation that directly calls showColumnSelectorUI
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (!activeSheet) {
      throw new Error('No active sheet found. Please select a sheet first.');
    }
    
    showColumnSelectorUI();
  } catch (e) {
    Logger.log(`Error in ColumnSelectorUI.showColumnSelector: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

/**
 * Gets the styles for the column selector dialog
 * @return {string} CSS styles
 */
ColumnSelectorUI.getStyles = function() {
  return `
    <style>
      :root {
        --primary-color: #4285f4;
        --primary-dark: #3367d6;
        --success-color: #0f9d58;
        --warning-color: #f4b400;
        --error-color: #db4437;
        --text-dark: #202124;
        --text-light: #5f6368;
        --bg-light: #f8f9fa;
        --border-color: #dadce0;
        --shadow: 0 1px 3px rgba(60,64,67,0.15);
        --shadow-hover: 0 4px 8px rgba(60,64,67,0.2);
        --transition: all 0.2s ease;
      }
      
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      
      body {
        font-family: 'Roboto', Arial, sans-serif;
        color: var(--text-dark);
        line-height: 1.5;
        margin: 0;
        padding: 16px;
        font-size: 14px;
      }
      
      .header {
        margin-bottom: 16px;
      }
      
      h3 {
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 12px;
        color: var(--text-dark);
      }
      
      .sheet-info {
        background-color: var(--bg-light);
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 14px;
        border-left: 4px solid var(--primary-color);
        display: flex;
        align-items: center;
      }
      
      .sheet-info svg {
        margin-right: 12px;
        fill: var(--primary-color);
      }
      
      .info {
        font-size: 13px;
        color: var(--text-light);
        margin-bottom: 16px;
      }
      
      .container {
        display: flex;
        height: 400px;
        gap: 16px;
      }
      
      .column {
        width: 50%;
        display: flex;
        flex-direction: column;
      }
      
      .column-header {
        font-weight: 500;
        margin-bottom: 8px;
        padding: 0 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .column-count {
        font-size: 12px;
        color: var(--text-light);
        font-weight: normal;
      }
      
      .search {
        padding: 10px 12px;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        margin-bottom: 12px;
        font-size: 14px;
        width: 100%;
        transition: var(--transition);
      }
      
      .search:focus {
        outline: none;
        border-color: var(--primary-color);
        box-shadow: 0 0 0 2px rgba(66,133,244,0.2);
      }
      
      .scrollable {
        flex-grow: 1;
        overflow-y: auto;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        background-color: white;
      }
      
      .item {
        padding: 8px 12px;
        margin: 2px 4px;
        cursor: pointer;
        border-radius: 4px;
        transition: var(--transition);
        display: flex;
        align-items: center;
      }
      
      .item:hover {
        background-color: rgba(66,133,244,0.08);
      }
      
      .item.selected {
        background-color: #e8f0fe;
        position: relative;
      }
      
      .item.selected:hover {
        background-color: #d2e3fc;
      }
      
      .category {
        font-weight: 500;
        padding: 8px 12px;
        background-color: var(--bg-light);
        margin: 8px 4px 4px 4px;
        border-radius: 4px;
        color: var(--text-dark);
        font-size: 13px;
      }
      
      .nested {
        margin-left: 16px;
        position: relative;
      }
      
      .nested::before {
        content: '';
        position: absolute;
        left: -8px;
        top: 50%;
        width: 6px;
        height: 1px;
        background-color: var(--border-color);
      }
      
      .footer {
        margin-top: 16px;
        display: flex;
        justify-content: space-between;
      }
      
      button {
        padding: 10px 16px;
        border: none;
        border-radius: 4px;
        font-weight: 500;
        cursor: pointer;
        font-size: 14px;
        transition: var(--transition);
      }
      
      button:focus {
        outline: none;
      }
      
      button:disabled {
        opacity: 0.7;
        cursor: not-allowed;
      }
      
      .primary-btn {
        background-color: var(--primary-color);
        color: white;
      }
      
      .primary-btn:hover {
        background-color: var(--primary-dark);
        box-shadow: var(--shadow-hover);
      }
      
      .secondary-btn {
        background-color: transparent;
        color: var(--primary-color);
      }
      
      .secondary-btn:hover {
        background-color: rgba(66,133,244,0.08);
      }
      
      .action-btns {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      
      .drag-handle {
        display: inline-block;
        width: 12px;
        height: 20px;
        background-image: radial-gradient(circle, var(--text-light) 1px, transparent 1px);
        background-size: 3px 3px;
        background-position: center;
        background-repeat: repeat;
        margin-right: 8px;
        cursor: grab;
        opacity: 0.5;
      }
      
      .selected:hover .drag-handle {
        opacity: 0.8;
      }
      
      .column-text {
        flex-grow: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      
      .add-btn, .remove-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
      }
      
      .add-btn {
        color: var(--success-color);
        background-color: rgba(15,157,88,0.1);
      }
      
      .remove-btn {
        color: var(--error-color);
        background-color: rgba(219,68,55,0.1);
      }
      
      .rename-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
        color: var(--primary-color);
        background-color: rgba(66,133,244,0.1);
      }
      
      .item:hover .add-btn,
      .item:hover .remove-btn,
      .item:hover .rename-btn {
        opacity: 1;
      }
      
      .loading {
        display: none;
        align-items: center;
        margin-right: 8px;
      }
      
      .loader {
        display: inline-block;
        width: 18px;
        height: 18px;
        border: 2px solid rgba(66,133,244,0.2);
        border-radius: 50%;
        border-top-color: var(--primary-color);
        animation: spin 1s ease-in-out infinite;
      }
      
      .dragging {
        opacity: 0.4;
        box-shadow: var(--shadow-hover);
      }
      
      .over {
        border-top: 2px solid var(--primary-color);
      }
      
      .indicator {
        display: none;
        padding: 12px 16px;
        border-radius: 4px;
        margin-bottom: 16px;
        font-weight: 500;
      }
      
      .indicator.success {
        background-color: rgba(15,157,88,0.1);
        color: var(--success-color);
        border-left: 4px solid var(--success-color);
      }
      
      .indicator.error {
        background-color: rgba(219,68,55,0.1);
        color: var(--error-color);
        border-left: 4px solid var(--error-color);
      }
      
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
  `;
};

/**
 * Gets the scripts for the column selector dialog
 * @return {string} JavaScript code
 */
ColumnSelectorUI.getScripts = function() {
  return `
    <script>
      // DOM elements
      const availableList = document.getElementById('availableList');
      const selectedList = document.getElementById('selectedList');
      const searchBox = document.getElementById('searchBox');
      const availableCountEl = document.getElementById('availableCount');
      const selectedCountEl = document.getElementById('selectedCount');
      
      // Create a deep copy of the original available columns to ensure we don't lose data
      const originalAvailableColumns = JSON.parse(JSON.stringify(availableColumns));
      
      // Render the lists
      function renderAvailableList(searchTerm = '') {
        availableList.innerHTML = '';
        
        // Group columns by parent key or top-level
        const topLevel = [];
        const nested = {};
        let availableCount = 0;
        
        // Always use the original available columns array as source of truth
        originalAvailableColumns.forEach(col => {
          // Check if this column is already selected
          if (!selectedColumns.some(selected => selected.key === col.key)) {
            availableCount++;
            
            if (!searchTerm || col.name.toLowerCase().includes(searchTerm.toLowerCase())) {
              if (!col.isNested) {
                topLevel.push(col);
              } else {
                const parentKey = col.parentKey || 'unknown';
                if (!nested[parentKey]) {
                  nested[parentKey] = [];
                }
                nested[parentKey].push(col);
              }
            }
          }
        });
        
        // Update available count
        availableCountEl.textContent = '(' + availableCount + ')';
        
        // Add top-level columns first
        if (topLevel.length > 0) {
          const topLevelHeader = document.createElement('div');
          topLevelHeader.className = 'category';
          topLevelHeader.textContent = 'Main Fields';
          availableList.appendChild(topLevelHeader);
          
          topLevel.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.dataset.key = col.key;
            
            const columnText = document.createElement('span');
            columnText.className = 'column-text';
            columnText.textContent = col.name;
            item.appendChild(columnText);
            
            const addBtn = document.createElement('span');
            addBtn.className = 'add-btn';
            addBtn.innerHTML = '+';
            addBtn.title = 'Add column';
            addBtn.onclick = (e) => {
              e.stopPropagation();
              addColumn(col);
            };
            item.appendChild(addBtn);
            
            item.onclick = () => addColumn(col);
            availableList.appendChild(item);
          });
        }
        
        // Then add nested columns by parent
        for (const parentKey in nested) {
          if (nested[parentKey].length > 0) {
            // Find the parent name from the original available columns
            const parentName = originalAvailableColumns.find(col => col.key === parentKey)?.name || parentKey;
            
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'category';
            categoryHeader.textContent = parentName;
            availableList.appendChild(categoryHeader);
            
            nested[parentKey].forEach(col => {
              const item = document.createElement('div');
              item.className = 'item nested';
              item.dataset.key = col.key;
              
              const columnText = document.createElement('span');
              columnText.className = 'column-text';
              columnText.textContent = col.name;
              item.appendChild(columnText);
              
              const addBtn = document.createElement('span');
              addBtn.className = 'add-btn';
              addBtn.innerHTML = '+';
              addBtn.title = 'Add column';
              addBtn.onclick = (e) => {
                e.stopPropagation();
                addColumn(col);
              };
              item.appendChild(addBtn);
              
              item.onclick = () => addColumn(col);
              availableList.appendChild(item);
            });
          }
        }
        
        // Show "no results" message if nothing matches search
        if (availableList.children.length === 0 && searchTerm) {
          const noResults = document.createElement('div');
          noResults.style.padding = '16px';
          noResults.style.textAlign = 'center';
          noResults.style.color = 'var(--text-light)';
          noResults.textContent = 'No matching columns found';
          availableList.appendChild(noResults);
        }
      }
      
      function renderSelectedList() {
        selectedList.innerHTML = '';
        
        // Update selected count
        selectedCountEl.textContent = '(' + selectedColumns.length + ')';
        
        if (selectedColumns.length === 0) {
          const emptyState = document.createElement('div');
          emptyState.style.padding = '16px';
          emptyState.style.textAlign = 'center';
          emptyState.style.color = 'var(--text-light)';
          emptyState.innerHTML = 'No columns selected yet<br>Select columns from the left panel';
          selectedList.appendChild(emptyState);
          return;
        }
        
        selectedColumns.forEach((col, index) => {
          const item = document.createElement('div');
          item.className = 'item selected';
          item.dataset.key = col.key;
          item.dataset.index = index;
          item.draggable = true;
          
          // Add drag handle
          const dragHandle = document.createElement('span');
          dragHandle.className = 'drag-handle';
          item.appendChild(dragHandle);
          
          // Add column name
          const columnText = document.createElement('span');
          columnText.className = 'column-text';
          columnText.textContent = col.customName || col.name;
          if (col.customName) {
            columnText.title = "Original field: " + col.name;
            columnText.style.fontStyle = 'italic';
          }
          item.appendChild(columnText);
          
          // Add remove button
          const removeBtn = document.createElement('span');
          removeBtn.className = 'remove-btn';
          removeBtn.innerHTML = '❌';
          removeBtn.title = 'Remove column';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeColumn(col);
          };
          item.appendChild(removeBtn);
          
          // Add rename button
          const renameBtn = document.createElement('span');
          renameBtn.className = 'rename-btn';
          renameBtn.innerHTML = '✏️';
          renameBtn.title = 'Rename column';
          renameBtn.onclick = (e) => {
            e.stopPropagation();
            renameColumn(index);
          };
          item.appendChild(renameBtn);
          
          // Set up drag events
          item.ondragstart = handleDragStart;
          item.ondragover = handleDragOver;
          item.ondrop = handleDrop;
          item.ondragend = handleDragEnd;
          
          selectedList.appendChild(item);
        });
      }
      
      // Drag and drop functionality
      let draggedItem = null;
      
      function handleDragStart(e) {
        draggedItem = this;
        this.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', this.dataset.index);
        
        // Add a small delay to make the visual change noticeable
        setTimeout(() => {
          this.style.opacity = '0.4';
        }, 0);
      }
      
      function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this.classList.add('over');
      }
      
      function handleDrop(e) {
        e.preventDefault();
        this.classList.remove('over');
        
        const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
        const toIndex = parseInt(this.dataset.index);
        
        if (fromIndex !== toIndex) {
          const item = selectedColumns[fromIndex];
          selectedColumns.splice(fromIndex, 1);
          selectedColumns.splice(toIndex, 0, item);
          renderSelectedList();
        }
      }
      
      function handleDragEnd() {
        this.classList.remove('dragging');
        document.querySelectorAll('.item').forEach(item => {
          item.classList.remove('over');
        });
      }
      
      // Column management
      function addColumn(column) {
        // Find the column in the original available columns array to ensure we have complete data
        const fullColumn = originalAvailableColumns.find(col => col.key === column.key);
        if (fullColumn) {
          // Add a deep copy to prevent reference issues
          selectedColumns.push(JSON.parse(JSON.stringify(fullColumn)));
          renderAvailableList(searchBox.value);
          renderSelectedList();
          showStatus('success', 'Column added: ' + fullColumn.name);
        }
      }
      
      function removeColumn(column) {
        selectedColumns = selectedColumns.filter(col => col.key !== column.key);
        renderAvailableList(searchBox.value);
        renderSelectedList();
        showStatus('success', 'Column removed: ' + column.name);
      }
      
      function renameColumn(columnIndex) {
        const column = selectedColumns[columnIndex];
        const currentName = column.customName || column.name;
        const newName = prompt("Enter custom header name for '" + column.name + "'", currentName);
        
        if (newName !== null) {
          // Update the column with the custom name
          selectedColumns[columnIndex].customName = newName;
          renderSelectedList();
          showStatus('success', 'Column renamed to: ' + newName);
        }
      }
      
      function showStatus(type, message) {
        const indicator = document.getElementById('statusIndicator');
        indicator.className = 'indicator ' + type;
        indicator.textContent = message;
        indicator.style.display = 'block';
        
        // Auto-hide after a delay
        setTimeout(function() {
          indicator.style.display = 'none';
        }, 2000);
      }
      
      // Event listeners
      document.getElementById('saveBtn').onclick = () => {
        if (selectedColumns.length === 0) {
          showStatus('error', 'Please select at least one column');
          return;
        }
        
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'flex';
        document.getElementById('saveBtn').disabled = true;
        document.getElementById('cancelBtn').disabled = true;
        
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('saveLoading').style.display = 'none';
            showStatus('success', 'Column preferences saved successfully!');
            
            // Close after a short delay to show the success message
            setTimeout(() => {
              google.script.host.close();
            }, 1500);
          })
          .withFailureHandler((error) => {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            document.getElementById('cancelBtn').disabled = false;
            showStatus('error', 'Error: ' + error.message);
          })
          .saveColumnPreferences(selectedColumns, entityType, sheetName);
      };
      
      document.getElementById('cancelBtn').onclick = () => {
        google.script.host.close();
      };
      
      document.getElementById('helpBtn').onclick = () => {
        const helpContent = 'Tips for selecting columns:\\n' +
                           '\\n• Search for specific columns using the search box' +
                           '\\n• Main fields are top-level Pipedrive fields' +
                           '\\n• Nested fields provide more detailed information' +
                           '\\n• Drag and drop to reorder selected columns' +
                           '\\n• The column order here determines the order in your sheet';
                           
        alert(helpContent);
      };
      
      searchBox.oninput = () => {
        renderAvailableList(searchBox.value);
      };
      
      // Initial render
      renderAvailableList();
      renderSelectedList();
    </script>
  `;
};

/**
 * Formats an entity type name for display
 * @param {string} entityType - The entity type to format
 * @return {string} The formatted entity type name
 */
ColumnSelectorUI.formatEntityTypeName = function(entityType) {
  if (!entityType) return '';
  
  // Remove any prefix/suffix and convert to title case
  const name = entityType.replace(/^ENTITY_TYPES\./, '').toLowerCase();
  return name.charAt(0).toUpperCase() + name.slice(1);
};

/**
 * Saves column preferences for a specific entity type and sheet
 * @param {string} entityType - Entity type (persons, deals, etc.)
 * @param {string} sheetName - Name of the sheet
 * @param {Array} columns - Array of column objects to save
 * @return {boolean} True if successful
 */
function saveColumnPreferences(entityType, sheetName, columns) {
  try {
    Logger.log(`Saving column preferences for ${entityType} in sheet "${sheetName}"`);
    Logger.log(`Received ${columns.length} columns to save`);
    
    // Log the first 10 columns to help debug
    if (columns.length > 0) {
      Logger.log('First 10 columns being saved:');
      columns.slice(0, 10).forEach((col, index) => {
        Logger.log(`Column ${index + 1}: key=${col.key}, name=${col.name || 'n/a'}, customName=${col.customName || 'n/a'}`);
      });
      
      // Specifically log any email and phone columns
      const emailColumns = columns.filter(col => col.key && (col.key === 'email' || col.key.startsWith('email.')));
      if (emailColumns.length > 0) {
        Logger.log(`Email columns (${emailColumns.length}): ${emailColumns.map(c => c.key).join(', ')}`);
      }
      
      const phoneColumns = columns.filter(col => col.key && (col.key === 'phone' || col.key.startsWith('phone.')));
      if (phoneColumns.length > 0) {
        Logger.log(`Phone columns (${phoneColumns.length}): ${phoneColumns.map(c => c.key).join(', ')}`);
      }
    }
    
    // Get Script Properties
    const properties = PropertiesService.getScriptProperties();
    
    // Make sure columns includes all necessary data
    const columnsToSave = columns.map(col => {
      return {
        key: col.key,
        name: col.name || formatColumnName(col.key),
        customName: col.customName || '',
        isNested: col.isNested,
        parentKey: col.parentKey
      };
    });
    
    // Generate key for saving the preferences - personalized for the current user
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log(`Saving preferences for user: ${userEmail}`);
    
    // Store full column objects
    const columnsKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
    const columnsJson = JSON.stringify(columnsToSave);
    
    Logger.log(`Saving with key: ${columnsKey}`);
    Logger.log(`JSON preview (first 100 chars): ${columnsJson.substring(0, 100)}...`);
    
    // Save to Script Properties
    properties.setProperty(columnsKey, columnsJson);
    
    // Verify the save
    try {
      const savedJson = properties.getProperty(columnsKey);
      const savedColumns = JSON.parse(savedJson);
      Logger.log(`Save verified: Retrieved ${savedColumns.length} columns`);
    } catch (verifyError) {
      Logger.log(`Error verifying saved data: ${verifyError.message}`);
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error saving column preferences: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    return false;
  }
}

/**
 * Gets column preferences for a specific entity type and sheet
 * @param {string} entityType - Entity type (persons, deals, etc.)
 * @param {string} sheetName - Name of the sheet
 * @return {Array} Array of column objects
 */
function getColumnPreferences(entityType, sheetName) {
  try {
    Logger.log(`Getting column preferences for ${entityType} in sheet "${sheetName}"`);
    
    // Get Script Properties
    const properties = PropertiesService.getScriptProperties();
    
    // Generate key for retrieving the preferences - personalized for the current user
    const userEmail = Session.getEffectiveUser().getEmail();
    const columnsKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
    
    Logger.log(`Looking for preferences with key: ${columnsKey}`);
    
    // Get from Script Properties
    const columnsJson = properties.getProperty(columnsKey);
    
    if (!columnsJson) {
      Logger.log(`No saved preferences found for key: ${columnsKey}`);
      return [];
    }
    
    try {
      const savedColumns = JSON.parse(columnsJson);
      Logger.log(`Found ${savedColumns.length} saved columns`);
      
      // Log first few columns to help debug
      if (savedColumns.length > 0) {
        Logger.log('First 5 saved columns:');
        savedColumns.slice(0, 5).forEach((col, index) => {
          Logger.log(`Column ${index + 1}: key=${col.key}, name=${col.name}`);
        });
      }
      
      return savedColumns;
    } catch (parseError) {
      Logger.log(`Error parsing saved columns: ${parseError.message}`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error getting column preferences: ${error.message}`);
    return [];
  }
}

/**
 * Handles saving column selections from the UI
 * @param {Array} selectedColumns - Array of selected column objects
 * @param {string} entityType - Entity type (persons, deals, etc.)
 * @param {string} sheetName - Name of the sheet
 * @return {Object} Response object
 */
function handleSaveColumns(selectedColumns, entityType, sheetName) {
  try {
    Logger.log(`Handling save columns request for ${entityType} in "${sheetName}"`);
    Logger.log(`Received ${selectedColumns.length} columns to save`);
    
    // Log a sample of columns being saved
    if (selectedColumns.length > 0) {
      Logger.log(`Sample columns being saved:`);
      selectedColumns.slice(0, 3).forEach((col, idx) => {
        Logger.log(`Column ${idx+1}: key=${col.key}, name=${col.name || 'unnamed'}`);
      });
    }
    
    // Fix the parameter order here to match the expected order in saveColumnPreferences
    const saveResult = saveColumnPreferences(entityType, sheetName, selectedColumns);
    
    if (saveResult) {
      Logger.log(`Successfully saved column preferences`);
      
      // If we're viewing this entity type in this sheet, refresh the data
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      
      if (activeSheet && activeSheet.getName() === sheetName) {
        Logger.log(`Active sheet matches saved sheet, refreshing data`);
        try {
          // Refresh data if the sheet is currently displaying this entity type
          const sheetEntityType = SyncService.getSheetEntityType(activeSheet);
          
          if (sheetEntityType === entityType) {
            Logger.log(`Sheet entity type matches, running sync`);
            SyncService.syncFromPipedrive();
          } else {
            Logger.log(`Sheet entity type (${sheetEntityType}) doesn't match saved entity type (${entityType})`);
          }
        } catch (refreshError) {
          Logger.log(`Error refreshing data: ${refreshError.message}`);
        }
      } else {
        Logger.log(`Active sheet does not match saved sheet, skipping refresh`);
      }
      
      return {
        success: true,
        message: `Saved ${selectedColumns.length} columns for ${entityType}`,
        columns: selectedColumns.length
      };
    } else {
      Logger.log(`Error saving column preferences`);
      return {
        success: false,
        message: "Error saving column preferences",
        error: "Save operation failed"
      };
    }
  } catch (error) {
    Logger.log(`Error in handleSaveColumns: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    
    return {
      success: false,
      message: "Error saving column preferences",
      error: error.message
    };
  }
}

// Export functions to be globally accessible
if (typeof ColumnSelectorUI === 'undefined') {
  // Create the ColumnSelectorUI namespace if it doesn't exist
  var ColumnSelectorUI = {};
}

// Export functions to the global namespace
ColumnSelectorUI.handleSaveColumns = handleSaveColumns;
ColumnSelectorUI.saveColumnPreferences = saveColumnPreferences;
ColumnSelectorUI.getColumnPreferences = getColumnPreferences;
ColumnSelectorUI.formatEntityTypeName = formatEntityTypeName;

/**
 * This global function is called from the menu
 */
function showColumnSelector() {
  try {
    // Call showColumnSelectorUI directly
    showColumnSelectorUI();
  } catch (e) {
    Logger.log(`Error in showColumnSelector: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Renders the HTML for the column selector
 * @param {Array} availableColumns - Available columns to select from
 * @param {Array} selectedColumns - Currently selected columns
 * @param {string} entityType - Entity type being edited
 * @param {string} sheetName - Sheet name
 * @return {string} HTML content
 */
function renderColumnSelectorHtml(availableColumns, selectedColumns, entityType, sheetName) {
  // Format entity type without causing circular reference
  function formatEntityTypeNameLocal(entityType) {
    if (!entityType) return '';
    
    // Remove any prefix/suffix and convert to title case
    const name = entityType.replace(/^ENTITY_TYPES\./, '').toLowerCase();
    return name.charAt(0).toUpperCase() + name.slice(1);
  }

  const formattedEntityType = formatEntityTypeNameLocal(entityType);
  
  // Create the data script that will be injected into the HTML template
  const dataScript = `<script>
    // Pass all available columns and selected columns to the front-end
    const availableColumns = ${JSON.stringify(availableColumns || [])};
    const selectedColumns = ${JSON.stringify(selectedColumns || [])};
    const entityType = "${entityType}";
    const sheetName = "${sheetName}";
    const entityTypeName = "${formattedEntityType}";
  </script>`;
  
  // Create the HTML template
  const template = HtmlService.createTemplateFromFile('ColumnSelector');
  
  // Pass data to template
  template.dataScript = dataScript;
  template.entityTypeName = formattedEntityType;
  template.sheetName = sheetName;
  
  // Evaluate template to HTML
  return template.evaluate().getContent();
}

/**
 * Global formatEntityTypeName function that delegates to the ColumnSelectorUI version
 * This provides compatibility with direct calls to formatEntityTypeName
 */
function formatEntityTypeName(entityType) {
  return ColumnSelectorUI.formatEntityTypeName(entityType);
} 