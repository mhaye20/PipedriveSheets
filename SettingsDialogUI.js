/**
 * Settings Dialog UI Helper
 * 
 * This module handles the settings dialog UI:
 * - Displaying the settings dialog
 * - Managing Pipedrive settings
 * - Handling settings persistence
 */

var SettingsDialogUI = SettingsDialogUI || {};

/**
 * Shows settings dialog where users can configure filter ID and entity type
 * @param {string} [initialTab='settings'] - Which tab to show initially ('settings' or 'columns')
 */
SettingsDialogUI.showSettings = function(initialTab = 'settings') {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get current settings from properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    // Only show auth dialog if we have no access token at all
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
    template.initialTab = initialTab; // Pass which tab should be initially active
    
    // Initialize column data variables for the column tab
    let availableColumnsData = [];
    let selectedColumnsData = [];
    const formattedEntityType = formatEntityTypeName(savedEntityType);
    
    // If we're opening the columns tab or need to preload column data, fetch it
    if (initialTab === 'columns') {
      try {
        Logger.log(`Attempting to load column data for ${savedEntityType}`);
        // Get the saved columns
        selectedColumnsData = getColumnPreferences(savedEntityType, activeSheetName) || [];
        
        // If in columns tab, try to get available columns (but don't fail if it doesn't work)
        // We'll fetch it in the UI if needed
        try {
          // This call is just to prime the cache, may fail but that's ok
          // The column selector tab has its own loading mechanism
          let sampleData = [];
          switch (savedEntityType) {
            case ENTITY_TYPES.DEALS:
              sampleData = PipedriveAPI.getDealsWithFilter(savedFilterId, 1);
              break;
            case ENTITY_TYPES.PERSONS:
              sampleData = PipedriveAPI.getPersonsWithFilter(savedFilterId, 1);
              break;
            case ENTITY_TYPES.ORGANIZATIONS:
              sampleData = PipedriveAPI.getOrganizationsWithFilter(savedFilterId, 1);
              break;
            case ENTITY_TYPES.ACTIVITIES:
              sampleData = PipedriveAPI.getActivitiesWithFilter(savedFilterId, 1);
              break;
            case ENTITY_TYPES.LEADS:
              sampleData = PipedriveAPI.getLeadsWithFilter(savedFilterId, 1);
              break;
            case ENTITY_TYPES.PRODUCTS:
              sampleData = PipedriveAPI.getProductsWithFilter(savedFilterId, 1);
              break;
          }
          
          if (sampleData && sampleData.length > 0) {
            Logger.log(`Got sample data for ${savedEntityType}, populating column data`);
            // We got sample data but won't process it here - the UI will handle that
            // This is just to make sure we have something for the template
            availableColumnsData = [
              { key: 'id', name: 'ID', isNested: false },
              { key: 'name', name: 'Name', isNested: false }
            ];
          }
        } catch (columnError) {
          Logger.log(`Non-critical error loading column data: ${columnError.message}`);
          // Don't fail the whole dialog for column data issues
        }
      } catch (dataError) {
        Logger.log(`Error preparing column data: ${dataError.message}`);
      }
    }
    
    // Create the data script with either real data or empty defaults
    template.dataScript = `<script>
      // Pass all available columns and selected columns to the front-end
      window.availableColumns = ${JSON.stringify(availableColumnsData || [])};
      window.selectedColumns = ${JSON.stringify(selectedColumnsData || [])};
      window.entityType = "${savedEntityType}";
      window.sheetName = "${activeSheetName}";
      window.entityTypeName = "${formattedEntityType}";
    </script>`;
    
    template.entityTypeName = formattedEntityType;
    template.sheetName = activeSheetName;
    
    // Create the HTML output from the template
    const htmlOutput = template.evaluate()
      .setWidth(600)
      .setHeight(600)
      .setTitle(`Pipedrive Settings for "${activeSheetName}"`)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    // Show a toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast('Opening settings dialog...', 'Pipedrive Sheets', 3);
    
    // Show modal dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
  } catch (error) {
    Logger.log('Error showing settings: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error showing settings: ' + error.message, 'Error', 5);
  }
};

/**
 * Shows the help and about dialog
 */
SettingsDialogUI.showHelp = function() {
  try {
    // Get current version from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const currentVersion = scriptProperties.getProperty('APP_VERSION') || '1.0.0';

    // Get version history if available
    const versionHistoryJson = scriptProperties.getProperty('VERSION_HISTORY');
    let versionHistory = [];
    
    if (versionHistoryJson) {
      try {
        versionHistory = JSON.parse(versionHistoryJson);
      } catch (e) {
        Logger.log('Error parsing version history: ' + e.message);
      }
    }

    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('Help');
    
    // Pass data to template
    template.currentVersion = currentVersion;
    template.versionHistory = versionHistory;
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(600)
      .setHeight(500)
      .setTitle('Help & About')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
    // Show a toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast('Opening help dialog...', 'Pipedrive Sheets', 2);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Help & About');
  } catch (error) {
    Logger.log('Error in showHelp: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error showing help: ' + error.message, 'Error', 5);
  }
};

// Export functions to be globally accessible
this.showSettings = function() { SettingsDialogUI.showSettings('settings'); };
this.showHelp = SettingsDialogUI.showHelp; 

// Add a function for the column selector that redirects to the combined dialog with columns tab
this.showColumnSelector = function() { SettingsDialogUI.showSettings('columns'); };




/**
 * Column Selector UI Helper
 * 
 * This module provides helper functions for the column selector UI:
 * - Styles for the column selector dialog
 * - Scripts for the column selector dialog
 */

var ColumnSelectorUI = SettingsDialogUI || {};

// Global variables
  let availableColumns = [];
  let selectedColumns = [];
  let originalAvailableColumns = [];

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
            
            // Enhanced display name handling for custom fields with hash IDs
            let displayName;
            
            // First try to use the field map which should contain proper names from the API
            if (fieldMap[key]) {
              displayName = fieldMap[key];
              Logger.log(`Using field map name for custom field ${key}: ${displayName}`);
            } 
            // If not in field map but looks like a hash ID, use a generic name + label
            else if (/^[a-f0-9]{20,}$/i.test(key)) {
              // Extract field type or label if possible from the value
              let fieldType = "Custom Field";
              if (typeof customFieldValue === 'object' && customFieldValue !== null) {
                if (customFieldValue.value !== undefined && customFieldValue.currency !== undefined) {
                  fieldType = "Currency Field";
                } else if (customFieldValue.value !== undefined && customFieldValue.formatted_address !== undefined) {
                  fieldType = "Address Field";
                } else if (customFieldValue.value !== undefined && customFieldValue.until !== undefined) {
                  fieldType = "Date Range Field";
                }
              }
              displayName = fieldType;
              Logger.log(`Generated generic name for hash ID custom field ${key}: ${displayName}`);
            }
            // Fall back to formatting the key
            else {
              displayName = formatColumnName(key);
              Logger.log(`Formatting custom field key ${key} to ${displayName}`);
            }
            
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
                
                // Add individual address components for more detailed control
                const addressComponents = [
                  'street_number', 'route', 'subpremise', 'sublocality', 
                  'locality', 'admin_area_level_1', 'admin_area_level_2',
                  'country', 'postal_code'
                ];
                
                addressComponents.forEach(component => {
                  if (customFieldValue[component] !== undefined) {
                    // Create a user-friendly display name
                    let componentName = component;
                    if (component === 'subpremise') componentName = 'Apartment/Suite';
                    if (component === 'street_number') componentName = 'Street Number';
                    if (component === 'route') componentName = 'Street';
                    if (component === 'sublocality') componentName = 'District/Borough';
                    if (component === 'locality') componentName = 'City';
                    if (component === 'admin_area_level_1') componentName = 'State/Province';
                    if (component === 'admin_area_level_2') componentName = 'County';
                    if (component === 'postal_code') componentName = 'ZIP/Postal Code';
                    
                    // Use the same user-friendly name in display names that will show in the UI
                    availableColumns.push({
                      key: `${currentPath}.${component}`,
                      name: `${displayName} (${componentName})`,
                      isNested: true,
                      parentKey: currentPath,
                    });
                  }
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
        
        // Skip arrays for IM fields - they're not commonly used and create clutter
        if (parentPath === 'im') {
          Logger.log(`Skipping IM array to reduce clutter`);
          return;
        }
        
        // For arrays of structured objects, like emails or phones
        if (obj.length > 0 && typeof obj[0] === 'object' && obj[0] !== null) {
          // Special handling for email/phone arrays
          if (obj[0].hasOwnProperty('value') && obj[0].hasOwnProperty('primary')) {
            
            if (obj[0].hasOwnProperty('label')) {
              // For contact fields like email/phone
              if (parentPath === 'email' || parentPath === 'phone') {
                // Don't create confusing array index entries
                // Instead, add useful named fields
                
                // Add primary field
                availableColumns.push({
                  key: `${parentPath}`,
                  name: `Primary ${formatColumnName(parentPath)}`,
                  isNested: true,
                  parentKey: parentPath
                });
                
                // Add common label types
                const commonLabels = ['work', 'home', 'mobile', 'other'];
                commonLabels.forEach(label => {
                  // Only add if this label type exists in the data
                  if (obj.some(item => item.label && item.label.toLowerCase() === label)) {
                    const columnName = `${formatColumnName(parentPath)} ${formatColumnName(label)}`;
                    
                    availableColumns.push({
                      key: `${parentPath}.${label}`,
                      name: columnName,
                      isNested: true,
                      parentKey: parentPath
                    });
                  }
                });
                
                return;
              }
            }
          } else if (parentPath !== 'im') {
            // For other object arrays (not email/phone/im), don't create confusing entries
            // Just skip them or handle specially
            Logger.log(`Skipping generic array processing for ${parentPath} to avoid confusion`);
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
        
        // Skip API-specific metadata fields that aren't useful
        if (['first_char', 'label', 'labels', 'visible_to', 'visible_from', 'in_visible_list'].includes(key)) {
          continue;
        }
        
        // Skip redundant owner fields (use owner_id instead)
        if (key === 'owner' && obj.hasOwnProperty('owner_id')) {
          continue;
        }
        
        // Skip duplicate name fields for owner/fields that will be handled by custom display names
        if (key === 'name' && parentPath && 
           (parentPath.endsWith('_id') || parentPath.includes('owner'))) {
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
    availableColumns = []; // Reset the global variable instead of redeclaring
    
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
      const addedKeys = new Set();
      
      // Keep track of important keys that we've already processed
      const processedTopLevelKeys = new Set();
      
      // First, build top-level column data
      Logger.log(`Building top-level columns first`);
      
      // Start with essential fields for every entity type
      const essentialFields = ['id', 'name', 'owner_id'];
      
      // Add entity-specific essential fields
      if (entityType === ENTITY_TYPES.DEALS) {
        essentialFields.push('title', 'value', 'currency', 'status', 'pipeline_id', 'stage_id', 'expected_close_date');
      } else if (entityType === ENTITY_TYPES.PERSONS) {
        essentialFields.push('email', 'phone', 'org_id');
      } else if (entityType === ENTITY_TYPES.ORGANIZATIONS) {
        essentialFields.push('address', 'owner_id', 'web');
      } else if (entityType === ENTITY_TYPES.ACTIVITIES) {
        essentialFields.push('type', 'due_date', 'due_time', 'note', 'deal_id', 'person_id', 'org_id');
      } else if (entityType === ENTITY_TYPES.PRODUCTS) {
        essentialFields.push('code', 'description', 'unit', 'cost', 'prices');
      }
      
      // Add common date fields for all entities
      essentialFields.push('add_time', 'update_time', 'created_at', 'updated_at');
      
      // Add the essential fields first
      for (const key of essentialFields) {
        if (sampleItem.hasOwnProperty(key) && !processedTopLevelKeys.has(key)) {
          processedTopLevelKeys.add(key);
          addedKeys.add(key);
          
          const displayName = fieldMap[key] || formatColumnName(key);
          
          availableColumns.push({
            key: key,
            name: displayName,
            isNested: false
          });
        }
      }
      
      // Then process the rest of top-level fields
      for (const key in sampleItem) {
        // Skip already processed keys, internal properties or functions
        if (processedTopLevelKeys.has(key) || key.startsWith('_') || typeof sampleItem[key] === 'function') {
          continue;
        }
        
        // Skip problematic fields like "im", "lm", "first_char", etc.
        if (['im', 'lm', 'first_char', 'label', 'labels', 'visible_to', 'visible_from'].includes(key)) {
          continue;
        }
        
        // Skip owner object if we have owner_id
        if (key === 'owner' && sampleItem.hasOwnProperty('owner_id')) {
          continue;
        }
        
        // Skip email and phone at top level if they're objects - handled specially
        if ((key === 'email' || key === 'phone') && typeof sampleItem[key] === 'object') {
          continue;
        }
        
        processedTopLevelKeys.add(key);
        addedKeys.add(key);
        
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
        // Add main email field (only if not already added)
        if (!availableColumns.some(col => col.key === 'email')) {
          availableColumns.push({
            key: 'email',
            name: 'Primary Email',
            isNested: false
          });
        }
        
        // Only process email array if it's actually an array
        if (Array.isArray(sampleItem.email)) {
          extractFields(sampleItem.email, 'email', 'Email');
        }
      }
      
      if (sampleItem.phone) {
        // Add main phone field (only if not already added)
        if (!availableColumns.some(col => col.key === 'phone')) {
          availableColumns.push({
            key: 'phone',
            name: 'Primary Phone', 
            isNested: false
          });
        }
        
        // Only process phone array if it's actually an array
        if (Array.isArray(sampleItem.phone)) {
          extractFields(sampleItem.phone, 'phone', 'Phone');
        }
      }
      
      // Skip IM field entirely to reduce clutter - don't even check for it
      
      // Then extract nested fields for all complex objects
      for (const key in sampleItem) {
        if (key.startsWith('_') || typeof sampleItem[key] === 'function') {
          continue;
        }
        
        // Skip problematic objects that cause clutter
        if (['im', 'lm', 'owner', 'first_char', 'label', 'labels'].includes(key)) {
          continue;
        }
        
        // Skip email/phone/im - already handled or intentionally skipped
        if (key === 'email' || key === 'phone' || key === 'im') {
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
    
    // Log all columns before filtering to debug
    Logger.log(`Total columns before filtering: ${availableColumns.length}`);
    const keylessColumns = availableColumns.filter(col => !col.key);
    Logger.log(`Keyless columns found: ${keylessColumns.length}`);
    keylessColumns.forEach(col => {
      Logger.log(`Keyless column: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
    });
    
    // Filter out problematic fields
    availableColumns = availableColumns.filter(col => {
      // CRITICAL: Filter out ALL entries that have no key
      // These problematic entries only have name and parentKey
      if (!col.key) {
        Logger.log(`Filtering out keyless field: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
        return false;
      }
      
      // Additional filters for fields that DO have keys
      // Filter out all IM-related fields
      if (col.parentKey && (col.parentKey === 'im' || col.parentKey.startsWith('im.'))) {
        Logger.log(`Filtering out IM field: ${col.name} (parent: ${col.parentKey})`);
        return false;
      }
      
      // Filter out fields with email.0, phone.0 parentKeys
      if (col.parentKey && col.parentKey.match(/^(email|phone)\.\d+$/)) {
        Logger.log(`Filtering out field with numeric parent: ${col.name} (parent: ${col.parentKey})`);
        return false;
      }
      
      // Filter out fields where parentKey is just 'email' or 'phone' but name contains " - 0"
      if (col.parentKey && (col.parentKey === 'email' || col.parentKey === 'phone') && 
          col.name && col.name.includes(' - 0')) {
        Logger.log(`Filtering out array field: ${col.name} (parent: ${col.parentKey})`);
        return false;
      }
      
      // Now handle fields that DO have keys
      if (col.key) {
        // Remove problematic fields like "im" and "lm"
        if (col.key.startsWith('im.') || 
            col.key.startsWith('lm.') ||
            col.key === 'im' || 
            col.key === 'lm' ||
            col.key === 'first_char' ||
            col.key === 'labels' ||
            col.key === 'label') {
          return false;
        }
        
        // Remove confusing array-indexed email/phone fields
        if (col.key.match(/^(email|phone|im)\.\d+/) || col.key.match(/\.\d+$/)) {
          Logger.log(`Filtering out confusing array field by key: ${col.key} (${col.name})`);
          return false;
        }
        
        // Remove nested array fields that have numeric indices
        if (col.key.includes('.') && col.key.match(/\.\d+\./)) {
          Logger.log(`Filtering out nested array field: ${col.key}`);
          return false;
        }
      }
      
      // Skip fields that typically contain no useful data
      if (['add_time', 'update_time'].includes(col.key) && 
          availableColumns.some(c => c.key === 'created_at' || c.key === 'updated_at')) {
        // Prefer created_at/updated_at over add_time/update_time
        return false;
      }
      
      // Handle duplicate organization fields - prefer org_id over org or organization
      if ((col.key === 'org' || col.key === 'organization') && 
          availableColumns.some(c => c.key === 'org_id')) {
        return false;
      }
      
      // Organization name duplicates
      if ((col.key === 'org.name' || col.key === 'organization.name') && 
          availableColumns.some(c => c.key === 'org_id.name')) {
        return false;
      }
      
      // Handle owner fields - only keep owner_id at top level and avoid duplicates
      if (col.key === 'owner' && availableColumns.some(c => c.key === 'owner_id')) {
        return false;
      }
      
      // Remove duplicates of owner_id.name vs owner.name
      if (col.key === 'owner.name' && availableColumns.some(c => c.key === 'owner_id.name')) {
        return false;
      }
      
      // Person duplicates - prefer person_id over person
      if (col.key === 'person' && availableColumns.some(c => c.key === 'person_id')) {
        return false;
      }
      
      // Person name duplicates
      if (col.key === 'person.name' && availableColumns.some(c => c.key === 'person_id.name')) {
        return false;
      }
      
      // Remove potentially confusing nested fields that duplicate info
      if (col.key.includes('.') && col.key.includes('id') && 
          col.key.endsWith('.id') && col.parentKey) {
        const parentField = col.parentKey;
        // If we have both parent.id and parent_id, keep parent_id
        if (availableColumns.some(c => c.key === parentField + '_id')) {
          return false;
        }
      }

      // Remove redundant address fields
      if (col.key === 'address_formatted_address' && availableColumns.some(c => c.key === 'address')) {
        return false;
      }
      
      // Also filter out any formatted_address fields that duplicate the main address
      if (col.key.includes('formatted_address') && availableColumns.some(c => c.key === 'address')) {
        return false;
      }

      // NEW: Filter out redundant read-only organization name fields
      if (col.key === 'org_name' && availableColumns.some(c => 
          (c.key === 'org_id' || c.key === 'organization_id' || c.key === 'organization'))) {
        return false;
      }
      
      // NEW: Filter out redundant read-only person name fields
      if (col.key === 'person_name' && availableColumns.some(c => 
          (c.key === 'person_id' || c.key === 'person'))) {
        return false;
      }
      
      // NEW: Filter out redundant owner name in favor of owner_id
      if (col.key === 'owner_name' && availableColumns.some(c => c.key === 'owner_id')) {
        return false;
      }
      
      // NEW: Filter out other common redundant name fields
      if ((col.key.endsWith('_name') || col.key.includes('_name_')) && 
          availableColumns.some(c => c.key === col.key.replace('_name', '_id'))) {
        return false;
      }
      
      // NEW: Filter out redundant date/time display variants
      if (col.key.startsWith('formatted_') && 
          availableColumns.some(c => c.key === col.key.replace('formatted_', ''))) {
        return false;
      }
      
      // NEW: Filter out redundant activity data
      if (col.key.includes('_activity_') && !col.key.includes('id') && 
          availableColumns.some(c => c.key.includes('_activity_id'))) {
        return false;
      }
      
      // NEW: Filter out redundant deal stage fields
      if (col.key === 'stage' && availableColumns.some(c => c.key === 'stage_id')) {
        return false;
      }
      
      // NEW: Filter out redundant pipeline fields
      if (col.key === 'pipeline' && availableColumns.some(c => c.key === 'pipeline_id')) {
        return false;
      }

      // NEW: Filter out redundant email/phone composite fields
      // Keep the specific type fields (like email.work, phone.mobile) instead of the generic composite
      if ((col.key === 'email' || col.key === 'phone') && 
          availableColumns.some(c => c.key.startsWith(col.key + '.'))) {
        return false;
      }
      
      // NEW: Filter out redundant label fields (keep label_ids for editing)
      if (col.key === 'label' && availableColumns.some(c => c.key === 'label_ids')) {
        return false;
      }
      
      // NEW: Filter out status fields that duplicate status_id
      if (col.key === 'status' && availableColumns.some(c => c.key === 'status_id')) {
        return false;
      }

      return true;
    });
    
    // Log results after filtering
    Logger.log(`Total columns after filtering: ${availableColumns.length}`);
    const remainingKeylessColumns = availableColumns.filter(col => !col.key);
    Logger.log(`Remaining keyless columns: ${remainingKeylessColumns.length}`);
    if (remainingKeylessColumns.length > 0) {
      Logger.log(`ERROR: Still have keyless columns after filtering!`);
      remainingKeylessColumns.forEach(col => {
        Logger.log(`Remaining keyless: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
      });
    }
    
    // Additional pass to remove redundant org/organization fields
    const orgKeys = availableColumns
      .filter(col => col.key === 'org_id' || col.key === 'org' || col.key === 'organization')
      .map(col => col.key);
    
    // If we have multiple org representations, keep only org_id
    if (orgKeys.length > 1 && orgKeys.includes('org_id')) {
      availableColumns = availableColumns.filter(col => 
        col.key !== 'org' && col.key !== 'organization' || col.key === 'org_id');
    }

    // Additional cleanup for organizations entity type
    if (entityType === ENTITY_TYPES.ORGANIZATIONS) {
      // When viewing organizations, remove redundant org/organization fields that reference self
      availableColumns = availableColumns.filter(col => 
        !(col.key.startsWith('org.') || col.key.startsWith('organization.') || 
          col.key === 'org' || col.key === 'organization'));
    }
    
    // Sort columns to prioritize important fields
    availableColumns.sort((a, b) => {
      // ID always comes first
      if (a.key === 'id') return -1;
      if (b.key === 'id') return 1;
      
      // Then name
      if (a.key === 'name') return -1;
      if (b.key === 'name') return 1;
      
      // Then owner
      if (a.key === 'owner_id') return -1;
      if (b.key === 'owner_id') return 1;
      
      // Main commonly used fields next
      const mainFields = ['title', 'status', 'value', 'currency', 'org_id', 'person_id', 'pipeline_id', 'stage_id'];
      const aIsMainField = mainFields.includes(a.key);
      const bIsMainField = mainFields.includes(b.key);
      if (aIsMainField && !bIsMainField) return -1;
      if (!aIsMainField && bIsMainField) return 1;
      if (aIsMainField && bIsMainField) {
        return mainFields.indexOf(a.key) - mainFields.indexOf(b.key);
      }
      
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
    
    // Post-process column names for better display
    availableColumns.forEach(col => {
      // Mark fields as read-only based on entity type context
      // Fields from related entities should be read-only
      if (col.key) {
        // Organization fields should be read-only when not in Organization view
        if ((col.key.startsWith('org.') || col.key.startsWith('org_') || 
             col.key.startsWith('organization.') || col.key.startsWith('organization_')) && 
            entityType !== ENTITY_TYPES.ORGANIZATIONS) {
          col.readOnly = true;
        }
        
        // Person fields should be read-only when not in Person view
        if ((col.key.startsWith('person.') || col.key.startsWith('person_')) && 
            entityType !== ENTITY_TYPES.PERSONS) {
          col.readOnly = true;
        }
        
        // Deal fields should be read-only when not in Deal view
        if ((col.key.startsWith('deal.') || col.key.startsWith('deal_')) && 
            entityType !== ENTITY_TYPES.DEALS) {
          col.readOnly = true;
        }
        
        // Activity fields should be read-only when not in Activity view
        if ((col.key.startsWith('activity.') || col.key.startsWith('activity_')) && 
            entityType !== ENTITY_TYPES.ACTIVITIES) {
          col.readOnly = true;
        }
        
        // Product fields should be read-only when not in Product view
        if ((col.key.startsWith('product.') || col.key.startsWith('product_')) && 
            entityType !== ENTITY_TYPES.PRODUCTS) {
          col.readOnly = true;
        }
        
        // Lead fields should be read-only when not in Lead view
        if ((col.key.startsWith('lead.') || col.key.startsWith('lead_')) && 
            entityType !== ENTITY_TYPES.LEADS) {
          col.readOnly = true;
        }
        
        // Address components should be read-only when not in Organization view
        if (col.key.includes('address.') && entityType !== ENTITY_TYPES.ORGANIZATIONS) {
          col.readOnly = true;
        }
      }

      // Mark read-only fields based on Pipedrive API documentation
      // These fields are not included in the update endpoints or are system-generated
      
      // Common read-only fields across all entity types
      if (col.key && (
          // System-generated fields
          col.key === 'creator_user_id' || 
          col.key === 'followers_count' ||
          col.key === 'participants_count' ||
          col.key === 'activities_count' ||
          col.key === 'done_activities_count' ||
          col.key === 'undone_activities_count' ||
          col.key === 'files_count' ||
          col.key === 'notes_count' ||
          col.key === 'email_messages_count' ||
          col.key === 'people_count' ||
          col.key === 'products_count' ||
          
          // Formatted or calculated fields
          col.key === 'formatted_value' || 
          col.key === 'weighted_value' || 
          col.key === 'formatted_weighted_value' ||
          col.key === 'weighted_value_currency' ||
          col.key.startsWith('formatted_') &&
          !col.key.includes('address'), // Don't mark address formatted fields as read-only
          
          // System fields
          col.key === 'first_char' ||
          col.key === 'active_flag' ||
          col.key === 'cc_email' ||
          col.key === 'next_activity_id' ||
          col.key === 'last_activity_id' ||
          col.key === 'last_incoming_mail_time' ||
          col.key === 'last_outgoing_mail_time' ||
          col.key === 'rotten_time' ||
          col.key === 'next_activity_date' ||
          col.key === 'next_activity_time' ||
          col.key === 'next_activity_type' ||
          col.key === 'next_activity_duration' ||
          col.key === 'next_activity_note' ||
          col.key === 'last_activity_date' ||
          col.key === 'archive_time' ||
          col.key === 'local_close_date' ||
          col.key === 'local_won_date' ||
          col.key === 'local_lost_date' ||
          
          // Deal-specific read-only fields
          col.key === 'stage_order_nr' ||
          col.key === 'person_name' ||
          col.key === 'org_name' ||
          col.key === 'origin' ||
          col.key === 'origin_id' ||
          
          // Person/Org specific
          col.key === 'has_pic' ||
          col.key === 'pic_hash' ||
          col.key === 'owner_name' ||
          col.key === 'org_hidden' ||
          col.key === 'person_hidden' ||
          
          // Additional count and statistical fields
          col.key === 'open_deals_count' ||
          col.key === 'closed_deals_count' ||
          col.key === 'related_open_deals_count' ||
          col.key === 'related_closed_deals_count' ||
          col.key === 'won_deals_count' ||
          col.key === 'lost_deals_count' ||
          col.key === 'related_won_deals_count' ||
          col.key === 'related_lost_deals_count' ||
          col.key === 'activities_count' ||
          col.key === 'done_activities_count' ||
          col.key === 'undone_activities_count' ||
          col.key === 'total_activities_count' ||
          col.key === 'company_id' ||
          col.key === 'last_activity' ||
          col.key === 'next_activity' ||
          col.key === 'picture' ||
          
          // Count fields in any format
          col.key.includes('_count') ||
          
          // Activity and notification fields
          col.key === 'last_notification_time' ||
          col.key === 'last_notification_user_id' ||
          col.key === 'notification_language_id' ||
          col.key === 'update_user_id' ||
          col.key === 'reference_type' ||
          col.key === 'reference_id' ||
          col.key === 'marked_as_done_time' ||
          col.key === 'series_id' ||
          col.key === 'conference_meeting_client' ||
          col.key === 'conference_meeting_url' ||
          col.key === 'rec_master_activity_id' ||
          col.key === 'original_start_time' ||
          
          // Deal specific fields
          col.key === 'first_won_time' ||
          col.key === 'last_stage_change_time' ||
          col.key === 'timezone' ||
          col.key.includes('timezone') ||
          
          // All auto-populated number fields for persons/organizations
          (col.key === 'people' || col.key === 'company') ||
          
          // Fields containing activity dates/activities which are calculated
          (col.key.includes('activity') && (col.key.includes('date') || col.key.includes('time') || col.key.includes('count')))
      )) {
        col.readOnly = true;
      }
      
      // Additional pass to mark more specific entity fields as read-only
      if (!col.readOnly) {
        // Organization fields
        if (entityType === ENTITY_TYPES.ORGANIZATIONS) {
          // People count and related fields are read-only for organizations
          if (col.key === 'people_count' || col.key === 'people') {
            col.readOnly = true;
          }
        }
        
        // Deal fields
        if (entityType === ENTITY_TYPES.DEALS) {
          // Status-related calculated fields
          if (col.key === 'status' && col.isNested) {
            col.readOnly = true;
          }
        }
        
        // Activity fields
        if (entityType === ENTITY_TYPES.ACTIVITIES) {
          // Reference and calendar fields
          if (col.key.includes('reference') || 
              col.key.includes('calendar') || 
              col.key.includes('series') ||
              col.key.includes('created_by') ||
              col.key.includes('conference')) {
            col.readOnly = true;
          }
        }
        
        // Person fields
        if (entityType === ENTITY_TYPES.PERSONS) {
          // Email and phone composite fields are read-only
          if ((col.key === 'email' || col.key === 'phone') && !col.key.includes('.')) {
            col.readOnly = true;
          }
        }
        
        // Check for participant fields
        if ((col.key.includes('participant') || col.key.includes('participants.')) && 
            (col.key.includes('person_id') || col.key.includes('primary_flag'))) {
          col.readOnly = true;
        }
        
        // Mark attendees fields as read-only
        if (col.key.includes('attendees.') || col.key === 'attendees') {
          col.readOnly = true;
        }
        
        // All participant interaction counts are read-only
        if (col.key.includes('deals') && col.key.includes('count')) {
          col.readOnly = true;
        }
        
        // All fields containing "lost", "won", "open", "closed" plus "deals" are likely read-only stats
        if ((col.key.includes('lost') || col.key.includes('won') || 
            col.key.includes('open') || col.key.includes('closed')) && 
            col.key.includes('deals')) {
          col.readOnly = true;
        }
        
        // Add owner-related fields that should be read-only
        if (col.key.startsWith('owner.') || 
            (col.key.startsWith('owner_') && col.key !== 'owner_id') ||
            col.key === 'owner_id.active_flag' ||
            col.key === 'owner_id.email' ||
            col.key === 'owner_id.has_pic' ||
            col.key === 'owner_id.id' ||
            col.key === 'owner_id.pic_hash' ||
            col.key === 'owner_id.value') {
          col.readOnly = true;
        }
        
        // Common read-only fields reported in UI
        const readOnlyFieldNames = [
          'active_flag',
          'activities_to_do',
          'cc_email',
          'done_activities',
          'email_messages_count',
          'files_count',
          'followers_count',
          'last_activity_date',
          'last_activity',
          'last_email_received',
          'last_email_sent',
          'next_activity_date',
          'next_activity',
          'next_activity_time',
          'notes_count',
          'org_name',
          'owner_name',
          'total_activities',
          'company',
          'people',
          'picture',
          'person_hidden',
          'org_hidden',
          'person_name',
          'creator_user',
          'source_channel',
          'channel',
          'source_origin',
          'origin',
          'stage_order_nr',
          'timezone',
          'weighted_value',
          'next_activity_duration',
          'next_activity_note',
          'next_activity_type',
          'participants_count',
          'products_count',
          'rotten_time',
          // Additional owner-related fields
          'owner_active_flag',
          'owner_email',
          'owner_has_pic',
          'owner_id.active_flag',
          'owner_id.email',
          'owner_id.has_pic',
          'owner_id.id',
          'owner_id.pic_hash',
          'owner_id.value',
          'owner_pic_hash',
          'owner_value'
        ];
        
        // Check if the field name exactly matches any of these fields
        if (readOnlyFieldNames.includes(col.key)) {
          col.readOnly = true;
        }
      }
      
      // Handle owner_id specially
      if (col.key === 'owner_id') {
        col.name = 'Owner';
      } else if (col.key === 'owner_id.name') {
        col.name = 'Owner Name';
      } else if (col.key.startsWith('owner_id.')) {
        const subfield = col.key.replace('owner_id.', '');
        col.name = 'Owner ' + formatColumnName(subfield);
      }
      
      // Handle organization fields
      if (col.key === 'org_id') {
        col.name = 'Organization';
      } else if (col.key === 'org_id.name') {
        col.name = 'Organization Name';
      } else if (col.key.startsWith('org_id.')) {
        const subfield = col.key.replace('org_id.', '');
        col.name = 'Organization ' + formatColumnName(subfield);
      }
      
      // Handle person fields
      if (col.key === 'person_id') {
        col.name = 'Person';
      } else if (col.key === 'person_id.name') {
        col.name = 'Person Name';
      } else if (col.key.startsWith('person_id.')) {
        const subfield = col.key.replace('person_id.', '');
        col.name = 'Person ' + formatColumnName(subfield);
      }
      
      // Handle deal fields
      if (col.key === 'deal_id') {
        col.name = 'Deal';
      } else if (col.key === 'deal_id.name' || col.key === 'deal_id.title') {
        col.name = 'Deal Title';
      } else if (col.key.startsWith('deal_id.')) {
        const subfield = col.key.replace('deal_id.', '');
        col.name = 'Deal ' + formatColumnName(subfield);
      }
      
      // Handle other IDs with better naming
      if (col.key.endsWith('_id') && !col.key.includes('.') &&
          !['owner_id', 'org_id', 'person_id', 'deal_id'].includes(col.key)) {
        const baseName = col.key.replace('_id', '');
        col.name = formatColumnName(baseName);
      }
      
      // For person.name, org.name, etc., make them clearer
      if (col.key.endsWith('.name')) {
        const parentName = col.key.replace('.name', '');
        if (['person', 'org', 'organization'].includes(parentName)) {
          col.name = formatColumnName(parentName) + ' Name';
        }
      }
    });
    
    // Additional pass to ensure primary entity name is clear
    const entityNameMap = {
      [ENTITY_TYPES.DEALS]: "Deal",
      [ENTITY_TYPES.PERSONS]: "Person",
      [ENTITY_TYPES.ORGANIZATIONS]: "Organization",
      [ENTITY_TYPES.ACTIVITIES]: "Activity",
      [ENTITY_TYPES.LEADS]: "Lead",
      [ENTITY_TYPES.PRODUCTS]: "Product"
    };

    // Make the entity name field clearer based on current entity type
    if (entityType && entityNameMap[entityType]) {
      availableColumns.forEach(col => {
        // Make "name" field explicit for the current entity type
        if (col.key === 'name' && !col.isNested) {
          col.name = `${entityNameMap[entityType]} Name`;
        }
        
        // Make "id" field explicit to distinguish it from any custom "ID" column
        if (col.key === 'id' && !col.isNested) {
          col.name = `Pipedrive ${entityNameMap[entityType]} ID`;
        }
        
        // Make "title" field explicit for deals
        if (entityType === ENTITY_TYPES.DEALS && col.key === 'title' && !col.isNested) {
          col.name = 'Deal Title';
        }
      });
    }
    
    // Log the structure of email and phone fields if they exist
    if (sampleItem.email) {
      Logger.log(`Email field structure: ${JSON.stringify(sampleItem.email)}`);
    }
    if (sampleItem.phone) {
      Logger.log(`Phone field structure: ${JSON.stringify(sampleItem.phone)}`);
    }
    
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
    // Implementation that directly calls the showSettings function with 'columns' tab active
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (!activeSheet) {
      throw new Error('No active sheet found. Please select a sheet first.');
    }
    
    SettingsDialogUI.showSettings('columns');
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
  return `<script>
    // Initialization handling
    document.addEventListener('DOMContentLoaded', function() {
      console.log('ColumnSelectorUI script loaded');
    });
    
    // Column selector specific script content
    function initColumnSelector() {
      // Initialize UI interactions
      console.log('Column selector initialized');
    }
  </script>`;
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
      // DEBUG: Log each column object being mapped
      // Logger.log(`SAVE_DEBUG: Mapping column: key=${col.key}, name=${col.name}, customName=${col.customName}`);
      return {
        key: col.key,
        name: col.name || formatColumnName(col.key),
        customName: col.customName || '',
        isNested: col.isNested,
        parentKey: col.parentKey
      };
    });

    // DEBUG: Log the final array being prepared for saving
    Logger.log(`SAVE_DEBUG: Final columnsToSave array prepared (first 5): ${JSON.stringify(columnsToSave.slice(0,5))}`);
    if (columnsToSave.length > 0) {
        Logger.log(`SAVE_DEBUG: First column in columnsToSave: key=${columnsToSave[0].key}, name=${columnsToSave[0].name}, customName=${columnsToSave[0].customName}`);
    }
    
    // Generate key for saving the preferences
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log(`Saving preferences for user: ${userEmail}`);
    
    // Instead of using user email, use team ID for shared preferences
    let keyIdentifier = userEmail;
    
    // Try to get user's team ID
    try {
      // Check if the getUserTeam function is available
      if (typeof getUserTeam === 'function') {
        // Get the user's team information
        const userTeam = getUserTeam(userEmail);
        if (userTeam && userTeam.teamId) {
          // Use team ID instead of email for shared preferences across team
          keyIdentifier = `TEAM_${userTeam.teamId}`;
          Logger.log(`User is part of team ${userTeam.name} (${userTeam.teamId}), using team-based storage`);
        } else {
          Logger.log(`User is not part of a team, using personal storage`);
        }
      } else {
        Logger.log('getUserTeam function not available, defaulting to personal storage');
      }
    } catch (teamError) {
      Logger.log(`Error getting team info: ${teamError.message}, defaulting to personal storage`);
    }
    
    // Store full column objects
    const columnsKey = `COLUMNS_${sheetName}_${entityType}_${keyIdentifier}`;
    const columnsJson = JSON.stringify(columnsToSave);
    
    Logger.log(`Saving with key: ${columnsKey}`);
    Logger.log(`JSON preview (first 100 chars): ${columnsJson.substring(0, 100)}...`);
    
    // Save to Script Properties
    properties.setProperty(columnsKey, columnsJson);
    
    // NEW: Create and store a header-to-field key mapping for pushChangesToPipedrive
    // This mapping will allow us to find the right Pipedrive field key regardless of header name changes
    const headerToFieldKeyMap = {};
    
    columnsToSave.forEach(col => {
      // If a column has a custom name, use that as the key in our mapping
      if (col.customName) {
        headerToFieldKeyMap[col.customName] = col.key;
      } else {
        // Otherwise use the standard formatted name
        headerToFieldKeyMap[col.name] = col.key;
      }
    });
    
    // Save the header-to-field mapping
    const mappingKey = `HEADER_TO_FIELD_MAP_${sheetName}_${entityType}`;
    const mappingJson = JSON.stringify(headerToFieldKeyMap);
    properties.setProperty(mappingKey, mappingJson);
    Logger.log(`Saved header-to-field mapping with key: ${mappingKey}`);
    
    // Verify the save
    try {
      const savedJson = properties.getProperty(columnsKey);
      const savedColumns = JSON.parse(savedJson);
      Logger.log(`Save verified: Retrieved ${savedColumns.length} columns`);
    } catch (verifyError) {
      Logger.log(`Error verifying saved data: ${verifyError.message}`);
    }
    
    // IMPORTANT: Notify TwoWaySyncSettingsUI about column changes
    // This ensures the Sync Status column gets properly repositioned on next sync
    if (typeof handleColumnPreferencesChange === 'function') {
      Logger.log(`Notifying TwoWaySyncSettingsUI about column changes for sheet: ${sheetName}`);
      handleColumnPreferencesChange(sheetName);
    } else {
      Logger.log(`handleColumnPreferencesChange function not found, skipping notification`);
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
    
    // Generate key for retrieving the preferences
    const userEmail = Session.getEffectiveUser().getEmail();
    
    // Try to get user's team ID first for team-shared preferences
    let keyIdentifier = userEmail;
    let columnsJson = null;
    
    try {
      // Check if the getUserTeam function is available
      if (typeof getUserTeam === 'function') {
        // Get the user's team information
        const userTeam = getUserTeam(userEmail);
        if (userTeam && userTeam.teamId) {
          // Check for team-based preferences first
          const teamKey = `COLUMNS_${sheetName}_${entityType}_TEAM_${userTeam.teamId}`;
          Logger.log(`Looking for team preferences with key: ${teamKey}`);
          
          columnsJson = properties.getProperty(teamKey);
          if (columnsJson) {
            Logger.log(`Found team-based preferences for ${sheetName}`);
            keyIdentifier = `TEAM_${userTeam.teamId}`;
          } else {
            // If no team preferences found, check for personal preferences 
            // (for backward compatibility with users who already set up preferences)
            const personalKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
            Logger.log(`No team preferences, checking personal preferences with key: ${personalKey}`);
            
            const personalJson = properties.getProperty(personalKey);
            if (personalJson) {
              Logger.log(`Found personal preferences, will migrate to team preferences`);
              columnsJson = personalJson;
              
              // Migrate personal preferences to team preferences
              properties.setProperty(teamKey, personalJson);
              Logger.log(`Migrated personal preferences to team preferences`);
              
              keyIdentifier = `TEAM_${userTeam.teamId}`;
            }
          }
        } else {
          Logger.log(`User is not part of a team, using personal storage`);
        }
      } else {
        Logger.log('getUserTeam function not available, defaulting to personal storage');
      }
    } catch (teamError) {
      Logger.log(`Error getting team info: ${teamError.message}, defaulting to personal storage`);
    }
    
    // If we couldn't find team preferences or personal preferences to migrate, 
    // fall back to personal preferences
    if (!columnsJson) {
      const personalKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
      Logger.log(`Looking for personal preferences with key: ${personalKey}`);
      
      columnsJson = properties.getProperty(personalKey);
    }
    
    if (!columnsJson) {
      Logger.log(`No saved preferences found for key: ${keyIdentifier}`);
      return [];
    }
    
    try {
      const savedColumns = JSON.parse(columnsJson);
      Logger.log(`Found ${savedColumns.length} saved columns using key identifier: ${keyIdentifier}`);
      
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
    // Use the SettingsDialogUI with columns tab
    SettingsDialogUI.showSettings('columns');
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
    window.availableColumns = ${JSON.stringify(availableColumns || [])};
    window.selectedColumns = ${JSON.stringify(selectedColumns || [])};
    window.entityType = "${entityType}";
    window.sheetName = "${sheetName}";
    window.entityTypeName = "${formattedEntityType}";
  </script>`;
  
  // Create the HTML template
  const template = HtmlService.createTemplateFromFile('SettingsDialog');
  
  // Pass data to template
  template.dataScript = dataScript;
  template.entityTypeName = formattedEntityType;
  template.sheetName = sheetName;
  // Set initialTab to 'columns' since this function is specifically for column selector
  template.initialTab = 'columns';
  
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

/**
 * Gets column data for a specific entity type - called from the UI
 * @param {string} entityType - The entity type to get columns for
 * @param {string} sheetName - The name of the sheet
 * @return {Object} Object with available and selected columns
 */
function getColumnsDataForEntity(entityType, sheetName) {
  Logger.log(`=== getColumnsDataForEntity CALLED with entityType: ${entityType}, sheetName: ${sheetName} ===`);
  
  // TEMPORARY TEST: Return empty data to see if function is called
  if (false) { // Set to true to test
    return {
      availableColumns: [],
      selectedColumns: []
    };
  }
  
  try {
    Logger.log(`Getting column data for entity type ${entityType} in sheet ${sheetName}`);
    
    let availableColumns = [];
    const selectedColumns = getColumnPreferences(entityType, sheetName) || [];
    
    // Get sample data to extract available columns
    try {
      // Initialize the custom fields cache to get proper field names
      const customFieldCache = initializeCustomFieldsCache(entityType);
      Logger.log(`Initialized custom field cache with ${Object.keys(customFieldCache || {}).length} fields`);
      
      // Get field definitions based on entity type to get proper field names
      let fieldDefinitions = [];
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          fieldDefinitions = getDealFields();
          break;
        case ENTITY_TYPES.PERSONS:
          fieldDefinitions = getPersonFields();
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          fieldDefinitions = getOrganizationFields();
          break;
        case ENTITY_TYPES.ACTIVITIES:
          fieldDefinitions = getActivityFields();
          break;
        case ENTITY_TYPES.LEADS:
          fieldDefinitions = getLeadFields();
          break;
        case ENTITY_TYPES.PRODUCTS:
          fieldDefinitions = getProductFields();
          break;
      }
      
      // Map field keys to their friendly names
      const fieldNameMap = {};
      fieldDefinitions.forEach(field => {
        if (field.key && field.name) {
          fieldNameMap[field.key] = field.name;
        }
      });
      
      // Get sample data based on filter
      const filterId = PropertiesService.getScriptProperties().getProperty(`FILTER_ID_${sheetName}`) || '';
      
      let sampleData = [];
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
      
      // Pre-process data to find date/time range custom fields
      const dateTimeRangeFields = new Map();
      if (sampleData && sampleData.length > 0) {
        const item = sampleData[0];
        // Look for fields with _until suffix which indicates a date/time range field
        for (const key in item) {
          if (key.match(/[a-f0-9]{20,}_until$/i)) {
            // This is an end date/time field
            const baseKey = key.replace(/_until$/, '');
            if (item[baseKey] !== undefined) {
              // We found both the start and end fields
              dateTimeRangeFields.set(baseKey, {
                startKey: baseKey,
                endKey: key
              });
            }
          }
        }
      }
      
      // Extract available columns from sample data
      if (sampleData && sampleData.length > 0) {
        availableColumns = extractColumnsFromData(sampleData[0], entityType);
        
        // Post-process columns for date/time range fields
        if (dateTimeRangeFields.size > 0) {
          // Create a map of available columns by key for easy lookup
          const columnsMap = new Map();
          availableColumns.forEach(col => {
            columnsMap.set(col.key, col);
          });
          
          // Add missing Start fields for date/time ranges
          dateTimeRangeFields.forEach((range, baseKey) => {
            // If we already have the base field but not explicitly as a Start field
            if (columnsMap.has(baseKey)) {
              // Create an explicit Start field
              const baseCol = columnsMap.get(baseKey);
              const startCol = {
                key: baseKey,
                name: `${baseCol.name} - Start`,
                isNested: false,
                parentKey: null,
                category: baseCol.category,
                readOnly: baseCol.readOnly
              };
              
              // Update the base field's name if it doesn't already indicate it's the Start field
              if (!baseCol.name.includes('Start')) {
                baseCol.name = startCol.name;
              }
            }
          });
        }
        
        // Log all columns before filtering to debug
        Logger.log(`[getColumnsDataForEntity] Total columns before filtering: ${availableColumns.length}`);
        availableColumns.forEach((col, idx) => {
          if (!col.key) {
            Logger.log(`[getColumnsDataForEntity] Column ${idx} has NO KEY: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
          }
        });
        
        // CRITICAL: Filter out ALL keyless entries BEFORE any other processing
        // These are the problematic "Email - 0", "Phone - 0", "Im - 0" entries
        const beforeFilterCount = availableColumns.length;
        availableColumns = availableColumns.filter(col => {
          if (!col.key) {
            Logger.log(`[getColumnsDataForEntity] REMOVING keyless field: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
            return false;
          }
          return true;
        });
        Logger.log(`[getColumnsDataForEntity] Filtered ${beforeFilterCount - availableColumns.length} keyless entries before improveColumnNamesForUI. Remaining: ${availableColumns.length}`);
        
        // Improve column names for the UI (only process entries that have keys)
        availableColumns = improveColumnNamesForUI(availableColumns, entityType);
      }
    } catch (error) {
      Logger.log(`Error getting available columns: ${error.message}`);
      Logger.log(`Stack: ${error.stack}`);
      throw error;
    }
    
    // Final check before returning
    const keylessCount = availableColumns.filter(col => !col.key).length;
    if (keylessCount > 0) {
      Logger.log(`[getColumnsDataForEntity] ERROR: Still have ${keylessCount} keyless columns before returning!`);
      availableColumns.forEach((col, idx) => {
        if (!col.key) {
          Logger.log(`[getColumnsDataForEntity] Keyless column ${idx}: name="${col.name}", parentKey="${col.parentKey || 'null'}"`);
        }
      });
      // Force filter them out one more time
      availableColumns = availableColumns.filter(col => col.key);
      Logger.log(`[getColumnsDataForEntity] Force filtered to ${availableColumns.length} columns`);
    }
    
    return {
      availableColumns,
      selectedColumns
    };
  } catch (error) {
    Logger.log(`Error in getColumnsDataForEntity: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    
    // Return a minimal set of columns as fallback
    return {
      availableColumns: getFallbackColumns(entityType),
      selectedColumns: []
    };
  }
}

/**
 * Extracts column information from a data object
 * @param {Object} data - The data object to extract columns from
 * @param {string} entityType - The entity type
 * @return {Array} Array of column objects
 */
function extractColumnsFromData(data, entityType) {
  try {
    const columns = [];
    const processedKeys = new Set();
    const mainCategoryName = getEntityCategoryName(entityType);
    const customFieldCategory = "Custom Fields";
    const metadataCategory = "System Fields";
    
    // Helper function to add a column with category
    function addColumn(key, name, isNested = false, parentKey = null, category = mainCategoryName, readOnly = false, required = false) {
      // Skip specific fields that shouldn't be shown
      if (shouldSkipField(key, processedKeys)) {
        return;
      }
      
      if (!processedKeys.has(key)) {
        columns.push({
          key, 
          name, 
          isNested, 
          parentKey, 
          category, // Add category
          readOnly: readOnly || isReadOnlyField(key, entityType), // Use central read-only check
          required 
        });
        processedKeys.add(key);
      }
    }
    
    // Helper to determine if a field should be skipped
    function shouldSkipField(key, processedKeys) {
      if (processedKeys.has(key)) return true;
      
      // Skip timezone fields which are not useful on their own
      if (key.includes('timezone') || key.endsWith('.timezone') || 
          key.includes('timezone_id') || key.endsWith('_timezone_id')) {
        return true;
      }
      
      // Skip currency fields which are part of money fields
      if (key.endsWith('.currency') || key.endsWith('_currency')) {
        return true;
      }
      
      // Skip "complete_address" or "formatted_address" in address components
      if (key.includes('complete_address') || key.includes('formatted_address') || 
          key.includes('_formatted_address') || key.includes('_complete') || 
          key.includes('.complete') || key.includes('.formatted')) {
        return true;
      }
      
      // Skip duplicate admin_area_level fields - we'll add them with clear naming
      if ((key.includes('admin_area_level') || key.includes('_admin_area_level')) && 
          !key.includes('admin_area_level_1') && !key.includes('admin_area_level_2')) {
        return true;
      }
      
      // Skip "End Time/Date" fields since we want to show both start and end together
      // But only skip them if they're part of a date/time range that we're handling specially
      if ((key.includes('_range') || key.includes('date_range') || key.includes('time_range')) &&
          (key.endsWith('.end') || key.endsWith('_end') || key.endsWith('.until') || key.endsWith('_until'))) {
        // Instead, ensure we have the parent field which contains both start and end
        const parentField = key.split(/[._](?:end|until)/)[0];
        if (!processedKeys.has(parentField)) {
          return false; // Let it through if we need to process the parent
        } else {
          return true; // Skip if parent is already processed
        }
      }
      
      return false;
    }

    // 1. Always include the ID column
    addColumn('id', 'Pipedrive ID', false, null, mainCategoryName, false, true);

    // 2. Process standard Pipedrive fields for this entity type
    const standardFieldKeys = getStandardFieldKeys(entityType);
    for (const key of standardFieldKeys) {
      if (data.hasOwnProperty(key) && key !== 'id') {
        const fieldName = formatColumnName(key); // Use the centralized formatter
        addColumn(key, fieldName, false, null, mainCategoryName);
      }
    }

    // 3. Process Custom Fields (Hash IDs)
    for (const key in data) {
      if (processedKeys.has(key)) continue;

      // Custom field (base hash or component)
      if (/[a-f0-9]{20,}/i.test(key)) {
        const fieldName = formatColumnName(key);
        addColumn(key, fieldName, false, null, customFieldCategory);
      } 
      // Skip objects for now, handle them in step 4
      else if (typeof data[key] === 'object' && data[key] !== null) {
        continue;
      } 
      // Other top-level fields not already processed (potential system/metadata fields)
      else {
        const fieldName = formatColumnName(key);
        
        // --- Refined Categorization Logic ---
        let category = metadataCategory; // Default to System/Metadata
        const lowerKey = key.toLowerCase();

        // Known System/Read-Only/Metadata fields
        const systemIndicators = ['_count', '_id', '_flag', '_char', 'cc_email', 'visible_', 'hidden', 'formatted_', 'weighted_', 'rotten_', 'last_', 'next_activity', 'first_won', 'stage_order', 'company_id'];
        const exactSystemKeys = ['active', 'deleted']; // Add more exact keys if needed

        if (exactSystemKeys.includes(lowerKey) || systemIndicators.some(ind => lowerKey.includes(ind))) {
            // Exceptions: Keep related IDs (_id) and dates (_time, _date) in their respective categories (handled later)
            if (!(lowerKey.endsWith('_id') && !lowerKey.startsWith('next_') && !lowerKey.startsWith('last_')) && 
                !lowerKey.includes('_time') && !lowerKey.includes('_date')) {
                category = metadataCategory;
            }
        }
        
        // Assign to Main Category if it's a standard field or looks like one (date/time/related id)
        if (standardFieldKeys.includes(key) || 
           (key.endsWith('_id') && !lowerKey.startsWith('next_') && !lowerKey.startsWith('last_')) || // Related IDs
           key.includes('_time') || key.includes('_date')) { // Dates/Times
           category = mainCategoryName; 
        } 
        // If not specifically categorized as Main or System, keep as Metadata for now
        // --- End Refined Categorization ---
        
        addColumn(key, fieldName, false, null, category);
      }
    }

    // 4. Process Nested/Related Entity Fields
    for (const key in data) {
      const value = data[key];
      if (processedKeys.has(key)) continue; // Already added as top-level (e.g., custom field hash)

      if (typeof value === 'object' && value !== null) {
        // Special handling for complex field types
        // For date/time range fields, ensure we add both start and end fields
        const isDateRange = key.includes('date_range') || key.includes('time_range');
        const isMoneyField = key.includes('money') && typeof value === 'object' && value.currency !== undefined;
        const isAddressField = key.includes('address') && typeof value === 'object';
        
        // Determine category based on the parent key
        let category = metadataCategory; // Default
        let relatedEntityType = null;

        if (key === 'person_id') { category = getEntityCategoryName(ENTITY_TYPES.PERSONS); relatedEntityType = ENTITY_TYPES.PERSONS; }
        else if (key === 'org_id') { category = getEntityCategoryName(ENTITY_TYPES.ORGANIZATIONS); relatedEntityType = ENTITY_TYPES.ORGANIZATIONS; }
        else if (key === 'owner_id') { category = "Owner Fields"; }
        else if (key === 'user_id') { category = "User Fields"; }
        else if (key === 'creator_user_id') { category = "Creator Fields"; }
        else if (key === 'deal_id') { category = getEntityCategoryName(ENTITY_TYPES.DEALS); relatedEntityType = ENTITY_TYPES.DEALS; } 
        else if (key === 'pipeline_id') { category = "Pipeline/Stage Fields"; }
        else if (key === 'stage_id') { category = "Pipeline/Stage Fields"; }
        // Add more related entities as needed

        // Add the parent object itself if it represents a link (like org_id)
        if (key.endsWith('_id') && !key.startsWith('next_') && !key.startsWith('last_')) {
            addColumn(key, formatColumnName(key), false, null, category); 
        }
        
        // For date/time range fields, ensure we add the parent field
        if (isDateRange) {
            // Add the parent field
            addColumn(key, formatColumnName(key), false, null, category);
            
            // Add start field if it exists
            if (value.start !== undefined) {
                const startPath = `${key}.start`;
                addColumn(startPath, `${formatColumnName(key)} - Start`, true, key, category);
            }
            
            // Add end field if it exists
            if (value.end !== undefined) {
                const endPath = `${key}.end`;
                addColumn(endPath, `${formatColumnName(key)} - End`, true, key, category);
            }
        }
        // For money fields, only add the amount value but not the currency
        else if (isMoneyField) {
            // Add the parent field for the money field
            addColumn(key, formatColumnName(key), false, null, category);
            
            // Only add the amount component
            if (value.amount !== undefined) {
                const amountPath = `${key}.amount`;
                addColumn(amountPath, `${formatColumnName(key)} - Amount`, true, key, category);
            }
        }
        // For address fields, properly handle all components
        else if (isAddressField) {
            // Add the parent address field
            addColumn(key, formatColumnName(key), false, null, category);
            
            // Add specific address components with clear names
            const addressComponents = [
                { key: 'street_number', name: 'Street Number' },
                { key: 'route', name: 'Street Name' },
                { key: 'sublocality', name: 'District' },
                { key: 'locality', name: 'City' },
                { key: 'admin_area_level_1', name: 'State/Province' },
                { key: 'admin_area_level_2', name: 'County' },
                { key: 'country', name: 'Country' },
                { key: 'postal_code', name: 'ZIP/Postal Code' },
                { key: 'subpremise', name: 'Suite/Apt' },
            ];
            
            // Add each address component with clear naming
            for (const component of addressComponents) {
                if (value[component.key] !== undefined || 
                    value['_' + component.key] !== undefined ||
                    value[component.key + '_'] !== undefined) {
                    const compPath = `${key}.${component.key}`;
                    addColumn(compPath, `${formatColumnName(key)} - ${component.name}`, true, key, category);
                }
            }
            
            // Special handling for custom field address components with key suffix naming
            if (key.match(/[a-f0-9]{20,}/i)) {
                // This is a custom field with hash ID format
                const hashId = key;
                
                // Look for components that follow the pattern hashId_component_name
                for (const component of addressComponents) {
                    const customCompKey = `${hashId}_${component.key}`;
                    if (data[customCompKey] !== undefined) {
                        addColumn(customCompKey, `${formatColumnName(key)} - ${component.name}`, true, key, category);
                    }
                }
            }
        }
        // For date/time range fields, ensure we add both start and end fields
        else if (isDateRange) {
            // Add the parent field
            addColumn(key, formatColumnName(key), false, null, category);
            
            // For date/time range fields, properly add start and end fields
            let hasStart = false;
            let hasEnd = false;
            
            // Check if this is a custom field with component fields in main data structure
            if (key.match(/[a-f0-9]{20,}/i)) {
                // Custom field - look for fields like hashId_until
                const hashId = key;
                
                // Look for _until field (used for end date/time)
                const untilKey = `${hashId}_until`;
                if (data[untilKey] !== undefined) {
                    hasEnd = true;
                    addColumn(untilKey, `${formatColumnName(key)} - End`, true, key, category);
                }
                
                // The main field itself is the start
                hasStart = true;
                addColumn(key, `${formatColumnName(key)} - Start`, false, null, category);
            }
            // For standard nested objects with start/end properties
            else {
                // Add start field if it exists
                if (value.start !== undefined) {
                    hasStart = true;
                    const startPath = `${key}.start`;
                    addColumn(startPath, `${formatColumnName(key)} - Start`, true, key, category);
                }
                
                // Add end field if it exists
                if (value.end !== undefined) {
                    hasEnd = true;
                    const endPath = `${key}.end`;
                    addColumn(endPath, `${formatColumnName(key)} - End`, true, key, category);
                }
                
                // Check for until field (alternate to end)
                if (value.until !== undefined) {
                    hasEnd = true;
                    const untilPath = `${key}.until`;
                    addColumn(untilPath, `${formatColumnName(key)} - End`, true, key, category);
                }
            }
            
            // If we only found end but not start, make sure to add start
            if (hasEnd && !hasStart) {
                // For date ranges, the main field itself is often the start date
                addColumn(key, `${formatColumnName(key)} - Start`, false, null, category);
            }
            
            // If we only found start but not end, add a placeholder for end
            if (hasStart && !hasEnd && key.match(/[a-f0-9]{20,}/i)) {
                // For custom fields, look for the _until variant
                const untilKey = `${key}_until`;
                if (data[untilKey] !== undefined) {
                    addColumn(untilKey, `${formatColumnName(key)} - End`, true, key, category);
                }
            }
        }
        // Standard nested object processing for other field types
        else {
            // Recursively process properties within the object
            processNestedObject(value, key, category, relatedEntityType, addColumn, 1);
        }
      }
    }
    
    // Sort columns: First by category order, then by name within category
    const categoryOrder = [
      mainCategoryName,
      "Owner Fields",
      getEntityCategoryName(ENTITY_TYPES.PERSONS),
      getEntityCategoryName(ENTITY_TYPES.ORGANIZATIONS),
      getEntityCategoryName(ENTITY_TYPES.DEALS),
      getEntityCategoryName(ENTITY_TYPES.ACTIVITIES),
      getEntityCategoryName(ENTITY_TYPES.LEADS),
      getEntityCategoryName(ENTITY_TYPES.PRODUCTS),
      "Pipeline/Stage Fields",
      "User Fields",
      "Creator Fields",
      customFieldCategory,
      metadataCategory
    ];

    columns.sort((a, b) => {
      const categoryIndexA = categoryOrder.indexOf(a.category);
      const categoryIndexB = categoryOrder.indexOf(b.category);

      if (categoryIndexA !== categoryIndexB) {
        // Handle categories not explicitly listed (put them last)
        if (categoryIndexA === -1) return 1;
        if (categoryIndexB === -1) return -1;
        return categoryIndexA - categoryIndexB;
      }
      
      // If same category, sort by nested status (non-nested first), then name
      if (a.isNested !== b.isNested) {
        return a.isNested ? 1 : -1;
      }
      return a.name.localeCompare(b.name);
    });

    Logger.log(`Extracted and categorized ${columns.length} columns for entity type ${entityType}`);
    return columns;
  } catch (e) {
    Logger.log(`Error extracting columns: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return getFallbackColumns(entityType); // Ensure fallback includes category
  }
}

/**
 * Helper to recursively process nested objects for column extraction
 */
function processNestedObject(obj, parentKey, parentCategory, relatedEntityType, addColumnFn, depth) {
   if (depth > 3) return; // Limit recursion depth

   // Special check for Owner fields
   const isOwnerField = parentKey === 'owner_id';
   
   // Special handling for complex field types
   const isDateRange = parentKey.includes('date_range') || parentKey.includes('time_range');
   const isMoneyField = parentKey.includes('money') && obj && obj.currency !== undefined;
   const isAddressField = parentKey.includes('address');
   
   // Skip processing for fields we handle specially in the extractColumnsFromData function
   if (isDateRange || isMoneyField || isAddressField) {
     return; // Skip processing as these are handled in the parent function
   }

   for (const key in obj) {
     if (key === 'id' && parentKey.endsWith('_id')) continue; // Avoid adding parent_id.id
     if (key.startsWith('_')) continue; // Skip internal props
     if (key === 'timezone' || key === 'currency') continue; // Skip timezone/currency fields
     if (key === 'complete_address' || key === 'formatted_address') continue; // Skip duplicates
     if (key === 'admin_area_level' && !key.match(/admin_area_level_[12]/)) continue; // Skip ambiguous admin levels
     
     const value = obj[key];
     const currentPath = `${parentKey}.${key}`;
     const name = formatColumnName(currentPath);
     const isNested = true;
     const isReadOnly = isReadOnlyField(currentPath, relatedEntityType || '');

     // Use "Owner Fields" category if processing owner_id properties
     const currentCategory = isOwnerField ? "Owner Fields" : parentCategory;

     if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
       // Recursively process deeper objects
       addColumnFn(currentPath, name, isNested, parentKey, currentCategory, isReadOnly);
       processNestedObject(value, currentPath, currentCategory, relatedEntityType, addColumnFn, depth + 1);
     } else if (!Array.isArray(value)) { // Add simple properties
       addColumnFn(currentPath, name, isNested, parentKey, currentCategory, isReadOnly);
     } 
     // DISABLED: Array processing for email/phone/im fields creates confusing "Email - 0" entries
     // Users should use the main email/phone fields instead
     /*
     else if (Array.isArray(value) && (parentKey === 'email' || parentKey === 'phone' || parentKey === 'im')) {
         // Add entries for different labels (work, home, etc.) - logic handled by formatColumnName
         // Add the base array field first
         addColumnFn(parentKey, formatColumnName(parentKey), false, null, currentCategory, isReadOnly); 
         value.forEach((item, index) => {
             if(item && typeof item === 'object'){
                 for(const prop in item){
                     const itemPath = `${parentKey}.${index}.${prop}`;
                     const itemName = formatColumnName(itemPath);
                     addColumnFn(itemPath, itemName, true, parentKey, currentCategory, isReadOnly);
                 }
             }
         });
     }
     */
   }
}

/**
 * Gets standard field keys for a given entity type.
 * Placeholder: In a real scenario, this might fetch from Pipedrive API or use constants.
 */
function getStandardFieldKeys(entityType) {
    // Fields available via API v2 (or common v1) based on documentation provided
    const common = ['id', 'owner_id', 'add_time', 'update_time', 'visible_to', 'label_ids']; // Common across multiple entities
    
    switch (entityType) {
        case ENTITY_TYPES.DEALS:
            // V2 PATCH /deals/{id}
            return [...common, 'title', 'person_id', 'org_id', 'pipeline_id', 'stage_id', 'value', 'currency', 'stage_change_time', 'status', 'probability', 'lost_reason', 'close_time', 'won_time', 'lost_time', 'expected_close_date'];
        case ENTITY_TYPES.PERSONS:
            // V2 PATCH /persons/{id}
            return [...common, 'name', 'org_id', 'emails', 'phones']; // emails/phones are arrays
        case ENTITY_TYPES.ORGANIZATIONS:
            // V2 PATCH /organizations/{id}
            return [...common, 'name', 'address']; // 'address' might be complex object, handled in extraction
        case ENTITY_TYPES.ACTIVITIES:
            // V2 PATCH /activities/{id}
            return [...common, 'subject', 'type', 'deal_id', 'lead_id', 'person_id', 'org_id', 'project_id', 'due_date', 'due_time', 'duration', 'busy', 'done', 'location', 'participants', 'attendees', 'public_description', 'priority', 'note'];
        case ENTITY_TYPES.LEADS:
            // V1 PATCH /leads/{id} - Note: Inherits Deal custom fields
            return [...common, 'title', 'person_id', 'organization_id', 'is_archived', 'value', 'expected_close_date', 'was_seen', 'channel', 'channel_id', 'note']; // 'value' is object {amount, currency}
        case ENTITY_TYPES.PRODUCTS:
             // V2 PATCH /products/{id}
            return [...common, 'name', 'code', 'description', 'unit', 'tax', 'category', 'is_linkable', 'prices', 'billing_frequency', 'billing_frequency_cycles']; // 'prices' is array
        default: 
            // Fallback with most common fields if type unknown
            return ['id', 'name', 'owner_id', 'org_id', 'person_id', 'add_time', 'update_time'];
    }
}

/**
 * Format a field name for display
 * @param {string} key - The field key
 * @return {string} Formatted field name
 */
function formatFieldName(key) {
  // Use the robust formatColumnName function from Utilities
  return formatColumnName(key);
}

/**
 * Provides fallback columns when no data is available
 * @param {string} entityType - The entity type
 * @return {Array} Array of basic column objects
 */
function getFallbackColumns(entityType) {
  const mainCategory = getEntityCategoryName(entityType);
  const columns = [
    { key: 'id', name: 'Pipedrive ID', isNested: false, category: mainCategory, required: true, readOnly: false }
  ];
  
  // Add common fields based on entity type
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      columns.push(
        { key: 'title', name: 'Deal Title', isNested: false, category: mainCategory },
        { key: 'value', name: 'Deal Value', isNested: false, category: mainCategory },
        { key: 'status', name: 'Status', isNested: false, category: mainCategory },
        { key: 'stage_id', name: 'Pipeline Stage', isNested: false, category: mainCategory }
      );
      break;
    case ENTITY_TYPES.PERSONS:
      columns.push(
        { key: 'name', name: 'Name', isNested: false, category: mainCategory },
        { key: 'email', name: 'Email', isNested: false, category: mainCategory },
        { key: 'phone', name: 'Phone', isNested: false, category: mainCategory }
      );
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      columns.push(
        { key: 'name', name: 'Name', isNested: false, category: mainCategory },
        { key: 'address', name: 'Address', isNested: false, category: mainCategory }
      );
      break;
    case ENTITY_TYPES.ACTIVITIES:
      columns.push(
        { key: 'subject', name: 'Subject', isNested: false, category: mainCategory },
        { key: 'type', name: 'Type', isNested: false, category: mainCategory },
        { key: 'due_date', name: 'Due Date', isNested: false, category: mainCategory }
      );
      break;
    case ENTITY_TYPES.LEADS:
      columns.push(
        { key: 'title', name: 'Title', isNested: false, category: mainCategory },
        { key: 'lead_value', name: 'Lead Value', isNested: false, category: mainCategory }
      );
      break;
    case ENTITY_TYPES.PRODUCTS:
      columns.push(
        { key: 'name', name: 'Name', isNested: false, category: mainCategory },
        { key: 'code', name: 'Code', isNested: false, category: mainCategory },
        { key: 'prices', name: 'Prices', isNested: false, category: mainCategory }
      );
      break;
    default:
      columns.push({ key: 'name', name: 'Name', isNested: false, category: mainCategory });
  }
  
  return columns;
}

/**
 * Check if a field is read-only based on field path and entity type
 * @param {string} key - The field key
 * @param {string} entityType - The current entity type
 * @return {boolean} True if the field is read-only
 */
function isReadOnlyField(key, entityType) {
  // Entity-specific read-only checks
  if (key) {
    // Organization fields are only editable when in Organization entity type
    if ((key.startsWith('org.') || key.startsWith('org_') ||
         key.startsWith('organization.') || key.startsWith('organization_')) &&
        entityType !== ENTITY_TYPES.ORGANIZATIONS) {
      return true;
    }
    
    // Person fields are only editable when in Person entity type
    if ((key.startsWith('person.') || key.startsWith('person_')) &&
        entityType !== ENTITY_TYPES.PERSONS) {
      return true;
    }
    
    // Deal fields are only editable when in Deal entity type
    if ((key.startsWith('deal.') || key.startsWith('deal_')) &&
        entityType !== ENTITY_TYPES.DEALS) {
      return true;
    }
    
    // Activity fields are only editable when in Activity entity type
    if ((key.startsWith('activity.') || key.startsWith('activity_')) &&
        entityType !== ENTITY_TYPES.ACTIVITIES) {
      return true;
    }
    
    // Product fields are only editable when in Product entity type
    if ((key.startsWith('product.') || key.startsWith('product_')) &&
        entityType !== ENTITY_TYPES.PRODUCTS) {
      return true;
    }
    
    // Lead fields are only editable when in Lead entity type
    if ((key.startsWith('lead.') || key.startsWith('lead_')) &&
        entityType !== ENTITY_TYPES.LEADS) {
      return true;
    }
    
    // Address components should be read-only when not in Organization/Person view
    if (key.includes('address.') && 
       !(entityType === ENTITY_TYPES.ORGANIZATIONS || entityType === ENTITY_TYPES.PERSONS)) {
      return true;
    }
    
    // User/Creator/Owner related fields are always read-only
    if (key.startsWith('user.') || key.startsWith('user_') ||
        key.startsWith('creator.') || key.startsWith('creator_') ||
        key.startsWith('owner.') || key.startsWith('owner_')) {
      // All user/creator/owner fields should be read-only since they're managed by Pipedrive
      return true;
    }
    
    // Explicit user/creator fields that should be read-only
    if (key === 'user_id' || key === 'creator_id' || key === 'owner_id' ||
        key.includes('.user_id') || key.includes('.creator_id') || key.includes('.owner_id')) {
      return true;
    }
    
    // User/Creator specific fields
    if ((key.includes('has_pic') || key.includes('pic_hash') || key.includes('value')) &&
        (key.includes('user') || key.includes('creator') || key.includes('owner'))) {
      return true;
    }
    
    // Timestamps - these are auto-generated by the system
    if (key === 'add_time' ||
        key === 'update_time' ||
        key === 'stage_change_time' ||
        key === 'lost_time' ||
        key === 'close_time' ||
        key === 'won_time' ||
        key === 'local_close_date' ||
        key === 'local_won_date' ||
        key === 'local_lost_date' ||
        key === 'marked_as_done_time' ||
        key === 'last_activity_date' ||
        key === 'next_activity_time' ||
        key === 'rotten_time' ||
        key === 'last_incoming_mail_time' ||
        key === 'last_outgoing_mail_time' ||
        key === 'archive_time') {
      return true;
    }
    
    // System generated fields and IDs
    if (key === 'creator_user_id' ||
        key === 'creator_id' ||
        key === 'user_id' ||
        key === 'id' ||
        key === 'is_deleted' ||
        key === 'cc_email' ||
        key === 'origin' ||
        key === 'origin_id' ||
        key === 'source_name') {
      return true;
    }
    
    // Counts and statistics
    if (key === 'followers_count' ||
        key === 'participants_count' ||
        key === 'activities_count' ||
        key === 'done_activities_count' ||
        key === 'undone_activities_count' ||
        key === 'files_count' ||
        key === 'notes_count' ||
        key === 'email_messages_count' ||
        key === 'people_count' ||
        key === 'products_count' ||
        key === 'open_deals_count' ||
        key === 'related_open_deals_count' ||
        key === 'closed_deals_count' ||
        key === 'related_closed_deals_count' ||
        key === 'participant_open_deals_count' ||
        key === 'participant_closed_deals_count' ||
        key === 'contacts_count') {
      return true;
    }
    
    // Formatted or calculated fields
    if (key === 'formatted_value' ||
        key === 'weighted_value' ||
        key === 'formatted_weighted_value' ||
        key === 'weighted_value_currency' ||
        key === 'first_char' ||
        (key.startsWith('formatted_') && !key.includes('address'))) {
      return true;
    }
    
    // Entity-specific read-only fields
    
    // Deal-specific read-only fields
    if (entityType === ENTITY_TYPES.DEALS && 
        (key === 'stage_order_nr' ||
         key === 'person_name' ||
         key === 'org_name' ||
         key === 'next_activity_id' ||
         key === 'last_activity_id' ||
         key === 'next_activity_type' ||
         key === 'next_activity_duration' ||
         key === 'next_activity_note' ||
         key === 'acv' ||
         key === 'arr' ||
         key === 'mrr')) {
      return true;
    }
    
    // Person-specific read-only fields
    if (entityType === ENTITY_TYPES.PERSONS && 
        (key === 'org_name' ||
         key === 'owner_name' ||
         key === 'has_pic' ||
         key === 'pic_hash' ||
         key === 'next_activity_id' ||
         key === 'last_activity_id')) {
      return true;
    }
    
    // Organization-specific read-only fields
    if (entityType === ENTITY_TYPES.ORGANIZATIONS && 
        (key === 'owner_name' ||
         key === 'next_activity_id' ||
         key === 'last_activity_id' ||
         key === 'has_pic' ||
         key === 'pic_hash' ||
         key === 'owner_name')) {
      return true;
    }
    
    // Lead-specific read-only fields
    if (entityType === ENTITY_TYPES.LEADS && 
        (key === 'was_seen' ||
         key === 'next_activity_id')) {
      return true;
    }
    
    // Product-specific read-only fields
    if (entityType === ENTITY_TYPES.PRODUCTS && 
        (key === 'first_char' ||
         key === 'active_flag' ||
         key === 'selectable')) {
      return true;
    }
    
    // Activity-specific read-only fields
    if (entityType === ENTITY_TYPES.ACTIVITIES && 
        (key === 'company_id' ||
         key === 'user_id' ||
         key === 'assigned_to_user_id' ||
         key === 'conference_meeting_client' ||
         key === 'conference_meeting_url' ||
         key === 'conference_meeting_id')) {
      return true;
    }
    
    // Check for patterns that indicate read-only fields
    if (/_name$/.test(key) || // Fields ending with _name (e.g., owner_name)
        /_email$/.test(key) || // Fields ending with _email
        /\.name$/.test(key) || // Nested name fields (e.g., owner_id.name)
        /\.email$/.test(key) || // Nested email fields
        /^cc_/.test(key) || // Fields starting with cc_
        /_count$/.test(key) || // Count fields
        /_flag$/.test(key) || // Flag fields
        /_hash$/.test(key) || // Hash fields (pic_hash)
        /has_pic/.test(key)) { // Has pic fields
      return true;
    }
    
    // Special cases that should be editable even though they might match above patterns
    if (key === 'name' || 
        key === 'first_name' || 
        key === 'last_name' ||
        key === 'label_ids') {
      return false;
    }
  }
  
  return false;
}

/**
 * Client-callable function to get a formatted column name
 * @param {string} key - The field key to format
 * @return {string} Formatted column name
 */
function getFormattedColumnName(key) {
  return formatColumnName(key);
}

/**
 * Improves column names for UI display
 * @param {Array} columns - Array of column objects
 * @param {string} entityType - Entity type for context
 * @return {Array} Array of columns with improved names
 */
function improveColumnNamesForUI(columns, entityType) {
  try {
    Logger.log(`Improving column names for ${columns.length} columns of type ${entityType}`);
    
    // First pass: identify date/time range fields
    const dateRangeFields = new Map();
    
    columns.forEach(column => {
      // Look for date/time range fields that might be part of a pair
      if (column.key.match(/[a-f0-9]{20,}$/i) && // Custom field hash
          (column.name.toLowerCase().includes('date range') || 
           column.name.toLowerCase().includes('time range'))) {
        // This might be a date/time range start field
        dateRangeFields.set(column.key, {
          startCol: column,
          endCol: null
        });
      }
      else if (column.key.match(/[a-f0-9]{20,}_until$/i)) {
        // This is an end field - find its corresponding start field
        const baseKey = column.key.replace(/_until$/, '');
        if (dateRangeFields.has(baseKey)) {
          // Link it with its start field
          const pair = dateRangeFields.get(baseKey);
          pair.endCol = column;
        }
      }
    });
    
    // Second pass: improve column names
    return columns.map(column => {
      // Skip if it already has a custom name set
      if (column.customName) {
        return column;
      }
      
      // Special handling for date/time range start fields
      if (dateRangeFields.has(column.key)) {
        const pair = dateRangeFields.get(column.key);
        if (pair.startCol === column) {
          // This is a start field - make sure it's labeled as such
          if (!column.name.includes('Start')) {
            column.name = `${formatColumnName(column.key)} - Start`;
          }
        }
      }
      
      // Special handling for date/time range end fields
      else if (column.key.match(/[a-f0-9]{20,}_until$/i)) {
        const baseKey = column.key.replace(/_until$/, '');
        if (dateRangeFields.has(baseKey)) {
          // This is an end field linked to a known start field
          const startField = dateRangeFields.get(baseKey).startCol;
          const baseName = startField.name.replace(/ - Start$/, '');
          column.name = `${baseName} - End`;
        } else {
          // End field without a linked start field
          column.name = `${formatColumnName(baseKey)} - End`;
        }
      }
      
      // Special handling for date/time range custom fields
      else if (column.key.match(/[a-f0-9]{20,}/i) && column.key.match(/_until$/i)) {
        // This is a custom field with _until suffix (end date/time)
        const baseKey = column.key.replace(/_until$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - End`;
      }
      // Special handling for date/time range fields with start/end components
      else if ((column.key.includes('date_range') || column.key.includes('time_range')) && 
               column.key.match(/\.(start|end)$/)) {
        const isStart = column.key.endsWith('.start');
        const baseKey = column.key.replace(/\.(start|end)$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - ${isStart ? 'Start' : 'End'}`;
      }
      
      // Special handling for address components
      else if (column.key.includes('address.') && column.isNested) {
        const parentName = formatColumnName(column.parentKey);
        
        // Map address components to user-friendly names
        const addressComponentNames = {
          'street_number': 'Street Number',
          'route': 'Street Name',
          'sublocality': 'District',
          'locality': 'City',
          'admin_area_level_1': 'State/Province',
          'admin_area_level_2': 'County',
          'country': 'Country',
          'postal_code': 'ZIP/Postal Code',
          'subpremise': 'Suite/Apt'
        };
        
        // Extract the component name from the key
        const parts = column.key.split('.');
        const component = parts[parts.length - 1];
        
        if (addressComponentNames[component]) {
          column.name = `${parentName} - ${addressComponentNames[component]}`;
        }
      }
      
      // Special handling for custom field address components
      else if (column.key.match(/[a-f0-9]{20,}_admin_area_level_1/i)) {
        // This is a custom field admin_area_level_1 (state/province)
        const baseKey = column.key.replace(/_admin_area_level_1$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - State/Province`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_admin_area_level_2/i)) {
        // This is a custom field admin_area_level_2 (county)
        const baseKey = column.key.replace(/_admin_area_level_2$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - County`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_locality/i)) {
        // This is a custom field locality (city)
        const baseKey = column.key.replace(/_locality$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - City`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_country/i)) {
        // This is a custom field country
        const baseKey = column.key.replace(/_country$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - Country`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_postal_code/i)) {
        // This is a custom field postal_code
        const baseKey = column.key.replace(/_postal_code$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - ZIP/Postal Code`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_route/i)) {
        // This is a custom field route (street name)
        const baseKey = column.key.replace(/_route$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - Street Name`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_street_number/i)) {
        // This is a custom field street_number
        const baseKey = column.key.replace(/_street_number$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - Street Number`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_sublocality/i)) {
        // This is a custom field sublocality (district)
        const baseKey = column.key.replace(/_sublocality$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - District`;
      }
      else if (column.key.match(/[a-f0-9]{20,}_subpremise/i)) {
        // This is a custom field subpremise (suite/apt)
        const baseKey = column.key.replace(/_subpremise$/, '');
        const parentName = formatColumnName(baseKey);
        column.name = `${parentName} - Suite/Apt`;
      }
      // Default handling for regular columns
      else {
        // Rely entirely on formatColumnName for the correct name
        column.name = formatColumnName(column.key);
        
        // For nested fields (not already handled by special cases)
        if (column.isNested) {
          // Update name to reflect parent relationship
          const parentKey = column.parentKey;
          
          if (parentKey) { // Check if parentKey is valid
            let parentName = '';
          
            // Get parent name using formatColumnName to ensure consistency
            parentName = formatColumnName(parentKey);
            
            // Avoid redundant parent name prefixes if already present
            if (!column.name.startsWith(parentName)) {
              column.name = `${parentName} - ${column.name}`;
            }
          } else {
            // If parent key exists but we couldn't format it, use a fallback
            column.name = `${formatBasicName(parentKey)} - ${column.name}`;
          }
        }
      }
      
      return column;
    });
  } catch (e) {
    Logger.log(`Error improving column names: ${e.message}`);
    // Return original columns if there's an error
    return columns;
  }
}