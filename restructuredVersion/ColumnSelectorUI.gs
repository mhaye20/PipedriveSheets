/**
 * Column Selector UI Helper
 * 
 * This module provides helper functions for the column selector UI:
 * - Styles for the column selector dialog
 * - Scripts for the column selector dialog
 */

var ColumnSelectorUI = ColumnSelectorUI || {};

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
                      readOnly: true // Address components are read-only
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
      if (sampleItem.email && typeof sampleItem.email === 'object') {
        availableColumns.push({
          key: 'email',
          name: 'Email',
          isNested: false
        });
        
        extractFields(sampleItem.email, 'email', 'Email');
      }
      
      if (sampleItem.phone && typeof sampleItem.phone === 'object') {
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
        
        // Skip problematic objects that cause clutter
        if (['im', 'lm', 'owner', 'first_char', 'label', 'labels'].includes(key)) {
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
      if (col.key.startsWith('im.') || 
          col.key.startsWith('lm.') ||
          col.key === 'im' || 
          col.key === 'lm' ||
          col.key === 'first_char' ||
          col.key === 'labels' ||
          col.key === 'label') {
        return false;
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

      return true;
    });
    
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
      // Mark address components as read-only based on Pipedrive API docs
      if (col.key && 
          (col.key.includes('.street_number') || 
           col.key.includes('.route') || 
           col.key.includes('.subpremise') || 
           col.key.includes('.sublocality') || 
           col.key.includes('.locality') || 
           col.key.includes('.admin_area_level_1') || 
           col.key.includes('.admin_area_level_2') || 
           col.key.includes('.country') || 
           col.key.includes('.postal_code') ||
           col.key.includes('_street_number') || 
           col.key.includes('_route') || 
           col.key.includes('_subpremise') || 
           col.key.includes('_sublocality') || 
           col.key.includes('_locality') || 
           col.key.includes('_admin_area_level_1') || 
           col.key.includes('_admin_area_level_2') || 
           col.key.includes('_country') || 
           col.key.includes('_postal_code'))) {
        col.readOnly = true;
      }
      
      // Mark read-only fields based on Pipedrive API documentation
      // These fields are not included in the update endpoints or are system-generated
      
      // Common read-only fields across all entity types
      if (col.key && (
          // ID fields (except the main ID used for identification)
          (col.key !== 'id' && col.key.endsWith('_id') && col.key !== 'owner_id' && 
           col.key !== 'person_id' && col.key !== 'org_id' && col.key !== 'organization_id' && 
           col.key !== 'deal_id' && col.key !== 'lead_id' && col.key !== 'stage_id' && 
           col.key !== 'pipeline_id' && col.key !== 'project_id' && col.key !== 'category') ||
          
          // System-generated fields
          col.key === 'creator_user_id' || 
          col.key === 'user_id' || 
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
          col.key.startsWith('formatted_') ||
          
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
          
          // Label information (can only be modified via label_ids)
          col.key === 'label' ||
          
          // Deal-specific read-only fields
          col.key === 'stage_order_nr' ||
          col.key === 'person_name' ||
          col.key === 'org_name' ||
          col.key === 'origin' ||
          col.key === 'origin_id' ||
          col.key === 'channel' ||
          
          // Person/Org specific
          col.key === 'has_pic' ||
          col.key === 'pic_hash' ||
          col.key === 'owner_name' ||
          col.key === 'org_hidden' ||
          col.key === 'person_hidden'
      )) {
        col.readOnly = true;
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
      return {
        key: col.key,
        name: col.name || formatColumnName(col.key),
        customName: col.customName || '',
        isNested: col.isNested,
        parentKey: col.parentKey
      };
    });
    
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