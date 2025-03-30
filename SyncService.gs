/**
 * Sync Service
 * 
 * This module handles the synchronization between Pipedrive and Google Sheets:
 * - Fetching data from Pipedrive and writing to sheets
 * - Tracking modifications and pushing changes back to Pipedrive
 * - Managing synchronization status and scheduling
 */

/**
 * Main synchronization function that syncs data from Pipedrive to the sheet
 */
function syncFromPipedrive() {
  // Show a loading message
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();

  detectColumnShifts();

  // Get the script properties
  const docProps = PropertiesService.getDocumentProperties();

  // Check if two-way sync is enabled for this sheet
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
  const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
  const twoWaySyncEnabled = docProps.getProperty(twoWaySyncEnabledKey) === 'true';

  // Get the current entity type for this specific sheet
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  const currentEntityType = docProps.getProperty(sheetEntityTypeKey);
  const lastEntityTypeKey = `LAST_ENTITY_TYPE_${activeSheetName}`;
  const lastEntityType = docProps.getProperty(lastEntityTypeKey);

  // Debug log to help troubleshoot
  Logger.log(`Syncing sheet ${activeSheetName}, entity type: ${currentEntityType}, sheet key: ${sheetEntityTypeKey}`);

  // Check if entity type has changed
  if (currentEntityType !== lastEntityType && currentEntityType && twoWaySyncEnabled) {
    // Entity type has changed - clear the Sync Status column letter
    Logger.log(`Entity type changed from ${lastEntityType || 'none'} to ${currentEntityType}. Clearing tracking column.`);
    docProps.deleteProperty(twoWaySyncTrackingColumnKey);

    // Add a flag to indicate that the Sync Status column should be repositioned
    const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${activeSheetName}`;
    docProps.setProperty(twoWaySyncColumnAtEndKey, 'true');

    // Store the new entity type as the last synced entity type
    docProps.setProperty(lastEntityTypeKey, currentEntityType);
  }

  // First show a confirmation dialog, including info about pushing changes if two-way sync is enabled
  let confirmMessage = `This will sync data from Pipedrive to the current sheet "${activeSheetName}". Any existing data in this sheet will be replaced.`;

  if (twoWaySyncEnabled) {
    confirmMessage += `\n\nTwo-way sync is enabled for this sheet. Modified rows will be pushed to Pipedrive before pulling new data.`;
  }

  confirmMessage += `\n\nDo you want to continue?`;

  const confirmation = ui.alert(
    'Sync Pipedrive Data',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );

  if (confirmation === ui.Button.NO) {
    return;
  }

  // If two-way sync is enabled, push changes to Pipedrive first
  if (twoWaySyncEnabled) {
    // Show a message that we're pushing changes
    spreadsheet.toast('Two-way sync enabled. Pushing modified rows to Pipedrive first...', 'Syncing', 5);

    try {
      // Call pushChangesToPipedrive() with true for isScheduledSync to suppress duplicate UI messages
      pushChangesToPipedrive(false, true);
    } catch (error) {
      // Log the error but continue with the sync
      Logger.log(`Error pushing changes: ${error.message}`);
      spreadsheet.toast(`Warning: Error pushing changes to Pipedrive: ${error.message}`, 'Sync Warning', 10);
    }
  }

  // Show a sync status UI
  showSyncStatus(activeSheetName);

  // Set the active sheet as the current sheet for this operation
  docProps.setProperty('SHEET_NAME', activeSheetName);

  try {
    if (!currentEntityType) {
      ui.alert(
        'No entity type configured',
        'Please configure Pipedrive settings for this sheet first.',
        ui.ButtonSet.OK
      );
      showSettings();
      return;
    }
    
    switch (currentEntityType) {
      case ENTITY_TYPES.DEALS:
        syncDealsFromFilter(true);
        break;
      case ENTITY_TYPES.PERSONS:
        syncPersonsFromFilter(true);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        syncOrganizationsFromFilter(true);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        syncActivitiesFromFilter(true);
        break;
      case ENTITY_TYPES.LEADS:
        syncLeadsFromFilter(true);
        break;
      case ENTITY_TYPES.PRODUCTS:
        syncProductsFromFilter(true);
        break;
      default:
        spreadsheet.toast('Unknown entity type. Please check settings.', 'Sync Error', 10);
        break;
    }

    // After successful sync, update the last entity type
    docProps.setProperty(lastEntityTypeKey, currentEntityType);
  } catch (error) {
    // If there's an error, show it
    spreadsheet.toast('Error: ' + error.message, 'Sync Error', 10);
    Logger.log('Sync error: ' + error.message);
    updateSyncStatus('error', 'Failed', error.message, 100);
  }
}

/**
 * Synchronizes Pipedrive data to the sheet based on entity type
 * @param {string} entityType - The type of entity to sync
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncPipedriveDataToSheet(entityType, skipPush = false) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const filterId = docProps.getProperty('PIPEDRIVE_FILTER_ID');
    const sheetName = docProps.getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    updateSyncStatus('retrieving', 'In Progress', `Retrieving ${entityType} from Pipedrive...`, 10);
    
    // Get filtered data based on entity type
    let items = [];
    let fields = [];
    
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        items = getDealsWithFilter(filterId);
        fields = getDealFields();
        break;
      case ENTITY_TYPES.PERSONS:
        items = getPersonsWithFilter(filterId);
        fields = getPersonFields();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        items = getOrganizationsWithFilter(filterId);
        fields = getOrganizationFields();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        items = getActivitiesWithFilter(filterId);
        fields = getActivityFields();
        break;
      case ENTITY_TYPES.LEADS:
        items = getLeadsWithFilter(filterId);
        fields = getLeadFields();
        break;
      case ENTITY_TYPES.PRODUCTS:
        items = getProductsWithFilter(filterId);
        fields = getProductFields();
        break;
      default:
        throw new Error(`Unsupported entity type: ${entityType}`);
    }
    
    if (!items || items.length === 0) {
      updateSyncStatus('complete', 'No Data', `No ${entityType} found with the selected filter.`, 100);
      return;
    }
    
    // Update status
    updateSyncStatus('processing', 'In Progress', `Processing ${items.length} ${entityType}...`, 30);
    
    // Get team-aware column preferences
    const selectedColumns = getTeamAwareColumnPreferences(entityType, sheetName);
    
    // If no columns are selected, select default ones or all
    if (!selectedColumns || selectedColumns.length === 0) {
      // Set default columns based on entity type
      let defaultColumns = ['id', 'name', 'owner_id', 'created_at', 'updated_at'];
      
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          defaultColumns = ['id', 'title', 'status', 'value', 'currency', 'owner_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.PERSONS:
          defaultColumns = ['id', 'name', 'email', 'phone', 'owner_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          defaultColumns = ['id', 'name', 'address', 'owner_id', 'created_at', 'updated_at'];
          break;
      }
      
      updateSyncStatus('processing', 'In Progress', `No columns selected, using defaults`, 40);
      
      // Save the default columns as preferences
      saveTeamAwareColumnPreferences(defaultColumns, entityType, sheetName);
    }
    
    // Process the data
    updateSyncStatus('formatting', 'In Progress', `Formatting data for spreadsheet...`, 50);
    
    // Get field option mappings
    const optionMappings = getFieldOptionMappingsForEntity(entityType);
    
    // Prepare for writing to sheet
    const isWriteTimestamp = docProps.getProperty('ENABLE_TIMESTAMP') === 'true';
    const isTwoWaySync = docProps.getProperty('ENABLE_TWO_WAY_SYNC') === 'true';
    
    // Determine options for writing data to sheet
    const writeOptions = {
      sheetName: sheetName,
      entityType: entityType,
      isWriteTimestamp: isWriteTimestamp,
      isTwoWaySync: isTwoWaySync,
      trackingColumn: docProps.getProperty('SYNC_TRACKING_COLUMN'),
      selectedColumns: selectedColumns || [],
      optionMappings: optionMappings,
      fields: fields
    };
    
    // Write data to sheet
    updateSyncStatus('writing', 'In Progress', `Writing data to sheet...`, 70);
    writeDataToSheet(items, writeOptions);
    
    // Update the sync status to complete
    updateSyncStatus('complete', 'Success', `Successfully synced ${items.length} ${entityType} to sheet.`, 100);
    
    // If two-way sync is enabled and we're not skipping push, refresh the sync status formatting
    if (isTwoWaySync && !skipPush) {
      refreshSyncStatusStyling();
    }
  } catch (e) {
    Logger.log(`Error in syncPipedriveDataToSheet: ${e.message}`);
    updateSyncStatus('error', 'Failed', e.message, 100);
    throw e;
  }
}

/**
 * Writes the data to the sheet
 * @param {Array} items - The Pipedrive items to write
 * @param {Object} options - Options for writing data
 */
function writeDataToSheet(items, options) {
  try {
    const { 
      sheetName, 
      entityType, 
      isWriteTimestamp, 
      isTwoWaySync,
      trackingColumn,
      selectedColumns,
      optionMappings,
      fields 
    } = options;
    
    // Get the active spreadsheet and destination sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Clear existing content
    sheet.clear();
    
    // Determine columns to use
    let columnsToPull = selectedColumns;
    
    // If no columns specified, use all available
    if (!columnsToPull || columnsToPull.length === 0) {
      columnsToPull = fields.map(field => field.key);
    }
    
    // Create header row
    const headerRow = ['ID'];
    
    // Add selected columns to header
    for (const columnPath of columnsToPull) {
      // Get user-friendly column name
      for (const field of fields) {
        if (field.key === columnPath) {
          headerRow.push(formatColumnName(field.name));
          break;
        }
      }
    }
    
    // Add timestamp column if enabled
    if (isWriteTimestamp) {
      headerRow.push('Last Sync Time');
    }
    
    // Add sync status column if two-way sync is enabled
    if (isTwoWaySync) {
      headerRow.push('Sync Status');
    }
    
    // Write header row
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Prepare data rows
    const dataRows = [];
    
    for (const item of items) {
      const row = [item.id.toString()];
      
      // Add values for selected columns
      for (const columnPath of columnsToPull) {
        let value = getValueByPath(item, columnPath);
        value = formatValue(value, columnPath, optionMappings);
        row.push(value);
      }
      
      // Add timestamp if enabled
      if (isWriteTimestamp) {
        row.push(new Date().toLocaleString());
      }
      
      // Add sync status if two-way sync is enabled
      if (isTwoWaySync) {
        row.push('Synced');
      }
      
      dataRows.push(row);
    }
    
    // Write data rows
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
    }
    
    // Apply formatting
    sheet.autoResizeColumns(1, headerRow.length);
    
    // Apply conditional formatting to sync status column if two-way sync is enabled
    if (isTwoWaySync) {
      const syncStatusColumnIndex = headerRow.length;
      const syncStatusColumnLetter = columnToLetter(syncStatusColumnIndex);
      
      // Create conditional formatting rules
      const range = sheet.getRange(`${syncStatusColumnLetter}2:${syncStatusColumnLetter}${dataRows.length + 1}`);
      
      // Clear existing rules
      const rules = sheet.getConditionalFormatRules();
      sheet.clearConditionalFormatRules();
      
      // Add rule for "Modified" status
      let rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FFEB3B')
        .setRanges([range])
        .build();
      rules.push(rule);
      
      // Add rule for "Synced" status
      rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#C8E6C9')
        .setRanges([range])
        .build();
      rules.push(rule);
      
      // Add rule for "Sync Failed" status
      rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Sync Failed')
        .setBackground('#FFCDD2')
        .setRanges([range])
        .build();
      rules.push(rule);
      
      // Set the rules
      sheet.setConditionalFormatRules(rules);
      
      // Set up onEdit trigger if it doesn't exist
      setupOnEditTrigger();
    }
    
    // Freeze the header row
    sheet.setFrozenRows(1);
    
    return true;
  } catch (e) {
    Logger.log(`Error in writeDataToSheet: ${e.message}`);
    throw e;
  }
}

/**
 * Updates the sync status in the UI
 * @param {string} phase - Current phase of sync
 * @param {string} status - Status text
 * @param {string} detail - Detailed status message
 * @param {number} progress - Progress percentage (0-100)
 */
function updateSyncStatus(phase, status, detail, progress) {
  try {
    // Get/create sync status object
    const syncStatus = getSyncStatus() || {
      phase: '',
      status: '',
      detail: '',
      progress: 0,
      startTime: new Date().getTime()
    };
    
    // Update fields
    syncStatus.phase = phase;
    syncStatus.status = status;
    syncStatus.detail = detail;
    syncStatus.progress = progress;
    syncStatus.lastUpdate = new Date().getTime();
    
    // Save to user properties
    PropertiesService.getUserProperties().setProperty('SYNC_STATUS', JSON.stringify(syncStatus));
    
    return syncStatus;
  } catch (e) {
    Logger.log(`Error updating sync status: ${e.message}`);
    return null;
  }
}

/**
 * Gets the current sync status
 * @return {Object} Sync status object or null if not available
 */
function getSyncStatus() {
  try {
    const statusJson = PropertiesService.getUserProperties().getProperty('SYNC_STATUS');
    return statusJson ? JSON.parse(statusJson) : null;
  } catch (e) {
    Logger.log(`Error getting sync status: ${e.message}`);
    return null;
  }
}

/**
 * Converts a column index to a letter (e.g., 1 = A, 27 = AA)
 * @param {number} columnIndex - 1-based column index
 * @return {string} Column letter
 */
function columnToLetter(columnIndex) {
  let temp;
  let letter = '';
  let col = columnIndex;
  
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  
  return letter;
}

/**
 * Sets up the onEdit trigger to track changes for two-way sync
 */
function setupOnEditTrigger() {
  try {
    // Check if the trigger already exists
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onEdit') {
        // Trigger already exists
        return;
      }
    }
    
    // Create the trigger
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
      
    Logger.log('onEdit trigger created');
  } catch (e) {
    Logger.log(`Error setting up onEdit trigger: ${e.message}`);
  }
}

/**
 * Main function to sync deals from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncDealsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.DEALS, skipPush);
}

/**
 * Main function to sync persons from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncPersonsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PERSONS, skipPush);
}

/**
 * Main function to sync organizations from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncOrganizationsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ORGANIZATIONS, skipPush);
}

/**
 * Main function to sync activities from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncActivitiesFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ACTIVITIES, skipPush);
}

/**
 * Main function to sync leads from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncLeadsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.LEADS, skipPush);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 */
function syncProductsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush);
}

/**
 * Pushes changes from the sheet back to Pipedrive
 * @param {boolean} isScheduledSync - Whether this is called from a scheduled sync
 * @param {boolean} suppressNoModifiedWarning - Whether to suppress the no modified rows warning
 */
function pushChangesToPipedrive(isScheduledSync = false, suppressNoModifiedWarning = false) {
  detectColumnShifts();
  try {
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    if (!twoWaySyncEnabled) {
      // Show an error message if two-way sync is not enabled, only for manual syncs
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Two-Way Sync Not Enabled',
          'Two-way sync is not enabled for this sheet. Please enable it in the Two-Way Sync Settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Get sheet-specific entity type
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Get API key from properties
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY');
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

    if (!apiKey) {
      // Show an error message if API key is not set, only for manual syncs
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'API Key Not Found',
          'Please set your Pipedrive API key in the Settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Get the original column configuration that maps headers to field keys
    const columnSettingsKey = `COLUMNS_${activeSheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let columnConfig = [];

    if (savedColumnsJson) {
      try {
        columnConfig = JSON.parse(savedColumnsJson);
        Logger.log(`Retrieved column configuration for ${entityType}`);
      } catch (e) {
        Logger.log(`Error parsing column configuration: ${e.message}`);
      }
    }

    // Create a mapping from display names to field keys
    const headerToFieldKeyMap = {};
    columnConfig.forEach(col => {
      const displayName = col.customName || col.name;
      headerToFieldKeyMap[displayName] = col.key;
    });

    // Get the tracking column
    let trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    let trackingColumnIndex;

    Logger.log(`Retrieved tracking column from properties: "${trackingColumn}"`);

    // Look for a column named "Sync Status"
    const headerRow = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
    let syncStatusColumnIndex = -1;

    // First, try to use the stored tracking column if available
    if (trackingColumn && trackingColumn.trim() !== '') {
      trackingColumnIndex = columnLetterToIndex(trackingColumn);
      // Verify the header matches "Sync Status"
      if (trackingColumnIndex >= 0 && trackingColumnIndex < headerRow.length) {
        if (headerRow[trackingColumnIndex] === "Sync Status") {
          Logger.log(`Using configured tracking column ${trackingColumn} (index: ${trackingColumnIndex})`);
          syncStatusColumnIndex = trackingColumnIndex;
        }
      }
    }

    // If not found by letter or the header doesn't match, search for "Sync Status" header
    if (syncStatusColumnIndex === -1) {
      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] === "Sync Status") {
          syncStatusColumnIndex = i;
          Logger.log(`Found Sync Status column at index ${syncStatusColumnIndex}`);

          // Update the stored tracking column letter
          trackingColumn = columnToLetter(syncStatusColumnIndex);
          scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
          break;
        }
      }
    }

    // Use the found column index
    trackingColumnIndex = syncStatusColumnIndex;

    // Validate tracking column index
    if (trackingColumnIndex < 0 || trackingColumnIndex >= activeSheet.getLastColumn()) {
      Logger.log(`Invalid tracking column index ${trackingColumnIndex}, cannot proceed with sync`);
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Sync Status Column Not Found',
          'The Sync Status column could not be found. Please check your two-way sync settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Double check the header of the tracking column
    const trackingHeader = activeSheet.getRange(1, trackingColumnIndex + 1).getValue();
    Logger.log(`Tracking column header: "${trackingHeader}" at column index ${trackingColumnIndex} (column ${trackingColumnIndex + 1})`);

    // Get the data range
    const dataRange = activeSheet.getDataRange();
    const values = dataRange.getValues();

    // Get column headers (first row)
    const headers = values[0];

    // Find the ID column index (usually first column)
    const idColumnIndex = 0; // Assuming ID is in column A (index 0)

    // Get field mappings based on entity type
    const fieldMappings = getFieldMappingsForEntity(entityType);

    // Track rows that need updating
    const modifiedRows = [];

    // Collect modified rows
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const syncStatus = row[trackingColumnIndex];

      // Only process rows marked as "Modified"
      if (syncStatus === 'Modified') {
        // Use let instead of const for rowId since we might need to change it
        let rowId = row[idColumnIndex];

        // Skip rows without an ID
        if (!rowId) {
          continue;
        }

        // For products, ensure we're using the correct ID field
        if (entityType === ENTITY_TYPES.PRODUCTS) {
          // If the first column contains a name instead of an ID, 
          // try to look for the ID in another column
          const idColumnName = 'ID'; // This should match your header for the ID column
          const idColumnIdx = headers.indexOf(idColumnName);
          if (idColumnIdx !== -1 && idColumnIdx !== idColumnIndex) {
            // Use the value from the specific ID column
            rowId = row[idColumnIdx]; // Using let above allows this reassignment
            Logger.log(`Using product ID ${rowId} from column ${idColumnName} instead of first column value ${row[idColumnIndex]}`);
          }
        }

        // Create an object with field values to update
        const updateData = {};

        // For API v2 custom fields
        if (!entityType.endsWith('Fields') && entityType !== ENTITY_TYPES.LEADS) {
          // Initialize custom fields container for API v2 as an object, not an array
          updateData.custom_fields = {};
        }

        // Map column values to API fields
        for (let j = 0; j < headers.length; j++) {
          // Skip the tracking column
          if (j === trackingColumnIndex) {
            continue;
          }

          const header = headers[j];
          const value = row[j];

          // Skip empty values
          if (value === '' || value === null || value === undefined) {
            continue;
          }

          // Get the field key for this header - first try the stored column config
          const fieldKey = headerToFieldKeyMap[header] || fieldMappings[header];

          if (fieldKey) {
            // Check if this is a multi-option field (this handles both custom and standard fields)
            if (isMultiOptionField(fieldKey, entityType)) {
              // Convert multi-option values (comma-separated in sheet) to array for API
              if (typeof value === 'string' && value.includes(',')) {
                const optionValues = value.split(',').map(option => option.trim());
                
                // Convert option labels to IDs
                const optionIds = optionValues.map(optionLabel => 
                  getOptionIdByLabel(fieldKey, optionLabel, entityType)
                ).filter(id => id !== null);
                
                // Only set the field if we have valid option IDs
                if (optionIds.length > 0) {
                  // Custom fields go in custom_fields object, standard fields at root
                  if (fieldKey.startsWith('custom_fields')) {
                    // For custom fields, extract the actual field key from the path
                    const customFieldKey = fieldKey.replace('custom_fields.', '');
                    updateData.custom_fields[customFieldKey] = optionIds;
                  } else {
                    updateData[fieldKey] = optionIds;
                  }
                }
              } else if (value) {
                // Handle single option for multi-option fields
                const optionId = getOptionIdByLabel(fieldKey, value, entityType);
                if (optionId !== null) {
                  if (fieldKey.startsWith('custom_fields')) {
                    const customFieldKey = fieldKey.replace('custom_fields.', '');
                    updateData.custom_fields[customFieldKey] = [optionId];
                  } else {
                    updateData[fieldKey] = [optionId];
                  }
                }
              }
            }
            // Handle date fields (convert from sheet format to API format)
            else if (isDateField(fieldKey)) {
              if (value instanceof Date) {
                // Convert to ISO format for API
                const isoDate = value.toISOString().split('T')[0];
                
                if (fieldKey.startsWith('custom_fields')) {
                  const customFieldKey = fieldKey.replace('custom_fields.', '');
                  updateData.custom_fields[customFieldKey] = isoDate;
                } else {
                  updateData[fieldKey] = isoDate;
                }
              } else if (typeof value === 'string') {
                // Try to parse the string as a date
                const apiDateFormat = convertToStandardDateFormat(value);
                
                if (apiDateFormat) {
                  if (fieldKey.startsWith('custom_fields')) {
                    const customFieldKey = fieldKey.replace('custom_fields.', '');
                    updateData.custom_fields[customFieldKey] = apiDateFormat;
                  } else {
                    updateData[fieldKey] = apiDateFormat;
                  }
                }
              }
            }
            // All other fields
            else {
              if (fieldKey.startsWith('custom_fields')) {
                const customFieldKey = fieldKey.replace('custom_fields.', '');
                updateData.custom_fields[customFieldKey] = value;
              } else {
                updateData[fieldKey] = value;
              }
            }
          }
        }

        // Only add rows with actual data to update
        if (Object.keys(updateData).length > 0 || (updateData.custom_fields && Object.keys(updateData.custom_fields).length > 0)) {
          modifiedRows.push({
            id: rowId,
            rowIndex: i,
            data: updateData
          });
        }
      }
    }

    // If there are no modified rows, show a message and exit
    if (modifiedRows.length === 0) {
      if (!suppressNoModifiedWarning && !isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'No Modified Rows',
          'No rows marked as "Modified" were found. Edit cells in rows where the Sync Status column shows "Synced" to mark them for update.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Show confirmation for manual syncs
    if (!isScheduledSync) {
      const ui = SpreadsheetApp.getUi();
      const result = ui.alert(
        'Confirm Push to Pipedrive',
        `You are about to push ${modifiedRows.length} modified row(s) to Pipedrive. Continue?`,
        ui.ButtonSet.YES_NO
      );

      if (result !== ui.Button.YES) {
        return;
      }
    }

    // Show a toast message
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Pushing ${modifiedRows.length} modified row(s) to Pipedrive...`,
      'Push to Pipedrive',
      30
    );

    // Build the base API URL
    const apiUrl = getPipedriveApiUrl();
    
    // Track success and failure counts
    let successCount = 0;
    let failureCount = 0;
    const failures = [];

    // Update each modified row in Pipedrive
    for (const rowData of modifiedRows) {
      try {
        // Set up the request URL based on entity type
        let updateUrl = `${apiUrl}/${entityType}/${rowData.id}?api_token=${apiKey}`;
        let method = 'PUT';
        
        // Special case for activities which use a different endpoint for updates
        if (entityType === ENTITY_TYPES.ACTIVITIES) {
          updateUrl = `${apiUrl}/${entityType}/${rowData.id}?api_token=${apiKey}`;
        }

        // Make the API call
        const options = {
          method: method,
          contentType: 'application/json',
          payload: JSON.stringify(rowData.data),
          muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(updateUrl, options);
        const statusCode = response.getResponseCode();
        const responseBody = response.getContentText();
        const responseJson = JSON.parse(responseBody);

        if (statusCode >= 200 && statusCode < 300 && responseJson.success) {
          // Update was successful
          successCount++;
          
          // Update the tracking column to "Synced"
          activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setValue('Synced');
          
          // Add a timestamp if desired
          if (scriptProperties.getProperty('ENABLE_TIMESTAMP') === 'true') {
            const timestamp = new Date().toLocaleString();
            activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setNote(`Last sync: ${timestamp}`);
          }
        } else {
          // Update failed
          failureCount++;
          
          // Set the sync status to "Error" with a note about the error
          activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setValue('Error');
          activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setNote(
            `API Error: ${statusCode} - ${responseJson.error || 'Unknown error'}`
          );
          
          // Track the failure
          failures.push({
            id: rowData.id,
            error: responseJson.error || `Status code: ${statusCode}`
          });
        }
      } catch (error) {
        // Handle exceptions
        failureCount++;
        Logger.log(`Error updating row ${rowData.rowIndex + 1}: ${error.message}`);
        
        // Update the cell to show the error
        activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setValue('Error');
        activeSheet.getRange(rowData.rowIndex + 1, trackingColumnIndex + 1).setNote(`Error: ${error.message}`);
        
        failures.push({
          id: rowData.id,
          error: error.message
        });
      }
    }

    // Update formatting for the tracking column
    refreshSyncStatusStyling();

    // Update the last sync time
    const now = new Date().toISOString();
    scriptProperties.setProperty(twoWaySyncLastSyncKey, now);

    // Show a summary message
    if (!isScheduledSync) {
      let message = `Push complete: ${successCount} row(s) updated successfully`;
      if (failureCount > 0) {
        message += `, ${failureCount} failed. See notes in the "Sync Status" column for error details.`;
      } else {
        message += '.';
      }

      SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Push to Pipedrive', 10);

      if (failureCount > 0) {
        // Show more details about failures
        let failureDetails = 'The following records had errors:\n\n';
        failures.forEach(failure => {
          failureDetails += `- ID ${failure.id}: ${failure.error}\n`;
        });

        const ui = SpreadsheetApp.getUi();
        ui.alert('Sync Errors', failureDetails, ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    Logger.log(`Error in pushChangesToPipedrive: ${error.message}`);
    
    // Show error message for manual syncs
    if (!isScheduledSync) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error', `Failed to push changes to Pipedrive: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Refreshes the styling of the Sync Status column
 * This is useful if the styling gets lost or if the user wants to reset it
 */
function refreshSyncStatusStyling() {
  try {
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();

    cleanupPreviousSyncStatusColumn(sheet, sheetName);

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    if (!twoWaySyncEnabled) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Two-way sync is not enabled for this sheet. Please enable it first in "Two-Way Sync Settings".',
        'Cannot Refresh Styling',
        5
      );
      return;
    }

    // Get the tracking column
    let trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    let trackingColumnIndex;

    if (trackingColumn) {
      // Convert column letter to index (0-based)
      trackingColumnIndex = columnLetterToIndex(trackingColumn);
    } else {
      // Try to find the Sync Status column by header name
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      trackingColumnIndex = headers.findIndex(header =>
        header && header.toString().toLowerCase().includes('sync status')
      );

      if (trackingColumnIndex === -1) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Could not find the Sync Status column. Please set up two-way sync again.',
          'Column Not Found',
          5
        );
        return;
      }
    }

    // Convert to 1-based index for getRange
    const columnPos = trackingColumnIndex + 1;

    // Style the header cell
    const headerCell = sheet.getRange(1, columnPos);
    headerCell.setBackground('#E8F0FE') // Light blue background
      .setFontWeight('bold')
      .setNote('This column tracks changes for two-way sync with Pipedrive');

    // Add only BORDERS to the entire status column (not background)
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const fullStatusColumn = sheet.getRange(1, columnPos, lastRow, 1);
    fullStatusColumn.setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

    // Add data validation for all cells in the status column (except header)
    if (lastRow > 1) {
      // Get all values from the first column to identify timestamps/separators
      const firstColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

      // Collect row indices of actual data rows (not timestamps/separators)
      const dataRowIndices = [];
      for (let i = 0; i < firstColumnValues.length; i++) {
        const value = firstColumnValues[i][0];
        const rowIndex = i + 2; // +2 because we start at row 2

        // Skip empty rows and rows that look like timestamps
        if (!value || (typeof value === 'string' &&
          (value.includes('Timestamp') || value.includes('Last synced')))) {
          continue;
        }

        dataRowIndices.push(rowIndex);
      }

      // Apply background color only to data rows
      dataRowIndices.forEach(rowIndex => {
        sheet.getRange(rowIndex, columnPos).setBackground('#F8F9FA'); // Light gray background
      });

      // Apply data validation only to actual data rows
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
        .build();

      // Apply validation to each data row individually
      dataRowIndices.forEach(rowIndex => {
        sheet.getRange(rowIndex, columnPos).setDataValidation(rule);
      });

      // Clear and recreate conditional formatting
      // Get all existing conditional formatting rules
      const rules = sheet.getConditionalFormatRules();
      // Clear any existing rules for the tracking column
      const newRules = rules.filter(rule => {
        const ranges = rule.getRanges();
        return !ranges.some(range => range.getColumn() === columnPos);
      });

      // Create conditional format for "Modified" status
      const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(modifiedRule);

      // Create conditional format for "Synced" status
      const syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA')  // Light green background
        .setFontColor('#137333')   // Green text
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(syncedRule);

      // Create conditional format for "Error" status
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Error')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setBold(true)             // Bold text for errors
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(errorRule);

      // Apply all rules
      sheet.setConditionalFormatRules(newRules);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Sync Status column styling has been refreshed successfully.',
      'Styling Updated',
      5
    );
  } catch (error) {
    Logger.log(`Error refreshing sync status styling: ${error.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error refreshing styling: ${error.message}`,
      'Error',
      5
    );
  }
}

/**
 * Cleans up previous Sync Status column formatting
 * @param {Sheet} sheet - The sheet containing the column
 * @param {string} sheetName - The name of the sheet
 */
function cleanupPreviousSyncStatusColumn(sheet, sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const previousColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const currentColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const previousColumn = scriptProperties.getProperty(previousColumnKey) || '';

    // Clean up the specifically tracked previous column
    if (previousColumn && previousColumn !== currentColumn) {
      cleanupColumnFormatting(sheet, previousColumn);
      scriptProperties.deleteProperty(previousColumnKey);
    }

    // NEW: Store current column positions for future comparison
    // This helps when columns are deleted and the position shifts
    const currentColumnIndex = currentColumn ? columnLetterToIndex(currentColumn) : -1;
    if (currentColumnIndex >= 0) {
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, currentColumnIndex.toString());
    }

    // IMPORTANT: Scan ALL columns for "Sync Status" headers and validation patterns
    scanAndCleanupAllSyncColumns(sheet, currentColumn);
  } catch (error) {
    Logger.log(`Error in cleanupPreviousSyncStatusColumn: ${error.message}`);
  }
}