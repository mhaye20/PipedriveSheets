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
  try {
    // Get the entity type and filter ID from document properties
    const docProps = PropertiesService.getDocumentProperties();
    const entityType = docProps.getProperty('PIPEDRIVE_ENTITY_TYPE') || ENTITY_TYPES.DEALS;
    const filterId = docProps.getProperty('PIPEDRIVE_FILTER_ID') || '';
    const sheetName = docProps.getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    if (!filterId) {
      SpreadsheetApp.getUi().alert(
        'No filter configured',
        'Please configure a Pipedrive filter in the settings before syncing data.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      showSettings();
      return false;
    }
    
    // Show sync status dialog
    showSyncStatus(sheetName);
    
    // Determine which entity type to sync
    syncPipedriveDataToSheet(entityType);
    
    return true;
  } catch (e) {
    Logger.log(`Error in syncFromPipedrive: ${e.message}`);
    updateSyncStatus('error', 'Failed', e.message, 100);
    
    // Show error to user
    SpreadsheetApp.getUi().alert(
      'Sync Error',
      'An error occurred during synchronization: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
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