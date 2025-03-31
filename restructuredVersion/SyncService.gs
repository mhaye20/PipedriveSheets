/**
 * Sync Service
 * 
 * This module handles the synchronization between Pipedrive and Google Sheets:
 * - Fetching data from Pipedrive and writing to sheets
 * - Tracking modifications and pushing changes back to Pipedrive
 * - Managing synchronization status and scheduling
 */

// Create SyncService namespace if it doesn't exist
var SyncService = SyncService || {};

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

  // Ensure we have OAuth authentication
  if (!refreshAccessTokenIfNeeded()) {
    ui.alert(
      'Authentication Failed',
      'Could not authenticate with Pipedrive. Please reconnect your account in Settings.',
      ui.ButtonSet.OK
    );
    return;
  }

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
    updateSyncStatus('3', 'error', error.message, 100);
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
    
    // Get the active sheet instead of using a property
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = activeSheet.getName();
    
    // Update phase 1: Connecting and initializing
    updateSyncStatus('1', 'active', 'Connecting to Pipedrive...', 50);
    
    // Get filtered data based on entity type
    let items = [];
    let fields = [];
    
    // Update phase 2: Retrieving data from Pipedrive
    updateSyncStatus('2', 'active', `Retrieving ${entityType} from Pipedrive...`, 10);
    
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
      updateSyncStatus('2', 'warning', `No ${entityType} found with the selected filter.`, 100);
      return;
    }
    
    // Update phase 2 completion
    updateSyncStatus('2', 'completed', `Retrieved ${items.length} ${entityType} from Pipedrive`, 100);
    
    // Update phase 3: Processing and writing data
    updateSyncStatus('3', 'active', `Processing ${items.length} ${entityType}...`, 30);
    
    // Get team-aware column preferences
    let selectedColumns = SyncService.getTeamAwareColumnPreferences(entityType, sheetName);
    
    // If no columns are selected, create default column objects
    if (!selectedColumns || selectedColumns.length === 0) {
      Logger.log(`No columns selected for ${entityType}, using defaults`);
      
      // Set default column keys based on entity type
      let defaultColumnKeys = ['id', 'name', 'owner_id', 'created_at', 'updated_at'];
      
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          defaultColumnKeys = ['id', 'title', 'status', 'value', 'currency', 'owner_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.PERSONS:
          defaultColumnKeys = ['id', 'name', 'email', 'phone', 'owner_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          defaultColumnKeys = ['id', 'name', 'address', 'owner_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.ACTIVITIES:
          defaultColumnKeys = ['id', 'type', 'due_date', 'duration', 'deal_id', 'person_id', 'org_id', 'note', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.LEADS:
          defaultColumnKeys = ['id', 'title', 'owner_id', 'person_id', 'organization_id', 'created_at', 'updated_at'];
          break;
        case ENTITY_TYPES.PRODUCTS:
          defaultColumnKeys = ['id', 'name', 'code', 'description', 'unit', 'tax', 'active_flag', 'created_at', 'updated_at'];
          break;
      }
      
      // Create column objects from default keys
      selectedColumns = defaultColumnKeys.map(key => ({ key: key }));
      updateSyncStatus('3', 'active', `No columns selected, using defaults`, 40);
    }
    
    // Format the data for the spreadsheet
    updateSyncStatus('3', 'active', `Formatting data for spreadsheet...`, 50);
    
    // Create header row from selected columns
    const headerRow = [];
    
    // Loop through selected columns to create header names
    selectedColumns.forEach(column => {
      // Column can be an object with key/customName or just a string
      let columnName = '';
      
      if (typeof column === 'object' && column.key) {
        // If column has a custom name defined, use that
        if (column.customName) {
          columnName = column.customName;
        } else {
          // Look for a friendly name in fields
          const matchingField = fields.find(f => f.key === column.key);
          if (matchingField) {
            columnName = formatColumnName(matchingField.name);
          } else {
            // Use the key as a fallback
            columnName = formatColumnName(column.key);
          }
        }
      } else if (typeof column === 'string') {
        // If column is just a string key, look up its name in fields
        const matchingField = fields.find(f => f.key === column);
        if (matchingField) {
          columnName = formatColumnName(matchingField.name);
        } else {
          columnName = formatColumnName(column);
        }
      }
      
      headerRow.push(columnName);
    });
    
    // Get field option mappings for dropdown/option fields
    let optionMappings = {};
    try {
      // Check if the function exists in the PipedriveAPI namespace
      if (typeof getFieldOptionMappingsForEntity === 'function') {
        optionMappings = getFieldOptionMappingsForEntity(entityType);
        Logger.log(`Retrieved option mappings for ${entityType}`);
      } else if (typeof PipedriveAPI !== 'undefined' && typeof PipedriveAPI.getFieldOptionMappingsForEntity === 'function') {
        optionMappings = PipedriveAPI.getFieldOptionMappingsForEntity(entityType);
        Logger.log(`Retrieved option mappings for ${entityType} from PipedriveAPI namespace`);
      } else {
        Logger.log('getFieldOptionMappingsForEntity function not found. Option values may not display correctly.');
      }
    } catch (e) {
      Logger.log(`Error getting field option mappings: ${e.message}`);
    }
    
    // Prepare data for writing to the sheet following the original structure
    const options = {
      entityType: entityType,
      sheetName: sheetName,
      headerRow: headerRow,    // Pass the prepared header row
      columns: selectedColumns, // Pass column objects/keys
      fields: fields,
      optionMappings: optionMappings, // Include option mappings for dropdown fields
      showTimestamp: docProps.getProperty('SHOW_TIMESTAMP') === 'true',
      enableTwoWaySync: docProps.getProperty(`TWOWAY_SYNC_ENABLED_${sheetName}`) === 'true',
      trackingColumn: docProps.getProperty(`TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`) || ''
    };
    
    // Write data to the spreadsheet
    updateSyncStatus('3', 'active', `Writing data to sheet...`, 70);
    writeDataToSheet(items, options);
    
    // Mark as completed
    updateSyncStatus('3', 'completed', `Successfully synced ${items.length} ${entityType} to sheet.`, 100);
    
    // Update the last sync time
    if (options.enableTwoWaySync) {
      const now = new Date().toISOString();
      docProps.setProperty(`TWOWAY_SYNC_LAST_SYNC_${sheetName}`, now);
    }

    // Refresh the sync status column styling - just like in the original implementation
    refreshSyncStatusStyling();
    
    // Show success toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${entityType} successfully synced from Pipedrive to "${sheetName}"! (${items.length} items total)`
    );
  } catch (e) {
    Logger.log(`Error in syncPipedriveDataToSheet: ${e.message}`);
    updateSyncStatus('3', 'error', e.message, 100);
    throw e;  // Re-throw to allow caller to handle
  }
}

/**
 * Writes data to the sheet with the specified entities and options
 * @param {Array} items - Array of Pipedrive entities to write
 * @param {Object} options - Options for writing data
 */
function writeDataToSheet(items, options) {
  try {
    // Extract options with defaults for safety
    const sheetName = options.sheetName || SpreadsheetApp.getActiveSheet().getName();
    const entityType = options.entityType || ENTITY_TYPES.DEALS;
    const columns = options.columns || [];
    const fields = options.fields || [];
    const isWriteTimestamp = options.showTimestamp === true;
    const isTwoWaySync = options.enableTwoWaySync === true;
    const trackingColumn = options.trackingColumn || '';
    const optionMappings = options.optionMappings || {};
    
    // Log the values we're working with for debugging
    Logger.log(`Writing data to sheet: ${sheetName}, entityType: ${entityType}`);
    Logger.log(`Selected columns: ${columns.length}, isWriteTimestamp: ${isWriteTimestamp}, isTwoWaySync: ${isTwoWaySync}`);
    
    // Get the active spreadsheet and destination sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Clear existing content
    sheet.clear();
    
    // Use the provided headerRow if available, otherwise generate from columns
    let headerRow = options.headerRow || [];
    
    // If no headerRow provided, generate from columns
    if (!headerRow || headerRow.length === 0) {
      // Extract column keys if selectedColumns contains objects
      const columnsToPull = columns.map(col => {
        return typeof col === 'object' && col.key ? col.key : col;
      });
      
      headerRow = [];
      
      // Add selected columns to header
      for (const columnKey of columnsToPull) {
        // Get user-friendly column name
        let columnName = columnKey; // Default to key if name not found
        
        // First look in columns for a matching object with name
        const columnObj = columns.find(col => {
          if (typeof col === 'object' && col.key === columnKey) return true;
          return col === columnKey;
        });
        
        if (columnObj && typeof columnObj === 'object' && columnObj.name) {
          columnName = formatColumnName(columnObj.name);
        } else if (fields && fields.length > 0) {
          // Fall back to searching fields
          for (const field of fields) {
            if (field && field.key === columnKey) {
              columnName = formatColumnName(field.name);
              break;
            }
          }
        }
        
        headerRow.push(columnName);
      }
    }
    
    // Make a copy of the headerRow for use in the full header
    const fullHeaderRow = [...headerRow];
    
    // Add timestamp column if enabled
    if (isWriteTimestamp) {
      fullHeaderRow.push('Last Sync Time');
    }
    
    // Add sync status column if two-way sync is enabled
    if (isTwoWaySync) {
      fullHeaderRow.push('Sync Status');
    }
    
    // Write header row
    sheet.getRange(1, 1, 1, fullHeaderRow.length).setValues([fullHeaderRow])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Prepare data rows
    const dataRows = [];
    
    // Handle empty items array
    if (!items || !Array.isArray(items) || items.length === 0) {
      // Just write headers, no data rows
      Logger.log('No items to write to sheet');
      return true;
    }
    
    for (const item of items) {
      if (!item) continue; // Skip null/undefined items
      
      // Create a row array with the right number of columns
      const row = new Array(fullHeaderRow.length).fill('');
      
      // If using original structure with headerRow from options
      if (options.headerRow && options.headerRow.length > 0 && columns && columns.length > 0) {
        // Process each column according to the original structure
        columns.forEach((column, index) => {
          try {
            // Get the column key (extract from object if needed)
            const columnKey = typeof column === 'object' && column.key ? column.key : column;
            
            // Get the value using path notation
            let value = getValueByPath(item, columnKey);
            
            // Format the value
            if (typeof formatValue === 'function') {
              value = formatValue(value, columnKey, optionMappings);
            } else {
              // Basic fallback
              value = value !== null && value !== undefined ? value.toString() : '';
            }
            
            row[index] = value;
          } catch (error) {
            Logger.log(`Error getting value for column ${JSON.stringify(column)}: ${error.message}`);
            row[index] = ''; // Push empty string for error cases
          }
        });
      } else {
        // Extract column keys if using the generated structure
        const columnsToPull = columns.map(col => {
          return typeof col === 'object' && col.key ? col.key : col;
        });
        
        // Fill data for each column
        columnsToPull.forEach((columnKey, index) => {
          try {
            let value = getValueByPath(item, columnKey);
            
            // Format the value
            if (typeof formatValue === 'function') {
              value = formatValue(value, columnKey, optionMappings);
            } else {
              // Basic fallback
              value = value !== null && value !== undefined ? value.toString() : '';
            }
            
            row[index] = value;
          } catch (error) {
            Logger.log(`Error getting value for column ${columnKey}: ${error.message}`);
            row[index] = ''; // Push empty string for error cases
          }
        });
      }
      
      // Add timestamp if enabled
      if (isWriteTimestamp) {
        row[headerRow.length] = new Date().toLocaleString();
      }
      
      // Add sync status if two-way sync is enabled
      if (isTwoWaySync) {
        row[fullHeaderRow.length - 1] = 'Synced';
      }
      
      dataRows.push(row);
    }
    
    // Check and log lengths to diagnose mismatch
    if (dataRows.length > 0) {
      Logger.log(`Header row length: ${fullHeaderRow.length}, Data row length: ${dataRows[0].length}`);
      
      // Ensure all data rows have the same length as the header row
      for (let i = 0; i < dataRows.length; i++) {
        // Adjust row length if it doesn't match headers
        if (dataRows[i].length < fullHeaderRow.length) {
          // Add empty cells to match header length
          while (dataRows[i].length < fullHeaderRow.length) {
            dataRows[i].push('');
          }
        } else if (dataRows[i].length > fullHeaderRow.length) {
          // Trim excess cells
          dataRows[i] = dataRows[i].slice(0, fullHeaderRow.length);
        }
      }
    }
    
    // Write data rows
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, fullHeaderRow.length).setValues(dataRows);
    }
    
    // Apply formatting
    sheet.autoResizeColumns(1, fullHeaderRow.length);
    
    // Apply conditional formatting to sync status column if two-way sync is enabled
    if (isTwoWaySync) {
      const syncStatusColumnIndex = fullHeaderRow.length - 1;
      const syncStatusColumnLetter = columnToLetter(syncStatusColumnIndex);
      
      // Save the tracking column to properties
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty(`TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`, syncStatusColumnLetter);
      
      // We'll let refreshSyncStatusStyling handle the formatting later
      // This ensures we use the same formatting as the original script
      
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
 * Updates the sync status properties
 * @param {string} phase - The phase number ('1', '2', or '3')
 * @param {string} status - Status of the phase ('active', 'completed', 'error', etc.)
 * @param {string} detail - Detailed message about the phase
 * @param {number} progress - Progress percentage (0-100)
 */
function updateSyncStatus(phase, status, detail, progress) {
  try {
    // Store both in our internal format and the format expected by the original implementation
    const scriptProperties = PropertiesService.getScriptProperties();
    const userProps = PropertiesService.getUserProperties();
    
    // Get/create sync status object for our internal format
    const syncStatus = {
      phase: phase,
      status: status,
      detail: detail,
      progress: progress,
      lastUpdate: new Date().getTime()
    };
    
    // Ensure progress is 100% for completed phases
    if (status === 'completed') {
      progress = 100;
      syncStatus.progress = 100;
    }
    
    // Save to user properties in our format
    userProps.setProperty('SYNC_STATUS', JSON.stringify(syncStatus));
    
    // Save in the original format for compatibility with the HTML
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_STATUS`, status);
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_DETAIL`, detail || '');
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_PROGRESS`, progress ? progress.toString() : '0');
    
    // Set current phase
    scriptProperties.setProperty('SYNC_CURRENT_PHASE', phase.toString());
    
    // If status is error, store the error
    if (status === 'error') {
      scriptProperties.setProperty('SYNC_ERROR', detail || 'An error occurred');
      scriptProperties.setProperty('SYNC_COMPLETED', 'true');
      syncStatus.error = detail || 'An error occurred';
    }
    
    // If this is the final phase completion, mark as completed
    if (status === 'completed' && phase === '3') {
      scriptProperties.setProperty('SYNC_COMPLETED', 'true');
    }
    
    // Also show a toast message for visibility
    let toastMessage = '';
    if (phase === '1') toastMessage = 'Connecting to Pipedrive...';
    else if (phase === '2') toastMessage = 'Retrieving data from Pipedrive...';
    else if (phase === '3') toastMessage = 'Writing data to spreadsheet...';
    
    if (status === 'error') {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${detail}`, 'Sync Error', 5);
    } else if (status === 'completed' && phase === '3') {
      SpreadsheetApp.getActiveSpreadsheet().toast('Sync completed successfully!', 'Sync Status', 3);
    } else if (detail) {
      SpreadsheetApp.getActiveSpreadsheet().toast(detail, toastMessage, 2);
    }
    
    return syncStatus;
  } catch (e) {
    Logger.log(`Error updating sync status: ${e.message}`);
    // Still show a toast message as backup
    SpreadsheetApp.getActiveSpreadsheet().toast(detail || 'Processing...', 'Sync Status', 2);
    return null;
  }
}

/**
 * Gets the current sync status for the dialog to poll
 * @return {Object} Sync status object or null if not available
 */
function getSyncStatus() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const userProps = PropertiesService.getUserProperties();
    const statusJson = userProps.getProperty('SYNC_STATUS');
    
    if (!statusJson) {
      // Return default format matching the expected structure in the HTML
      return {
        phase1: {
          status: scriptProperties.getProperty('SYNC_PHASE_1_STATUS') || 'active',
          detail: scriptProperties.getProperty('SYNC_PHASE_1_DETAIL') || 'Connecting to Pipedrive...',
          progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_1_PROGRESS') || '0')
        },
        phase2: {
          status: scriptProperties.getProperty('SYNC_PHASE_2_STATUS') || 'pending',
          detail: scriptProperties.getProperty('SYNC_PHASE_2_DETAIL') || 'Waiting to start...',
          progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_2_PROGRESS') || '0')
        },
        phase3: {
          status: scriptProperties.getProperty('SYNC_PHASE_3_STATUS') || 'pending',
          detail: scriptProperties.getProperty('SYNC_PHASE_3_DETAIL') || 'Waiting to start...',
          progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_3_PROGRESS') || '0')
        },
        currentPhase: scriptProperties.getProperty('SYNC_CURRENT_PHASE') || '1',
        completed: scriptProperties.getProperty('SYNC_COMPLETED') || 'false',
        error: scriptProperties.getProperty('SYNC_ERROR') || ''
      };
    }
    
    // Convert from our internal format to the format expected by the HTML
    const status = JSON.parse(statusJson);
    
    // Identify which phase is active based on the phase field
    const activePhase = status.phase || '1';
    const statusValue = status.status || 'active';
    const detailValue = status.detail || '';
    const progressValue = status.progress || 0;
    
    // Create the response in the format expected by the HTML
    const response = {
      phase1: {
        status: activePhase === '1' ? statusValue : (activePhase > '1' ? 'completed' : 'pending'),
        detail: activePhase === '1' ? detailValue : (activePhase > '1' ? 'Completed' : 'Waiting to start...'),
        progress: activePhase === '1' ? progressValue : (activePhase > '1' ? 100 : 0)
      },
      phase2: {
        status: activePhase === '2' ? statusValue : (activePhase > '2' ? 'completed' : 'pending'),
        detail: activePhase === '2' ? detailValue : (activePhase > '2' ? 'Completed' : 'Waiting to start...'),
        progress: activePhase === '2' ? progressValue : (activePhase > '2' ? 100 : 0)
      },
      phase3: {
        status: activePhase === '3' ? statusValue : (activePhase > '3' ? 'completed' : 'pending'),
        detail: activePhase === '3' ? detailValue : (activePhase > '3' ? 'Completed' : 'Waiting to start...'),
        progress: activePhase === '3' ? progressValue : (activePhase > '3' ? 100 : 0)
      },
      currentPhase: activePhase,
      completed: activePhase === '3' && status.status === 'completed' ? 'true' : 'false',
      error: status.error || ''
    };
    
    return response;
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

    // Ensure we have OAuth authentication
    if (!refreshAccessTokenIfNeeded()) {
      // Show an error message if authentication fails
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Authentication Failed',
          'Could not authenticate with Pipedrive. Please reconnect your account in Settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

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
        // Ensure we have a valid token
        if (!refreshAccessTokenIfNeeded()) {
          throw new Error('Not authenticated with Pipedrive. Please connect your account first.');
        }
        
        const scriptProperties = PropertiesService.getScriptProperties();
        const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
        
        // Set up the request URL based on entity type
        let updateUrl = `${apiUrl}/${entityType}/${rowData.id}`;
        let method = 'PUT';
        
        // Special case for activities which use a different endpoint for updates
        if (entityType === ENTITY_TYPES.ACTIVITIES) {
          updateUrl = `${apiUrl}/${entityType}/${rowData.id}`;
        }

        // Make the API call with proper OAuth authentication
        const options = {
          method: method,
          contentType: 'application/json',
          headers: {
            'Authorization': 'Bearer ' + accessToken
          },
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

/**
 * Saves settings to script properties
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Save global settings (only if provided)
  if (apiKey) scriptProperties.setProperty('PIPEDRIVE_API_KEY', apiKey);
  if (subdomain) scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
  
  // Save sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  
  scriptProperties.setProperty(sheetFilterIdKey, filterId);
  scriptProperties.setProperty(sheetEntityTypeKey, entityType);
  scriptProperties.setProperty('SHEET_NAME', sheetName);
}

/**
 * Saves column preferences to script properties
 * @param {Array} columns - Array of column objects with key and name properties
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @param {string} userEmail - Email of the user saving the preferences
 */
SyncService.saveColumnPreferences = function(columns, entityType, sheetName, userEmail) {
  try {
    Logger.log(`SyncService.saveColumnPreferences for ${entityType} in sheet "${sheetName}" for user ${userEmail}`);
    
    // Store the full column objects to preserve names and other properties
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Store columns based on both entity type and sheet name for sheet-specific preferences
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    scriptProperties.setProperty(columnSettingsKey, JSON.stringify(columns));
    
    // Also store in user-specific property if email is provided
    if (userEmail) {
      const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
      scriptProperties.setProperty(userColumnSettingsKey, JSON.stringify(columns));
      Logger.log(`Saved user-specific column preferences with key: ${userColumnSettingsKey}`);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in SyncService.saveColumnPreferences: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    throw e;
  }
}

/**
 * Logs debug information about the Pipedrive data
 */
function logDebugInfo() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetName = scriptProperties.getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;
  
  // Get sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  
  const filterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Show which column selections are available for the current entity type and sheet
  const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
  const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
  
  if (savedColumnsJson) {
    Logger.log(`\n===== COLUMN SETTINGS FOR ${sheetName} - ${entityType} =====`);
    try {
      const selectedColumns = JSON.parse(savedColumnsJson);
      Logger.log(`Number of selected columns: ${selectedColumns.length}`);
      Logger.log(JSON.stringify(selectedColumns, null, 2));
    } catch (e) {
      Logger.log(`Error parsing column settings: ${e.message}`);
    }
  } else {
    Logger.log(`\n===== NO COLUMN SETTINGS FOUND FOR ${sheetName} - ${entityType} =====`);
  }
  
  // Get a sample item to see what data is available
  let sampleData = [];
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
  }
  
  if (sampleData && sampleData.length > 0) {
    const sampleItem = sampleData[0];
    
    // Log filter ID and entity type
    Logger.log('===== DEBUG INFORMATION =====');
    Logger.log(`Entity Type: ${entityType}`);
    Logger.log(`Filter ID: ${filterId}`);
    Logger.log(`Sheet Name: ${sheetName}`);
    
    // Log complete raw deal data for inspection
    Logger.log(`\n===== COMPLETE RAW ${entityType.toUpperCase()} DATA =====`);
    Logger.log(JSON.stringify(sampleItem, null, 2));
    
    // Extract all fields including nested ones
    Logger.log('\n===== ALL AVAILABLE FIELDS =====');
    const allFields = {};
    
    // Recursive function to extract all fields with their paths
    function extractAllFields(obj, path = '') {
      if (!obj || typeof obj !== 'object') return;
      
      if (Array.isArray(obj)) {
        // For arrays, log the length and extract fields from first item if exists
        Logger.log(`${path} (Array with ${obj.length} items)`);
        if (obj.length > 0 && typeof obj[0] === 'object') {
          extractAllFields(obj[0], `${path}[0]`);
        }
      } else {
        // For objects, extract each property
        for (const key in obj) {
          const value = obj[key];
          const newPath = path ? `${path}.${key}` : key;
          
          if (value === null) {
            allFields[newPath] = 'null';
            continue;
          }
          
          const type = typeof value;
          
          if (type === 'object') {
            if (Array.isArray(value)) {
              allFields[newPath] = `array[${value.length}]`;
              Logger.log(`${newPath}: array[${value.length}]`);
              
              // Special case for custom fields with options
              if (key === 'options' && value.length > 0 && value[0] && value[0].label) {
                Logger.log(`  - Multiple options field with values: ${value.map(opt => opt.label).join(', ')}`);
              }
              
              // For small arrays with objects, recursively extract from the first item
              if (value.length > 0 && typeof value[0] === 'object') {
                extractAllFields(value[0], `${newPath}[0]`);
              }
            } else {
              allFields[newPath] = 'object';
              Logger.log(`${newPath}: object`);
              extractAllFields(value, newPath);
            }
          } else {
            allFields[newPath] = type;
            
            // Log a preview of the value unless it's a string longer than 50 chars
            const preview = type === 'string' && value.length > 50 
              ? value.substring(0, 50) + '...' 
              : value;
              
            Logger.log(`${newPath}: ${type} = ${preview}`);
          }
        }
      }
    }
    
    // Start extraction from the top level
    extractAllFields(sampleItem);
    
    // Specifically focus on custom fields section if it exists
    if (sampleItem.custom_fields) {
      Logger.log('\n===== CUSTOM FIELDS DETAIL =====');
      for (const key in sampleItem.custom_fields) {
        const field = sampleItem.custom_fields[key];
        const fieldType = typeof field;
        
        if (fieldType === 'object' && Array.isArray(field)) {
          Logger.log(`${key}: array[${field.length}]`);
          // Check if this is a multiple options field
          if (field.length > 0 && field[0] && field[0].label) {
            Logger.log(`  - Multiple options with values: ${field.map(opt => opt.label).join(', ')}`);
          }
        } else {
          const preview = fieldType === 'string' && field.length > 50 
            ? field.substring(0, 50) + '...' 
            : field;
          Logger.log(`${key}: ${fieldType} = ${preview}`);
        }
      }
    }
    
    // Count unique fields
    const fieldPaths = Object.keys(allFields).sort();
    Logger.log(`\nTotal unique fields found: ${fieldPaths.length}`);
    
    // Log all field paths in alphabetical order for easy lookup
    Logger.log('\n===== ALPHABETICAL LIST OF ALL FIELD PATHS =====');
    fieldPaths.forEach(path => {
      Logger.log(`${path}: ${allFields[path]}`);
    });
    
  } else {
    Logger.log(`No ${entityType} found with this filter. Please check the filter ID.`);
  }
}

/**
 * Gets team-aware column preferences - wrapper for UI.gs function
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of column keys
 */
SyncService.getTeamAwareColumnPreferences = function(entityType, sheetName) {
  try {
    // Call the function in UI.gs that handles retrieving from both storage locations
    return UI.getTeamAwareColumnPreferences(entityType, sheetName);
  } catch (e) {
    Logger.log(`Error in SyncService.getTeamAwareColumnPreferences: ${e.message}`);
    
    // Fallback to local implementation if UI.getTeamAwareColumnPreferences fails
    const scriptProperties = PropertiesService.getScriptProperties();
    const oldFormatKey = `COLUMNS_${sheetName}_${entityType}`;
    const oldFormatJson = scriptProperties.getProperty(oldFormatKey);
    
    if (oldFormatJson) {
      try {
        return JSON.parse(oldFormatJson);
      } catch (parseError) {
        Logger.log(`Error parsing column preferences: ${parseError.message}`);
      }
    }
    
    return [];
  }
}

/**
 * Saves team-aware column preferences - wrapper for UI.gs function
 * @param {Array} columns - Column objects or keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
SyncService.saveTeamAwareColumnPreferences = function(columns, entityType, sheetName) {
  try {
    // Keep full column objects intact to preserve names
    // Call the function in UI.gs that handles saving to both storage locations
    return UI.saveTeamAwareColumnPreferences(columns, entityType, sheetName);
  } catch (e) {
    Logger.log(`Error in SyncService.saveTeamAwareColumnPreferences: ${e.message}`);
    
    // Fallback to local implementation if UI.saveTeamAwareColumnPreferences fails
    const scriptProperties = PropertiesService.getScriptProperties();
    const key = `COLUMNS_${sheetName}_${entityType}`;
    
    // Store the full column objects
    scriptProperties.setProperty(key, JSON.stringify(columns));
  }
}

/**
 * Gets field mappings for a specific entity type
 * @param {string} entityType - The entity type to get mappings for
 * @return {Object} An object mapping column headers to API field keys
 */
function getFieldMappingsForEntity(entityType) {
  // Basic field mappings for each entity type
  const commonMappings = {
    'ID': 'id',
    'Name': 'name',
    'Owner': 'owner_id',
    'Organization': 'org_id',
    'Person': 'person_id',
    'Added': 'add_time',
    'Updated': 'update_time'
  };

  // Entity-specific mappings
  const entityMappings = {
    [ENTITY_TYPES.DEALS]: {
      'Value': 'value',
      'Currency': 'currency',
      'Title': 'title',
      'Pipeline': 'pipeline_id',
      'Stage': 'stage_id',
      'Status': 'status',
      'Expected Close Date': 'expected_close_date'
    },
    [ENTITY_TYPES.PERSONS]: {
      'Email': 'email',
      'Phone': 'phone',
      'First Name': 'first_name',
      'Last Name': 'last_name',
      'Organization': 'org_id'
    },
    [ENTITY_TYPES.ORGANIZATIONS]: {
      'Address': 'address',
      'Website': 'web'
    },
    [ENTITY_TYPES.ACTIVITIES]: {
      'Type': 'type',
      'Due Date': 'due_date',
      'Due Time': 'due_time',
      'Duration': 'duration',
      'Deal': 'deal_id',
      'Person': 'person_id',
      'Organization': 'org_id',
      'Note': 'note'
    },
    [ENTITY_TYPES.PRODUCTS]: {
      'Code': 'code',
      'Description': 'description',
      'Unit': 'unit',
      'Tax': 'tax',
      'Category': 'category',
      'Active': 'active_flag',
      'Selectable': 'selectable',
      'Visible To': 'visible_to',
      'First Price': 'first_price',
      'Cost': 'cost',
      'Prices': 'prices',
      'Owner Name': 'owner_id.name'  // Map "Owner Name" to owner_id.name so we can detect this field
    }
  };

  // Combine common mappings with entity-specific mappings
  return { ...commonMappings, ...(entityMappings[entityType] || {}) };
}