/**
 * Sync Service Code
 *
 * This module handles the synchronization between Pipedrive and Google Sheets™:
 * - Fetching data from Pipedrive and writing to sheets
 * - Tracking modifications and pushing changes back to Pipedrive
 * - Managing synSchronization status and scheduling
 */

// Create SyncService namespace if it doesn't exist
var SyncService = SyncService || {};

/**
 * Process and format field values based on their field type definitions
 * @param {Object} fields - Object containing field key-value pairs
 * @param {Object} fieldDefinitions - Field definitions from Pipedrive
 * @param {string} entityType - The entity type (for specific handling)
 * @return {Object} Processed fields with proper formatting
 */
function processFieldTypes(fields, fieldDefinitions, entityType) {
  const processedFields = {};
  
  Object.keys(fields).forEach((key) => {
    let value = fields[key];
    const fieldDef = fieldDefinitions[key];
    
    // Skip if value is null or undefined
    if (value === null || value === undefined) {
      processedFields[key] = value;
      return;
    }
    
    // Check field type and apply appropriate formatting
    if (fieldDef) {
      // Date fields - convert to YYYY-MM-DD format
      if (fieldDef.field_type === 'date' && value) {
        if (typeof value === 'string' && value.includes('T')) {
          value = value.split('T')[0];
        }
      }
      
      // Phone fields - ensure they are strings
      else if (fieldDef.field_type === 'phone') {
        value = String(value);
      }
      
      // Time fields - ensure proper HH:MM:SS format
      else if (fieldDef.field_type === 'time') {
        if (typeof value === 'string' && value.includes('T')) {
          // Extract time from datetime
          const timePart = value.split('T')[1];
          if (timePart) {
            value = timePart.split('.')[0]; // Remove milliseconds
          }
        }
      }
      
      // Monetary fields - ensure they are numbers
      else if (fieldDef.field_type === 'monetary' && typeof value === 'string') {
        value = parseFloat(value) || 0;
      }
    }
    
    // Handle boolean conversions for specific fields
    if ((key === 'is_archived' || key === 'was_seen' || key === 'active_flag' || key === 'done') && 
        typeof value === 'string') {
      value = value.toLowerCase() === 'yes' || value.toLowerCase() === 'true';
    }
    
    // Handle visibility field conversion
    if (key === 'visible_to' && typeof value === 'string') {
      const visibilityMap = {
        'owner only': '1',
        'owner & followers': '1',
        "owner's visibility group": '3',
        "owner's visibility group and sub-groups": '5',
        'entire company': '7',
        'all users': '3' // For Essential/Advanced plans
      };
      const mappedValue = visibilityMap[value.toLowerCase()];
      if (mappedValue) {
        value = mappedValue;
      }
    }
    
    processedFields[key] = value;
  });
  
  return processedFields;
}

/**
 * Checks if a sync operation is currently running
 * @return {boolean} True if a sync is running, false otherwise
 */
function isSyncRunning() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty("SYNC_RUNNING") === "true";
}

/**
 * Sets the sync running status
 * @param {boolean} isRunning - Whether the sync is running
 */
function setSyncRunning(isRunning) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("SYNC_RUNNING", isRunning ? "true" : "false");
}

/**
 * Main synchronization function that syncs data from Pipedrive to the sheet
 */
function syncFromPipedrive() {
  try {
    
    // Check subscription status before syncing
    const plan = PaymentService.getCurrentPlan();
    if (plan.plan === 'free') {
      // For free plan, check if this is a scheduled sync
      const triggers = ScriptApp.getProjectTriggers();
      const isScheduled = triggers.some(trigger => 
        trigger.getHandlerFunction() === 'syncFromPipedrive' ||
        trigger.getHandlerFunction() === 'scheduledSync'
      );
      
      if (isScheduled) {
        SpreadsheetApp.getUi().alert(
          'Scheduled sync is not available on the Free plan. Please upgrade to Pro or Team to enable automatic syncing.'
        );
        return;
      }
    }
    
    // Get active sheet info
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();

    // IMPORTANT: Detect any column shifts that may have occurred since last sync
    detectColumnShifts();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();

    // Check for two-way sync
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled =
      scriptProperties.getProperty(twoWaySyncEnabledKey) === "true";
      
    // Check if user has access to two-way sync
    if (twoWaySyncEnabled && !PaymentService.hasFeatureAccess('two_way_sync')) {
      SpreadsheetApp.getUi().alert(
        'Two-way sync is only available on Pro and Team plans. Please upgrade to enable this feature.'
      );
      return;
    }

    // Show confirmation dialog
    const ui = SpreadsheetApp.getUi();
    let confirmMessage = `This will sync data from Pipedrive to the current sheet "${sheetName}". Any existing data in this sheet will be replaced.`;

    if (twoWaySyncEnabled) {
      confirmMessage += `\n\nTwo-way sync is enabled for this sheet. Modified rows will be pushed to Pipedrive before pulling new data.`;
    }

    const response = ui.alert(
      "Confirm Sync",
      confirmMessage,
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) {
      return;
    }

    // Prevent multiple syncs running at once
    if (isSyncRunning()) {
      ui.alert(
        "A sync operation is already running. Please wait for it to complete."
      );
      return;
    }

    // Get configuration
    const entityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const filterIdKey = `FILTER_ID_${sheetName}`;

    const entityType = scriptProperties.getProperty(entityTypeKey);
    const filterId = scriptProperties.getProperty(filterIdKey);


    // Check for required settings
    if (!entityType) {
      SpreadsheetApp.getUi().alert(
        "No Pipedrive entity type set for this sheet. Please configure your filter settings first."
      );
      return;
    }

    // Show sync status dialog
    showSyncStatus(sheetName);

    // Mark sync as running
    setSyncRunning(true);

    // Start the sync process
    updateSyncStatus("1", "active", "Connecting to Pipedrive...", 50);

    // Perform sync with skip push parameter as false
    syncPipedriveDataToSheet(entityType, false, sheetName, filterId);

    // Show completion message
    SpreadsheetApp.getUi().alert("Sync completed successfully!");
  } catch (error) {

    // Update sync status
    updateSyncStatus("3", "error", `Error: ${error.message}`, 0);

    // Show error message
    SpreadsheetApp.getUi().alert(`Error syncing data: ${error.message}`);
  } finally {
    // Always release sync lock
    setSyncRunning(false);
  }
}

/**
 * Synchronizes Pipedrive data to the sheet based on entity type
 * @param {string} entityType - The type of entity to sync
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 * @param {string} filterId - The filter ID to use for retrieving data
 */
function syncPipedriveDataToSheet(
  entityType,
  skipPush = false,
  sheetName = null,
  filterId = null
) {
  try {

    // Get sheet name if not provided
    sheetName = sheetName || SpreadsheetApp.getActiveSheet().getName();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();

    // If no filter ID provided, try to get from script properties
    if (!filterId) {
      const filterIdKey = `FILTER_ID_${sheetName}`;
      filterId = scriptProperties.getProperty(filterIdKey);
    }

    // Show UI that we are retrieving data
    updateSyncStatus("2", "active", "Retrieving data from Pipedrive...", 10);

    // Check for two-way sync settings
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled =
      scriptProperties.getProperty(twoWaySyncEnabledKey) === "true";

    // Key for tracking column
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    // If two-way sync is enabled and we're not skipping push, automatically push changes
    if (!skipPush && twoWaySyncEnabled) {
      // Push changes to Pipedrive first without showing additional dialogs
      pushChangesToPipedrive(true, true); // true for scheduled sync, true for suppress warning
    }

    // Get data from Pipedrive based on entity type
    let items = [];

    // Update status to show we're connecting to API
    updateSyncStatus(
      "2",
      "active",
      `Retrieving ${entityType} from Pipedrive...`,
      20
    );


    // Use appropriate function based on entity type
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        items = getDealsWithFilter(filterId);
        break;
      case ENTITY_TYPES.PERSONS:
        items = getPersonsWithFilter(filterId);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        items = getOrganizationsWithFilter(filterId);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        items = getActivitiesWithFilter(filterId);
        break;
      case ENTITY_TYPES.LEADS:
        items = getLeadsWithFilter(filterId);
        break;
      case ENTITY_TYPES.PRODUCTS:
        items = getProductsWithFilter(filterId);
        break;
      default:
        throw new Error(`Unknown entity type: ${entityType}`);
    }

    
    // Check row limits based on subscription plan
    try {
      PaymentService.enforceRowLimit(items.length);
    } catch (error) {
      updateSyncStatus("3", "error", error.message, 0);
      SpreadsheetApp.getUi().alert(error.message);
      PaymentService.showUpgradeDialog();
      return;
    }

    // Log first item structure for debugging
    if (items.length > 0) {

      // Specifically log email and phone fields if they exist
      if (items[0].email) {
      }

      if (items[0].phone) {
      }

      // Log address fields for organizations
      if (entityType === ENTITY_TYPES.ORGANIZATIONS && items[0].address) {
      }
    }

    // Build comprehensive entity mappings for all linked fields
    const entityMappings = buildEntityMappings(items, entityType);
    
    
    // Backwards compatibility - keep personIdToNameMap for existing code
    const personIdToNameMap = entityMappings.personIdToName;
    
    // Special handling for address fields in organizations
    if (entityType === ENTITY_TYPES.ORGANIZATIONS) {
      for (let i = 0; i < items.length; i++) {
        const org = items[i];

        // Process address components for organizations
        if (org.address) {
          // Create explicit fields for each address component if they don't exist
          if (typeof org.address === "object") {
            // Full address components (use dot notation to extract later)
            if (org.address.street_number) {
              org["address.street_number"] = org.address.street_number;
            }
            if (org.address.route) {
              org["address.route"] = org.address.route;
            }
            if (org.address.sublocality) {
              org["address.sublocality"] = org.address.sublocality;
            }
            if (org.address.locality) {
              org["address.locality"] = org.address.locality;
            }
            if (org.address.admin_area_level_1) {
              org["address.admin_area_level_1"] =
                org.address.admin_area_level_1;
            }
            if (org.address.admin_area_level_2) {
              org["address.admin_area_level_2"] =
                org.address.admin_area_level_2;
            }
            if (org.address.country) {
              org["address.country"] = org.address.country;
            }
            if (org.address.postal_code) {
              org["address.postal_code"] = org.address.postal_code;
            }
            if (org.address.formatted_address) {
              org["address.formatted_address"] = org.address.formatted_address;
            }
            // The "apartment" or "suite" is often in the subpremise field
            if (org.address.subpremise) {
              org["address.subpremise"] = org.address.subpremise;
            }

            // Log the extracted address components
            for (const key in org) {
              if (key.startsWith("address.")) {
              }
            }
          }
        }
      }
    }

    // Check if we have any data
    if (items.length === 0) {
      throw new Error(
        `No ${entityType} found. Please check your filter settings.`
      );
    }

    // Update status to show data retrieval is complete
    updateSyncStatus(
      "2",
      "completed",
      `Retrieved ${items.length} ${entityType} from Pipedrive`,
      100
    );

    // Get field options for handling picklists/enums
    let optionMappings = {};

    try {
      optionMappings = getFieldOptionMappingsForEntity(entityType);

      // Sample logging of one option mapping if available
      const sampleField = Object.keys(optionMappings)[0];
      if (sampleField) {
      }
    } catch (e) {
    }

    // Start writing to sheet
    updateSyncStatus("3", "active", "Writing data to spreadsheet...", 10);

    // Get the active spreadsheet and sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Get column preferences
    let columns = SyncService.getTeamAwareColumnPreferences(
      entityType,
      sheetName
    );

    if (columns.length === 0) {
      // If no column preferences, use default columns

      if (DEFAULT_COLUMNS[entityType]) {
        DEFAULT_COLUMNS[entityType].forEach((key) => {
          columns.push({
            key: key,
            name: formatColumnName(key),
          });
        });
      } else {
        DEFAULT_COLUMNS.COMMON.forEach((key) => {
          columns.push({
            key: key,
            name: formatColumnName(key),
          });
        });
      }

    } else {
    }

    // Create header row from column names
    const headers = columns.map((column) => {
      if (typeof column === "object" && column.customName) {
        return column.customName;
      }

      if (typeof column === "object" && column.name) {
        return column.name;
      }

      // Default to formatted key
      return formatColumnName(column.key || column);
    });

    // DEBUG: Log the headers created directly from column preferences

    // Use the makeHeadersUnique function to ensure header uniqueness
    const uniqueHeaders = makeHeadersUnique(headers, columns);

    // Filter out any empty or undefined headers
    const finalHeaders = uniqueHeaders.filter(
      (header) => header && header.trim()
    );


    // Options for writing data
    const options = {
      sheetName: sheetName,
      columns: columns,
      headerRow: finalHeaders,
      entityType: entityType,
      optionMappings: optionMappings,
      twoWaySyncEnabled: twoWaySyncEnabled,
      personIdToNameMap: personIdToNameMap, // Backwards compatibility
      entityMappings: entityMappings, // Comprehensive entity mappings
    };

    // Store original data for undo detection when two-way sync is enabled
    if (twoWaySyncEnabled) {
      try {
        storeOriginalData(items, options);
      } catch (storageError) {
        // Continue with sync even if storage fails
      }
    }

    // Write data to the sheet
    writeDataToSheet(items, options);

    // IMPORTANT: If two-way sync is enabled, make sure we thoroughly clean up previous Sync Status columns
    if (twoWaySyncEnabled) {
      try {
        // Get the current Sync Status column position
        const currentStatusColumn = scriptProperties.getProperty(
          twoWaySyncTrackingColumnKey
        );

        // Get the sheet object
        const sheet =
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

        if (sheet && currentStatusColumn) {
          cleanupPreviousSyncStatusColumn(sheet, currentStatusColumn);

          // Also perform a complete cell-by-cell scan for any sync status validation that might have been missed
          scanAllCellsForSyncStatusValidation(sheet);
        }
      } catch (cleanupError) {
        // Continue with sync even if cleanup has issues
      }
    }

    // Note: Status is now updated in writeDataToSheet when data is actually written
    // We no longer need to update sync status here

    // Store sync timestamp
    const timestamp = new Date().toISOString();
    scriptProperties.setProperty(`LAST_SYNC_${sheetName}`, timestamp);

    // Ensure we have a valid header-to-field mapping for future pushChangesToPipedrive calls
    ensureHeaderFieldMapping(sheetName, entityType);


    // Log sync activity for team members
    try {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        // Check if user is in a team
        const userTeam = getUserTeam(userEmail);
        if (userTeam) {
          // Log the sync activity
          logTeamActivity(userTeam.teamId, 'sync_performed', userEmail, {
            entityType: entityType,
            sheetName: sheetName,
            itemCount: items.length
          });
        }
      }
    } catch (activityError) {
      // Don't fail the sync if activity logging fails
    }

    // Mark Phase 3 as completed
    updateSyncStatus(
      "3",
      "completed",
      "Data successfully written to sheet",
      100
    );

    setSyncRunning(false);

    // Check if we need to recreate triggers after column changes
    checkAndRecreateTriggers();

    return true;
  } catch (error) {

    // Update sync status
    updateSyncStatus("3", "error", `Error: ${error.message}`, 0);

    // Show error in UI
    SpreadsheetApp.getUi().alert(`Error syncing data: ${error.message}`);

    throw error;
  }
}

/**
 * Writes data to the sheet with columns matching the data
 * @param {Array} items - The items to write to the sheet
 * @param {Object} options - Options for writing data
 */
function writeDataToSheet(items, options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get sheet name from options
    const sheetName = options.sheetName;

    // Get script properties for tracking previous column positions
    const scriptProperties = PropertiesService.getScriptProperties();
    const previousSyncColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const previousColumnPosition = scriptProperties.getProperty(
      previousSyncColumnKey
    );

    // Get sheet
    let sheet = ss.getSheetByName(sheetName);

    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      // When resyncing, ALWAYS clear the entire sheet first to prevent old data persisting
      // This ensures we don't have leftover values from previous syncs
      sheet.clear();
      sheet.clearFormats();
      // Fix: Use getDataRange() to clear data validations
      if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
        sheet.getDataRange().clearDataValidations();
      }
    }

    // CRITICAL: Debug what headers we're getting from Pipedrive

    // Get headers from options - ALWAYS make a copy to avoid modifying the original
    const headers = options.headerRow ? [...options.headerRow] : [];

    // Check if two-way sync is enabled
    const twoWaySyncEnabled = options.twoWaySyncEnabled || false;

    // Key for tracking column
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    // Get current tracking column
    const currentSyncColumn = scriptProperties.getProperty(
      twoWaySyncTrackingColumnKey
    );

    // Store previous position if we have a current one
    if (currentSyncColumn) {
      scriptProperties.setProperty(previousSyncColumnKey, currentSyncColumn);
    }

    // Add Sync Status column if two-way sync is enabled
    let syncStatusColumn = -1;

    if (twoWaySyncEnabled) {
      headers.push("Sync Status");
      syncStatusColumn = headers.length;

      // Store the new position for future reference
      const colLetter = columnToLetter(syncStatusColumn);
      scriptProperties.setProperty(twoWaySyncTrackingColumnKey, colLetter);
    }

    // Ensure all headers are unique and properly descriptive
    const uniqueHeaders = makeHeadersUnique(headers, options.columns);

    // If we have data and headers, write them to the sheet
    if (items.length > 0 && uniqueHeaders.length > 0) {
      // First, write the header row (bold)
      sheet.getRange(1, 1, 1, uniqueHeaders.length).setValues([uniqueHeaders]);
      sheet.getRange(1, 1, 1, uniqueHeaders.length).setFontWeight("bold");

      // Format the headers nicely
      sheet.setFrozenRows(1);

      // Prepare the data rows
      const dataRows = [];

      // Process each item
      for (const item of items) {
        const rowData = [];

        // Get data for each column
        for (let i = 0; i < options.columns.length; i++) {
          const column = options.columns[i];
          let key, value;

          if (typeof column === "object" && column.key) {
            key = column.key;
          } else {
            key = column;
          }

          // Get the value using our helper function
          value = getValueByPath(item, key);

          // Special handling for participants field in activities
          if (key === 'participants' && options.entityType === 'activities' && Array.isArray(value)) {
            // Build list of participant names
            const participantNames = [];
            
            for (const participant of value) {
              if (participant && participant.person_id) {
                // Try to get name from mapping
                const personName = options.entityMappings && options.entityMappings.personIdToName && options.entityMappings.personIdToName[participant.person_id];
                if (personName) {
                  participantNames.push(personName);
                } else {
                  // Fallback to showing ID if name not found
                  participantNames.push(`Person ${participant.person_id}`);
                }
              }
            }
            
            // Join names with comma
            rowData.push(participantNames.join(', '));
          } else {
            // Format the value if needed - pass entity mappings
            const formattedValue = formatValue(
              value,
              key,
              options.optionMappings,
              options.entityMappings
            );
            rowData.push(formattedValue);
          }
        }

        // Add Sync Status cell for two-way sync
        if (twoWaySyncEnabled) {
          rowData.push("Synced"); // Default status
        }

        // Add the row to our data array
        dataRows.push(rowData);
      }

      // Write all data at once (much faster than row by row)
      if (dataRows.length > 0) {
        sheet
          .getRange(2, 1, dataRows.length, uniqueHeaders.length)
          .setValues(dataRows);
      }

      // Auto-resize columns for better visibility
      try {
        sheet.autoResizeColumns(1, uniqueHeaders.length);
      } catch (e) {
      }
    }
    
    // Check if timestamp is enabled for this sheet
    const timestampEnabledKey = `TIMESTAMP_ENABLED_${sheetName}`;
    const timestampEnabled = scriptProperties.getProperty(timestampEnabledKey) === 'true';
    
    // Add timestamp of last sync if enabled
    if (timestampEnabled && items.length > 0) {
      try {
        const lastRow = sheet.getLastRow();
        const timestampRow = lastRow + 2; // Add one blank row, then timestamp
        
        // Add the timestamp
        sheet.getRange(timestampRow, 1).setValue('Last synced:');
        sheet.getRange(timestampRow, 2).setValue(new Date()).setNumberFormat('yyyy-MM-dd HH:mm:ss');
        
        // Format the timestamp row for better visibility
        sheet.getRange(timestampRow, 1, 1, 2).setFontWeight('bold');
        sheet.getRange(timestampRow, 1, 1, 2).setBackground('#f1f3f4'); // Light gray background
        
      } catch (e) {
      }
    }

    // Ensure we have a valid header-to-field mapping based on the current headers
    try {
      ensureHeaderFieldMapping(sheetName, options.entityType);
    } catch (mappingError) {
    }

    // Load saved Sync Status values if needed
    if (twoWaySyncEnabled && items.length > 0) {
      try {
        // Get the saved sync status data
        const savedStatusData = getSavedSyncStatusData(sheetName);

        if (savedStatusData && Object.keys(savedStatusData).length > 0) {

          // Get the ID column (always first column)
          const idColumnIdx = 0;

          // Get the sync status column
          const statusColumnIdx = syncStatusColumn - 1;

          // Get all current data
          const dataRange = sheet.getRange(
            2,
            1,
            items.length,
            uniqueHeaders.length
          );
          const currentValues = dataRange.getValues();

          // Track which values were updated
          let updateCount = 0;
          let statusCellsToUpdate = [];

          // Loop through all rows
          for (let i = 0; i < currentValues.length; i++) {
            const row = currentValues[i];
            const id = row[idColumnIdx];

            // Check if we have a saved status for this ID
            if (savedStatusData[id]) {
              row[statusColumnIdx] = savedStatusData[id];
              updateCount++;

              // Add to the batch update
              statusCellsToUpdate.push([savedStatusData[id]]);
            }
          }

          // If we have cells to update, do it in a batch
          if (updateCount > 0) {
            const statusRange = sheet.getRange(
              2,
              syncStatusColumn,
              updateCount,
              1
            );
            statusRange.setValues(statusCellsToUpdate);
          }
        }
      } catch (e) {
      }
    }

    return true;
  } catch (error) {
    return false;
  }
}

/**
 * Makes header names unique and descriptive
 * @param {Array} headers - Original header names
 * @param {Array} columns - Column configuration objects
 * @return {Array} Unique header names
 */
function makeHeadersUnique(headers, columns) {
  // Map to track header occurrences and corresponding column keys
  const headerMap = new Map();
  const resultHeaders = [];

  headers.forEach((header, index) => {
    // Check if this is a custom name from column config
    let isCustomName = false;
    if (columns && columns[index] && columns[index].customName) {
      isCustomName = true;
    }

    // Skip entirely empty headers
    if (!header) {
      resultHeaders.push(`Column ${index + 1}`);
      return;
    }

    let originalHeader = header;

    // Get the column key if available (to help with categorization)
    let columnKey = "";
    if (columns && columns[index] && columns[index].key) {
      columnKey = columns[index].key;
    }

    // If this is the first occurrence of the header
    if (!headerMap.has(header)) {
      headerMap.set(header, {
        count: 1,
        columnKeys: [columnKey],
      });
      resultHeaders.push(header);
    } else {
      // This is a duplicate header
      const headerInfo = headerMap.get(header);
      headerInfo.count++;
      headerInfo.columnKeys.push(columnKey);
      headerMap.set(header, headerInfo);

      // For custom names, preserve them but add a number suffix to make unique
      if (isCustomName) {
        resultHeaders.push(`${header} (${headerInfo.count})`);
        return;
      }

      // Attempt to make a more descriptive name based on column key
      if (columnKey && columnKey.includes("_")) {
        // For keys with components like hash_component
        if (columnKey.includes("_subpremise")) {
          resultHeaders.push(`${header} - Suite/Apt`);
        } else if (columnKey.includes("_street_number")) {
          resultHeaders.push(`${header} - Street Number`);
        } else if (columnKey.includes("_route")) {
          resultHeaders.push(`${header} - Street Name`);
        } else if (columnKey.includes("_locality")) {
          resultHeaders.push(`${header} - City`);
        } else if (columnKey.includes("_country")) {
          resultHeaders.push(`${header} - Country`);
        } else if (columnKey.includes("_postal_code")) {
          resultHeaders.push(`${header} - ZIP/Postal`);
        } else if (columnKey.includes("_formatted_address")) {
          resultHeaders.push(`${header} - Full Address`);
        } else if (columnKey.includes("_timezone_id")) {
          resultHeaders.push(`${header} - Timezone`);
        } else if (columnKey.includes("_until")) {
          resultHeaders.push(`${header} - End Time/Date`);
        } else if (columnKey.includes("_currency")) {
          resultHeaders.push(`${header} - Currency`);
        } else {
          // Default numbering if no special case applies
          resultHeaders.push(`${header} (${headerInfo.count})`);
        }
      } else if (columnKey && columnKey.includes(".")) {
        // For nested fields like person_id.name
        const parts = columnKey.split(".");
        if (parts.length === 2) {
          resultHeaders.push(`${header} - ${formatBasicName(parts[1])}`);
        } else {
          resultHeaders.push(`${header} (${headerInfo.count})`);
        }
      } else {
        // Default numbering if no special case applies
        resultHeaders.push(`${header} (${headerInfo.count})`);
      }
    }
  });

  return resultHeaders;
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
      lastUpdate: new Date().getTime(),
    };

    // Ensure progress is 100% for completed phases
    if (status === "completed") {
      progress = 100;
      syncStatus.progress = 100;
    }

    // Save to user properties in our format
    userProps.setProperty("SYNC_STATUS", JSON.stringify(syncStatus));

    // Save in the original format for compatibility with the HTML
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_STATUS`, status);
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_DETAIL`, detail || "");
    scriptProperties.setProperty(
      `SYNC_PHASE_${phase}_PROGRESS`,
      progress ? progress.toString() : "0"
    );

    // Set current phase
    scriptProperties.setProperty("SYNC_CURRENT_PHASE", phase.toString());

    // If status is error, store the error
    if (status === "error") {
      scriptProperties.setProperty("SYNC_ERROR", detail || "An error occurred");
      scriptProperties.setProperty("SYNC_COMPLETED", "true");
      syncStatus.error = detail || "An error occurred";
    }

    // If this is the final phase completion, mark as completed
    if (status === "completed" && phase === "3") {
      scriptProperties.setProperty("SYNC_COMPLETED", "true");
    }

    // Also show a toast message for visibility
    let toastMessage = "";
    if (phase === "1") toastMessage = "Connecting to Pipedrive...";
    else if (phase === "2") toastMessage = "Retrieving data from Pipedrive...";
    else if (phase === "3") toastMessage = "Writing data to spreadsheet...";

    if (status === "error") {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Error: ${detail}`,
        "Sync Error",
        5
      );
    } else if (status === "completed" && phase === "3") {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Sync completed successfully!",
        "Sync Status",
        3
      );
    } else if (detail) {
      SpreadsheetApp.getActiveSpreadsheet().toast(detail, toastMessage, 2);
    }

    return syncStatus;
  } catch (e) {
    // Still show a toast message as backup
    SpreadsheetApp.getActiveSpreadsheet().toast(
      detail || "Processing...",
      "Sync Status",
      2
    );
    return null;
  }
}

/**
 * Shows the sync status dialog to the user
 * @param {string} sheetName - The name of the sheet being synced
 */
function showSyncStatus(sheetName) {
  try {
    // Reset any previous sync status
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("SYNC_COMPLETED", "false");
    scriptProperties.setProperty("SYNC_ERROR", "");

    // Initialize status for each phase
    for (let phase = 1; phase <= 3; phase++) {
      const status = phase === 1 ? "active" : "pending";
      const detail =
        phase === 1 ? "Connecting to Pipedrive..." : "Waiting to start...";
      const progress = phase === 1 ? 0 : 0;

      scriptProperties.setProperty(`SYNC_PHASE_${phase}_STATUS`, status);
      scriptProperties.setProperty(`SYNC_PHASE_${phase}_DETAIL`, detail);
      scriptProperties.setProperty(
        `SYNC_PHASE_${phase}_PROGRESS`,
        progress.toString()
      );
    }

    // Set current phase to 1
    scriptProperties.setProperty("SYNC_CURRENT_PHASE", "1");

    // Get entity type for the sheet
    const entityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType =
      scriptProperties.getProperty(entityTypeKey) || ENTITY_TYPES.DEALS;
    const entityTypeName = formatEntityTypeName(entityType);

    // Create the dialog
    const htmlTemplate = HtmlService.createTemplateFromFile("SyncStatus");
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.entityType = entityType;
    htmlTemplate.entityTypeName = entityTypeName;

    const html = htmlTemplate
      .evaluate()
      .setWidth(400)
      .setHeight(400)
      .setTitle("Sync Status");

    // Show dialog
    SpreadsheetApp.getUi().showModalDialog(html, "Sync Status");

    // Return true to indicate success
    return true;
  } catch (error) {

    // Show a fallback toast message instead
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Starting sync operation...",
      "Pipedrive Sync",
      3
    );
    return false;
  }
}

/**
 * Helper function to format entity type name for display
 * @param {string} entityType - The entity type to format
 * @return {string} Formatted entity type name
 */
function formatEntityTypeName(entityType) {
  if (!entityType) return "";

  const entityMap = {
    deals: "Deals",
    persons: "Persons",
    organizations: "Organizations",
    activities: "Activities",
    leads: "Leads",
    products: "Products",
  };

  return (
    entityMap[entityType] ||
    entityType.charAt(0).toUpperCase() + entityType.slice(1)
  );
}

/**
 * Gets the current sync status for the dialog to poll
 * @return {Object} Sync status object or null if not available
 */
function getSyncStatus() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const userProps = PropertiesService.getUserProperties();
    const statusJson = userProps.getProperty("SYNC_STATUS");

    if (!statusJson) {
      // Return default format matching the expected structure in the HTML
      return {
        phase1: {
          status:
            scriptProperties.getProperty("SYNC_PHASE_1_STATUS") || "active",
          detail:
            scriptProperties.getProperty("SYNC_PHASE_1_DETAIL") ||
            "Connecting to Pipedrive...",
          progress: parseInt(
            scriptProperties.getProperty("SYNC_PHASE_1_PROGRESS") || "0"
          ),
        },
        phase2: {
          status:
            scriptProperties.getProperty("SYNC_PHASE_2_STATUS") || "pending",
          detail:
            scriptProperties.getProperty("SYNC_PHASE_2_DETAIL") ||
            "Waiting to start...",
          progress: parseInt(
            scriptProperties.getProperty("SYNC_PHASE_2_PROGRESS") || "0"
          ),
        },
        phase3: {
          status:
            scriptProperties.getProperty("SYNC_PHASE_3_STATUS") || "pending",
          detail:
            scriptProperties.getProperty("SYNC_PHASE_3_DETAIL") ||
            "Waiting to start...",
          progress: parseInt(
            scriptProperties.getProperty("SYNC_PHASE_3_PROGRESS") || "0"
          ),
        },
        currentPhase: scriptProperties.getProperty("SYNC_CURRENT_PHASE") || "1",
        completed: scriptProperties.getProperty("SYNC_COMPLETED") || "false",
        error: scriptProperties.getProperty("SYNC_ERROR") || "",
      };
    }

    // Convert from our internal format to the format expected by the HTML
    const status = JSON.parse(statusJson);

    // Identify which phase is active based on the phase field
    const activePhase = status.phase || "1";
    const statusValue = status.status || "active";
    const detailValue = status.detail || "";
    const progressValue = status.progress || 0;

    // Create the response in the format expected by the HTML
    const response = {
      phase1: {
        status:
          activePhase === "1"
            ? statusValue
            : activePhase > "1"
            ? "completed"
            : "pending",
        detail:
          activePhase === "1"
            ? detailValue
            : activePhase > "1"
            ? "Completed"
            : "Waiting to start...",
        progress:
          activePhase === "1" ? progressValue : activePhase > "1" ? 100 : 0,
      },
      phase2: {
        status:
          activePhase === "2"
            ? statusValue
            : activePhase > "2"
            ? "completed"
            : "pending",
        detail:
          activePhase === "2"
            ? detailValue
            : activePhase > "2"
            ? "Completed"
            : "Waiting to start...",
        progress:
          activePhase === "2" ? progressValue : activePhase > "2" ? 100 : 0,
      },
      phase3: {
        status:
          activePhase === "3"
            ? statusValue
            : activePhase > "3"
            ? "completed"
            : "pending",
        detail:
          activePhase === "3"
            ? detailValue
            : activePhase > "3"
            ? "Completed"
            : "Waiting to start...",
        progress:
          activePhase === "3" ? progressValue : activePhase > "3" ? 100 : 0,
      },
      currentPhase: activePhase,
      completed:
        activePhase === "3" && status.status === "completed" ? "true" : "false",
      error: status.error || "",
    };

    return response;
  } catch (e) {
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
  let letter = "";
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
    // First, remove any existing onEdit triggers to avoid duplicates
    removeOnEditTrigger();

    // Then create a new trigger
    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  } catch (e) {
  }
}

/**
 * Removes the onEdit trigger
 */
function removeOnEditTrigger() {
  try {
    // Get all triggers
    const triggers = ScriptApp.getProjectTriggers();

    // Find and delete the onEdit trigger
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      if (trigger.getHandlerFunction() === "onEdit") {
        ScriptApp.deleteTrigger(trigger);
        return;
      }
    }

  } catch (e) {
  }
}

/**
 * Handles edits to the sheet and marks rows as modified for two-way sync
 * This function is automatically triggered when a user edits the sheet
 * @param {Object} e The edit event object
 */
function onEdit(e) {
  try {
    // Get the edited sheet
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled =
      scriptProperties.getProperty(twoWaySyncEnabledKey) === "true";

    // If two-way sync is not enabled, exit
    if (!twoWaySyncEnabled) {
      return;
    }

    // Get the edited range
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();

    // Create a unique execution ID to prevent duplicate processing
    const executionId = Utilities.getUuid();
    const lockKey = `EDIT_LOCK_${sheetName}`;

    try {
      // Try to acquire a lock using properties
      const currentLock = scriptProperties.getProperty(lockKey);

      // If there's an active lock, exit
      if (currentLock) {
        const lockData = JSON.parse(currentLock);
        const now = new Date().getTime();

        // If the lock is less than 5 seconds old, exit
        if (now - lockData.timestamp < 5000) {
          return;
        }

        // Lock is old, we can override it
      }

      // Set a new lock
      scriptProperties.setProperty(
        lockKey,
        JSON.stringify({
          id: executionId,
          timestamp: new Date().getTime(),
          row: row,
          col: column,
        })
      );
    } catch (lockError) {
      // Continue execution even if lock fails
    }

    // Check if the edit is in the header row - if it is, we might need to update tracking
    const headerRow = 1;
    if (row === headerRow) {
      // If someone renamed the Sync Status column, we'd handle that here
      // For now, just exit as we don't need special handling
      releaseLock(executionId, lockKey);
      return;
    }

    // Find the "Sync Status" column using our helper function
    const syncStatusColIndex = findSyncStatusColumn(sheet, sheetName);

    // Exit if no Sync Status column found
    if (syncStatusColIndex === -1) {
      releaseLock(executionId, lockKey);
      return;
    }

    // Get headers for later use
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // Convert to 1-based for sheet functions
    const syncStatusColPos = syncStatusColIndex + 1;

    // Check if the edit is in the Sync Status column itself (to avoid loops)
    if (column === syncStatusColPos) {
      releaseLock(executionId, lockKey);
      return;
    }

    // Get the row content to check if it's a real data row or a timestamp/blank row
    const rowContent = sheet
      .getRange(row, 1, 1, Math.min(10, sheet.getLastColumn()))
      .getValues()[0];

    // Check if this is a timestamp row
    const firstCell = String(rowContent[0] || "").toLowerCase();
    const isTimestampRow =
      firstCell.includes("last") ||
      firstCell.includes("updated") ||
      firstCell.includes("synced") ||
      firstCell.includes("date");

    // Count non-empty cells to determine if this is a data row
    const nonEmptyCells = rowContent.filter(
      (cell) => cell !== "" && cell !== null && cell !== undefined
    ).length;

    // Skip if this is a timestamp row or has too few cells with data
    if (isTimestampRow || nonEmptyCells < 3) {
      releaseLock(executionId, lockKey);
      return;
    }

    // Get the row ID from the first column
    const idColumnIndex = 0;
    const id = rowContent[idColumnIndex];

    // Skip rows without an ID (likely empty rows)
    if (!id) {
      releaseLock(executionId, lockKey);
      return;
    }

    // Get the sync status cell
    const syncStatusCell = sheet.getRange(row, syncStatusColPos);
    const currentStatus = syncStatusCell.getValue();

    // Get the cell state key for this cell
    const cellStateKey = `CELL_STATE_${sheetName}_${row}_${id}`;
    let cellState;

    try {
      const cellStateJson = scriptProperties.getProperty(cellStateKey);
      cellState = cellStateJson
        ? JSON.parse(cellStateJson)
        : {
            status: null,
            lastChanged: 0,
            originalValues: {},
          };
    } catch (parseError) {
      cellState = {
        status: null,
        lastChanged: 0,
        originalValues: {},
      };
    }

    // Get the current time
    const now = new Date().getTime();

    // Check for recent changes to prevent toggling
    if (
      cellState.lastChanged &&
      now - cellState.lastChanged < 5000 &&
      cellState.status === currentStatus
    ) {
      releaseLock(executionId, lockKey);
      return;
    }

    // Get the original data from Pipedrive
    const originalDataKey = `ORIGINAL_DATA_${sheetName}`;
    let originalData;

    try {
      // Try to get original data from document properties
      const originalDataJson = scriptProperties.getProperty(originalDataKey);
      originalData = originalDataJson ? JSON.parse(originalDataJson) : {};
    } catch (parseError) {
      originalData = {};
    }

    // Check if we have original data for this row
    const rowKey = id.toString();

    // Log for debugging

    // Get the column header name for the edited column
    const headerName = headers[column - 1]; // Adjust for 0-based array

    // Enhanced debug logging

    // If this row is already Modified, check if we should undo the status
    if (currentStatus === "Modified" && originalData[rowKey]) {
      // Get the original value for the field that was just edited
      const originalValue = originalData[rowKey][headerName];
      const currentValue = e.value;


      // First try direct comparison for exact matches
      let valuesMatch = originalValue === currentValue;

      // If values don't match exactly, try string conversion and trimming
      if (!valuesMatch) {
        const origString =
          originalValue === null || originalValue === undefined
            ? ""
            : String(originalValue).trim();
        const currString =
          currentValue === null || currentValue === undefined
            ? ""
            : String(currentValue).trim();
        valuesMatch = origString === currString;

      }

      // If the values match (original = current), check if all other values in the row match their originals
      if (valuesMatch) {

        // Get the current row values
        const rowValues = sheet
          .getRange(row, 1, 1, headers.length)
          .getValues()[0];

        // Flag to track if all values match
        let allMatch = true;

        // Check each stored original value
        for (const field in originalData[rowKey]) {
          // Skip checking the same field we just verified
          if (field === headerName) continue;

          // Skip non-data fields
          if (field === "Sync Status") continue;

          // Find the column index for this field
          const fieldIndex = headers.indexOf(field);
          if (fieldIndex === -1) {
            continue;
          }

          // Get the original and current values
          const origValue = originalData[rowKey][field];
          const currValue = rowValues[fieldIndex];

          // First try direct comparison
          let fieldMatch = origValue === currValue;

          // If direct comparison fails, try string conversion
          if (!fieldMatch) {
            const origStr =
              origValue === null || origValue === undefined
                ? ""
                : String(origValue).trim();
            const currStr =
              currValue === null || currValue === undefined
                ? ""
                : String(currValue).trim();
            fieldMatch = origStr === currStr;

            // Special handling for numbers
            if (!fieldMatch && !isNaN(origValue) && !isNaN(currValue)) {
              // Try numeric comparison with potential floating point issues
              const origNum = parseFloat(origValue);
              const currNum = parseFloat(currValue);

              // Check if numbers are close enough (within a small epsilon)
              const epsilon = 0.0001;
              if (Math.abs(origNum - currNum) < epsilon) {
                fieldMatch = true;
              }
            }

            // Special handling for dates
            if (
              !fieldMatch &&
              (origStr.match(/^\d{4}-\d{2}-\d{2}/) ||
                currStr.match(/^\d{4}-\d{2}-\d{2}/))
            ) {
              try {
                // Try date parsing
                const origDate = new Date(origStr);
                const currDate = new Date(currStr);

                if (!isNaN(origDate.getTime()) && !isNaN(currDate.getTime())) {
                  // Compare dates
                  fieldMatch = origDate.getTime() === currDate.getTime();
                }
              } catch (e) {
              }
            }
          }


          // If any field doesn't match, set flag to false and break
          if (!fieldMatch) {
            allMatch = false;
            break;
          }
        }

        // If all fields match their original values, set status back to "Not Modified"
        if (allMatch) {

          // Mark as not modified
          syncStatusCell.setValue("Not modified");

          // Update cell state
          cellState.status = "Not modified";
          cellState.lastChanged = now;

          try {
            scriptProperties.setProperty(
              cellStateKey,
              JSON.stringify(cellState)
            );
          } catch (saveError) {
          }

          // Apply correct formatting
          syncStatusCell.setBackground("#F8F9FA").setFontColor("#000000");

          // Skip data validation - removed to avoid formatting issues

          releaseLock(executionId, lockKey);
          return;
        }
      }
    }

    // Handle first-time edit case (current status is not Modified)
    if (currentStatus !== "Modified") {
      // Store the original value before marking as modified
      if (!originalData[rowKey]) {
        originalData[rowKey] = {};
      }

      if (headerName) {
        // Make sure we store the original value with the proper type and formatting
        const oldValue = e.oldValue !== undefined ? e.oldValue : null;

        // Store original value and extra debug info
        originalData[rowKey][headerName] = oldValue;

        // Log detailed information about the stored original value

        // Save updated original data
        try {
          scriptProperties.setProperty(
            originalDataKey,
            JSON.stringify(originalData)
          );
        } catch (saveError) {
        }

        // Mark as modified (with special prevention of change-back)
        syncStatusCell.setValue("Modified");

        // Save new cell state to prevent toggling back
        cellState.status = "Modified";
        cellState.lastChanged = now;
        cellState.originalValues[headerName] = oldValue;

        try {
          scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
        } catch (saveError) {
        }

        // Skip data validation - removed to avoid formatting issues

        // Make sure the styling is consistent
        syncStatusCell.setBackground("#FCE8E6").setFontColor("#D93025");
      }
    } else {
      // Row is already modified - check if this edit reverts to original value
      if (
        headerName &&
        originalData[rowKey] &&
        originalData[rowKey][headerName] !== undefined
      ) {
        const originalValue = originalData[rowKey][headerName];
        const currentValue = e.value;


        // Improved equality check - try to normalize values for comparison regardless of type
        let originalString =
          originalValue === null || originalValue === undefined
            ? ""
            : String(originalValue).trim();
        let currentString =
          currentValue === null || currentValue === undefined
            ? ""
            : String(currentValue).trim();

        // Special handling for email fields - normalize domains for comparison
        if (headerName.toLowerCase().includes("email")) {
          // Apply email normalization rules
          if (originalString.includes("@")) {
            const origParts = originalString.split("@");
            const origUsername = origParts[0].toLowerCase();
            let origDomain = origParts[1].toLowerCase();

            // Fix common domain typos
            if (origDomain === "gmail.comm") origDomain = "gmail.com";
            if (origDomain === "gmail.con") origDomain = "gmail.com";
            if (origDomain === "gmial.com") origDomain = "gmail.com";
            if (origDomain === "hotmail.comm") origDomain = "hotmail.com";
            if (origDomain === "hotmail.con") origDomain = "hotmail.com";
            if (origDomain === "yahoo.comm") origDomain = "yahoo.com";
            if (origDomain === "yahoo.con") origDomain = "yahoo.com";

            // Reassemble normalized email
            originalString = origUsername + "@" + origDomain;
          }

          if (currentString.includes("@")) {
            const currParts = currentString.split("@");
            const currUsername = currParts[0].toLowerCase();
            let currDomain = currParts[1].toLowerCase();

            // Fix common domain typos
            if (currDomain === "gmail.comm") currDomain = "gmail.com";
            if (currDomain === "gmail.con") currDomain = "gmail.com";
            if (currDomain === "gmial.com") currDomain = "gmail.com";
            if (currDomain === "hotmail.comm") currDomain = "hotmail.com";
            if (currDomain === "hotmail.con") currDomain = "hotmail.com";
            if (currDomain === "yahoo.comm") currDomain = "yahoo.com";
            if (currDomain === "yahoo.con") currDomain = "yahoo.com";

            // Reassemble normalized email
            currentString = currUsername + "@" + currDomain;
          }

        }
        // Special handling for name fields - normalize common typos
        else if (headerName.toLowerCase().includes("name")) {
          // Check for common name typos like extra letter at the end
          if (originalString.length > 0 && currentString.length > 0) {
            // Check if one string is the same as the other with an extra character at the end
            if (originalString.length === currentString.length + 1) {
              if (originalString.startsWith(currentString)) {
                originalString = currentString;
              }
            } else if (currentString.length === originalString.length + 1) {
              if (currentString.startsWith(originalString)) {
                currentString = originalString;
              }
            }
            // Check for single character mismatch at the end (e.g., "Simpson" vs "Simpsonm")
            else if (originalString.length === currentString.length) {
              // Find the first character that differs
              let diffIndex = -1;
              for (let i = 0; i < originalString.length; i++) {
                if (originalString[i] !== currentString[i]) {
                  diffIndex = i;
                  break;
                }
              }

              // If the difference is near the end
              if (diffIndex > 0 && diffIndex >= originalString.length - 2) {
                // Normalize by taking the shorter version up to the differing character
                const normalizedName = originalString.substring(0, diffIndex);
                originalString = normalizedName;
                currentString = normalizedName;
              }
            }
          }

        }

        // For numeric values, try to normalize scientific notation and number formats
        if (
          !isNaN(parseFloat(originalString)) &&
          !isNaN(parseFloat(currentString))
        ) {
          // Convert both to numbers and back to strings for comparison
          try {
            const origNum = parseFloat(originalString);
            const currNum = parseFloat(currentString);

            // If both are integers, compare as integers
            if (
              Math.floor(origNum) === origNum &&
              Math.floor(currNum) === currNum
            ) {
              originalString = Math.floor(origNum).toString();
              currentString = Math.floor(currNum).toString();
            } else {
              // Compare with fixed decimal places for floating point numbers
              originalString = origNum.toString();
              currentString = currNum.toString();
            }
          } catch (numError) {
          }
        }

        // Check if this is a structural field with complex nested structure
        if (
          originalValue &&
          typeof originalValue === "object" &&
          originalValue.__isStructural
        ) {

          // Simple direct comparison before complex checks
          if (originalString === currentString) {

            // Check if all edited values in the row now match original values
            const allMatch = checkAllValuesMatchOriginal(
              sheet,
              row,
              headers,
              originalData[rowKey]
            );


            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");

              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true; // Special flag to indicate this is an undo operation

              try {
                scriptProperties.setProperty(
                  cellStateKey,
                  JSON.stringify(cellState)
                );
              } catch (saveError) {
              }

              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(
                  noChangeLockKey,
                  JSON.stringify({
                    timestamp: now,
                    expiry: now + 10000, // 10 seconds
                    status: "Not modified",
                  })
                );
              } catch (lockError) {
              }

              // Skip data validation - removed to avoid formatting issues

              // Reset formatting
              syncStatusCell.setBackground("#F8F9FA").setFontColor("#000000");
            }

            releaseLock(executionId, lockKey);
            return;
          }

          // Create a data object that mimics Pipedrive structure
          const dataObj = {
            id: id,
          };

          // Try to reconstruct the structure based on the __key
          const key = originalValue.__key;
          const parts = key.split(".");
          const structureType = parts[0];

          if (["phone", "email"].includes(structureType)) {
            // Handle phone/email fields
            dataObj[structureType] = [];

            // If it's a label-based path (e.g., phone.mobile)
            if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
              dataObj[structureType].push({
                label: parts[1],
                value: currentValue,
              });
            }
            // If it's an array index path (e.g., phone.0.value)
            else if (parts.length === 3 && parts[2] === "value") {
              const idx = parseInt(parts[1]);
              while (dataObj[structureType].length <= idx) {
                dataObj[structureType].push({});
              }
              dataObj[structureType][idx].value = currentValue;
            }
          }
          // Custom fields
          else if (structureType === "custom_fields") {
            dataObj.custom_fields = {};

            if (parts.length === 2) {
              // Simple custom field
              dataObj.custom_fields[parts[1]] = currentValue;
            } else if (parts.length > 2) {
              // Nested custom field like address or currency
              dataObj.custom_fields[parts[1]] = {};

              // Handle complex types
              if (parts[2] === "formatted_address") {
                dataObj.custom_fields[parts[1]].formatted_address =
                  currentValue;
              } else if (parts[2] === "currency") {
                dataObj.custom_fields[parts[1]].currency = currentValue;
              } else {
                dataObj.custom_fields[parts[1]][parts[2]] = currentValue;
              }
            }
          } else {
            // Other nested fields not covered above

            // Build a generic nested structure
            let current = dataObj;
            for (let i = 0; i < parts.length - 1; i++) {
              if (current[parts[i]] === undefined) {
                if (!isNaN(parseInt(parts[i + 1]))) {
                  current[parts[i]] = [];
                } else {
                  current[parts[i]] = {};
                }
              }
              current = current[parts[i]];
            }
            current[parts[parts.length - 1]] = currentValue;
          }

          // Compare using the normalized values
          const normalizedOriginal = originalValue.__normalized || "";
          const normalizedCurrent = getNormalizedFieldValue(dataObj, key);


          // Check if values match
          const valuesMatch = normalizedOriginal === normalizedCurrent;

          // If values match, check all fields
          if (valuesMatch) {
            // Check if all edited values in the row now match original values
            const allMatch = checkAllValuesMatchOriginal(
              sheet,
              row,
              headers,
              originalData[rowKey]
            );


            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");

              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true; // Special flag to indicate this is an undo operation

              try {
                scriptProperties.setProperty(
                  cellStateKey,
                  JSON.stringify(cellState)
                );
              } catch (saveError) {
              }

              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(
                  noChangeLockKey,
                  JSON.stringify({
                    timestamp: now,
                    expiry: now + 10000, // 10 seconds
                    status: "Not modified",
                  })
                );
              } catch (lockError) {
              }

              // Skip data validation - removed to avoid formatting issues

              // Reset formatting
              syncStatusCell.setBackground("#F8F9FA").setFontColor("#000000");
            }
          }
        } else {
          // This is a regular field, not a structural field

          // Special handling for null/empty values
          if (
            (originalValue === null || originalValue === "") &&
            (currentValue === null || currentValue === "")
          ) {
          }

          // Simple direct comparison before complex checks
          if (originalString === currentString) {

            // Check if all edited values in the row now match original values
            const allMatch = checkAllValuesMatchOriginal(
              sheet,
              row,
              headers,
              originalData[rowKey]
            );


            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");

              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true; // Special flag to indicate this is an undo operation

              try {
                scriptProperties.setProperty(
                  cellStateKey,
                  JSON.stringify(cellState)
                );
              } catch (saveError) {
              }

              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(
                  noChangeLockKey,
                  JSON.stringify({
                    timestamp: now,
                    expiry: now + 10000, // 10 seconds
                    status: "Not modified",
                  })
                );
              } catch (lockError) {
              }

              // Skip data validation - removed to avoid formatting issues

              // Reset formatting
              syncStatusCell.setBackground("#F8F9FA").setFontColor("#000000");
            }

            releaseLock(executionId, lockKey);
            return;
          }

          // Create a data object that mimics Pipedrive structure
          const dataObj = {
            id: id,
          };

          // Populate the field being edited
          if (headerName.includes(".")) {
            // Handle nested structure
            const parts = headerName.split(".");

            if (["phone", "email"].includes(parts[0])) {
              // Handle phone/email fields
              dataObj[parts[0]] = [];

              // If it's a label-based path (e.g., phone.mobile)
              if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
                dataObj[parts[0]].push({
                  label: parts[1],
                  value: currentValue,
                });
              }
              // If it's an array index path (e.g., phone.0.value)
              else if (parts.length === 3 && parts[2] === "value") {
                const idx = parseInt(parts[1]);
                while (dataObj[parts[0]].length <= idx) {
                  dataObj[parts[0]].push({});
                }
                dataObj[parts[0]][idx].value = currentValue;
              }
            }
            // Custom fields
            else if (parts[0] === "custom_fields") {
              dataObj.custom_fields = {};

              if (parts.length === 2) {
                // Simple custom field
                dataObj.custom_fields[parts[1]] = currentValue;
              } else if (parts.length > 2) {
                // Nested custom field like address or currency
                dataObj.custom_fields[parts[1]] = {};

                // Handle complex types
                if (parts[2] === "formatted_address") {
                  dataObj.custom_fields[parts[1]].formatted_address =
                    currentValue;
                } else if (parts[2] === "currency") {
                  dataObj.custom_fields[parts[1]].currency = currentValue;
                } else {
                  dataObj.custom_fields[parts[1]][parts[2]] = currentValue;
                }
              }
            } else {
              // Other nested fields not covered above

              // Build a generic nested structure
              let current = dataObj;
              for (let i = 0; i < parts.length - 1; i++) {
                if (!current[parts[i]]) {
                  current[parts[i]] = {};
                }
                current = current[parts[i]];
              }
              current[parts[parts.length - 1]] = currentValue;

            }
          } else {
            // Regular top-level field
            dataObj[headerName] = currentValue;
          }

          // Dump the constructed data object

          // Use the generalized field value normalization for comparison
          const normalizedOriginal = getNormalizedFieldValue(
            {
              [headerName]: originalValue,
            },
            headerName
          );
          const normalizedCurrent = getNormalizedFieldValue(
            dataObj,
            headerName
          );


          // Check if values match
          const valuesMatch = normalizedOriginal === normalizedCurrent;

          // If values match, check all fields
          if (valuesMatch) {
            // Check if all edited values in the row now match original values
            const allMatch = checkAllValuesMatchOriginal(
              sheet,
              row,
              headers,
              originalData[rowKey]
            );


            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");

              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true; // Special flag to indicate this is an undo operation

              try {
                scriptProperties.setProperty(
                  cellStateKey,
                  JSON.stringify(cellState)
                );
              } catch (saveError) {
              }

              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(
                  noChangeLockKey,
                  JSON.stringify({
                    timestamp: now,
                    expiry: now + 10000, // 10 seconds
                    status: "Not modified",
                  })
                );
              } catch (lockError) {
              }

              // Skip data validation - removed to avoid formatting issues

              // Reset formatting
              syncStatusCell.setBackground("#F8F9FA").setFontColor("#000000");
            }
          }
        }
      } else if (e.oldValue !== undefined && headerName) {
        // Store the first known value as original if we don't have it yet
        if (!originalData[rowKey]) {
          originalData[rowKey] = {};
        }

        if (!originalData[rowKey][headerName]) {
          originalData[rowKey][headerName] = e.oldValue;

          // Save updated original data
          try {
            scriptProperties.setProperty(
              originalDataKey,
              JSON.stringify(originalData)
            );
          } catch (saveError) {
          }
        }
      }
    }

    // Release the lock at the end
    releaseLock(executionId, lockKey);
  } catch (error) {
    // Silent fail for onEdit triggers
  }
}

/**
 * Helper function to release the lock
 * @param {string} executionId - The ID of the execution that set the lock
 * @param {string} lockKey - The key used for the lock
 */
function releaseLock(executionId, lockKey) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const currentLock = scriptProperties.getProperty(lockKey);

    if (currentLock) {
      const lockData = JSON.parse(currentLock);

      // Only release if this execution set the lock
      if (lockData.id === executionId) {
        scriptProperties.deleteProperty(lockKey);
      }
    }
  } catch (error) {
  }
}

/**
 * Helper function to check if all edited values in a row match their original values
 * @param {Sheet} sheet - The sheet containing the row
 * @param {number} row - The row number to check
 * @param {Array} headers - The column headers
 * @param {Object} originalValues - The original values for the row, keyed by header name
 * @return {boolean} True if all values match original, false otherwise
 */
function checkAllValuesMatchOriginal(sheet, row, headers, originalValues) {
  try {
    // If no original values stored, can't verify
    if (!originalValues || Object.keys(originalValues).length === 0) {
      return false;
    }


    // Get current values for the entire row
    const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

    // Get the first value (ID) to use for retrieving the original data
    const id = rowValues[0];

    // Create a data object that mimics Pipedrive structure for nested field handling
    const dataObj = {
      id: id,
    };

    // Create a mapping of header names to their column indices for faster lookup
    const headerIndices = {};
    headers.forEach((header, index) => {
      headerIndices[header] = index;
    });

    // Populate the data object with values from the row
    headers.forEach((header, index) => {
      if (index < rowValues.length) {
        // Use dot notation to create nested objects
        if (header.includes(".")) {
          const parts = header.split(".");

          // Common nested structures to handle specially
          if (["phone", "email"].includes(parts[0])) {
            // Handle phone/email specially
            if (!dataObj[parts[0]]) {
              dataObj[parts[0]] = [];
            }

            // If it's a label-based path (e.g., phone.mobile)
            if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
              dataObj[parts[0]].push({
                label: parts[1],
                value: rowValues[index],
              });
            }
            // If it's an array index path (e.g., phone.0.value)
            else if (parts.length === 3 && parts[2] === "value") {
              const idx = parseInt(parts[1]);
              while (dataObj[parts[0]].length <= idx) {
                dataObj[parts[0]].push({});
              }
              dataObj[parts[0]][idx].value = rowValues[index];
            }
          }
          // Custom fields
          else if (parts[0] === "custom_fields") {
            if (!dataObj.custom_fields) {
              dataObj.custom_fields = {};
            }

            if (parts.length === 2) {
              // Simple custom field
              dataObj.custom_fields[parts[1]] = rowValues[index];
            } else if (parts.length > 2) {
              // Nested custom field
              if (!dataObj.custom_fields[parts[1]]) {
                dataObj.custom_fields[parts[1]] = {};
              }
              // Handle complex types like address
              if (parts[2] === "formatted_address") {
                dataObj.custom_fields[parts[1]].formatted_address =
                  rowValues[index];
              } else if (parts[2] === "currency") {
                dataObj.custom_fields[parts[1]].currency = rowValues[index];
              } else {
                dataObj.custom_fields[parts[1]][parts[2]] = rowValues[index];
              }
            }
          } else {
            // Other nested paths - build the structure
            let current = dataObj;
            for (let i = 0; i < parts.length - 1; i++) {
              if (current[parts[i]] === undefined) {
                // If part is numeric, create an array
                if (!isNaN(parseInt(parts[i + 1]))) {
                  current[parts[i]] = [];
                } else {
                  current[parts[i]] = {};
                }
              }
              current = current[parts[i]];
            }

            // Set the value at the final level
            current[parts[parts.length - 1]] = rowValues[index];
          }
        } else {
          // Regular top-level field
          dataObj[header] = rowValues[index];
        }
      }
    });

    // Debug log

    // Check each column that has a stored original value
    let matchCount = 0;
    let mismatchCount = 0;

    for (const headerName in originalValues) {
      // Skip the Sync Status column itself
      if (headerName === "Sync Status") {
        continue;
      }

      // Find the column index for this header
      const colIndex = headerIndices[headerName];
      if (colIndex === undefined) {

        // Even if the header is not found, we'll try to compare by field name
        // This handles cases where the column position changed but the header name is the same
        let foundMatch = false;
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] === headerName) {
            foundMatch = true;

            const originalValue = originalValues[headerName];
            const currentValue = rowValues[i];

            // Compare the values
            const originalString =
              originalValue === null || originalValue === undefined
                ? ""
                : String(originalValue).trim();
            const currentString =
              currentValue === null || currentValue === undefined
                ? ""
                : String(currentValue).trim();


            if (originalString === currentString) {
              matchCount++;
            } else {
              mismatchCount++;
              return false; // Early exit on mismatch
            }

            break;
          }
        }

        if (!foundMatch) {
          // The header is truly missing, we can't compare
        }
        continue;
      }

      const originalValue = originalValues[headerName];
      const currentValue = rowValues[colIndex];

      // Special handling for null/empty values
      if (
        (originalValue === null || originalValue === "") &&
        (currentValue === null || currentValue === "")
      ) {
        matchCount++;
        continue; // Both empty, consider a match
      }

      // Check if this is a structural field with complex nested structure
      if (
        originalValue &&
        typeof originalValue === "object" &&
        originalValue.__isStructural
      ) {

        // Use the pre-computed normalized value for comparison
        const normalizedOriginal = originalValue.__normalized || "";
        const normalizedCurrent = getNormalizedFieldValue(
          dataObj,
          originalValue.__key
        );


        // If the normalized values don't match, return false
        if (normalizedOriginal !== normalizedCurrent) {
          mismatchCount++;
          return false;
        }

        matchCount++;
        // Skip to the next field
        continue;
      }

      // Use the generalized field value normalization for regular fields
      const normalizedOriginal = getNormalizedFieldValue(
        {
          [headerName]: originalValue,
        },
        headerName
      );
      const normalizedCurrent = getNormalizedFieldValue(dataObj, headerName);


      // If the normalized values don't match, return false
      if (normalizedOriginal !== normalizedCurrent) {
        mismatchCount++;
        return false;
      }

      matchCount++;
    }

    // If we reach here, all values with stored originals match
    return mismatchCount === 0 && matchCount > 0;
  } catch (error) {
    return false;
  }
}

/**
 * Normalizes a phone number by removing all non-digit characters
 * and handling scientific notation
 * @param {*} value - The phone number value to normalize
 * @return {string} The normalized phone number
 */
function normalizePhoneNumber(value) {
  try {
    // Handle null/undefined
    if (value === null || value === undefined) {
      return "";
    }

    // Handle array/object (like Pipedrive phone fields)
    if (typeof value === "object") {
      // If it's an array of phone objects (Pipedrive format)
      if (Array.isArray(value) && value.length > 0) {
        if (value[0] && value[0].value) {
          // Use the first phone number's value
          value = value[0].value;
        } else if (typeof value[0] === "string") {
          // It's just an array of strings, use the first one
          value = value[0];
        }
      } else if (value.value) {
        // It's a single phone object with a value property
        value = value.value;
      } else {
        // Try to extract a phone number from the object
        // This covers cases like complex nested objects
        const objStr = JSON.stringify(value);
        const phoneMatch = objStr.match(/"value":"([^"]+)"/);
        if (phoneMatch && phoneMatch[1]) {
          return normalizeDigitsOnly(phoneMatch[1]);
        } else {
          // Just stringify the object as a fallback
          value = objStr;
        }
      }
    }

    return normalizeDigitsOnly(value);
  } catch (e) {
    return String(value); // Return as string in case of error
  }
}

/**
 * Extracts only the digits from a value, handling scientific notation
 * and various number formats
 * @param {*} value - The value to normalize
 * @return {string} The normalized digits-only string
 */
function normalizeDigitsOnly(value) {
  try {
    // Handle null/undefined
    if (value === null || value === undefined) {
      return "";
    }

    // If it's a number or scientific notation, convert to a regular number string
    if (
      typeof value === "number" ||
      (typeof value === "string" && value.includes("E"))
    ) {
      // Parse it as a number first to handle scientific notation
      let numValue;
      try {
        numValue = Number(value);
        if (!isNaN(numValue)) {
          // Convert to regular number format without scientific notation
          return numValue.toFixed(0);
        }
      } catch (e) {
        // If parsing fails, continue with string handling
      }
    }

    // Convert to string if not already
    const strValue = String(value);

    // Remove all non-digit characters
    return strValue.replace(/\D/g, "");
  } catch (e) {
    return String(value);
  }
}

/**
 * Gets the value from a phone number field regardless of format
 * This is crucial for handling the different ways phone numbers appear in Pipedrive
 * @param {Object} data - The data object containing the phone field
 * @param {string} key - The key or path to the phone field
 * @return {string} The normalized phone number
 */
function getPhoneNumberFromField(data, key) {
  try {
    if (!data || !key) return "";

    // If key is already a value, just normalize it
    if (typeof key !== "string") {
      return normalizePhoneNumber(key);
    }

    // Handle different path formats for phone numbers

    // Case 1: Direct phone field (e.g., "phone")
    if (key === "phone" && data.phone) {
      return normalizePhoneNumber(data.phone);
    }

    // Case 2: Specific label format (e.g., "phone.mobile")
    if (key.startsWith("phone.") && key.split(".").length === 2) {
      const label = key.split(".")[1];

      // Handle the array of phone objects with labels
      if (Array.isArray(data.phone)) {
        // Try to find a phone with the matching label
        const match = data.phone.find(
          (p) => p && p.label && p.label.toLowerCase() === label.toLowerCase()
        );

        if (match && match.value) {
          return normalizePhoneNumber(match.value);
        }

        // If not found but we were looking for primary, try to find primary flag
        if (label === "primary") {
          const primary = data.phone.find((p) => p && p.primary);
          if (primary && primary.value) {
            return normalizePhoneNumber(primary.value);
          }
        }

        // If nothing found, try first phone
        if (data.phone.length > 0 && data.phone[0] && data.phone[0].value) {
          return normalizePhoneNumber(data.phone[0].value);
        }
      }
    }

    // Case 3: Array index format (e.g., "phone.0.value")
    if (key.startsWith("phone.") && key.includes(".value")) {
      const parts = key.split(".");
      const index = parseInt(parts[1]);

      if (
        !isNaN(index) &&
        Array.isArray(data.phone) &&
        data.phone.length > index &&
        data.phone[index]
      ) {
        return normalizePhoneNumber(data.phone[index].value);
      }
    }

    // Case 4: Use getValueByPath as a fallback
    // This handles other complex nested paths
    try {
      const value = getValueByPath(data, key);
      return normalizePhoneNumber(value);
    } catch (e) {
    }

    // If all else fails, return empty string
    return "";
  } catch (e) {
    return "";
  }
}

/**
 * Main function to sync deals from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncDealsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.DEALS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync persons from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncPersonsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PERSONS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync organizations from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncOrganizationsFromFilter(
  filterId,
  skipPush = false,
  sheetName = null
) {
  syncPipedriveDataToSheet(
    ENTITY_TYPES.ORGANIZATIONS,
    skipPush,
    sheetName,
    filterId
  );
}

/**
 * Main function to sync activities from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncActivitiesFromFilter(
  filterId,
  skipPush = false,
  sheetName = null
) {
  syncPipedriveDataToSheet(
    ENTITY_TYPES.ACTIVITIES,
    skipPush,
    sheetName,
    filterId
  );
}

/**
 * Main function to sync leads from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncLeadsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.LEADS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet™
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncProductsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(
    ENTITY_TYPES.PRODUCTS,
    skipPush,
    sheetName,
    filterId
  );
}

/**
 * Pushes changes from the sheet back to Pipedrive
 * Refactored version with focus on reliable address component handling
 * @param {boolean} isScheduledSync - Whether this is called from a scheduled sync
 * @param {boolean} suppressNoModifiedWarning - Whether to suppress the no modified rows warning
 */
async function pushChangesToPipedrive(
  isScheduledSync = false,
  suppressNoModifiedWarning = false
) {
  detectColumnShifts();
  try {
    // Make sure the user is authenticated
    const accessToken = ScriptApp.getOAuthToken();
    if (!accessToken) {
      // Show an auth error
      SpreadsheetApp.getUi().alert(
        "Authentication Error",
        'You need to authorize the script to access Pipedrive. Click the "Connect to Pipedrive" button in the Pipedrive menu.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get active sheet and validate settings
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    const scriptProperties = PropertiesService.getScriptProperties();

    // Verify two-way sync is enabled
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncEnabled =
      scriptProperties.getProperty(twoWaySyncEnabledKey) === "true";

    if (!twoWaySyncEnabled) {
      if (!isScheduledSync) {
        SpreadsheetApp.getUi().alert(
          "Two-Way Sync Not Enabled",
          "Two-way sync is not enabled for this sheet. Please enable it in the Two-Way Sync Settings.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }

    // Get entity type for the active sheet
    const entityTypeProperty = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(entityTypeProperty);

    // Get subdomain from script properties
    const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN");

    if (!entityType) {
      if (!isScheduledSync) {
        SpreadsheetApp.getUi().alert(
          "Configuration Missing",
          'This sheet has not been configured for Pipedrive synchronization. Please use the "Configure Sync" button in the Pipedrive menu.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }

    // Log entity type for debugging

    // Make sure we have an OAuth token that is still valid
    try {
      // Get the current OAuth token - no need for tokenManager
      const accessToken = ScriptApp.getOAuthToken();
      if (!accessToken) {
        if (!isScheduledSync) {
          SpreadsheetApp.getUi().alert(
            "Authentication Error",
            'There was a problem with your Pipedrive authentication. Please re-authorize by clicking the "Connect to Pipedrive" button in the Pipedrive menu.',
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        }
        return;
      }
    } catch (tokenError) {
      if (!isScheduledSync) {
        SpreadsheetApp.getUi().alert(
          "Authentication Error",
          'There was a problem with your Pipedrive authentication. Please re-authorize by clicking the "Connect to Pipedrive" button in the Pipedrive menu.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }

    // Create API configuration
    const apiBasePath = `https://${subdomain}.pipedrive.com/v1`;
    const apiConfig = {
      basePath: apiBasePath,
      apiKey: accessToken,
    };

    // For improved handling of Deals updates, use the enhanced DealsApi
    let enhancedDealsApi = null;
    if (
      entityType === "deals" &&
      typeof AppLib !== "undefined" &&
      AppLib.getUpdatedDealsApi
    ) {
      try {
        enhancedDealsApi = AppLib.getUpdatedDealsApi(accessToken, apiBasePath);
      } catch (dealsApiErr) {
      }
    }

    // Log the configuration object

    // Initialize Pipedrive client using the npm package
    let apiClient;
    try {
      // Get the Pipedrive library through AppLib
      const pipedriveLib = AppLib.getPipedriveLib();

      // Log the structure to understand what's available
      logObjectStructure(pipedriveLib, "pipedriveLib");

      // The library has v1 and v2 namespaces - we'll use v1
      const pipedriveV1 = pipedriveLib.v1;
      logObjectStructure(pipedriveV1, "pipedriveV1");

      // For improved handling of Deals updates, use the enhanced DealsApi
      if (
        entityType === "deals" &&
        typeof AppLib !== "undefined" &&
        AppLib.getUpdatedDealsApi
      ) {
        try {
          apiClient = AppLib.getUpdatedDealsApi(accessToken, apiBasePath);
        } catch (dealsApiErr) {
          apiClient = new pipedriveV1.DealsApi(apiConfig);
        }
      } else {
        // Select the appropriate API client based on entity type
        switch (entityType) {
          case "deals":
            apiClient = new pipedriveV1.DealsApi(apiConfig);
            logObjectStructure(apiClient, "deals API client");
            break;
          case "persons":
            apiClient = new pipedriveV1.PersonsApi(apiConfig);
            logObjectStructure(apiClient, "persons API client");
            break;
          case "organizations":
            apiClient = new pipedriveV1.OrganizationsApi(apiConfig);
            logObjectStructure(apiClient, "organizations API client");
            break;
          case "activities":
            apiClient = new pipedriveV1.ActivitiesApi(apiConfig);
            logObjectStructure(apiClient, "activities API client");
            break;
          case "leads":
            apiClient = new pipedriveV1.LeadsApi(apiConfig);
            logObjectStructure(apiClient, "leads API client");
            break;
          case "products":
            apiClient = new pipedriveV1.ProductsApi(apiConfig);
            logObjectStructure(apiClient, "products API client");
            break;
          default:
            throw new Error(`Unknown entity type: ${entityType}`);
        }
      }
    } catch (libError) {
      throw libError;
    }

    // Configure the API client with our custom adapter
    try {
      // First try to use the new dedicated function
      if (typeof configureApiClientAdapter === "function") {
        const configured = configureApiClientAdapter(apiClient);

        // If the configuration function exists but failed, try direct configuration
        if (!configured && apiClient && apiClient.axios) {
          const gasAdapter = AppLib.getGASAxiosAdapter();
          apiClient.axios.defaults.adapter = gasAdapter;
        }

        // Inspect the client to verify configuration
        if (typeof inspectApiClientAxios === "function") {
          inspectApiClientAxios(apiClient);
        }
      }
      // Fallback for direct configuration if the functions don't exist
      else if (apiClient && apiClient.axios && AppLib.getGASAxiosAdapter) {
        apiClient.axios.defaults.adapter = AppLib.getGASAxiosAdapter();
      }
    } catch (adapterError) {
      // Continue anyway, we'll handle errors during API calls
    }

    // Get header-to-field mapping
    const headerToFieldKeyMap = ensureHeaderFieldMapping(
      activeSheetName,
      entityType
    );

    // Get sync status column position
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const syncStatusColumnPos = scriptProperties.getProperty(
      twoWaySyncTrackingColumnKey
    );

    if (!syncStatusColumnPos) {
      if (!isScheduledSync) {
        SpreadsheetApp.getUi().alert(
          "Sync Tracking Column Not Found",
          "Could not find the sync tracking column. Please configure two-way sync settings again.",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }

    // Convert column letter to index
    const syncStatusColumnIndex = columnLetterToIndex(syncStatusColumnPos) - 1; // 0-based index

    // Get sheet data
    const dataRange = activeSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    // Find the ID column index
    let idColumnIndex = headers.findIndex(
      (header) =>
        header === "Pipedrive ID" ||
        header.match(/^Pipedrive .* ID$/) ||
        header === "ID"
    );

    if (idColumnIndex === -1) {
      idColumnIndex = 0; // Fallback to first column
    }

    // Collect modified rows
    const modifiedRows = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const syncStatus = row[syncStatusColumnIndex];

      // Only process rows marked as "Modified"
      if (syncStatus === "Modified") {
        // Get row ID
        let rowId = row[idColumnIndex];
        if (!rowId) continue;

        // Create data object for this row
        const updateData = {
          id: rowId,
          data: {},
          sheetRowIndex: i + 1, // Store the actual row index in the sheet (1-based for getRange)
        };        

        // Add special fields container for API v2
        // Note: Leads also need custom_fields as they inherit custom fields from deals
        if (!entityType.endsWith("Fields")) {
          updateData.data.custom_fields = {};
        }

        // Store phone and email data separately for proper formatting
        const phoneData = [];
        const emailData = [];

        // Store address components separately for proper handling
        const addressComponents = {};

        // Map column values to API fields
        for (let j = 0; j < headers.length; j++) {
          if (j === syncStatusColumnIndex) continue;

          const header = headers[j];
          const value = row[j];

          // Skip empty values
          if (value === "" || value === null || value === undefined) continue;

          // Get field key from mapping
          const fieldKey = headerToFieldKeyMap[header];
          if (!fieldKey) continue;

          // Handle different field types
          if (fieldKey === "id") {
            // Skip ID field - already handled
            continue;
          } else if (fieldKey === "email" || fieldKey.startsWith("email.")) {
            // Handle email fields
            const label = fieldKey.includes(".")
              ? fieldKey.split(".")[1]
              : "work";
            emailData.push({
              label: label,
              value: value,
              primary: label === "work" || label === "primary",
            });
          } else if (fieldKey === "phone" || fieldKey.startsWith("phone.")) {
            // Handle phone fields
            const label = fieldKey.includes(".")
              ? fieldKey.split(".")[1]
              : "work";
            phoneData.push({
              label: label,
              value: String(value), // Ensure phone is a string
              primary: label === "work" || label === "primary",
            });
          } else if (fieldKey === "participants") {
            // Handle participants field for activities
            // Store as is - will be parsed in the activities case
            updateData.data[fieldKey] = value;
          }
          // Handle address components (which are working correctly)
          else if (fieldKey.match(/^[a-f0-9]{20,}_[a-z_]+$/i)) {
            // This is an address component (e.g., custom field ID + component name)
            const parts = fieldKey.split("_");
            const fieldId = parts[0];
            const component = parts.slice(1).join("_"); // In case component name has multiple parts

            // Initialize address components for this field
            if (!addressComponents[fieldId]) {
              addressComponents[fieldId] = {};
            }

            // Store the component value
            addressComponents[fieldId][component] = String(value || ""); // Ensure it's a string
          }
          // Handle custom fields (including time range _until fields)
          else if (fieldKey.match(/^[a-f0-9]{20,}(_until)?$/i)) {
            // This is a custom field ID (possibly with _until suffix for time ranges)
            
            // Special handling for time fields to prevent timezone conversion
            let processedValue = value;
            
            // Check if this is a time field by examining the value
            if (value instanceof Date) {
              // Check if this is a time-only value (Excel epoch date)
              if (value.getFullYear() === 1899 && value.getMonth() === 11 && value.getDate() === 30) {
                // This is a time-only value, extract the time without timezone conversion
                const hours = value.getHours();
                const minutes = value.getMinutes();
                const seconds = value.getSeconds();
                processedValue = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
              }
            }
            
            // CRITICAL: Ensure we never store objects for time fields
            // If processedValue is an object with a 'value' property, extract just the value
            if (typeof processedValue === 'object' && processedValue !== null && processedValue.value !== undefined) {
              processedValue = processedValue.value;
            }
            
            updateData.data.custom_fields[fieldKey] = processedValue;
            
            // Log field detection with value

            // Log time range field detection
            if (fieldKey.endsWith("_until")) {
            }
          }
          // Handle all other fields
          else {
            // Regular field
            
            // Apply time conversion for regular fields that might contain time values
            let processedValue = value;
            if (value instanceof Date) {
              // Check if this is a time-only value (Excel epoch date)
              if (value.getFullYear() === 1899 && value.getMonth() === 11 && value.getDate() === 30) {
                // This is a time-only value, extract the time without timezone conversion
                const hours = value.getHours();
                const minutes = value.getMinutes();
                const seconds = value.getSeconds();
                processedValue = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
              }
            }
            
            updateData.data[fieldKey] = processedValue;
          }
        }

        // Add email and phone data if collected
        if (emailData.length > 0) {
          updateData.data.email = emailData;
        }

        if (phoneData.length > 0) {
          updateData.data.phone = phoneData;
        }

        // Add address components to custom fields
        for (const fieldId in addressComponents) {
          // If this field already exists in custom_fields, update it with components
          if (updateData.data.custom_fields[fieldId]) {
            // If the field is a string, convert to object first
            if (typeof updateData.data.custom_fields[fieldId] === "string") {
              const addressValue = updateData.data.custom_fields[fieldId];
              updateData.data.custom_fields[fieldId] = {
                value: addressValue || "Address", // Ensure value isn't empty
              };
            } else if (
              updateData.data.custom_fields[fieldId] === null ||
              updateData.data.custom_fields[fieldId] === undefined
            ) {
              // Initialize if null/undefined
              updateData.data.custom_fields[fieldId] = {
                value: "Address",
              };
            } else if (
              typeof updateData.data.custom_fields[fieldId] === "object" &&
              !updateData.data.custom_fields[fieldId].value
            ) {
              // Ensure value property exists
              updateData.data.custom_fields[fieldId].value = "Address";
            }

            // Add all components to the address object as strings
            for (const component in addressComponents[fieldId]) {
              updateData.data.custom_fields[fieldId][component] = String(
                addressComponents[fieldId][component] || ""
              );
            }
          }
          // Otherwise, create a new field with components
          else {
            // Create object with components
            updateData.data.custom_fields[fieldId] = addressComponents[fieldId];
            // Ensure it has a value property (required by Pipedrive)
            if (!updateData.data.custom_fields[fieldId].value) {
              updateData.data.custom_fields[fieldId].value = "Address";
            }

            // Ensure all address components are strings
            for (const key in updateData.data.custom_fields[fieldId]) {
              updateData.data.custom_fields[fieldId][key] = String(
                updateData.data.custom_fields[fieldId][key] || ""
              );
            }
          }
        }

        // Store row with email and phone data
        updateData.emailData = emailData;
        updateData.phoneData = phoneData;

        // Log the complete updateData to see what fields we have

        // Check for time range fields that might be missing their pair
        for (const fieldKey in updateData.data.custom_fields) {
          if (!fieldKey.endsWith("_until")) {
            const untilKey = fieldKey + "_until";
            // Check if we have a time range start without end
            if (!updateData.data.custom_fields[untilKey]) {
              // Look for the end time in headers
              for (let k = 0; k < headers.length; k++) {
                if (headerToFieldKeyMap[headers[k]] === untilKey && row[k]) {
                  let endValue = row[k];
                  
                  // Apply same time conversion logic for end time
                  if (endValue instanceof Date) {
                    // Check if this is a time-only value (Excel epoch date)
                    if (endValue.getFullYear() === 1899 && endValue.getMonth() === 11 && endValue.getDate() === 30) {
                      // This is a time-only value, extract the time without timezone conversion
                      const hours = endValue.getHours();
                      const minutes = endValue.getMinutes();
                      const seconds = endValue.getSeconds();
                      endValue = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
                    }
                  }
                  
                  updateData.data.custom_fields[untilKey] = endValue;
                  break;
                }
              }
            }
          }
        }

        modifiedRows.push(updateData);
      }
    }

    // Check if we have any modified rows
    if (modifiedRows.length === 0) {
      if (!suppressNoModifiedWarning && !isScheduledSync) {
        SpreadsheetApp.getUi().alert(
          "No Modified Rows",
          'No rows marked as "Modified" were found. Edit cells in rows to mark them for update.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return;
    }

    // Show confirmation for manual syncs
    if (!isScheduledSync) {
      const result = SpreadsheetApp.getUi().alert(
        "Confirm Push to Pipedrive",
        `You are about to push ${modifiedRows.length} modified row(s) to Pipedrive. Continue?`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );

      if (result !== SpreadsheetApp.getUi().Button.YES) {
        return;
      }
    }

    // Show progress toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Pushing ${modifiedRows.length} modified row(s) to Pipedrive...`,
      "Push to Pipedrive",
      30
    );

    // Track results
    let successCount = 0;
    let failureCount = 0;
    const failures = [];

    // Update each modified row
    for (let rowIndex = 0; rowIndex < modifiedRows.length; rowIndex++) {
      try {
        const rowData = modifiedRows[rowIndex];

        // Final processed payload ready to send
        let payloadToSend = rowData.data; // This would be the final processed data in your original code

        // IMPORTANT: ADD THIS CODE HERE - Process date and time fields in the payload
        const fieldDefinitions = getFieldDefinitionsMap(entityType);

        // Enhanced logging to debug time range fields

        // Check for any fields that might be time range fields
        const potentialTimeRangeFields = [];
        for (const key in rowData.data) {
          const value = rowData.data[key];
          // Check if this looks like a time value
          if (
            typeof value === "string" &&
            (value.includes(":") || value.includes("1899-12-30"))
          ) {
            potentialTimeRangeFields.push({ key, value });
          }
        }

        // Look for _until fields in headers
        const headers = activeSheet
          .getRange(1, 1, 1, activeSheet.getLastColumn())
          .getValues()[0];
        const untilHeaders = headers.filter(
          (h) => h && h.toString().toLowerCase().includes("end time")
        );

        // Check headerToFieldKeyMap for time range pairs
        const timeRangePairs = {};
        for (const [header, fieldKey] of Object.entries(headerToFieldKeyMap)) {
          if (fieldKey && fieldKey.endsWith("_until")) {
            const baseFieldKey = fieldKey.replace(/_until$/, "");
            timeRangePairs[baseFieldKey] = fieldKey;

            // Check if we have data for this end time field
            if (rowData.data[header]) {
              // Ensure this is included in the payload
              if (!payloadToSend[fieldKey]) {
                payloadToSend[fieldKey] = rowData.data[header];
              }
              if (!payloadToSend.custom_fields) {
                payloadToSend.custom_fields = {};
              }
              if (!payloadToSend.custom_fields[fieldKey]) {
                payloadToSend.custom_fields[fieldKey] = rowData.data[header];
              }
            }
          }
        }


        if (fieldDefinitions && Object.keys(fieldDefinitions).length > 0) {
          payloadToSend = processDateTimeFields(
            payloadToSend,
            rowData,
            fieldDefinitions,
            headerToFieldKeyMap
          );

          // Now ensure time range fields are properly handled
          payloadToSend = ensureTimeRangeFieldsForPipedrive(
            payloadToSend,
            rowData.data,
            headerToFieldKeyMap
          );

          // Add detailed logging of any time range fields in the final payload
          Object.keys(payloadToSend).forEach((key) => {
            if (key.endsWith("_until")) {
              const baseKey = key.replace(/_until$/, "");
            }
          });
          if (payloadToSend.custom_fields) {
            Object.keys(payloadToSend.custom_fields).forEach((key) => {
              if (key.endsWith("_until")) {
                const baseKey = key.replace(/_until$/, "");
              }
            });
          }
        }

        // Use the npm client to update the entity
        let responseBody;
        let responseCode = 200;
        let success = false;

        try {

          // Extract key data from the payload for logging
          const payloadKeys = Object.keys(payloadToSend);

          // Deep inspect custom fields for debugging
          if (payloadToSend.custom_fields) {
            // Check each custom field has correct format
            Object.keys(payloadToSend.custom_fields).forEach((fieldKey) => {
            });
          }

          // Validate payload format - ensure custom_fields is properly formatted
          if (
            payloadToSend.custom_fields &&
            typeof payloadToSend.custom_fields === "object"
          ) {
            // Create a standard custom fields object
            let customFields = {};

            // Clean up any potentially malformed fields
            Object.keys(payloadToSend.custom_fields).forEach((fieldKey) => {
              const value = payloadToSend.custom_fields[fieldKey];

              // If the value is an address object with components but missing 'value' property
              if (
                typeof value === "object" &&
                (value.hasOwnProperty("locality") ||
                  value.hasOwnProperty("route") ||
                  value.hasOwnProperty("street_number") ||
                  value.hasOwnProperty("postal_code"))
              ) {
                // Ensure it has a value property
                if (!value.value || value.value === "") {
                  value.value = "Address";
                }
              }

              // Add to the customFields object - Pipedrive expects these in a specific format
              customFields[fieldKey] = value;
            });

            // Replace the existing custom_fields with the properly formatted version
            // For Pipedrive API v1, we keep them in the custom_fields object
            payloadToSend.custom_fields = customFields;

            // Try a simpler approach - only include standard fields and a couple of custom fields
            const simplifiedPayload = {
              title: payloadToSend.title,
              name: payloadToSend.name, // Organizations use 'name' not 'title'
            };
            
            // Include email and phone fields for persons
            if (entityType === 'persons') {
              if (payloadToSend.email) {
                simplifiedPayload.email = payloadToSend.email;
              }
              if (payloadToSend.phone) {
                simplifiedPayload.phone = payloadToSend.phone;
              }
            }

            // Only include a few simple custom fields first to test
            const simpleCustomFields = {};
            Object.keys(customFields).forEach((key) => {
              // Include address fields, date/time fields (including _until fields), and simple fields
              if (
                typeof customFields[key] === "string" ||
                typeof customFields[key] === "number" ||
                key.endsWith("_until") || // Include time range end fields
                key.match(/[a-f0-9]{20,}/) || // Include all custom fields by ID pattern
                (typeof customFields[key] === "object" &&
                  customFields[key] !== null &&
                  (customFields[key].street_number ||
                    customFields[key].route ||
                    customFields[key].locality ||
                    customFields[key].postal_code))
              ) {
                simpleCustomFields[key] = customFields[key];
              }
            });

            if (Object.keys(simpleCustomFields).length > 0) {
              simplifiedPayload.custom_fields = simpleCustomFields;
            }

          }

          // Handle user_id.email format - convert to user_id if possible
          if (payloadToSend["user_id.email"]) {
            // Look up user ID by email if possible
            try {
              const userEmail = payloadToSend["user_id.email"];
              // Use a user lookup function if available, or skip this field
              if (typeof lookupUserIdByEmail === "function") {
                const userId = lookupUserIdByEmail(userEmail);
                if (userId) {
                  payloadToSend.user_id = userId;
                }
              }
              // Remove the original email field
              delete payloadToSend["user_id.email"];
            } catch (e) {
              delete payloadToSend["user_id.email"];
            }
          }

          // Handle org_id field - convert organization name to ID if needed
          if (payloadToSend.org_id && typeof payloadToSend.org_id === 'string' && isNaN(payloadToSend.org_id)) {
            
            delete payloadToSend.org_id;
          }
          
          // Convert option labels to IDs for all fields with options (for Persons, Organizations, Deals, and Leads)
          if (entityType === 'persons' || entityType === 'organizations' || entityType === 'deals' || entityType === 'leads') {
            try {
              // Get field definitions for the entity type
              const fieldDefinitions = getFieldDefinitionsMap(entityType);
              
              // Helper function to process option fields
              const processOptionField = (fieldKey, fieldValue, targetObject) => {
                const fieldDef = fieldDefinitions[fieldKey];
                
                // Check if this is an enum or set field with options
                if (fieldDef && (fieldDef.field_type === 'enum' || fieldDef.field_type === 'set') && fieldDef.options) {
                  // Create a mapping of option labels to IDs
                  const optionLabelToId = {};
                  fieldDef.options.forEach(option => {
                    if (option.label && option.id !== undefined) {
                      optionLabelToId[option.label.toLowerCase()] = option.id;
                    }
                  });
                  
                  // Process single enum field
                  if (fieldDef.field_type === 'enum' && typeof fieldValue === 'string') {
                    // Check if it's already an ID (number as string)
                    if (!isNaN(fieldValue)) {
                      // It's a number, keep it as is but convert to number
                      targetObject[fieldKey] = Number(fieldValue);
                    } else {
                      // It's a label, try to find the ID
                      const optionId = optionLabelToId[fieldValue.toLowerCase()];
                      if (optionId !== undefined) {
                        targetObject[fieldKey] = optionId;
                      } else {
                        // Label not found, could be a new option - keep as string for now
                      }
                    }
                  }
                  
                  // Process multi-select set field
                  else if (fieldDef.field_type === 'set') {
                    let processedOptions = [];
                    let valueArray = [];
                    
                    // Convert string to array if needed
                    if (typeof fieldValue === 'string') {
                      // Remove brackets and quotes, split by comma
                      const cleanValue = fieldValue.replace(/[\[\]"']/g, '');
                      valueArray = cleanValue.split(',').map(s => s.trim()).filter(s => s);
                    } else if (Array.isArray(fieldValue)) {
                      valueArray = fieldValue;
                    }
                    
                    // Process each value
                    valueArray.forEach(val => {
                      if (typeof val === 'number' || !isNaN(val)) {
                        // It's already an ID
                        processedOptions.push(Number(val));
                      } else if (typeof val === 'string') {
                        // It's a label, try to find the ID
                        const optionId = optionLabelToId[val.toLowerCase()];
                        if (optionId !== undefined) {
                          processedOptions.push(optionId);
                        } else {
                          // Label not found, could be a new option
                        }
                      }
                    });
                    
                    if (processedOptions.length > 0) {
                      targetObject[fieldKey] = processedOptions;
                    }
                  }
                }
              };
              
              // Process standard fields
              for (const fieldKey in payloadToSend) {
                if (fieldKey !== 'custom_fields' && !fieldKey.startsWith('__')) {
                  const fieldValue = payloadToSend[fieldKey];
                  processOptionField(fieldKey, fieldValue, payloadToSend);
                }
              }
              
              // Process custom fields
              if (payloadToSend.custom_fields) {
                for (const fieldKey in payloadToSend.custom_fields) {
                  const fieldValue = payloadToSend.custom_fields[fieldKey];
                  processOptionField(fieldKey, fieldValue, payloadToSend.custom_fields);
                }
              }
            } catch (e) {
              // Keep original values if conversion fails
            }
          }
          
          // Convert label names to IDs for leads
          if (entityType === 'leads' && payloadToSend.label_ids) {
            try {
              // Get lead labels
              const leadLabels = getLeadLabels();
              
              // Create a mapping of label names to IDs
              const labelNameToId = {};
              leadLabels.forEach(label => {
                if (label.name && label.id) {
                  labelNameToId[label.name.toLowerCase()] = label.id;
                }
              });
              
              // Process label_ids - could be string, array, or already formatted
              let processedLabels = [];
              
              if (typeof payloadToSend.label_ids === 'string') {
                // Could be comma-separated names or IDs, or a stringified array
                const labelString = payloadToSend.label_ids.replace(/[\[\]"']/g, '');
                const labelParts = labelString.split(',').map(s => s.trim()).filter(s => s);
                
                labelParts.forEach(part => {
                  // Check if it's a UUID (label ID)
                  if (part.match(/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i)) {
                    processedLabels.push(part);
                  } else {
                    // It's a name, try to find the ID
                    const labelId = labelNameToId[part.toLowerCase()];
                    if (labelId) {
                      processedLabels.push(labelId);
                    } else {
                    }
                  }
                });
              } else if (Array.isArray(payloadToSend.label_ids)) {
                // Already an array, process each element
                payloadToSend.label_ids.forEach(item => {
                  if (typeof item === 'string') {
                    // Check if it's a UUID (label ID)
                    if (item.match(/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i)) {
                      processedLabels.push(item);
                    } else {
                      // It's a name, try to find the ID
                      const labelId = labelNameToId[item.toLowerCase()];
                      if (labelId) {
                        processedLabels.push(labelId);
                      } else {
                      }
                    }
                  } else {
                    processedLabels.push(item);
                  }
                });
              }
              
              if (processedLabels.length > 0) {
                payloadToSend.label_ids = processedLabels;
              } else {
                delete payloadToSend.label_ids;
              }
            } catch (e) {
              // Keep original value if conversion fails
            }
          }
          
          // Convert label_ids from string to array if needed (for other entity types)
          else if (payloadToSend.label_ids && typeof payloadToSend.label_ids === 'string') {
            try {
              // Remove brackets and parse
              const labelString = payloadToSend.label_ids.replace(/[\[\]]/g, '');
              if (labelString) {
                payloadToSend.label_ids = labelString.split(',').map(id => parseInt(id.trim()));
              } else {
                delete payloadToSend.label_ids;
              }
            } catch (e) {
              delete payloadToSend.label_ids;
            }
          }

          // Remove read-only fields that cannot be updated
          const readOnlyFields = [
            "add_time",
            "update_time",
            "id",
            "creator_user_id",
            "creator_id", // Read-only for leads
            "related_open_deals_count", // This is also read-only
            "related_closed_deals_count",
            "related_won_deals_count",
            "related_lost_deals_count",
            "cc_email", // System-generated for leads
            "source_name", // May be read-only
            "origin", // May be read-only
            "origin_id" // May be read-only
          ];
          readOnlyFields.forEach((field) => {
            if (payloadToSend[field] !== undefined) {
              delete payloadToSend[field];
            }
          });

          // Log the final modified payload

          // Show some key parameters for the API call
          const requestParams = {
            id: rowData.id,
            entityType: entityType,
            apiBasePath: apiClient.basePath,
            accessToken: accessToken
              ? accessToken.substring(0, 5) + "..."
              : "undefined",
          };

          switch (entityType) {
            case "deals":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties =
                  PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty(
                  "PIPEDRIVE_ACCESS_TOKEN"
                );

                if (!pipedriveToken) {
                  throw new Error(
                    "Pipedrive API token not found in script properties"
                  );
                }

                // Get subdomain from properties
                const subdomain =
                  scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";

                // Get field definitions for processing
                const fieldDefinitions = getFieldDefinitionsMap(entityType);
                
                // Prepare payload
                const finalPayload = {};

                // Process standard fields with type conversions
                const standardFields = {};
                Object.keys(payloadToSend).forEach((key) => {
                  if (key !== "custom_fields") {
                    standardFields[key] = payloadToSend[key];
                  }
                });
                
                // Apply field type processing to standard fields
                const processedStandardFields = processFieldTypes(standardFields, fieldDefinitions, entityType);
                Object.assign(finalPayload, processedStandardFields);

                // Process and move custom fields to top level
                if (payloadToSend.custom_fields) {
                  // Apply field type processing to custom fields
                  const processedCustomFields = processFieldTypes(payloadToSend.custom_fields, fieldDefinitions, entityType);
                  
                  // Special handling for address fields in deals
                  Object.keys(processedCustomFields).forEach((key) => {
                    const value = processedCustomFields[key];

                    // Special handling for address objects
                    if (
                      value &&
                      typeof value === "object" &&
                      (value.locality ||
                        value.route ||
                        value.street_number ||
                        value.postal_code)
                    ) {
                      // Fetch the current address object using getCurrentAddressData
                      let currentAddress = {};
                      try {
                        currentAddress =
                          getCurrentAddressData("deals", rowData.id, key) || {};
                      } catch (e) {
                      }

                      // Merge updated components into the original address object
                      const mergedAddress = {
                        ...currentAddress,
                        ...value,
                      };

                      // Rebuild the full address string from all components
                      const addressParts = [];
                      if (mergedAddress.street_number)
                        addressParts.push(mergedAddress.street_number);
                      if (mergedAddress.route) {
                        if (addressParts.length > 0) {
                          addressParts[0] =
                            addressParts[0] + " " + mergedAddress.route;
                        } else {
                          addressParts.push(mergedAddress.route);
                        }
                      }
                      if (mergedAddress.locality)
                        addressParts.push(mergedAddress.locality);
                      if (mergedAddress.admin_area_level_1)
                        addressParts.push(mergedAddress.admin_area_level_1);
                      if (mergedAddress.postal_code)
                        addressParts.push(mergedAddress.postal_code);
                      // Optionally add more components (country, etc.) as needed

                      mergedAddress.value = addressParts.join(", ");

                      // Set merged address to the payload in all required places
                      finalPayload[key] = mergedAddress.value;
                      if (!finalPayload.custom_fields)
                        finalPayload.custom_fields = {};
                      finalPayload.custom_fields[key] = {
                        ...mergedAddress,
                      };

                      // Add individual components as separate fields if required by Pipedrive
                      if (mergedAddress.street_number)
                        finalPayload[`${key}_street_number`] =
                          mergedAddress.street_number;
                      if (mergedAddress.route)
                        finalPayload[`${key}_route`] = mergedAddress.route;
                      if (mergedAddress.locality)
                        finalPayload[`${key}_locality`] =
                          mergedAddress.locality;
                      if (mergedAddress.postal_code)
                        finalPayload[`${key}_postal_code`] =
                          mergedAddress.postal_code;
                      if (mergedAddress.admin_area_level_1)
                        finalPayload[`${key}_admin_area_level_1`] =
                          mergedAddress.admin_area_level_1;
                      if (mergedAddress.admin_area_level_2)
                        finalPayload[`${key}_admin_area_level_2`] =
                          mergedAddress.admin_area_level_2;
                      if (mergedAddress.country)
                        finalPayload[`${key}_country`] = mergedAddress.country;
                      if (mergedAddress.sublocality)
                        finalPayload[`${key}_sublocality`] =
                          mergedAddress.sublocality;
                      if (mergedAddress.subpremise)
                        finalPayload[`${key}_subpremise`] =
                          mergedAddress.subpremise;
                      if (mergedAddress.formatted_address)
                        finalPayload[`${key}_formatted_address`] =
                          mergedAddress.formatted_address;

                    } else {
                      // For all other custom fields, add them directly to the final payload
                      finalPayload[key] = value;
                    }
                  });
                  
                }

                // Handle linked field name-to-ID conversions for deals
                // Handle person_id field - convert person name to ID if provided as string
                if (finalPayload.person_id && typeof finalPayload.person_id === 'string' && isNaN(finalPayload.person_id)) {
                  const nameToIdMap = searchEntitiesByName('persons', [finalPayload.person_id]);
                  if (nameToIdMap[finalPayload.person_id]) {
                    finalPayload.person_id = nameToIdMap[finalPayload.person_id];
                  } else {
                    delete finalPayload.person_id;
                  }
                }
                
                // Handle org_id field - convert organization name to ID if provided as string
                if (finalPayload.org_id && typeof finalPayload.org_id === 'string' && isNaN(finalPayload.org_id)) {
                  const nameToIdMap = searchEntitiesByName('organizations', [finalPayload.org_id]);
                  if (nameToIdMap[finalPayload.org_id]) {
                    finalPayload.org_id = nameToIdMap[finalPayload.org_id];
                  } else {
                    delete finalPayload.org_id;
                  }
                }
                
                // Handle owner_id field - convert user name to ID if provided as string  
                if (finalPayload.owner_id && typeof finalPayload.owner_id === 'string' && isNaN(finalPayload.owner_id)) {
                  const nameToIdMap = searchEntitiesByName('users', [finalPayload.owner_id]);
                  if (nameToIdMap[finalPayload.owner_id]) {
                    finalPayload.owner_id = nameToIdMap[finalPayload.owner_id];
                  } else {
                    delete finalPayload.owner_id;
                  }
                }
                
                // Handle stage_id field - convert stage name to ID if provided as string
                if (finalPayload.stage_id && typeof finalPayload.stage_id === 'string' && isNaN(finalPayload.stage_id)) {
                  const nameToIdMap = searchEntitiesByName('stages', [finalPayload.stage_id]);
                  if (nameToIdMap[finalPayload.stage_id]) {
                    finalPayload.stage_id = nameToIdMap[finalPayload.stage_id];
                  } else {
                    delete finalPayload.stage_id;
                  }
                }
                
                // Handle pipeline_id field - convert pipeline name to ID if provided as string
                if (finalPayload.pipeline_id && typeof finalPayload.pipeline_id === 'string' && isNaN(finalPayload.pipeline_id)) {
                  const nameToIdMap = searchEntitiesByName('pipelines', [finalPayload.pipeline_id]);
                  if (nameToIdMap[finalPayload.pipeline_id]) {
                    finalPayload.pipeline_id = nameToIdMap[finalPayload.pipeline_id];
                  } else {
                    delete finalPayload.pipeline_id;
                  }
                }

                // Log the final payload

                // Special handling for time range fields
                if (
                  finalPayload.__hasTimeRangeFields ||
                  Object.keys(finalPayload).some((key) =>
                    key.endsWith("_until")
                  )
                ) {

                  // First, build a map of time range pairs
                  const timeRangePairs = {};

                  // Look for _until fields in custom_fields
                  if (finalPayload.custom_fields) {
                    Object.keys(finalPayload.custom_fields).forEach((key) => {
                      if (key.endsWith("_until")) {
                        const baseKey = key.replace(/_until$/, "");
                        timeRangePairs[baseKey] = key;
                      }
                    });
                  }

                  // Also look for _until fields at root level
                  Object.keys(finalPayload).forEach((key) => {
                    if (key.endsWith("_until")) {
                      const baseKey = key.replace(/_until$/, "");
                      timeRangePairs[baseKey] = key;
                    }
                  });

                  // Now ensure both parts are present in both root and custom_fields
                  for (const baseKey in timeRangePairs) {
                    const untilKey = timeRangePairs[baseKey];

                    // Get values from wherever they are
                    let startValue = finalPayload[baseKey];
                    let endValue = finalPayload[untilKey];

                    if (
                      !startValue &&
                      finalPayload.custom_fields &&
                      finalPayload.custom_fields[baseKey]
                    ) {
                      startValue = finalPayload.custom_fields[baseKey];
                    }

                    if (
                      !endValue &&
                      finalPayload.custom_fields &&
                      finalPayload.custom_fields[untilKey]
                    ) {
                      endValue = finalPayload.custom_fields[untilKey];
                    }


                    // Now make sure both values are in both places
                    if (startValue) {
                      finalPayload[baseKey] = startValue;
                      if (!finalPayload.custom_fields)
                        finalPayload.custom_fields = {};
                      finalPayload.custom_fields[baseKey] = startValue;
                    }

                    if (endValue) {
                      finalPayload[untilKey] = endValue;
                      if (!finalPayload.custom_fields)
                        finalPayload.custom_fields = {};
                      finalPayload.custom_fields[untilKey] = endValue;
                    }
                  }
                }

                // CRITICAL: Final check for time range fields before API call
                if (
                  finalPayload.__preserveTimeRangePairs &&
                  finalPayload.__timeRangePairs
                ) {

                  for (const baseKey in finalPayload.__timeRangePairs) {
                    const pair = finalPayload.__timeRangePairs[baseKey];
                    // Ensure both start and end values are in the root payload
                    if (pair.startValue) {
                      finalPayload[pair.startKey] = pair.startValue;
                    }
                    if (pair.endValue) {
                      finalPayload[pair.endKey] = pair.endValue;
                    }
                  }

                  // Clean up our temporary fields
                  delete finalPayload.__preserveTimeRangePairs;
                  delete finalPayload.__timeRangePairs;
                }

                // Enhanced time range field detection
                let timeRangeFieldsDetected =
                  finalPayload.__hasTimeRangeFields === true;
                let timeRangePairs = {};

                // More aggressive time range detection - check for any fields ending with _until
                // First collect all possible time range fields
                for (const key in finalPayload) {
                  if (key.endsWith("_until")) {
                    const baseKey = key.replace(/_until$/, "");
                    timeRangeFieldsDetected = true;
                    timeRangePairs[baseKey] = {
                      startKey: baseKey,
                      endKey: key,
                      startValue: finalPayload[baseKey],
                      endValue: finalPayload[key],
                    };
                  }
                }

                // Also check all custom fields for time range endings
                if (finalPayload.custom_fields) {
                  for (const key in finalPayload.custom_fields) {
                    if (key.endsWith("_until")) {
                      const baseKey = key.replace(/_until$/, "");

                      // If we already detected this pair at root level, update with any values from custom_fields
                      if (timeRangePairs[baseKey]) {
                        if (
                          !timeRangePairs[baseKey].startValue &&
                          finalPayload.custom_fields[baseKey]
                        ) {
                          timeRangePairs[baseKey].startValue =
                            finalPayload.custom_fields[baseKey];
                        }
                        if (
                          !timeRangePairs[baseKey].endValue &&
                          finalPayload.custom_fields[key]
                        ) {
                          timeRangePairs[baseKey].endValue =
                            finalPayload.custom_fields[key];
                        }
                      } else {
                        // This is a new time range pair
                        timeRangeFieldsDetected = true;
                        timeRangePairs[baseKey] = {
                          startKey: baseKey,
                          endKey: key,
                          startValue: finalPayload.custom_fields[baseKey],
                          endValue: finalPayload.custom_fields[key],
                        };
                      }
                    }
                  }
                }

                // Check any field with time or hour in the name that might be time fields
                if (!timeRangeFieldsDetected) {
                  for (const key in finalPayload) {
                    if (
                      (key.includes("time_") || key.includes("hour")) &&
                      typeof finalPayload[key] === "string" &&
                      finalPayload[key].match(/\d{1,2}:\d{2}/)
                    ) {
                      // This looks like a time field but doesn't follow the _until pattern
                      // Check if there's a corresponding end field
                      const possibleEndKeys = [
                        `${key}_until`,
                        `${key}_end`,
                        `${key}_to`,
                      ];

                      for (const endKey of possibleEndKeys) {
                        if (
                          finalPayload[endKey] ||
                          (finalPayload.custom_fields &&
                            finalPayload.custom_fields[endKey])
                        ) {
                          timeRangeFieldsDetected = true;
                          const endValue =
                            finalPayload[endKey] ||
                            (finalPayload.custom_fields &&
                              finalPayload.custom_fields[endKey]);
                          break;
                        }
                      }

                      if (timeRangeFieldsDetected) {
                        break;
                      }
                    }
                  }
                }

                // Ensure all time range pairs have both start and end values
                if (Object.keys(timeRangePairs).length > 0) {
                  // Flag that we have time range fields to handle
                  finalPayload.__hasTimeRangeFields = true;
                  finalPayload.__timeRangePairs = timeRangePairs;

                  // Ensure both parts of each pair are properly formatted and present
                  for (const baseKey in timeRangePairs) {
                    const pair = timeRangePairs[baseKey];

                    // If we have a value for one part but not the other, copy it
                    if (pair.startValue && !pair.endValue) {
                      // If we have start but no end, use start as end
                      pair.endValue = pair.startValue;

                      // Set in both places
                      finalPayload[pair.endKey] = pair.endValue;
                      if (!finalPayload.custom_fields)
                        finalPayload.custom_fields = {};
                      finalPayload.custom_fields[pair.endKey] = pair.endValue;

                    }

                    if (!pair.startValue && pair.endValue) {
                      // If we have end but no start, use end as start
                      pair.startValue = pair.endValue;

                      // Set in both places
                      finalPayload[pair.startKey] = pair.startValue;
                      if (!finalPayload.custom_fields)
                        finalPayload.custom_fields = {};
                      finalPayload.custom_fields[pair.startKey] =
                        pair.startValue;

                    }
                  }

                }

                // Use direct API for deals with time range fields
                if (timeRangeFieldsDetected) {
                  try {
                    // Get the field definitions map for deals
                    const fieldDefinitions = getFieldDefinitionsMap(entityType);

                    // Call updateDealDirect with field definitions
                    const directResponse = updateDealDirect(
                      Number(rowData.id),
                      finalPayload,
                      pipedriveToken,
                      `https://${subdomain}.pipedrive.com/api/v1`,
                      fieldDefinitions // Pass field definitions
                    );

                    // Process the response similar to your existing code
                    responseCode = directResponse.responseCode || 200;
                    const directResponseText = JSON.stringify(directResponse);

                    // Parse the response
                    responseBody = directResponse;
                    success = responseBody && responseBody.success === true;

                    // Skip the regular API call
                    break;
                  } catch (dealError) {
                    responseBody = {
                      error: dealError.message,
                    };
                    success = false;
                    break;
                  }
                }

                // Construct URL and options
                const apiUrl = `https://${subdomain}.pipedrive.com/v1/deals/${Number(
                  rowData.id
                )}`;
                const options = {
                  method: "PUT",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${pipedriveToken}`,
                  },
                  payload: JSON.stringify(finalPayload),
                  muteHttpExceptions: true,
                };

                // Make the request
                const response = UrlFetchApp.fetch(apiUrl, options);
                responseCode = response.getResponseCode();
                const responseText = response.getContentText();

                // Parse the response
                responseBody = JSON.parse(responseText);
                success = responseBody && responseBody.success === true;
              } catch (dealError) {
                responseBody = {
                  error: dealError.message,
                };
                success = false;
              }
              break;

            case "persons":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties = PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
                
                if (!pipedriveToken) {
                  throw new Error("Pipedrive API token not found in script properties");
                }
                
                // Get subdomain from properties
                const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";
                
                // Use direct API call to avoid URL constructor issue
                
                // Get field definitions for processing
                const fieldDefinitions = getFieldDefinitionsMap(entityType);
                
                // Prepare final payload - move custom fields to top level for API v1
                const finalPayload = {};
                
                // Process standard fields with type conversions
                const standardFields = {};
                Object.keys(payloadToSend).forEach((key) => {
                  if (key !== "custom_fields") {
                    standardFields[key] = payloadToSend[key];
                  }
                });
                
                // Apply field type processing to standard fields
                const processedStandardFields = processFieldTypes(standardFields, fieldDefinitions, entityType);
                Object.assign(finalPayload, processedStandardFields);
                
                // Process and move custom fields to top level
                if (payloadToSend.custom_fields) {
                  // Apply field type processing to custom fields
                  const processedCustomFields = processFieldTypes(payloadToSend.custom_fields, fieldDefinitions, entityType);
                  
                  // Move to top level for API v1
                  Object.keys(processedCustomFields).forEach((key) => {
                    finalPayload[key] = processedCustomFields[key];
                  });
                  
                }
                
                // Handle linked field name-to-ID conversions for persons
                // Handle org_id field - convert organization name to ID if provided as string
                if (finalPayload.org_id && typeof finalPayload.org_id === 'string' && isNaN(finalPayload.org_id)) {
                  const nameToIdMap = searchEntitiesByName('organizations', [finalPayload.org_id]);
                  if (nameToIdMap[finalPayload.org_id]) {
                    finalPayload.org_id = nameToIdMap[finalPayload.org_id];
                  } else {
                    delete finalPayload.org_id;
                  }
                }
                
                // Handle owner_id field - convert user name to ID if provided as string
                if (finalPayload.owner_id && typeof finalPayload.owner_id === 'string' && isNaN(finalPayload.owner_id)) {
                  const nameToIdMap = searchEntitiesByName('users', [finalPayload.owner_id]);
                  if (nameToIdMap[finalPayload.owner_id]) {
                    finalPayload.owner_id = nameToIdMap[finalPayload.owner_id];
                  } else {
                    delete finalPayload.owner_id;
                  }
                }
                
                // Log the final payload
                
                // Construct URL and options
                const apiUrl = `https://${subdomain}.pipedrive.com/v1/persons/${Number(rowData.id)}`;
                const options = {
                  method: "PUT",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${pipedriveToken}`,
                  },
                  payload: JSON.stringify(finalPayload),
                  muteHttpExceptions: true,
                };
                
                
                // Make the request
                const response = UrlFetchApp.fetch(apiUrl, options);
                responseCode = response.getResponseCode();
                const responseText = response.getContentText();
                
                // Parse the response
                responseBody = JSON.parse(responseText);
                success = responseBody && responseBody.success === true;
                
              } catch (personError) {
                responseBody = {
                  error: personError.message,
                };
                success = false;
              }
              break;

            case "organizations":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties = PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
                
                if (!pipedriveToken) {
                  throw new Error("Pipedrive API token not found in script properties");
                }
                
                // Get subdomain from properties
                const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";
                
                // Get field definitions for processing
                const fieldDefinitions = getFieldDefinitionsMap(entityType);
                
                // Use direct API call to avoid URL constructor issue
                
                // Prepare final payload - move custom fields to top level for API v1
                const finalPayload = {};
                
                // Process standard fields with type conversions
                const standardFields = {};
                Object.keys(payloadToSend).forEach((key) => {
                  if (key !== "custom_fields") {
                    standardFields[key] = payloadToSend[key];
                  }
                });
                
                // Apply field type processing to standard fields
                const processedStandardFields = processFieldTypes(standardFields, fieldDefinitions, entityType);
                Object.assign(finalPayload, processedStandardFields);
                
                // Process and move custom fields to top level
                if (payloadToSend.custom_fields) {
                  // Apply field type processing to custom fields
                  const processedCustomFields = processFieldTypes(payloadToSend.custom_fields, fieldDefinitions, entityType);
                  
                  // Move to top level for API v1
                  Object.keys(processedCustomFields).forEach((key) => {
                    finalPayload[key] = processedCustomFields[key];
                  });
                  
                }
                
                // Handle linked field name-to-ID conversions for organizations
                // Handle owner_id field - convert user name to ID if provided as string
                if (finalPayload.owner_id && typeof finalPayload.owner_id === 'string' && isNaN(finalPayload.owner_id)) {
                  const nameToIdMap = searchEntitiesByName('users', [finalPayload.owner_id]);
                  if (nameToIdMap[finalPayload.owner_id]) {
                    finalPayload.owner_id = nameToIdMap[finalPayload.owner_id];
                  } else {
                    delete finalPayload.owner_id;
                  }
                }
                
                // Log the final payload
                
                // Construct URL and options
                const apiUrl = `https://${subdomain}.pipedrive.com/v1/organizations/${Number(rowData.id)}`;
                const options = {
                  method: "PUT",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${pipedriveToken}`,
                  },
                  payload: JSON.stringify(finalPayload),
                  muteHttpExceptions: true,
                };
                
                
                // Make the request
                const response = UrlFetchApp.fetch(apiUrl, options);
                responseCode = response.getResponseCode();
                const responseText = response.getContentText();
                
                // Parse the response
                responseBody = JSON.parse(responseText);
                success = responseBody && responseBody.success === true;
                
              } catch (orgError) {
                responseBody = {
                  error: orgError.message,
                };
                success = false;
              }
              break;

            case "activities":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties = PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
                
                if (!pipedriveToken) {
                  throw new Error("Pipedrive API token not found in script properties");
                }
                
                // Get subdomain from properties
                const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";
                
                // Use direct API call to avoid URL constructor issue
                
                // Get field definitions for processing
                const fieldDefinitions = getFieldDefinitionsMap(entityType);
                
                // Prepare final payload - move custom fields to top level for API v1
                const finalPayload = {};
                
                // Process standard fields with type conversions
                const standardFields = {};
                Object.keys(payloadToSend).forEach((key) => {
                  if (key !== "custom_fields") {
                    standardFields[key] = payloadToSend[key];
                  }
                });
                
                // Apply field type processing to standard fields
                const processedStandardFields = processFieldTypes(standardFields, fieldDefinitions, entityType);
                Object.assign(finalPayload, processedStandardFields);
                
                // Special handling for participants field
                if (finalPayload.participants && typeof finalPayload.participants === 'string') {
                  try {
                    // First try to parse as JSON array string
                    const parsed = JSON.parse(finalPayload.participants);
                    if (Array.isArray(parsed)) {
                      finalPayload.participants = parsed;
                    }
                  } catch (e) {
                    // If not JSON, check if it's comma-separated values
                    const participantValues = finalPayload.participants.split(',').map(v => v.trim()).filter(v => v);
                    
                    if (participantValues.length > 0) {
                      // Check if all values are numeric (person IDs)
                      const allNumeric = participantValues.every(v => /^\d+$/.test(v));
                      
                      if (allNumeric) {
                        // Convert numeric IDs to Pipedrive format
                        finalPayload.participants = participantValues.map((id, index) => ({
                          person_id: parseInt(id),
                          primary_flag: index === 0 // First one is primary
                        }));
                      } else {
                        // Values contain names - need to look up person IDs
                        
                        // Search for person IDs by name
                        const nameToIdMap = searchPersonsByName(participantValues);
                        
                        // Convert found persons to participants array
                        const participants = [];
                        for (let i = 0; i < participantValues.length; i++) {
                          const name = participantValues[i];
                          const personId = nameToIdMap[name];
                          
                          if (personId) {
                            participants.push({
                              person_id: personId,
                              primary_flag: i === 0 // First one is primary
                            });
                          } else {
                          }
                        }
                        
                        if (participants.length > 0) {
                          finalPayload.participants = participants;
                        } else {
                          delete finalPayload.participants;
                        }
                      }
                    } else {
                      delete finalPayload.participants;
                    }
                  }
                }
                
                // Handle other linked fields for activities
                
                // Handle person_id field
                if (finalPayload.person_id && typeof finalPayload.person_id === 'string' && isNaN(finalPayload.person_id)) {
                  const nameToIdMap = searchEntitiesByName('persons', [finalPayload.person_id]);
                  if (nameToIdMap[finalPayload.person_id]) {
                    finalPayload.person_id = nameToIdMap[finalPayload.person_id];
                  } else {
                    delete finalPayload.person_id;
                  }
                }
                
                // Handle org_id field
                if (finalPayload.org_id && typeof finalPayload.org_id === 'string' && isNaN(finalPayload.org_id)) {
                  const nameToIdMap = searchEntitiesByName('organizations', [finalPayload.org_id]);
                  if (nameToIdMap[finalPayload.org_id]) {
                    finalPayload.org_id = nameToIdMap[finalPayload.org_id];
                  } else {
                    delete finalPayload.org_id;
                  }
                }
                
                // Handle deal_id field
                if (finalPayload.deal_id && typeof finalPayload.deal_id === 'string' && isNaN(finalPayload.deal_id)) {
                  const nameToIdMap = searchEntitiesByName('deals', [finalPayload.deal_id]);
                  if (nameToIdMap[finalPayload.deal_id]) {
                    finalPayload.deal_id = nameToIdMap[finalPayload.deal_id];
                  } else {
                    delete finalPayload.deal_id;
                  }
                }
                
                // Process and move custom fields to top level
                if (payloadToSend.custom_fields) {
                  // Apply field type processing to custom fields
                  const processedCustomFields = processFieldTypes(payloadToSend.custom_fields, fieldDefinitions, entityType);
                  
                  // Move to top level for API v1
                  Object.keys(processedCustomFields).forEach((key) => {
                    finalPayload[key] = processedCustomFields[key];
                  });
                  
                }
                
                // Log the final payload
                
                // Construct URL and options
                const apiUrl = `https://${subdomain}.pipedrive.com/v1/activities/${Number(rowData.id)}`;
                const options = {
                  method: "PUT",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${pipedriveToken}`,
                  },
                  payload: JSON.stringify(finalPayload),
                  muteHttpExceptions: true,
                };
                
                
                // Make the request
                const response = UrlFetchApp.fetch(apiUrl, options);
                responseCode = response.getResponseCode();
                const responseText = response.getContentText();
                
                // Parse the response
                responseBody = JSON.parse(responseText);
                success = responseBody && responseBody.success === true;
                
              } catch (activityError) {
                responseBody = {
                  error: activityError.message,
                };
                success = false;
              }
              break;

            case "leads":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties = PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
                
                if (!pipedriveToken) {
                  throw new Error("Pipedrive API token not found in script properties");
                }
                
                // Get subdomain from properties
                const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";
                
                // Use direct API call for leads
                
                // Prepare final payload - move custom fields to top level for API v1
                const finalPayload = {};
                
                // Copy all non-custom fields with type conversions
                Object.keys(payloadToSend).forEach((key) => {
                  if (key !== "custom_fields") {
                    let value = payloadToSend[key];
                    
                    // Convert boolean strings to actual booleans for leads
                    if (entityType === 'leads' && (key === 'is_archived' || key === 'was_seen')) {
                      if (typeof value === 'string') {
                        value = value.toLowerCase() === 'yes' || value.toLowerCase() === 'true';
                      }
                    }
                    
                    // Convert visible_to text to number for leads
                    if (entityType === 'leads' && key === 'visible_to' && typeof value === 'string') {
                      const visibilityMap = {
                        'owner only': '1',
                        'owner & followers': '1',
                        "owner's visibility group": '3',
                        "owner's visibility group and sub-groups": '5',
                        'entire company': '7',
                        'all users': '3' // For Essential/Advanced plans
                      };
                      const mappedValue = visibilityMap[value.toLowerCase()];
                      if (mappedValue) {
                        value = mappedValue;
                      }
                    }
                    
                    finalPayload[key] = value;
                  }
                });
                
                // Move custom fields to top level with date formatting
                if (payloadToSend.custom_fields) {
                  // Get field definitions to check field types
                  const fieldDefinitions = getFieldDefinitionsMap(entityType);
                  
                  Object.keys(payloadToSend.custom_fields).forEach((key) => {
                    let value = payloadToSend.custom_fields[key];
                    
                    // Check if this is a date field and format it
                    const fieldDef = fieldDefinitions[key];
                    if (fieldDef && fieldDef.field_type === 'date' && value) {
                      // Convert ISO date string to YYYY-MM-DD format
                      if (typeof value === 'string' && value.includes('T')) {
                        value = value.split('T')[0];
                      }
                    }
                    
                    // Check if this is a phone field and ensure it's a string
                    if (fieldDef && fieldDef.field_type === 'phone' && value !== null && value !== undefined) {
                      value = String(value);
                    }
                    
                    // For Pipedrive API v1, custom fields should be at top level
                    finalPayload[key] = value;
                  });
                }
                
                // Log the final payload
                
                // Construct URL and options - leads use PATCH method
                const apiUrl = `https://${subdomain}.pipedrive.com/v1/leads/${String(rowData.id)}`;
                const options = {
                  method: "PATCH",
                  headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${pipedriveToken}`,
                  },
                  payload: JSON.stringify(finalPayload),
                  muteHttpExceptions: true,
                };
                
                
                // Make the request
                const response = UrlFetchApp.fetch(apiUrl, options);
                responseCode = response.getResponseCode();
                const responseText = response.getContentText();
                
                // Parse the response
                responseBody = JSON.parse(responseText);
                success = responseBody && responseBody.success === true;
                
              } catch (leadError) {
                responseBody = {
                  error: leadError.message,
                };
                success = false;
              }
              break;

            case "products":
              try {
                // Get the Pipedrive OAuth token from script properties
                const scriptProperties = PropertiesService.getScriptProperties();
                const pipedriveToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");
                
                if (!pipedriveToken) {
                  throw new Error("Pipedrive API token not found in script properties");
                }
                
                // Get subdomain from properties
                const subdomain = scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") || "api";
                
                // Get field definitions for products
                let fieldDefinitions = {};
                try {
                  fieldDefinitions = getEntityFields('products');
                } catch (fieldErr) {
                }
                
                // Use direct API call to avoid URL constructor issue
                
                const directResponse = updateProductDirect(
                  rowData.id,
                  payloadToSend,
                  pipedriveToken,
                  `https://${subdomain}.pipedrive.com/v1`,
                  fieldDefinitions
                );
                
                // Process the response
                responseCode = directResponse.responseCode || 200;
                responseBody = directResponse;
                success = responseBody && responseBody.success === true;
                
              } catch (prodError) {
                responseBody = {
                  error: prodError.message,
                };
                success = false;
              }
              break;

            default:
              throw new Error(`Unknown entity type: ${entityType}`);
          }

        } catch (apiError) {
          success = false;
          responseCode = apiError.status || 500;

          // Better error handling with more details
          if (apiError.response) {
            // The request was made and the server responded with a status code
            responseBody = apiError.response.data || {
              error: "API error with response",
            };
          } else if (apiError.request) {
            // The request was made but no response was received
            responseBody = {
              error: "No response from server",
            };
          } else {
            // Something happened in setting up the request that triggered an Error
            responseBody = {
              error: apiError.message,
            };
          }
        }

        // Handle response
        if (success) {
          // Update was successful
          successCount++;

          // Update the cell status to "Synced"
          const row = rowData.sheetRowIndex; // Use the stored sheet row index
          
          // Validate that we have a valid row index
          if (!row || row < 1) {
            console.error(`Invalid row index for ID ${rowData.id}: ${row}`);
            continue;
          }
                    
          const statusCell = activeSheet.getRange(
            row,
            syncStatusColumnIndex + 1
          ); // +1 for 1-based sheet indexes
          statusCell.setValue("Synced");
          statusCell.setBackground("#E6F4EA").setFontColor("#137333");
          statusCell.clearNote();
        } else {
          // Update failed
          failureCount++;

          // Get error message
          let errorMessage = "Unknown error";
          if (responseBody.error) {
            errorMessage = responseBody.error;
          } else if (responseBody.message) {
            errorMessage = responseBody.message;
          } else if (responseBody.errors && responseBody.errors.length > 0) {
            errorMessage = responseBody.errors[0].message || "API error";
          } else if (responseBody.data && responseBody.data.errors) {
            // Handle nested errors
            const errorData = [];
            for (const field in responseBody.data.errors) {
              errorData.push(`${field}: ${responseBody.data.errors[field]}`);
            }
            errorMessage = errorData.join("; ");
          }


          // Store failure details
          failures.push({
            id: rowData.id,
            error: errorMessage,
            row: rowData.sheetRowIndex,
          });

          // Update cell status to "Error"
          const row = rowData.sheetRowIndex; // Use the stored sheet row index
          const statusCell = activeSheet.getRange(
            row,
            syncStatusColumnIndex + 1
          );
          statusCell.setValue("Error");
          statusCell.setBackground("#FCE8E6").setFontColor("#D93025");
        }
      } catch (error) {
        // Handle exceptions
        failureCount++;

        failures.push({
          id: modifiedRows[rowIndex].id,
          error: error.message,
          row: modifiedRows[rowIndex].sheetRowIndex,
        });

        // Update cell status to "Error"
        const row = modifiedRows[rowIndex].sheetRowIndex; // Use the stored sheet row index
        const statusCell = activeSheet.getRange(row, syncStatusColumnIndex + 1);
        statusCell.setValue("Error");
        statusCell.setBackground("#FCE8E6").setFontColor("#D93025");
      }
    }

    // Show completion message with diagnostic info
    if (failureCount > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Completed with ${successCount} success(es) and ${failureCount} failure(s)`,
        "Push to Pipedrive",
        5
      );

      if (!isScheduledSync) {
        let errorMessage = "The following errors occurred:\n\n";
        for (let i = 0; i < Math.min(failures.length, 5); i++) {
          errorMessage += `- Row ${failures[i].row}: ${failures[i].error}\n`;
        }

        if (failures.length > 5) {
          errorMessage += `\n... and ${
            failures.length - 5
          } more errors. See cell notes for details.`;
        }

        SpreadsheetApp.getUi().alert(
          "Errors Occurred",
          errorMessage,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Successfully pushed ${successCount} row(s) to Pipedrive`,
        "Push to Pipedrive",
        5
      );
    }

    return {
      success: successCount,
      failures: failureCount,
      details: failures,
    };
  } catch (error) {
    // Handle overall function errors

    if (!isScheduledSync) {
      SpreadsheetApp.getUi().alert(
        "Error",
        `An error occurred while pushing changes to Pipedrive: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

    return {
      success: 0,
      failures: 0,
      error: error.message,
    };
  }
}

/**
 * Filters out read-only fields from the data before sending to Pipedrive API
 * @param {Object} data - The data object to filter
 * @param {string} entityType - The entity type
 * @return {Object} Filtered data object
 */
function filterReadOnlyFields(data, entityType) {
  if (!data) return data;

  const filteredData = {};

  // Copy custom_fields to the filtered data if it exists
  if (data.custom_fields) {
    filteredData.custom_fields = JSON.parse(JSON.stringify(data.custom_fields));

    // Format custom fields to match Pipedrive API requirements
    formatCustomFields(filteredData.custom_fields);
  }

  // List of read-only fields by entity type according to Pipedrive API documentation
  const commonReadOnlyFields = [
    // Timestamps
    "add_time",
    "update_time",
    "last_activity_date",
    "next_activity_time",
    "last_activity_id",
    "stage_change_time",
    "lost_time",
    "close_time",

    // System generated fields
    "id",
    "creator_user_id",
    "user_id",
    "org_id.name",
    "person_id.name",
    "owner_id.name",
    "stage_id.name",
    "pipeline_id.name",
    "is_deleted",
    "visible_to",
    "was_seen",
    "cc_email",
    "origin",

    // Counts and stats
    "activities_count",
    "done_activities_count",
    "undone_activities_count",
    "files_count",
    "notes_count",
    "followers_count",
    "weighted_value",
    "formatted_value",
    "rotten_time",
  ];

  // Additional read-only fields specific to entity types
  const entitySpecificReadOnlyFields = {
    deals: [
      "status",
      "probability",
      "lost_reason",
      "contacts_count",
      "products_count",
    ],
    persons: [
      "last_name",
      "first_name",
      "org_name",
      "owner_name",
      "cc_email",
      "open_deals_count",
      "related_open_deals_count",
      "closed_deals_count",
      "related_closed_deals_count",
      "participant_open_deals_count",
      "participant_closed_deals_count",
    ],
    organizations: [
      "owner_name",
      "cc_email",
      "open_deals_count",
      "related_open_deals_count",
      "closed_deals_count",
      "related_closed_deals_count",
      "people_count",
    ],
    activities: ["company_id", "user_id", "note", "assigned_to_user_id"],
    leads: [
      "creator_user_id",
      "add_time",
      "update_time",
      "visible_to",
      "cc_email",
    ],
    products: [
      "first_char",
      "active_flag",
      "selectable",
      "files_count",
      "followers_count",
      "add_time",
      "update_time",
    ],
  };

  // Create a combined list of read-only fields for this entity type
  const readOnlyFields = [...commonReadOnlyFields];
  if (entitySpecificReadOnlyFields[entityType]) {
    readOnlyFields.push(...entitySpecificReadOnlyFields[entityType]);
  }

  // Also match patterns for read-only fields
  const readOnlyPatterns = [
    /_name$/, // Fields ending with _name (e.g., owner_name)
    /_email$/, // Fields ending with _email
    /\.name$/, // Nested name fields (e.g., owner_id.name)
    /\.email$/, // Nested email fields
    /^cc_/, // Fields starting with cc_
    /_count$/, // Count fields
    /_flag$/, // Flag fields
  ];

  // Copy fields that are not read-only to the filtered data
  for (const key in data) {
    // Skip the custom_fields object, which we've already handled
    if (key === "custom_fields") continue;

    // Skip fields that are in the read-only list
    if (readOnlyFields.includes(key)) {
      continue;
    }

    // Skip fields that match read-only patterns
    let isReadOnly = false;
    for (const pattern of readOnlyPatterns) {
      if (pattern.test(key)) {
        isReadOnly = true;
        break;
      }
    }
    if (isReadOnly) continue;

    // Special handling for won_time field - this needs to be properly formatted
    if (key === "won_time") {
      const formattedDate = formatDateField(data[key]);
      if (formattedDate) {
        filteredData[key] = formattedDate;
      } else {
        continue;
      }
    }
    // Include this field in the filtered data
    else {
      filteredData[key] = data[key];
    }
  }

  // For organizations, special handling for address components
  if (entityType === "organizations") {
    // Don't include individual address components at root level
    const addressFieldsRegex =
      /^address_(street_number|route|sublocality|locality|admin_area_level_[12]|country|postal_code|formatted_address)$/;
    for (const key in filteredData) {
      if (addressFieldsRegex.test(key)) {
        delete filteredData[key];
      }
    }
  }

  // For email and phone fields, ensure they are formatted correctly
  if (entityType === "persons") {
    if (data.email && Array.isArray(data.email)) {
      filteredData.email = data.email;
    }

    if (data.phone && Array.isArray(data.phone)) {
      filteredData.phone = data.phone;
    }
  }

  // Format datetime fields properly
  const timeFields = [
    "won_time",
    "lost_time",
    "close_time",
    "expected_close_date",
    "next_activity_date",
  ];
  for (const field of timeFields) {
    if (filteredData[field]) {
      filteredData[field] = formatDateTimeField(filteredData[field]);
    }
  }

  // Log how many fields were filtered
  const originalFieldCount =
    Object.keys(data).length +
    (data.custom_fields ? Object.keys(data.custom_fields).length : 0);
  const filteredFieldCount =
    Object.keys(filteredData).length +
    (filteredData.custom_fields
      ? Object.keys(filteredData.custom_fields).length
      : 0);
  const topLevelFieldCount = Object.keys(filteredData).length;
  const customFieldCount = filteredData.custom_fields
    ? Object.keys(filteredData.custom_fields).length
    : 0;


  // CRITICAL: Ensure all fields are properly formatted before sending to API
  // Apply final formatting to the entire data structure
  const formattedData = ensureCriticalFieldFormats(filteredData, entityType);

  // Validate won_time specifically as it's causing issues - Use ISO format for won_time
  if (formattedData.won_time) {
    // Won time requires ISO datetime format, not just YYYY-MM-DD
    const isoDateTime = formatDateTimeField(formattedData.won_time);
    if (isoDateTime) {
      formattedData.won_time = isoDateTime;
    } else {
      delete formattedData.won_time;
    }
  }

  // Same for lost_time and close_time
  if (formattedData.lost_time) {
    const isoDateTime = formatDateTimeField(formattedData.lost_time);
    if (isoDateTime) {
      formattedData.lost_time = isoDateTime;
    } else {
      delete formattedData.lost_time;
    }
  }

  if (formattedData.close_time) {
    const isoDateTime = formatDateTimeField(formattedData.close_time);
    if (isoDateTime) {
      formattedData.close_time = isoDateTime;
    } else {
      delete formattedData.close_time;
    }
  }

  // Double-check custom fields formatting
  if (formattedData.custom_fields) {
    // Perform a final check for each custom field type
    for (const fieldId in formattedData.custom_fields) {
      try {
        const value = formattedData.custom_fields[fieldId];

        // Skip null values
        if (value === null || value === undefined) {
          delete formattedData.custom_fields[fieldId]; // Remove null values completely
          continue;
        }

        // Skip empty strings
        if (value === "") {
          delete formattedData.custom_fields[fieldId]; // Remove empty strings completely
          continue;
        }

        // FINAL SANITY CHECK BASED ON FIELD NAME PATTERNS
        // This ensures each field matches Pipedrive's expected format

        // 1. Date custom fields - must be YYYY-MM-DD
        if (
          fieldId.includes("date") &&
          !fieldId.includes("date_range") &&
          !fieldId.includes("datetime")
        ) {
          if (
            typeof value !== "string" ||
            !value.match(/^\d{4}-\d{2}-\d{2}$/)
          ) {
            // Try to reformat
            const formattedDate = formatDateField(value);
            if (formattedDate) {
              formattedData.custom_fields[fieldId] = formattedDate;
            } else {
              delete formattedData.custom_fields[fieldId];
            }
          }
        }

        // 2. Date range fields - must be object with start/end dates
        else if (fieldId.includes("date_range")) {
          if (
            typeof value !== "object" ||
            value === null ||
            Array.isArray(value)
          ) {
            delete formattedData.custom_fields[fieldId];
          } else {
            // Ensure each property is properly formatted
            const rangeObj = {
              start: null,
              end: null,
            };

            if (value.start) {
              const formattedDate = formatDateField(value.start);
              if (formattedDate) {
                rangeObj.start = formattedDate;
              }
            }

            if (value.end) {
              const formattedDate = formatDateField(value.end);
              if (formattedDate) {
                rangeObj.end = formattedDate;
              }
            }

            formattedData.custom_fields[fieldId] = rangeObj;
          }
        }

        // 3. Multi options fields - must be arrays
        else if (
          fieldId.includes("options") ||
          fieldId.includes("multi") ||
          fieldId.includes("multiple")
        ) {
          if (!Array.isArray(value)) {
            try {
              // Try to convert to array
              let optionsArray = [];

              if (typeof value === "string" && value.includes(",")) {
                optionsArray = value.split(",").map((item) => item.trim());
              } else {
                optionsArray = [value];
              }

              // Convert to numbers if possible
              optionsArray = optionsArray.map((item) => {
                if (typeof item === "string" && !isNaN(Number(item))) {
                  return Number(item);
                } else if (typeof item === "object" && item.id) {
                  return Number(item.id);
                }
                return item;
              });

              formattedData.custom_fields[fieldId] = optionsArray;
            } catch (e) {
              delete formattedData.custom_fields[fieldId];
            }
          }
        }

        // 4. Organization fields - must be numbers
        else if (fieldId.includes("org") || fieldId.includes("company")) {
          if (typeof value !== "number") {
            try {
              if (typeof value === "string" && !isNaN(Number(value))) {
                formattedData.custom_fields[fieldId] = Number(value);
              } else if (typeof value === "object" && value.id) {
                formattedData.custom_fields[fieldId] = Number(value.id);
              } else {
                delete formattedData.custom_fields[fieldId];
              }
            } catch (e) {
              delete formattedData.custom_fields[fieldId];
            }
          }
        }

        // 5. Phone fields - must be strings
        else if (fieldId.includes("phone")) {
          if (typeof value !== "string") {
            try {
              formattedData.custom_fields[fieldId] = String(value);
            } catch (e) {
              delete formattedData.custom_fields[fieldId];
            }
          }
        }

        // 6. Time fields - must be objects with hour/minute
        else if (
          fieldId.includes("time") &&
          !fieldId.includes("range") &&
          !fieldId.includes("date")
        ) {
          if (
            typeof value !== "object" ||
            value === null ||
            Array.isArray(value)
          ) {
            let hour = 0,
              minute = 0;

            if (typeof value === "string" && value.includes(":")) {
              const parts = value.split(":");
              hour = parseInt(parts[0], 10) || 0;
              minute = parseInt(parts[1], 10) || 0;
            }

            filteredData.custom_fields[fieldId] = {
              hour,
              minute,
            };
          }
        }

        // TIME RANGE FIELDS - force object format with start/end
        else if (fieldId.includes("time_range")) {
          if (
            typeof value !== "object" ||
            value === null ||
            Array.isArray(value)
          ) {
            filteredData.custom_fields[fieldId] = {
              start: {
                hour: 0,
                minute: 0,
              },
              end: {
                hour: 0,
                minute: 0,
              },
            };
          } else {
            // Ensure start and end properties exist and are properly formatted
            if (!value.start || typeof value.start !== "object") {
              filteredData.custom_fields[fieldId].start = {
                hour: 0,
                minute: 0,
              };
            }
            if (!value.end || typeof value.end !== "object") {
              filteredData.custom_fields[fieldId].end = {
                hour: 0,
                minute: 0,
              };
            }
          }
        }

        // USER FIELDS - force number format
        else if (fieldId.includes("user")) {
          if (typeof value !== "number") {
            if (typeof value === "string" && !isNaN(Number(value))) {
              filteredData.custom_fields[fieldId] = Number(value);
            } else if (
              typeof value === "object" &&
              value !== null &&
              value.id
            ) {
              filteredData.custom_fields[fieldId] = Number(value.id);
            } else {
              delete filteredData.custom_fields[fieldId];
            }
          }
        }
      } catch (e) {
        // If all else fails, remove the field
        delete formattedData.custom_fields[fieldId];
      }
    }
  }

  // Handle entity-specific fields
  switch (entityType) {
    case "persons":
    case "person":
      try {
        // Handle phone field - ensure it's a string
        if ("phone" in data && data.phone !== null) {
          if (Array.isArray(data.phone)) {
            // Make sure each phone object has string values
            data.phone.forEach((phone) => {
              if (phone && phone.value) {
                phone.value = String(phone.value);
              }
            });
          } else {
            data.phone = String(data.phone);
          }
        }

        // Handle email field - ensure it's a string
        if ("email" in data && data.email !== null) {
          if (Array.isArray(data.email)) {
            // Make sure each email object has string values
            data.email.forEach((email) => {
              if (email && email.value) {
                email.value = String(email.value);
              }
            });
          } else {
            data.email = String(data.email);
          }
        }
      } catch (e) {
      }
      break;

    case "organizations":
    case "organization":
      try {
        // Handle address field - ensure it's a string
        if ("address" in data && data.address !== null) {
          if (typeof data.address !== "string") {
            if (typeof data.address === "object" && data.address !== null) {
              if (data.address.formatted_address) {
                data.address = data.address.formatted_address;
              } else if (data.address.value) {
                data.address = data.address.value;
              } else {
                data.address = String(data.address);
              }
            } else {
              data.address = String(data.address);
            }
          }
        }
      } catch (e) {
      }
      break;
  }

  return data;
}

/**
 * Ensure critical fields are properly formatted for Pipedrive API
 * @param {Object} data - The data to be sent to Pipedrive
 * @param {string} entityType - The type of entity (deal, person, organization, etc.)
 * @return {Object} - The properly formatted data
 */
function ensureCriticalFieldFormats(data, entityType) {
  if (!data) {
    return data;
  }


  // Handle deal-specific fields
  if (entityType === "deals") {
    // Handle won_time field
    if ("won_time" in data && data.won_time) {
      var formattedDate = formatDateField(data.won_time);
      if (formattedDate) {
        data.won_time = formattedDate;
      } else {
        delete data.won_time;
      }
    }

    // Handle lost_time field
    if ("lost_time" in data && data.lost_time) {
      var formattedDate = formatDateField(data.lost_time);
      if (formattedDate) {
        data.lost_time = formattedDate;
      } else {
        delete data.lost_time;
      }
    }

    // Handle expected_close_date field
    if ("expected_close_date" in data && data.expected_close_date) {
      var formattedDate = formatDateField(data.expected_close_date);
      if (formattedDate) {
        data.expected_close_date = formattedDate;
      } else {
        delete data.expected_close_date;
      }
    }

    // Handle close_time field
    if ("close_time" in data && data.close_time) {
      var formattedDate = formatDateField(data.close_time);
      if (formattedDate) {
        data.close_time = formattedDate;
      } else {
        delete data.close_time;
      }
    }
  }

  // Format timestamp fields
  var timestampFields = [
    "add_time",
    "update_time",
    "first_name_update_time",
    "last_name_update_time",
  ];
  for (var i = 0; i < timestampFields.length; i++) {
    var field = timestampFields[i];
    if (field in data && data[field]) {
      var formattedDate = formatDateField(data[field]);
      if (formattedDate) {
        data[field] = formattedDate;
      } else {
        delete data[field];
      }
    }
  }

  // Process custom fields
  if (data.custom_fields && typeof data.custom_fields === "object") {
    for (var fieldId in data.custom_fields) {
      if (!data.custom_fields.hasOwnProperty(fieldId)) continue;

      // Skip if field is null, undefined, or empty string
      if (
        data.custom_fields[fieldId] === null ||
        data.custom_fields[fieldId] === undefined ||
        data.custom_fields[fieldId] === ""
      ) {
        delete data.custom_fields[fieldId];
        continue;
      }

      var fieldValue = data.custom_fields[fieldId];

      // DATE CUSTOM FIELD
      if (
        typeof fieldValue === "string" &&
        (fieldValue.includes("/") ||
          fieldValue.includes("-") ||
          fieldValue.includes("."))
      ) {
        var formattedDate = formatDateField(fieldValue);
        if (formattedDate) {
          data.custom_fields[fieldId] = formattedDate;
        } else {
          delete data.custom_fields[fieldId];
        }
      }

      // MULTI OPTIONS FIELDS
      else if (
        fieldValue.toString().includes(",") ||
        Array.isArray(fieldValue)
      ) {
        var optionsArray = [];

        if (Array.isArray(fieldValue)) {
          optionsArray = fieldValue.slice(); // Clone array
        } else if (typeof fieldValue === "string") {
          optionsArray = fieldValue.split(",");
          for (var i = 0; i < optionsArray.length; i++) {
            optionsArray[i] = optionsArray[i].trim();
          }
        } else {
          optionsArray = [fieldValue]; // Convert single value to array
        }

        data.custom_fields[fieldId] = optionsArray;
      }

      // PHONE FIELDS
      else if (fieldId.includes("phone")) {
        data.custom_fields[fieldId] = String(fieldValue);
      }
    }
  }

  // Handle entity-specific fields
  if (entityType === "persons" || entityType === "person") {
    // Handle phone field
    if ("phone" in data && data.phone !== null) {
      if (Array.isArray(data.phone)) {
        for (var i = 0; i < data.phone.length; i++) {
          if (data.phone[i] && data.phone[i].value) {
            data.phone[i].value = String(data.phone[i].value);
          }
        }
      } else {
        data.phone = String(data.phone);
      }
    }

    // Handle email field
    if ("email" in data && data.email !== null) {
      if (Array.isArray(data.email)) {
        for (var i = 0; i < data.email.length; i++) {
          if (data.email[i] && data.email[i].value) {
            data.email[i].value = String(data.email[i].value);
          }
        }
      } else {
        data.email = String(data.email);
      }
    }
  } else if (entityType === "organizations" || entityType === "organization") {
    // Handle address field
    if ("address" in data && data.address !== null) {
      if (typeof data.address !== "string") {
        if (typeof data.address === "object" && data.address !== null) {
          if (data.address.formatted_address) {
            data.address = data.address.formatted_address;
          } else if (data.address.value) {
            data.address = data.address.value;
          } else {
            data.address = String(data.address);
          }
        } else {
          data.address = String(data.address);
        }
      }
    }
  }

  return data;
}

/**
 * Formats custom fields to match Pipedrive API requirements
 * @param {Object} customFields - Object containing custom fields
 */
function formatCustomFields(customFields) {
  if (!customFields) return;

  let processedCount = 0;

  // Loop through each field
  for (var fieldId in customFields) {
    if (!customFields.hasOwnProperty(fieldId)) continue;

    var value = customFields[fieldId];

    // Skip null/undefined/empty values
    if (value === null || value === undefined || value === "") {
      delete customFields[fieldId];
      continue;
    }

    // DATE RANGE FIELDS - Must be an object with {start, end} properties for Pipedrive API
    if (fieldId.includes("date") && fieldId.includes("range")) {
      try {
        let start = null,
          end = null;
        if (typeof value === "object" && value !== null) {
          // Accept both {start, end} and {value, until}
          start = value.start || value.value || null;
          end = value.end || value.until || null;
        } else if (
          typeof value === "string" &&
          (value.includes("-") || value.includes("to"))
        ) {
          let dates = value.includes("to")
            ? value.split("to")
            : value.split("-");
          if (dates.length === 2) {
            start = dates[0].trim();
            end = dates[1].trim();
          }
        } else if (typeof value === "string") {
          // Single date, use for both start and end
          start = end = value.trim();
        }
        // Format dates
        start = formatDateField(start);
        end = formatDateField(end);
        if (start && end) {
          customFields[fieldId] = {
            start: start,
            end: end,
          };
        } else {
          delete customFields[fieldId];
        }
        processedCount++;
        continue;
      } catch (e) {
        delete customFields[fieldId];
        continue;
      }
    }

    // REGULAR DATE FIELDS
    if (fieldId.includes("date") && !fieldId.includes("range")) {
      var formattedDate = formatDateField(value);
      if (formattedDate) {
        customFields[fieldId] = formattedDate;
      } else {
        delete customFields[fieldId];
      }
      processedCount++;
      continue;
    }

    // MULTI-OPTION FIELDS - Must be array of numeric IDs
    if (
      fieldId.includes("option") ||
      fieldId.includes("multi") ||
      (typeof value === "string" && value.includes(","))
    ) {

      try {
        // Convert to array if it's not already
        var optionsArray = [];

        if (Array.isArray(value)) {
          optionsArray = value.slice(); // Clone the array
        } else if (typeof value === "string") {
          optionsArray = value.split(",").map(function (item) {
            return item.trim();
          });
        } else {
          optionsArray = [value];
        }

        // Convert all option IDs to numbers for Pipedrive API
        optionsArray = optionsArray.map(function (option) {
          // If it's already a number, return it
          if (typeof option === "number") {
            return option;
          }

          // If it's a string that can be converted to a number, convert it
          if (typeof option === "string" && !isNaN(option)) {
            return Number(option);
          }

          // Otherwise, try to find the option ID from the field definitions
          // This would require additional API logic to look up option IDs by label
          // For now, just log a warning and return the original value
          return option;
        });

        customFields[fieldId] = optionsArray;
      } catch (e) {
        delete customFields[fieldId];
      }
      processedCount++;
      continue;
    }

    // TIME FIELDS - Must be string in format HH:MM
    if (fieldId.includes("time") && !fieldId.includes("date")) {
      try {
        // Ensure it's a string
        if (typeof value !== "string") {
          value = String(value);
        }

        // Match HH:MM format
        if (/^\d{1,2}:\d{2}$/.test(value)) {
          // Already in correct format
          customFields[fieldId] = value;
        }
        // Try to format from other common formats
        else {
          // Try to extract hours and minutes
          let timeValue = null;

          // Try to parse as Date object if it has date components
          if (value.includes("/") || value.includes("-")) {
            try {
              const date = new Date(value);
              if (!isNaN(date.getTime())) {
                // Format as HH:MM
                timeValue =
                  padZero(date.getHours()) + ":" + padZero(date.getMinutes());
              }
            } catch (e) {
              // Failed to parse as date
            }
          }
          // Try extracting from formats like "13h30m" or "1:30 PM"
          else {
            // Extract hours and minutes with regex
            const match = value.match(/(\d{1,2})[h:](\d{2})/i);
            if (match) {
              timeValue = padZero(parseInt(match[1])) + ":" + match[2];
            }

            // Handle AM/PM format
            if (value.match(/\d{1,2}:\d{2}\s*(am|pm)/i)) {
              const parts = value.match(/(\d{1,2}):(\d{2})\s*(am|pm)/i);
              if (parts) {
                let hours = parseInt(parts[1]);
                const minutes = parts[2];
                const ampm = parts[3].toLowerCase();

                if (ampm === "pm" && hours < 12) {
                  hours += 12;
                } else if (ampm === "am" && hours === 12) {
                  hours = 0;
                }

                timeValue = padZero(hours) + ":" + minutes;
              }
            }
          }

          if (timeValue) {
            customFields[fieldId] = timeValue;
          } else {
            delete customFields[fieldId];
          }
        }
      } catch (e) {
        delete customFields[fieldId];
      }
      processedCount++;
      continue;
    }

    // ORGANIZATION FIELDS
    if (
      typeof value === "object" &&
      value !== null &&
      (value.name !== undefined || value.id !== undefined)
    ) {
      // Just need the ID for API
      if (value.id) {
        customFields[fieldId] = value.id;
      } else {
        delete customFields[fieldId];
      }
      processedCount++;
      continue;
    }

    // PHONE FIELDS
    if (
      typeof value === "object" &&
      value !== null &&
      value.value !== undefined &&
      value.code !== undefined
    ) {
      // Pipedrive expects just the phone number
      customFields[fieldId] = value.value;
      processedCount++;
      continue;
    }

    // Assume any other object type fields should be normalized
    if (typeof value === "object" && value !== null) {
      // If it has a value property, use that
      if (value.value !== undefined) {
        customFields[fieldId] = value.value;
      }
      // If it's empty, remove it
      else if (Object.keys(value).length === 0) {
        delete customFields[fieldId];
      }
      // Otherwise keep it as is
      processedCount++;
      continue;
    }

    processedCount++;
  }

}

// Helper function to pad numbers with leading zeros
function padZero(num) {
  return num < 10 ? "0" + num : num;
}

/**
 * Formats a date value to YYYY-MM-DD format
 * @param {string|Date} dateValue - The date value to format
 * @return {string|null} - The formatted date string or null if invalid
 */
function formatDateField(dateValue) {
  if (!dateValue) return null;

  try {
    // If already in YYYY-MM-DD format, validate and return
    if (
      typeof dateValue === "string" &&
      dateValue.match(/^\d{4}-\d{2}-\d{2}$/)
    ) {
      // Validate if it's a valid date
      const parts = dateValue.split("-");
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // JS months are 0-based
      const day = parseInt(parts[2], 10);

      const testDate = new Date(year, month, day);
      if (
        testDate.getFullYear() === year &&
        testDate.getMonth() === month &&
        testDate.getDate() === day
      ) {
        return dateValue; // Valid date in correct format
      }
      return null; // Invalid date
    }

    // Handle Date objects
    if (dateValue instanceof Date) {
      const year = dateValue.getFullYear();
      const month = (dateValue.getMonth() + 1).toString().padStart(2, "0");
      const day = dateValue.getDate().toString().padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    // Try to parse various string formats
    if (typeof dateValue === "string") {
      // Handle MM/DD/YYYY format
      if (dateValue.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        const parts = dateValue.split("/");
        const month = parseInt(parts[0], 10).toString().padStart(2, "0");
        const day = parseInt(parts[1], 10).toString().padStart(2, "0");
        const year = parts[2];
        return `${year}-${month}-${day}`;
      }

      // Handle DD/MM/YYYY format
      if (dateValue.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        const parts = dateValue.split(".");
        const day = parseInt(parts[0], 10).toString().padStart(2, "0");
        const month = parseInt(parts[1], 10).toString().padStart(2, "0");
        const year = parts[2];
        return `${year}-${month}-${day}`;
      }

      // Try creating a Date object and formatting
      const date = new Date(dateValue);
      if (!isNaN(date.getTime())) {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, "0");
        const day = date.getDate().toString().padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
    }

    // If number, assume it's a timestamp
    if (typeof dateValue === "number") {
      const date = new Date(dateValue);
      if (!isNaN(date.getTime())) {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, "0");
        const day = date.getDate().toString().padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
    }

    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Formats a date value to ISO 8601 datetime format with timezone (required for won_time and other special fields)
 * @param {string|Date} dateValue - The date value to format
 * @return {string|null} - The formatted ISO date string or null if invalid
 */
function formatDateTimeField(dateValue) {
  if (!dateValue) return null;

  try {
    let date;

    // If already a Date object
    if (dateValue instanceof Date) {
      date = dateValue;
    }
    // If string in YYYY-MM-DD format
    else if (
      typeof dateValue === "string" &&
      dateValue.match(/^\d{4}-\d{2}-\d{2}$/)
    ) {
      const parts = dateValue.split("-");
      date = new Date(
        parseInt(parts[0], 10),
        parseInt(parts[1], 10) - 1, // JS months are 0-based
        parseInt(parts[2], 10)
      );
    }
    // If other string formats
    else if (typeof dateValue === "string") {
      // Try parsing various formats
      if (dateValue.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        const parts = dateValue.split("/");
        date = new Date(
          parseInt(parts[2], 10),
          parseInt(parts[0], 10) - 1,
          parseInt(parts[1], 10)
        );
      } else if (dateValue.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        const parts = dateValue.split(".");
        date = new Date(
          parseInt(parts[2], 10),
          parseInt(parts[1], 10) - 1,
          parseInt(parts[0], 10)
        );
      } else {
        // Try standard Date parsing
        date = new Date(dateValue);
      }
    }
    // If timestamp number
    else if (typeof dateValue === "number") {
      date = new Date(dateValue);
    }

    // Check if we got a valid date
    if (date instanceof Date && !isNaN(date.getTime())) {
      // Format as ISO 8601 string
      return date.toISOString();
    }

    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Converts a column letter to an index (e.g., A -> 1, AA -> 27)
 * @param {string} columnLetter - The column letter (e.g., 'A', 'BC')
 * @return {number} The column index (1-based)
 */
function columnLetterToIndex(columnLetter) {
  let result = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    result = result * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Gets field mappings for a specific entity type
 * @param {string} entityType - The entity type to get mappings for
 * @return {Object} An object mapping column headers to API field keys
 */
function getFieldMappingsForEntity(entityType) {
  // Basic field mappings for each entity type
  const commonMappings = {
    ID: "id",
    Name: "name",
    Owner: "owner_id",
    Organization: "org_id",
    Person: "person_id",
    Added: "add_time",
    Updated: "update_time",
  };

  // Entity-specific mappings
  const entityMappings = {
    [ENTITY_TYPES.DEALS]: {
      Value: "value",
      Currency: "currency",
      Title: "title",
      Pipeline: "pipeline_id",
      Stage: "stage_id",
      Status: "status",
      "Expected Close Date": "expected_close_date",
    },
    [ENTITY_TYPES.PERSONS]: {
      Email: "email",
      Phone: "phone",
      "First Name": "first_name",
      "Last Name": "last_name",
      Organization: "org_id",
    },
    [ENTITY_TYPES.ORGANIZATIONS]: {
      Address: "address",
      Website: "web",
    },
    [ENTITY_TYPES.ACTIVITIES]: {
      Type: "type",
      "Due Date": "due_date",
      "Due Time": "due_time",
      Duration: "duration",
      Deal: "deal_id",
      Person: "person_id",
      Organization: "org_id",
      Note: "note",
    },
    [ENTITY_TYPES.PRODUCTS]: {
      Code: "code",
      Description: "description",
      Unit: "unit",
      Tax: "tax",
      Category: "category",
      Active: "active_flag",
      Selectable: "selectable",
      "Visible To": "visible_to",
      "First Price": "first_price",
      Cost: "cost",
      Prices: "prices",
      "Owner Name": "owner_id.name", // Map "Owner Name" to owner_id.name so we can detect this field
    },
  };

  // Combine common mappings with entity-specific mappings
  return {
    ...commonMappings,
    ...(entityMappings[entityType] || {}),
  };
}

/**
 * Detects if columns in the sheet have shifted and updates tracking accordingly
 */
function detectColumnShifts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    const scriptProperties = PropertiesService.getScriptProperties();

    // Get current and previous positions
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const currentColLetter =
      scriptProperties.getProperty(trackingColumnKey) || "";
    const previousPosStr =
      scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`) ||
      "-1";
    const previousPos = parseInt(previousPosStr, 10);

    // Find all "Sync Status" headers in the sheet
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    let syncStatusColumns = [];

    // Find ALL instances of "Sync Status" headers
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        syncStatusColumns.push(i);
      }
    }

    // If we have multiple "Sync Status" columns, clean up all but the rightmost one
    if (syncStatusColumns.length > 1) {
      // Keep only the rightmost column
      const rightmostIndex = Math.max(...syncStatusColumns);

      // Clean up all other columns
      for (const colIndex of syncStatusColumns) {
        if (colIndex !== rightmostIndex) {
          const colLetter = columnToLetter(colIndex + 1);
          cleanupColumnFormatting(sheet, colLetter);
        }
      }

      // Update the tracking to the rightmost column
      const rightmostColLetter = columnToLetter(rightmostIndex + 1);
      scriptProperties.setProperty(trackingColumnKey, rightmostColLetter);
      scriptProperties.setProperty(
        `CURRENT_SYNCSTATUS_POS_${sheetName}`,
        rightmostIndex.toString()
      );
      return; // Exit after handling duplicates
    }

    let actualSyncStatusIndex =
      syncStatusColumns.length > 0 ? syncStatusColumns[0] : -1;

    if (actualSyncStatusIndex >= 0) {
      const actualColLetter = columnToLetter(actualSyncStatusIndex + 1);

      // If there's a mismatch, columns might have shifted
      if (currentColLetter && actualColLetter !== currentColLetter) {

        // If the actual position is less than the recorded position, columns were removed
        if (actualSyncStatusIndex < previousPos) {

          // Clean ALL columns to be safe
          for (let i = 0; i < sheet.getLastColumn(); i++) {
            if (i !== actualSyncStatusIndex) {
              // Skip current Sync Status column
              cleanupColumnFormatting(sheet, columnToLetter(i + 1));
            }
          }
        }

        // Clean up all potential previous locations
        scanAndCleanupAllSyncColumns(sheet, actualColLetter);

        // Update the tracking column property
        scriptProperties.setProperty(trackingColumnKey, actualColLetter);
        scriptProperties.setProperty(
          `CURRENT_SYNCSTATUS_POS_${sheetName}`,
          actualSyncStatusIndex.toString()
        );
      }
    }
  } catch (error) {
  }
}

/**
 * Cleans up formatting in a column
 * @param {Sheet} sheet - The sheet containing the column
 * @param {string} columnLetter - The letter of the column to clean up
 * @param {boolean} isCurrentColumn - Whether this is the current active Sync Status column
 */
function cleanupColumnFormatting(sheet, columnLetter, isCurrentColumn = false) {
  try {

    // Convert column letter to index
    const columnIndex = letterToColumn(columnLetter);

    // Get the data range to know the last row
    const lastRow = Math.max(sheet.getLastRow(), 2);

    // Clean the header first - only if it's not the current column
    if (!isCurrentColumn) {
      const headerCell = sheet.getRange(1, columnIndex);
      const headerValue = headerCell.getValue();
      const note = headerCell.getNote();


      // Check if this column has 'Sync Status' header or sync-related note
      if (
        headerValue === "Sync Status" ||
        headerValue === "Sync Status (hidden)" ||
        headerValue === "Status" ||
        (note &&
          (note.includes("sync") ||
            note.includes("track") ||
            note.includes("Pipedrive")))
      ) {
        headerCell.setValue("");
        headerCell.clearNote();
        headerCell.clearFormat();
      }
    }

    // Always clean data validations for the entire column (even for current column)
    if (lastRow > 1) {
      // Clear data validations in the entire column
      sheet.getRange(2, columnIndex, lastRow - 1, 1).clearDataValidations();

      // For previous columns, thoroughly clean all sync-related values
      if (!isCurrentColumn) {
        // Get all values to check for specific sync status values
        const values = sheet
          .getRange(2, columnIndex, lastRow - 1, 1)
          .getValues();
        const newValues = [];
        let cleanedCount = 0;

        // Clear only cells containing specific sync status values
        for (let i = 0; i < values.length; i++) {
          const value = values[i][0];

          // Check if the value is a known sync status value
          if (
            value === "Modified" ||
            value === "Not modified" ||
            value === "Synced" ||
            value === "Error"
          ) {
            newValues.push([""]); // Clear known status values
            cleanedCount++;
          } else {
            newValues.push([value]); // Keep other values
          }
        }

        // Set the cleaned values back to the sheet
        sheet
          .getRange(2, columnIndex + 1, values.length, 1)
          .setValues(newValues);

        // Remove conditional formatting for this column
        removeConditionalFormattingForColumn(sheet, columnIndex);
      }
    }

  } catch (error) {
  }
}

// Function to identify and clean up previous Sync Status columns
function cleanupPreviousSyncStatusColumn(sheet, currentSyncColumn) {
  try {

    // Show a toast to let users know that post-processing is happening
    // This helps users understand that data is already written but cleanup is still in progress
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Performing post-sync cleanup and formatting. Your data is already written.",
      "Finalizing Sync",
      5
    );

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetName = sheet.getName();
    const previousSyncColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const previousSyncColumn = scriptProperties.getProperty(
      previousSyncColumnKey
    );

    // IMPORTANT: We do NOT want to override column headers when Pipedrive adds new fields
    // If the previous column is converted to a Pipedrive data column, we need to keep it

    // Clean up the known previous column first if it exists and is different from current
    if (previousSyncColumn && previousSyncColumn !== currentSyncColumn) {

      try {
        // Convert letter to column index
        const previousColumnIndex = letterToColumn(previousSyncColumn);

        // Clean up all Sync Status-specific formatting and validation from this column
        // but do NOT clear the header cell - let the main sync function handle headers

        // First, clear any sync-specific formatting and validation in the data cells
        if (sheet.getLastRow() > 1) {
          const dataRange = sheet.getRange(
            2,
            previousColumnIndex,
            Math.max(sheet.getLastRow() - 1, 1),
            1
          );

          // Clear all formatting and validation from data cells
          dataRange.clearFormat();
          dataRange.clearDataValidations();

          // Check for and clear status-specific values ONLY
          const values = dataRange.getValues();
          const newValues = values.map((row) => {
            const value = String(row[0]).trim();

            // Only clear if it's one of the specific status values
            if (
              value === "Modified" ||
              value === "Not modified" ||
              value === "Synced" ||
              value === "Error"
            ) {
              return [""];
            }
            return [value]; // Keep any other values
          });

          // Write the cleaned values back
          dataRange.setValues(newValues);
        }

        // Remove any sync-specific formatting or notes from the header
        // but KEEP the header cell itself for Pipedrive data
        const headerCell = sheet.getRange(1, previousColumnIndex);
        headerCell.clearFormat();
        headerCell.clearNote();
        // Do NOT call setValue() - let the main sync function set the header

      } catch (e) {
      }
    }

    // Scan for any other columns that might have sync status formatting
    const lastColumn = sheet.getLastColumn();

    for (let i = 1; i <= lastColumn; i++) {
      const colLetter = columnToLetter(i);

      // Skip the current sync column
      if (colLetter === currentSyncColumn) {
        continue;
      }

      // Check if this might be a sync status column by inspecting formatting and values
      try {
        // Check the header for sync status indicators
        const headerCell = sheet.getRange(1, i);
        const headerValue = headerCell.getValue();
        const headerNote = headerCell.getNote();

        const isSyncStatusHeader =
          headerValue === "Sync Status" ||
          headerValue === "Sync Status (hidden)" ||
          headerValue === "Status" ||
          (headerNote &&
            (headerNote.includes("sync") ||
              headerNote.includes("track") ||
              headerNote.includes("Pipedrive")));

        // Also check for sync status values in the data cells
        let hasSyncStatusValues = false;
        if (sheet.getLastRow() > 1) {
          // Sample a few cells to check for status values
          const sampleSize = Math.min(10, sheet.getLastRow() - 1);
          const sampleRange = sheet.getRange(2, i, sampleSize, 1);
          const sampleValues = sampleRange.getValues();

          hasSyncStatusValues = sampleValues.some((row) => {
            const value = String(row[0]).trim();
            return (
              value === "Modified" ||
              value === "Not modified" ||
              value === "Synced" ||
              value === "Error"
            );
          });
        }

        // If this column has sync status indicators, clean it
        if (isSyncStatusHeader || hasSyncStatusValues) {

          // Clean any sync-specific formatting and validation but preserve the header cell
          if (sheet.getLastRow() > 1) {
            const dataRange = sheet.getRange(
              2,
              i,
              Math.max(sheet.getLastRow() - 1, 1),
              1
            );

            // Clear all formatting and validation
            dataRange.clearFormat();
            dataRange.clearDataValidations();

            // Only clear specific status values
            const values = dataRange.getValues();
            const newValues = values.map((row) => {
              const value = String(row[0]).trim();
              if (
                value === "Modified" ||
                value === "Not modified" ||
                value === "Synced" ||
                value === "Error"
              ) {
                return [""];
              }
              return [value]; // Keep any other values
            });

            dataRange.setValues(newValues);
          }

          // Remove sync-specific formatting and notes from header
          // but preserve the header cell itself for Pipedrive data
          headerCell.clearFormat();
          headerCell.clearNote();
          // Do NOT clear header text - let main sync function set it

        }
      } catch (e) {
      }
    }

    // Clear the previous column tracking since we've cleaned it up
    scriptProperties.deleteProperty(previousSyncColumnKey);
  } catch (error) {
  }
}

// Function to remove conditional formatting rules for a specific column
function removeConditionalFormattingForColumn(sheet, columnIndex) {
  try {
    const rules = sheet.getConditionalFormatRules();
    const colA1Notation = columnToLetter(columnIndex + 1);
    const newRules = [];

    // Only keep rules that don't apply to this column
    for (let i = 0; i < rules.length; i++) {
      const rule = rules[i];
      const ranges = rule.getRanges();
      let keepRule = true;

      // Check if this rule applies to the column we're cleaning
      for (let j = 0; j < ranges.length; j++) {
        const range = ranges[j];
        const a1Notation = range.getA1Notation();

        // If the rule includes this column, don't keep it
        if (a1Notation.includes(colA1Notation)) {
          keepRule = false;
          break;
        }
      }

      if (keepRule) {
        newRules.push(rule);
      }
    }

    // Update the conditional formatting rules
    sheet.setConditionalFormatRules(newRules);
  } catch (error) {
  }
}

/**
 * Helper function to convert a column letter to a 1-based index
 * @param {string} letter - The column letter (e.g., "A", "B", "AA")
 * @return {number} The 1-based index
 */
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * Helper function to convert a column number to a letter
 * @param {number} column - The column number (1-based)
 * @return {string} The column letter
 */
function columnToLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Formats a date or date-time field for Pipedrive API
 * @param {*} value - The date/time value to format
 * @param {Object} fieldDefinition - The field definition from Pipedrive
 * @return {string} Properly formatted date/time string for Pipedrive API
 */
function formatDateTimeForPipedrive(value, fieldDefinition) {
  try {
    if (!value) return null;

    // Check if this is a time-only field (special format required)
    if (fieldDefinition && fieldDefinition.field_type === "time") {

      // If value is already in HH:MM:SS format, use it directly
      if (
        typeof value === "string" &&
        value.match(/^\d{1,2}:\d{2}(:\d{2})?$/)
      ) {
        // Ensure it has seconds if not present
        if (!value.includes(":")) {
          return value + ":00";
        }
        return value;
      }

      // Handle time values in date objects or various formats
      let hours,
        minutes,
        seconds = "00";

      if (value instanceof Date) {
        // Extract time components from Date object
        hours = value.getHours();
        minutes = value.getMinutes();
      } else if (typeof value === "string") {
        // Try to parse various time formats

        // Format: "4:30 PM" or "4:30 AM"
        const amPmMatch = value.match(
          /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)/i
        );
        if (amPmMatch) {
          hours = parseInt(amPmMatch[1], 10);
          minutes = parseInt(amPmMatch[2], 10);
          if (amPmMatch[3]) seconds = amPmMatch[3];

          // Adjust for PM
          if (amPmMatch[4].toUpperCase() === "PM" && hours < 12) {
            hours += 12;
          }
          // Adjust for AM 12-hour
          if (amPmMatch[4].toUpperCase() === "AM" && hours === 12) {
            hours = 0;
          }
        }
        // Format: "16:30" or "16:30:00"
        else if (value.match(/^\d{1,2}:\d{2}(:\d{2})?$/)) {
          const parts = value.split(":");
          hours = parseInt(parts[0], 10);
          minutes = parseInt(parts[1], 10);
          if (parts.length > 2) seconds = parts[2];
        }
        // Try to parse as full date and extract time
        else {
          try {
            const dateObj = new Date(value);
            if (!isNaN(dateObj.getTime())) {
              hours = dateObj.getHours();
              minutes = dateObj.getMinutes();
              seconds = String(dateObj.getSeconds()).padStart(2, "0");
            }
          } catch (e) {
            return null;
          }
        }
      }

      // If we couldn't extract time components, return null
      if (hours === undefined || minutes === undefined) {
        return null;
      }

      // Format time as HH:MM:SS - THIS IS THE KEY PART FOR PIPEDRIVE API
      return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
        2,
        "0"
      )}:${seconds}`;
    }

    // For regular date fields, continue with the existing logic
    let dateObj;
    if (typeof value === "string") {
      // Handle Excel/Sheets date formats
      if (value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        const parts = value.split("/");
        dateObj = new Date(parts[2], parts[0] - 1, parts[1]);
      } else {
        dateObj = new Date(value);
      }
    } else if (value instanceof Date) {
      dateObj = value;
    } else {
      // For numbers (Excel serial dates) or other formats
      try {
        dateObj = new Date(value);
      } catch (e) {
        return null;
      }
    }

    // Check if the date is valid
    if (isNaN(dateObj.getTime())) {
      return null;
    }

    // Determine if the field is a date-only field
    let isDateOnly = false;
    if (fieldDefinition && fieldDefinition.field_type) {
      isDateOnly = fieldDefinition.field_type === "date";
    } else {
      // If no field definition, try to guess based on time part
      isDateOnly =
        dateObj.getHours() === 0 &&
        dateObj.getMinutes() === 0 &&
        dateObj.getSeconds() === 0;
    }

    // Format appropriately
    if (isDateOnly) {
      // Format as YYYY-MM-DD for date fields
      return dateObj.toISOString().split("T")[0];
    } else {
      // Format as full ISO string for datetime fields
      return dateObj.toISOString();
    }
  } catch (error) {
    return null;
  }
}

/**
 * Ensures time range fields (with _until suffix) are properly handled for Pipedrive API
 * @param {Object} payload - The payload to process
 * @param {Object} rowData - The original row data that might contain end time values
 * @param {Object} headerToFieldKeyMap - Mapping of headers to field keys
 * @return {Object} - The processed payload with properly defined time range fields
 */
function ensureTimeRangeFieldsForPipedrive(
  payload,
  rowData,
  headerToFieldKeyMap
) {
  try {

    // Create a copy of the payload to avoid modification issues
    const updatedPayload = JSON.parse(JSON.stringify(payload));

    // Make sure custom_fields exists
    if (!updatedPayload.custom_fields) {
      updatedPayload.custom_fields = {};
    }

    // CRITICAL: First check rowData for any end time fields that might not be in the payload
    if (rowData && headerToFieldKeyMap) {
      for (const [header, value] of Object.entries(rowData)) {
        if (header && value && header.toLowerCase().includes("end time")) {
          const fieldKey = headerToFieldKeyMap[header];
          if (fieldKey && fieldKey.endsWith("_until")) {

            // Add to payload if missing
            if (!updatedPayload[fieldKey]) {
              updatedPayload[fieldKey] = value;
            }
            if (!updatedPayload.custom_fields[fieldKey]) {
              updatedPayload.custom_fields[fieldKey] = value;
            }

            // Also ensure the base field exists
            const baseKey = fieldKey.replace(/_until$/, "");
            const baseHeader = Object.keys(headerToFieldKeyMap).find(
              (h) => headerToFieldKeyMap[h] === baseKey
            );
            if (baseHeader && rowData[baseHeader]) {
              if (!updatedPayload[baseKey]) {
                updatedPayload[baseKey] = rowData[baseHeader];
              }
              if (!updatedPayload.custom_fields[baseKey]) {
                updatedPayload.custom_fields[baseKey] = rowData[baseHeader];
              }
            }
          }
        }
      }
    }

    // 1. First identify all potential time range field pairs (base fields and their _until counterparts)
    const timeRangeFields = {};

    // Look for fields with _until suffix at root level
    for (const key of Object.keys(updatedPayload)) {
      if (key.endsWith("_until")) {
        const baseKey = key.replace(/_until$/, "");
        timeRangeFields[baseKey] = {
          baseKey: baseKey,
          untilKey: key,
          startTime: updatedPayload[baseKey],
          endTime: updatedPayload[key],
        };
      }
    }

    // Look for fields with _until suffix in custom_fields
    if (updatedPayload.custom_fields) {
      for (const key of Object.keys(updatedPayload.custom_fields)) {
        if (key.endsWith("_until")) {
          const baseKey = key.replace(/_until$/, "");
          // If already found at root level, update with custom_fields values if they exist
          if (timeRangeFields[baseKey]) {
            timeRangeFields[baseKey].startTime =
              timeRangeFields[baseKey].startTime ||
              updatedPayload.custom_fields[baseKey];
            timeRangeFields[baseKey].endTime =
              timeRangeFields[baseKey].endTime ||
              updatedPayload.custom_fields[key];
          } else {
            timeRangeFields[baseKey] = {
              baseKey: baseKey,
              untilKey: key,
              startTime: updatedPayload.custom_fields[baseKey],
              endTime: updatedPayload.custom_fields[key],
            };
          }
        }
      }
    }

    // 2. Process each time range field pair to ensure both start and end times are set
    for (const baseKey in timeRangeFields) {
      const field = timeRangeFields[baseKey];
      const untilKey = field.untilKey;

      // Format values if they exist
      // Check if this is a date range or time range based on the values
      // Special case: "1899-12-30" is Excel/Sheets' way of storing time-only values, so treat it as time
      const isExcelTime = (field.startTime && String(field.startTime).includes('1899-12-30')) ||
                         (field.endTime && String(field.endTime).includes('1899-12-30'));
      
      const isDateRange = !isExcelTime && (
        (field.startTime && (String(field.startTime).includes('T') || String(field.startTime).match(/^\d{4}-\d{2}-\d{2}/))) ||
        (field.endTime && (String(field.endTime).includes('T') || String(field.endTime).match(/^\d{4}-\d{2}-\d{2}/)))
      );
      
      let startTimeFormatted, endTimeFormatted;
      
      if (isDateRange) {
        // Format as dates for date range fields
        startTimeFormatted = field.startTime ? formatDateField(field.startTime) : null;
        endTimeFormatted = field.endTime ? formatDateField(field.endTime) : null;
      } else {
        // Format as times for time range fields
        startTimeFormatted = field.startTime ? formatTimeValue(field.startTime) : null;
        endTimeFormatted = field.endTime ? formatTimeValue(field.endTime) : null;
      }


      // If we have a start time but no end time, use the start time for both
      if (startTimeFormatted && !endTimeFormatted) {
        endTimeFormatted = startTimeFormatted;
      }

      // If we have an end time but no start time, use the end time for both
      if (!startTimeFormatted && endTimeFormatted) {
        startTimeFormatted = endTimeFormatted;
      }

      // Set values in both root and custom_fields if at least one time value exists
      if (startTimeFormatted || endTimeFormatted) {
        // Use whichever value is not null, prioritizing the formatted values
        const finalStartTime = startTimeFormatted || endTimeFormatted;
        const finalEndTime = endTimeFormatted || startTimeFormatted;

        // Set at root level
        updatedPayload[baseKey] = finalStartTime;
        updatedPayload[untilKey] = finalEndTime;

        // Set in custom_fields
        updatedPayload.custom_fields[baseKey] = finalStartTime;
        updatedPayload.custom_fields[untilKey] = finalEndTime;

      }
    }

    return updatedPayload;
  } catch (error) {
    return payload; // Return original payload if error occurs
  }
}

/**
 * Processes date and time fields in the data before sending to Pipedrive API
 * This is especially important for time range fields (_until pairs)
 * @param {Object} data - The data object to process
 * @param {Object} fieldDefinitions - Optional field definitions from Pipedrive
 * @return {Object} - The processed data object
 */
function processDateTimeFields(
  payload,
  rowData,
  fieldDefinitions,
  headerToFieldKeyMap
) {
  try {
    const fieldKeys = Object.keys(payload);

    // Create reverse mapping for fieldKey to header
    const fieldKeyToHeader = {};
    for (const header in headerToFieldKeyMap) {
      if (headerToFieldKeyMap.hasOwnProperty(header)) {
        fieldKeyToHeader[headerToFieldKeyMap[header]] = header;
      }
    }

    // Track which fields were processed as date/time
    const processedDateTimeFields = [];

    // First identify date/time range pairs directly from field keys (not headers)
    const timeRangePairs = {};

    // Check both root level and custom_fields for time range pairs
    const checkFields = (fields, source) => {
      for (const fieldKey of Object.keys(fields)) {
        if (fieldKey.endsWith("_until")) {
          const baseFieldKey = fieldKey.replace(/_until$/, "");
          if (
            fieldDefinitions[baseFieldKey] ||
            payload[baseFieldKey] ||
            (payload.custom_fields && payload.custom_fields[baseFieldKey])
          ) {
            timeRangePairs[baseFieldKey] = fieldKey;
          }
        }
      }
    };

    // Check root level fields
    checkFields(payload, "root level");

    // Check custom_fields
    if (payload.custom_fields) {
      checkFields(payload.custom_fields, "custom_fields");
    }

    // Process time range pairs
    Object.keys(timeRangePairs).forEach((baseKey) => {
      const untilKey = timeRangePairs[baseKey];

      // Find the headers that map to these field keys
      const baseHeader = fieldKeyToHeader[baseKey];
      const untilHeader = fieldKeyToHeader[untilKey];


      // Get values from payload or rowData
      let startValue = payload[baseKey];
      let endValue = payload[untilKey];

      // If values not in payload, try to get from rowData
      if (!startValue && baseHeader && rowData[baseHeader]) {
        startValue = rowData[baseHeader];
      }

      if (!endValue && untilHeader && rowData[untilHeader]) {
        endValue = rowData[untilHeader];
      }

      // Update both values in payload and custom_fields
      if (startValue || endValue) {
        if (!payload.custom_fields) payload.custom_fields = {};

        if (startValue) {
          payload[baseKey] = formatTimeValue(startValue);
          payload.custom_fields[baseKey] = formatTimeValue(startValue);
        }

        if (endValue) {
          payload[untilKey] = formatTimeValue(endValue);
          payload.custom_fields[untilKey] = formatTimeValue(endValue);
        }

        // Flag payload as having time range fields
        payload.__hasTimeRangeFields = true;
      }
    });

    return payload;
  } catch (error) {
    return payload;
  }
}

/**
 * Checks if headerFieldMap contains corresponding _until fields for time ranges
 * If a time range is detected, ensure both start and end times are included in the payload
 * @param {Object} payload - The payload to be sent to Pipedrive
 * @param {Object} rowData - The row data from the sheet
 * @param {Object} headerFieldMap - Mapping of headers to field keys
 * @return {Object} Updated payload with time range fields properly handled
 */
function ensureTimeRangePairs(payload, rowData, headerFieldMap) {
  // Create a map of all time range field pairs
  const timeRangePairs = {};

  // First find all fields with _until suffix in the header map
  for (const header in headerFieldMap) {
    const fieldKey = headerFieldMap[header];
    if (fieldKey && fieldKey.endsWith("_until")) {
      const baseFieldKey = fieldKey.replace(/_until$/, "");
      // Check if the base field also exists in the header map
      for (const baseHeader in headerFieldMap) {
        if (headerFieldMap[baseHeader] === baseFieldKey) {
          // Found a time range pair
          timeRangePairs[baseFieldKey] = {
            baseHeader: baseHeader,
            untilHeader: header,
            untilKey: fieldKey,
          };
          break;
        }
      }
    }
  }

  // Now ensure both start and end times are in the payload for each time range
  for (const baseKey in timeRangePairs) {
    const pair = timeRangePairs[baseKey];

    // Get the values from row data using headers
    const startValue = rowData[pair.baseHeader];
    const endValue = rowData[pair.untilHeader];


    // At least one value needs to be present for a time range
    if (startValue || endValue) {
      // Initialize custom_fields if needed
      if (!payload.custom_fields) payload.custom_fields = {};

      // Add start time to payload if present
      if (startValue) {
        payload[baseKey] = startValue;
        payload.custom_fields[baseKey] = startValue;
      }

      // Add end time to payload if present - ALWAYS include end time if start time is present
      if (endValue || startValue) {
        // If we have a start time but no end time, the API still needs the until field
        const effectiveEndValue = endValue || (startValue ? startValue : null);
        if (effectiveEndValue) {
          payload[pair.untilKey] = effectiveEndValue;
          payload.custom_fields[pair.untilKey] = effectiveEndValue;
        }
      }

      // Flag that we have time range fields - this will ensure direct API is used
      payload.__hasTimeRangeFields = true;
    }
  }

  return payload;
}

/**
 * Helper function to find the Sync Status column index
 * @param {Sheet} sheet - The sheet to search in
 * @param {string} sheetName - The name of the sheet
 * @return {number} The 0-based index of the Sync Status column, or -1 if not found
 */
function findSyncStatusColumn(sheet, sheetName) {
  try {
    // First try to find by header name
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        return i;
      }
    }

    // If not found by header, try to get from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const trackingKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const trackingColumn = scriptProperties.getProperty(trackingKey);

    if (trackingColumn) {
      // Convert column letter to index (0-based)
      const index = letterToColumn(trackingColumn) - 1;
      return index;
    }

    // If still not found, check if there's a column with sync status values
    const lastRow = Math.min(sheet.getLastRow(), 10); // Check first 10 rows max
    if (lastRow > 1) {
      for (let i = 0; i < headers.length; i++) {
        // Get values in this column for the first few rows
        const colValues = sheet
          .getRange(2, i + 1, lastRow - 1, 1)
          .getValues()
          .map((row) => row[0]);

        // Check if any cell contains a typical sync status value
        const containsSyncStatus = colValues.some(
          (value) =>
            value === "Modified" ||
            value === "Not modified" ||
            value === "Synced" ||
            value === "Error"
        );

        if (containsSyncStatus) {
          return i;
        }
      }
    }

    // Not found
    return -1;
  } catch (error) {
    return -1;
  }
}

/**
 * Debug function to check the original values stored for a sheet
 * This can be called manually to troubleshoot two-way sync issues
 * @param {string} sheetName - The name of the sheet to check
 */
function debugTwoWaySyncOriginalValues(sheetName) {
  try {
    if (!sheetName) {
      // Use active sheet if no sheet name provided
      sheetName = SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet()
        .getName();
    }


    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();

    // Check if two-way sync is enabled
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled =
      scriptProperties.getProperty(twoWaySyncEnabledKey) === "true";


    if (!twoWaySyncEnabled) {
      return;
    }

    // Get tracking column
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const trackingColumn = scriptProperties.getProperty(trackingColumnKey);


    // Get original data
    const originalDataKey = `ORIGINAL_DATA_${sheetName}`;
    const originalDataJson = scriptProperties.getProperty(originalDataKey);

    if (!originalDataJson) {
      return;
    }

    // Parse original data
    try {
      const originalData = JSON.parse(originalDataJson);
      const rowCount = Object.keys(originalData).length;


      // Log details for each row
      for (const rowKey in originalData) {
        const rowData = originalData[rowKey];
        const fieldCount = Object.keys(rowData).length;


        // Log each field and its original value
        for (const field in rowData) {
          const value = rowData[field];
        }
      }
    } catch (parseError) {
    }
  } catch (error) {
  }
}

// Find where data is written to the sheet in the syncPipedriveDataToSheet function
// Add logging around where values are extracted from Pipedrive data and written to the sheet

// Look for a section where it's processing each column and add logging for address fields:
// ... existing code ...

// Inside the item processing loop where it extracts values
function getFieldValue(item, fieldKey) {
  // Add logging for address-related fields
  if (fieldKey && (fieldKey === "address" || fieldKey.startsWith("address."))) {

    // Log the full address structure if available
    if (
      fieldKey === "address" &&
      item.address &&
      typeof item.address === "object"
    ) {
    }

    // Special handling for address subfields
    if (fieldKey.startsWith("address.") && item.address) {
      const addressComponent = fieldKey.replace("address.", "");

      if (item.address[addressComponent] !== undefined) {
        return item.address[addressComponent];
      } else {
      }
    }
  }

  // Add special handling for custom field address components
  if (
    fieldKey &&
    (fieldKey.includes("_subpremise") ||
      fieldKey.includes("_locality") ||
      fieldKey.includes("_formatted_address") ||
      fieldKey.includes("_street_number") ||
      fieldKey.includes("_route") ||
      fieldKey.includes("_admin_area") ||
      fieldKey.includes("_postal_code") ||
      fieldKey.includes("_country"))
  ) {

    // Custom field address components are stored directly at the item's top level
    if (item[fieldKey] !== undefined) {
      return item[fieldKey];
    } else {
    }
  }

  // Original getFieldValue logic continues below
  let value = null;

  try {
    // Special handling for nested keys like "org_id.name"
    if (fieldKey.includes(".")) {
      // Split the key into parts
      const keyParts = fieldKey.split(".");
      let currentObj = item;

      // Navigate through the object hierarchy
      for (let i = 0; i < keyParts.length; i++) {
        const part = keyParts[i];

        // Special handling for email.work, phone.mobile etc.
        if (
          (keyParts[0] === "email" || keyParts[0] === "phone") &&
          i === 1 &&
          isNaN(parseInt(part))
        ) {
          // This is a label-based lookup like email.work or phone.mobile
          const itemArray = currentObj; // The array of email/phone objects
          if (Array.isArray(itemArray)) {
            // Find the item with the matching label
            const foundItem = itemArray.find(
              (item) =>
                item &&
                item.label &&
                item.label.toLowerCase() === part.toLowerCase()
            );

            // If found, use its value
            if (foundItem) {
              currentObj = foundItem;
              continue;
            } else {
              // If label not found, check if we're looking for primary
              if (part.toLowerCase() === "primary") {
                const primaryItem = itemArray.find(
                  (item) => item && item.primary
                );
                if (primaryItem) {
                  currentObj = primaryItem;
                  continue;
                }
              }
              // Fallback to first item if available
              if (itemArray.length > 0) {
                currentObj = itemArray[0];
                continue;
              }
            }
          }
          // If we get here, we couldn't find a matching item
          currentObj = null;
          break;
        }

        // Handle array access - check if part is a number
        if (!isNaN(parseInt(part))) {
          const index = parseInt(part);
          if (Array.isArray(currentObj) && index < currentObj.length) {
            currentObj = currentObj[index];
          } else {
            currentObj = null;
            break;
          }
        } else {
          // Regular object property access
          if (
            currentObj &&
            typeof currentObj === "object" &&
            currentObj[part] !== undefined
          ) {
            currentObj = currentObj[part];
          } else {
            currentObj = null;
            break;
          }
        }
      }

      value = currentObj;
    } else {
      // Direct property access
      value = item[fieldKey];
    }
  } catch (error) {
    value = null;
  }

  return value;
}

/**
 * Checks if a field is a date field
 * @param {string} fieldKey - The field key to check
 * @param {string} entityType - The entity type for context-specific checks
 * @return {boolean} True if the field is a date field
 */
function isDateField(fieldKey, entityType) {
  // Common date fields across all entity types
  const commonDateFields = [
    "add_time",
    "update_time",
    "created_at",
    "updated_at",
    "last_activity_date",
    "next_activity_date",
    "due_date",
    "expected_close_date",
    "won_time",
    "lost_time",
    "close_time",
    "last_incoming_mail_time",
    "last_outgoing_mail_time",
    "start_date",
    "end_date",
    "date",
  ];

  // Check if it's a known date field
  if (commonDateFields.includes(fieldKey)) {
    return true;
  }

  // Entity-specific date fields
  if (entityType === ENTITY_TYPES.DEALS) {
    const dealDateFields = [
      "close_date",
      "lost_reason_changed_time",
      "dropped_time",
      "rotten_time",
    ];
    if (dealDateFields.includes(fieldKey)) {
      return true;
    }
  } else if (entityType === ENTITY_TYPES.ACTIVITIES) {
    const activityDateFields = [
      "due_date",
      "due_time",
      "marked_as_done_time",
      "last_notification_time",
    ];
    if (activityDateFields.includes(fieldKey)) {
      return true;
    }
  }

  // Check if it looks like a date field by name
  return (
    fieldKey.endsWith("_date") ||
    fieldKey.endsWith("_time") ||
    fieldKey.includes("date_") ||
    fieldKey.includes("time_")
  );
}

/**
 * Validates and converts ID fields to ensure they are integers
 * @param {Object} data - The data object containing fields to validate
 * @param {string} fieldName - The name of the field to validate
 * @returns {boolean} True if the field was valid or fixed, false if it was removed
 */
function validateIdField(data, fieldName) {
  if (data[fieldName] !== undefined) {
    // Check if the field is not an integer
    if (
      isNaN(parseInt(data[fieldName])) ||
      !/^\d+$/.test(String(data[fieldName]))
    ) {
      // Remove the invalid field from the request to prevent API errors
      delete data[fieldName];
      return false;
    } else {
      // Convert to integer if it's a valid number
      data[fieldName] = parseInt(data[fieldName]);
      return true;
    }
  }
  return true; // Field not present, so no validation needed
}

/**
 * Gets team-aware column preferences
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of column keys
 */
SyncService.getTeamAwareColumnPreferences = function (entityType, sheetName) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const userEmail = Session.getEffectiveUser().getEmail();
    let columnsJson = null;
    let usedKey = "";

    // 1. Try Team Key
    try {
      const userTeam = getUserTeam(userEmail); // Assumes getUserTeam is available
      if (userTeam && userTeam.teamId) {
        const teamKey = `COLUMNS_${sheetName}_${entityType}_TEAM_${userTeam.teamId}`;
        columnsJson = properties.getProperty(teamKey);
        if (columnsJson) {
          usedKey = teamKey;
        } else {
        }
      } else {
      }
    } catch (teamError) {
    }

    // 2. Try Personal Key if Team Key failed
    if (!columnsJson) {
      const personalKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
      columnsJson = properties.getProperty(personalKey);
      if (columnsJson) {
        usedKey = personalKey;
      } else {
      }
    }

    // 3. Log raw JSON and attempt parse
    if (columnsJson) {
      try {
        const savedColumns = JSON.parse(columnsJson);
        // Log details of the first column to check for customName
        if (savedColumns.length > 0) {
        }
        return savedColumns;
      } catch (parseError) {
        return []; // Return empty on parse error
      }
    } else {
      return []; // Return empty if no JSON found
    }
  } catch (error) {
    return []; // Return empty on general error
  }
};

/**
 * Saves team-aware column preferences - wrapper for UI.gs function
 * @param {Array} columns - Column objects or keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
SyncService.saveTeamAwareColumnPreferences = function (
  columns,
  entityType,
  sheetName
) {
  try {
    // Keep full column objects intact to preserve names
    // Call the function in UI.gs that handles saving to both storage locations
    return UI.saveTeamAwareColumnPreferences(columns, entityType, sheetName);
  } catch (e) {

    // Fallback to local implementation if UI.saveTeamAwareColumnPreferences fails
    const scriptProperties = PropertiesService.getScriptProperties();
    const key = `COLUMNS_${sheetName}_${entityType}`;

    // Store the full column objects
    scriptProperties.setProperty(key, JSON.stringify(columns));
  }
};

/**
 * Validate that ID fields contain numeric values
 * @param {Object} data - The data object
 * @param {string} fieldName - Field name to validate
 */
function validateIdField(data, fieldName) {
  // Skip if field doesn't exist
  if (!data || data[fieldName] === undefined || data[fieldName] === null) {
    return;
  }

  // Handle common ID fields that must be numeric
  const numericIdFields = [
    "owner_id",
    "person_id",
    "org_id",
    "organization_id",
    "pipeline_id",
    "stage_id",
    "user_id",
    "creator_user_id",
    "category_id",
    "tax_id",
    "unit_id",
  ];

  // If this is a field that requires numeric ID
  if (numericIdFields.includes(fieldName)) {
    try {
      const idValue = data[fieldName];
      // If it's already a number, we're good
      if (typeof idValue === "number") {
        return;
      }

      // If it's a string, try to convert to number
      if (typeof idValue === "string") {
        const numericId = parseInt(idValue, 10);
        // If it's a valid number, update the field
        if (!isNaN(numericId) && /^\d+$/.test(idValue.trim())) {
          data[fieldName] = numericId;
        } else {
          // Not a valid number, delete the field to prevent API errors
          delete data[fieldName];
        }
      } else {
        // Not a number or string, delete the field
        delete data[fieldName];
      }
    } catch (e) {
      // If any errors, delete the field to be safe
      delete data[fieldName];
    }
  }
}

/**
 * Saves settings to script properties
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Save global settings (only if provided)
  if (apiKey) scriptProperties.setProperty("PIPEDRIVE_API_KEY", apiKey);
  if (subdomain) scriptProperties.setProperty("PIPEDRIVE_SUBDOMAIN", subdomain);

  // Save sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;

  scriptProperties.setProperty(sheetFilterIdKey, filterId);
  scriptProperties.setProperty(sheetEntityTypeKey, entityType);
  scriptProperties.setProperty("SHEET_NAME", sheetName);
}

/**
 * Sets up the onEdit trigger to track changes for two-way sync
 */
function setupOnEditTrigger() {
  try {
    // First, remove any existing onEdit triggers to avoid duplicates
    removeOnEditTrigger();

    // Then create a new trigger
    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  } catch (e) {
  }
}

/**
 * Removes the onEdit trigger
 */
function removeOnEditTrigger() {
  try {
    // Get all triggers
    const triggers = ScriptApp.getProjectTriggers();

    // Find and delete the onEdit trigger
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      if (trigger.getHandlerFunction() === "onEdit") {
        ScriptApp.deleteTrigger(trigger);
        return;
      }
    }

  } catch (e) {
  }
}

/**
 * Logs debug information about the Pipedrive data
 */
function logDebugInfo() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetName =
    scriptProperties.getProperty("SHEET_NAME") || DEFAULT_SHEET_NAME;

  // Get sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;

  const filterId = scriptProperties.getProperty(sheetFilterIdKey) || "";
  const entityType =
    scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

  // Show which column selections are available for the current entity type and sheet
  const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
  const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);

  if (savedColumnsJson) {
    try {
      const selectedColumns = JSON.parse(savedColumnsJson);
    } catch (e) {
    }
  } else {
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

    // Log complete raw deal data for inspection

    // Extract all fields including nested ones
    const allFields = {};

    // Recursive function to extract all fields with their paths
    function extractAllFields(obj, path = "") {
      if (!obj || typeof obj !== "object") return;

      if (Array.isArray(obj)) {
        // For arrays, log the length and extract fields from first item if exists
        if (obj.length > 0 && typeof obj[0] === "object") {
          extractAllFields(obj[0], `${path}[0]`);
        }
      } else {
        // For objects, extract each property
        for (const key in obj) {
          const value = obj[key];
          const newPath = path ? `${path}.${key}` : key;

          if (value === null) {
            allFields[newPath] = "null";
            continue;
          }

          const type = typeof value;

          if (type === "object") {
            if (Array.isArray(value)) {
              allFields[newPath] = `array[${value.length}]`;

              // Special case for custom fields with options
              if (
                key === "options" &&
                value.length > 0 &&
                value[0] &&
                value[0].label
              ) {
              }

              // For small arrays with objects, recursively extract from the first item
              if (value.length > 0 && typeof value[0] === "object") {
                extractAllFields(value[0], `${newPath}[0]`);
              }
            } else {
              allFields[newPath] = "object";
              extractAllFields(value, newPath);
            }
          } else {
            allFields[newPath] = type;

            // Log a preview of the value unless it's a string longer than 50 chars
            const preview =
              type === "string" && value.length > 50
                ? value.substring(0, 50) + "..."
                : value;

          }
        }
      }
    }

    // Start extraction from the top level
    extractAllFields(sampleItem);

    // Specifically focus on custom fields section if it exists
    if (sampleItem.custom_fields) {
      for (const key in sampleItem.custom_fields) {
        const field = sampleItem.custom_fields[key];
        const fieldType = typeof field;

        if (fieldType === "object" && Array.isArray(field)) {
          // Check if this is a multiple options field
          if (field.length > 0 && field[0] && field[0].label) {
          }
        } else {
          const preview =
            fieldType === "string" && field.length > 50
              ? field.substring(0, 50) + "..."
              : field;
        }
      }
    }

    // Count unique fields
    const fieldPaths = Object.keys(allFields).sort();

    // Log all field paths in alphabetical order for easy lookup
    fieldPaths.forEach((path) => {
    });
  } else {
  }
}

/**
 * Detects if columns have been shifted, renamed, or reordered in a sheet
 * and updates the header-to-field mapping accordingly
 * @return {boolean} True if shift detected and fixed, false otherwise
 */
function detectColumnShifts() {
  try {
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();

    // Get sheet-specific entity type
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey);

    // If no entity type set for this sheet, exit
    if (!entityType) {
      return false;
    }

    // Get the current headers from the sheet
    const headers = activeSheet
      .getRange(1, 1, 1, activeSheet.getLastColumn())
      .getValues()[0];

    // Make sure we have a valid header-to-field mapping
    const headerToFieldMap = ensureHeaderFieldMapping(
      activeSheetName,
      entityType
    );

    // Track if we've updated the mapping
    let updated = false;

    // Check if the current headers match the stored mapping
    // We'll log each header to see if it exists in our mapping

    // Count headers found in mapping
    let headersFoundInMapping = 0;

    // Identify headers not found in mapping
    const headersNotInMapping = [];

    headers.forEach((header) => {
      if (header && typeof header === "string") {
        if (headerToFieldMap[header]) {
          // This header is already in our mapping
          headersFoundInMapping++;
        } else {
          // This header is not in our mapping - could be a renamed column
          headersNotInMapping.push(header);

          // Try case-insensitive match
          const headerLower = header.toLowerCase();
          let matchFound = false;

          // First look for exact case-insensitive match
          for (const mappedHeader in headerToFieldMap) {
            if (mappedHeader.toLowerCase() === headerLower) {
              // Found a case-insensitive match, update the mapping
              const fieldKey = headerToFieldMap[mappedHeader];
              headerToFieldMap[header] = fieldKey;
              updated = true;
              matchFound = true;
              break;
            }
          }

          // If no exact match, try normalized match (remove spaces, punctuation, etc.)
          if (!matchFound) {
            const normalizedHeader = headerLower
              .replace(/\s+/g, "") // Remove all whitespace
              .replace(/[^\w\d]/g, ""); // Remove non-alphanumeric characters

            for (const mappedHeader in headerToFieldMap) {
              const normalizedMappedHeader = mappedHeader
                .toLowerCase()
                .replace(/\s+/g, "")
                .replace(/[^\w\d]/g, "");

              if (normalizedHeader === normalizedMappedHeader) {
                // Found a normalized match
                const fieldKey = headerToFieldMap[mappedHeader];
                headerToFieldMap[header] = fieldKey;
                updated = true;
                break;
              }
            }
          }
        }
      }
    });


    if (headersNotInMapping.length > 0) {
    }

    // If we updated the mapping, save it
    if (updated) {
      const mappingKey = `HEADER_TO_FIELD_MAP_${activeSheetName}_${entityType}`;
      scriptProperties.setProperty(
        mappingKey,
        JSON.stringify(headerToFieldMap)
      );
      return true;
    }

    return false;
  } catch (error) {
    return false;
  }
}

/**
 * Ensures that a valid header-to-field mapping exists for the given sheet and entity type
 * @param {string} sheetName - The name of the sheet
 * @param {string} entityType - The entity type (deals, persons, etc.)
 * @return {Object} The header-to-field mapping
 */
function ensureHeaderFieldMapping(sheetName, entityType) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const mappingKey = `HEADER_TO_FIELD_MAP_${sheetName}_${entityType}`;
    const mappingJson = scriptProperties.getProperty(mappingKey);
    let headerToFieldKeyMap = {};

    if (mappingJson) {
      try {
        headerToFieldKeyMap = JSON.parse(mappingJson);
      } catch (e) {
        headerToFieldKeyMap = {};
      }
    }

    // If mapping exists and has entries, use it as a base but check for missing address components
    if (Object.keys(headerToFieldKeyMap).length > 0) {
      // Check if we need to update any address component mappings (fix for admin_area_level_2 with trailing space)
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (sheet) {
        const headers = sheet
          .getRange(1, 1, 1, sheet.getLastColumn())
          .getValues()[0];

        // Find all address field IDs in the mapping
        const addressFieldIds = new Set();
        for (const [header, fieldKey] of Object.entries(headerToFieldKeyMap)) {
          // Check if this is an address field key (non-component)
          if (
            fieldKey &&
            !fieldKey.includes("_locality") &&
            !fieldKey.includes("_route") &&
            !fieldKey.includes("_street_number") &&
            !fieldKey.includes("_postal_code") &&
            !fieldKey.includes("_admin_area_level_1") &&
            !fieldKey.includes("_admin_area_level_2") &&
            !fieldKey.includes("_country")
          ) {
            // Look for address component headers that reference this address field
            for (const h of headers) {
              if (h && typeof h === "string") {
                const headerTrimmed = h.trim();
                // Check if this header is for an address component of the current field
                if (
                  headerTrimmed.startsWith(header) &&
                  headerTrimmed.includes(" - ")
                ) {
                  const componentPart = headerTrimmed.split(" - ")[1].trim();

                  // Map the component based on its name
                  let component = "";
                  if (componentPart.includes("City")) component = "locality";
                  else if (componentPart.includes("Street Name"))
                    component = "route";
                  else if (componentPart.includes("Street Number"))
                    component = "street_number";
                  else if (
                    componentPart.includes("ZIP") ||
                    componentPart.includes("Postal")
                  )
                    component = "postal_code";
                  else if (
                    componentPart.includes("State") ||
                    componentPart.includes("Province")
                  )
                    component = "admin_area_level_1";
                  else if (
                    componentPart.includes("County") ||
                    componentPart.includes("Admin Area Level")
                  )
                    component = "admin_area_level_2";
                  else if (componentPart.includes("Country"))
                    component = "country";

                  if (component) {
                    const componentFieldKey = `${fieldKey}_${component}`;
                    if (!headerToFieldKeyMap[headerTrimmed]) {
                      headerToFieldKeyMap[headerTrimmed] = componentFieldKey;

                      // Also add mapping for header with possible trailing space
                      if (h !== headerTrimmed) {
                        headerToFieldKeyMap[h] = componentFieldKey;
                      }
                    }

                    // Double-check if we have the "Admin Area Level" field with a space
                    if (component === "admin_area_level_2" && h.endsWith(" ")) {
                      const exactHeader = h;
                      headerToFieldKeyMap[exactHeader] = componentFieldKey;
                    }
                  }
                }
              }
            }
          }
        }

        // Save any updates to the mapping
        if (Object.keys(headerToFieldKeyMap).length > 0) {
          scriptProperties.setProperty(
            mappingKey,
            JSON.stringify(headerToFieldKeyMap)
          );
        }
      }

      return headerToFieldKeyMap;
    }


    // Get column preferences for this sheet/entity
    const columnConfig = getColumnPreferences(entityType, sheetName);

    if (!columnConfig || columnConfig.length === 0) {
      // Try with SyncService method if available
      try {
        if (
          typeof SyncService !== "undefined" &&
          typeof SyncService.getTeamAwareColumnPreferences === "function"
        ) {
          const teamColumns = SyncService.getTeamAwareColumnPreferences(
            entityType,
            sheetName
          );
          if (teamColumns && teamColumns.length > 0) {

            // Create mapping from these columns
            teamColumns.forEach((col) => {
              if (col.key) {
                const displayName =
                  col.customName || col.name || formatColumnName(col.key);
                headerToFieldKeyMap[displayName] = col.key;

                // For address fields, also add mappings for their components
                if (col.type === "address" || col.key.endsWith("address")) {
                  const baseHeader = displayName;

                  // Add mappings for common address components
                  const components = [
                    "locality",
                    "route",
                    "street_number",
                    "postal_code",
                    "admin_area_level_1",
                    "admin_area_level_2",
                    "country",
                  ];

                  components.forEach((component) => {
                    const componentFieldKey = `${col.key}_${component}`;
                    let componentHeader = "";

                    // Format component header based on component type
                    switch (component) {
                      case "locality":
                        componentHeader = `${baseHeader} - City`;
                        break;
                      case "route":
                        componentHeader = `${baseHeader} - Street Name`;
                        break;
                      case "street_number":
                        componentHeader = `${baseHeader} - Street Number`;
                        break;
                      case "postal_code":
                        componentHeader = `${baseHeader} - ZIP/Postal Code`;
                        break;
                      case "admin_area_level_1":
                        componentHeader = `${baseHeader} - State/Province`;
                        break;
                      case "admin_area_level_2":
                        componentHeader = `${baseHeader} - Admin Area Level`;
                        break;
                      case "country":
                        componentHeader = `${baseHeader} - Country`;
                        break;
                    }

                    if (componentHeader) {
                      headerToFieldKeyMap[componentHeader] = componentFieldKey;

                      // Also add with trailing space for admin_area_level_2 to handle the specific issue
                      if (component === "admin_area_level_2") {
                        const spacedHeader = `${baseHeader} - Admin Area Level `;
                        headerToFieldKeyMap[spacedHeader] = componentFieldKey;
                      }
                    }
                  });
                }
              }
            });

            // Save the mapping
            if (Object.keys(headerToFieldKeyMap).length > 0) {
              scriptProperties.setProperty(
                mappingKey,
                JSON.stringify(headerToFieldKeyMap)
              );
              return headerToFieldKeyMap;
            }
          }
        }
      } catch (teamError) {
      }

      // If still no columns, use default columns
      const defaultColumns = getDefaultColumns(entityType);

      // Create mapping from default columns
      defaultColumns.forEach((col) => {
        // For default columns, key and display name are the same
        const key = typeof col === "object" ? col.key : col;
        const displayName = formatColumnName(key);
        headerToFieldKeyMap[displayName] = key;
      });
    } else {
      // Create mapping from column config

      columnConfig.forEach((col) => {
        if (col.key) {
          const displayName =
            col.customName || col.name || formatColumnName(col.key);
          headerToFieldKeyMap[displayName] = col.key;

          // For address fields, also add mappings for their components
          if (col.type === "address" || col.key.endsWith("address")) {
            const baseHeader = displayName;

            // Add mappings for common address components
            const components = [
              "locality",
              "route",
              "street_number",
              "postal_code",
              "admin_area_level_1",
              "admin_area_level_2",
              "country",
            ];

            components.forEach((component) => {
              const componentFieldKey = `${col.key}_${component}`;
              let componentHeader = "";

              // Format component header based on component type
              switch (component) {
                case "locality":
                  componentHeader = `${baseHeader} - City`;
                  break;
                case "route":
                  componentHeader = `${baseHeader} - Street Name`;
                  break;
                case "street_number":
                  componentHeader = `${baseHeader} - Street Number`;
                  break;
                case "postal_code":
                  componentHeader = `${baseHeader} - ZIP/Postal Code`;
                  break;
                case "admin_area_level_1":
                  componentHeader = `${baseHeader} - State/Province`;
                  break;
                case "admin_area_level_2":
                  componentHeader = `${baseHeader} - Admin Area Level`;
                  break;
                case "country":
                  componentHeader = `${baseHeader} - Country`;
                  break;
              }

              if (componentHeader) {
                headerToFieldKeyMap[componentHeader] = componentFieldKey;

                // Also add with trailing space for admin_area_level_2 to handle the specific issue
                if (component === "admin_area_level_2") {
                  const spacedHeader = `${baseHeader} - Admin Area Level `;
                  headerToFieldKeyMap[spacedHeader] = componentFieldKey;
                }
              }
            });
          }
        }
      });
    }

    // Also add common field mappings that might not be in column config
    const commonMappings = {
      ID: "id",
      "Pipedrive ID": "id",
      "Deal Title": "title",
      Organization: "org_id",
      "Organization Name": "org_id.name",
      Person: "person_id",
      "Person Name": "person_id.name",
      Owner: "owner_id",
      "Owner Name": "owner_id.name",
      Value: "value",
      Stage: "stage_id",
      Pipeline: "pipeline_id",
    };

    Object.keys(commonMappings).forEach((displayName) => {
      if (!headerToFieldKeyMap[displayName]) {
        headerToFieldKeyMap[displayName] = commonMappings[displayName];
      }
    });

    // Save the mapping
    scriptProperties.setProperty(
      mappingKey,
      JSON.stringify(headerToFieldKeyMap)
    );

    return headerToFieldKeyMap;
  } catch (error) {
    return {}; // Return empty mapping in case of error
  }
}

/**
 * Gets default columns for a given entity type
 * @param {string} entityType - The entity type
 * @return {Array} Array of default column keys
 */
function getDefaultColumns(entityType) {
  // Use the DEFAULT_COLUMNS constant if it exists
  if (typeof DEFAULT_COLUMNS !== "undefined") {
    if (DEFAULT_COLUMNS[entityType]) {
      return DEFAULT_COLUMNS[entityType];
    } else if (DEFAULT_COLUMNS.COMMON) {
      return DEFAULT_COLUMNS.COMMON;
    }
  }

  // Fallback default columns by entity type
  switch (entityType) {
    case "deals":
      return [
        "id",
        "title",
        "value",
        "currency",
        "status",
        "stage_id",
        "pipeline_id",
        "person_id",
        "org_id",
        "owner_id",
      ];
    case "persons":
      return ["id", "name", "email", "phone", "org_id", "owner_id"];
    case "organizations":
      return ["id", "name", "address", "owner_id"];
    case "activities":
      return [
        "id",
        "subject",
        "type",
        "due_date",
        "due_time",
        "person_id",
        "org_id",
        "deal_id",
        "owner_id",
      ];
    case "leads":
      return [
        "id",
        "title",
        "value",
        "person_id",
        "organization_id",
        "owner_id",
      ];
    case "products":
      return ["id", "name", "code", "unit", "price", "owner_id"];
    default:
      return ["id", "name", "owner_id"];
  }
}

/**
 * Filters out read-only fields from the data before sending to Pipedrive API
 * @param {Object} data - The data object to filter
 * @param {string} entityType - The entity type
 * @return {Object} Filtered data object
 */
function filterReadOnlyFields(data, entityType) {
  if (!data) return data;

  const filteredData = {};
  // Initialize custom_fields object
  filteredData.custom_fields = {};

  // Fields ending with these patterns are typically read-only
  const readOnlyPatterns = [
    /\.name$/,
    /\.email$/,
    /\.value$/,
    /\.active_flag$/,
    /\.phone$/,
    /\.pic_hash$/,
  ];
  const allowedNameFields = ["first_name", "last_name"]; // These are explicitly allowed

  // Custom read-only fields by entity type
  const readOnlyFields = [
    // User/owner related
    "owner_name",
    "owner_email",
    "user_id.email",
    "user_id.name",
    "user_id.active_flag",
    "creator_user_id.email",
    "creator_user_id.name",

    // Organization/Person related
    "org_name",
    "person_name",

    // System fields
    "cc_email",
    "weighted_value",
    "formatted_value",
    "source_channel",
    "source_origin",
    "origin", // Added based on the error
    "channel",

    // Activities and stats
    "next_activity_time",
    "next_activity_id",
    "last_activity_id",
    "last_activity_date",
    "activities_count",
    "done_activities_count",
    "undone_activities_count",
    "files_count",
    "notes_count",
    "followers_count",

    // Timestamps and system IDs
    "add_time",
    "update_time",
    "stage_order_nr",
    "rotten_time",
  ];

  // Track nested relationship fields
  const relationships = {
    "owner_id.name": "owner_id",
    "org_id.name": "org_id",
    "person_id.name": "person_id",
    "creator_user_id.name": "creator_user_id",
    "user_id.name": "user_id",
    "deal_id.name": "deal_id",
    "deal_id.title": "deal_id",
    "stage_id.name": "stage_id",
    "pipeline_id.name": "pipeline_id",
  };

  // Pattern for custom field IDs (long hex string)
  const customFieldPattern = /^[a-f0-9]{20,}$/i;
  // Pattern for custom field components (like _locality, _street, etc.)
  const customFieldComponentPattern = /^[a-f0-9]{20,}_[a-z_]+$/i;

  // First, extract IDs from relationship fields if possible
  for (const nestedField in relationships) {
    const parentField = relationships[nestedField];

    // If we have the relationship field but not the parent field, look up the ID
    if (data[nestedField] && !data[parentField]) {

      // Try to extract ID from the name based on entity type and field
      // This is complex and may need to be enhanced with API lookups
      if (nestedField.includes("owner_id") || nestedField.includes("user_id")) {
        // For user-related fields, try to find a user ID matching this name
        const userId = lookupUserIdByName(data[nestedField]);
        if (userId) {
          data[parentField] = userId;
        }
      }
      // Add other entity lookups as needed
    }
  }

  // First pass - identify custom fields and organize data
  const customFields = {};
  const addressComponents = {};
  const addressValues = {};

  for (const key in data) {
    // First, handle custom fields with their special format
    if (customFieldPattern.test(key)) {
      // It's a base custom field (could be an address field base value)

      // Store both in customFields for regular fields and in addressValues for address fields
      customFields[key] = data[key];
      addressValues[key] = data[key]; // Store the full address value
      continue;
    } else if (customFieldComponentPattern.test(key)) {
      // It's a custom field component (like address_locality)
      const parts = key.match(/^([a-f0-9]{20,})_(.+)$/i);
      if (parts && parts.length === 3) {
        const fieldId = parts[1];
        const component = parts[2];

        // Special handling for admin_area_level_2 to ensure it's not filtered out
        if (component === "admin_area_level_2") {

          // Skip the read-only check for admin_area_level_2
          if (!addressComponents[fieldId]) {
            addressComponents[fieldId] = {};
          }
          addressComponents[fieldId][component] = data[key];
          continue;
        }

        // Skip if this is in read-only fields
        if (readOnlyFields.includes(key)) {
          continue;
        }

        // Store address components separately for proper handling
        if (!addressComponents[fieldId]) {
          addressComponents[fieldId] = {};
        }
        addressComponents[fieldId][component] = data[key];
      }
      continue;
    }
  }

  // Now filter out the read-only fields
  for (const key in data) {
    // Skip custom fields and address components - we handled those separately
    if (customFieldPattern.test(key) || customFieldComponentPattern.test(key)) {
      continue;
    }

    // Skip if this is a known read-only field
    if (readOnlyFields.includes(key)) {
      continue;
    }

    // Check if this field matches any read-only pattern
    let isReadOnly = false;
    for (const pattern of readOnlyPatterns) {
      if (pattern.test(key) && !allowedNameFields.includes(key)) {
        isReadOnly = true;
        break;
      }
    }

    if (isReadOnly) {
      const parentField = Object.keys(relationships).find((rel) => rel === key);

      if (parentField) {
      } else {
      }
      continue;
    }

    // Special handling for nested properties
    if (key.includes(".")) {
      // Only allow specific nested fields that are known to be updatable
      const allowedNestedFields = [
        "address.street",
        "address.city",
        "address.state",
        "address.postal_code",
        "address.country",
      ];

      if (!allowedNestedFields.includes(key)) {
        continue;
      }
    }

    // Include this field in the filtered data
    filteredData[key] = data[key];
  }

  // Handle custom fields
  if (Object.keys(customFields).length > 0) {
    for (const fieldId in customFields) {
      // Special handling for address fields - preserve them as objects
      if (
        (typeof customFields[fieldId] === "object" &&
          customFields[fieldId] !== null) ||
        (addressComponents[fieldId] &&
          Object.keys(addressComponents[fieldId]).length > 0)
      ) {
        // This appears to be an address field or another object-type field
        // Don't overwrite it - we'll handle it in the address components section
        continue;
      }

      // Regular custom fields
      filteredData.custom_fields[fieldId] = customFields[fieldId];
    }
  }

  // Handle address components - for address fields, we may need special handling
  if (Object.keys(addressComponents).length > 0) {

    // For address fields, Pipedrive expects an object structure with a 'value' property
    for (const fieldId in addressComponents) {
      // Create an address object with components
      const addressObject = addressComponents[fieldId];

      // First, try to fetch current address data from Pipedrive
      let currentAddressData = {};
      if (entityId) {
        currentAddressData = getCurrentAddressData(
          entityType,
          entityId,
          fieldId
        );
      }

      // Check if we have a direct update to the full address field
      const hasFullAddressUpdate =
        addressValues[fieldId] !== undefined && addressValues[fieldId] !== "";

      // If we have a full address update, prioritize it but still include components
      if (hasFullAddressUpdate) {
        // Prioritize the full address value - ensure it's a string
        addressObject.value = String(addressValues[fieldId]);

        // First preserve any current components that aren't being updated
        if (currentAddressData && typeof currentAddressData === "object") {
          for (const component in currentAddressData) {
            if (component !== "value" && !addressObject[component]) {
              addressObject[component] = String(currentAddressData[component]);
            }
          }
        }

        // Then ensure all components are strings
        for (const component in addressObject) {
          if (component !== "value") {
            addressObject[component] = String(addressObject[component]);
          }
        }
      }
      // If we don't have a full address update but have component changes, build the address from components
      else if (Object.keys(addressObject).length > 0 && !addressObject.value) {
        // First, add all existing components and value from Pipedrive
        if (currentAddressData && typeof currentAddressData === "object") {
          // Use the current full address value
          if (currentAddressData.value) {
            addressObject.value = String(currentAddressData.value);
          }

          // Add all current components that aren't already being updated
          for (const component in currentAddressData) {
            if (component !== "value" && !addressObject[component]) {
              addressObject[component] = String(currentAddressData[component]);
            }
          }
        }
        // Fall back to address value if no current data but we have it
        else if (addressValues[fieldId]) {
          addressObject.value = String(addressValues[fieldId]);
        }
        // If no current data or original value, construct a new address from components
        else {
          // Construct a new address from all available components (both preserved and updated)
          let newAddress = "";

          // Start with street number and route (street name)
          if (addressObject.street_number && addressObject.route) {
            newAddress = `${addressObject.street_number} ${addressObject.route}`;
          } else if (addressObject.route) {
            newAddress = addressObject.route;
          }

          // Add city (locality)
          if (addressObject.locality) {
            if (newAddress) newAddress += `, ${addressObject.locality}`;
            else newAddress = addressObject.locality;
          }

          // Add state/province with comma
          if (addressObject.admin_area_level_1) {
            if (newAddress)
              newAddress += `, ${addressObject.admin_area_level_1}`;
            else newAddress = addressObject.admin_area_level_1;
          }

          // Add postal code (no comma before postal code)
          if (addressObject.postal_code) {
            if (newAddress) newAddress += ` ${addressObject.postal_code}`;
            else newAddress = addressObject.postal_code;
          }

          // Add country with comma
          if (addressObject.country) {
            if (newAddress) newAddress += `, ${addressObject.country}`;
            else newAddress = addressObject.country;
          }

          // Set the value property to the constructed address
          if (newAddress) {
            addressObject.value = newAddress;
          }
        }

        // Ensure all components are strings
        for (const component in addressObject) {
          if (component !== "value") {
            addressObject[component] = String(addressObject[component]);
          }
        }
      }

      // Add the complete address object to custom_fields using the field ID as the key
      // IMPORTANT: We must send the address as an object, not a string
      filteredData.custom_fields[fieldId] = {
        ...addressObject,
      };

      // Make sure admin_area_level_2 is included directly in the object
      if (addressObject.admin_area_level_2) {
        filteredData.custom_fields[fieldId].admin_area_level_2 = String(
          addressObject.admin_area_level_2
        );
      }

      // Final check to ensure all values are strings in the address object
      for (const component in filteredData.custom_fields[fieldId]) {
        if (
          component !== "value" &&
          filteredData.custom_fields[fieldId][component] !== undefined
        ) {
          filteredData.custom_fields[fieldId][component] = String(
            filteredData.custom_fields[fieldId][component]
          );
        }
      }

      // Log the final object to confirm it's properly structured
    }
  }

  // Remove custom_fields if empty
  if (Object.keys(filteredData.custom_fields).length === 0) {
    delete filteredData.custom_fields;
  }

  return filteredData;
}

/**
 * Helper function to look up a user ID by name
 * This is a placeholder - in a full implementation, you might cache user data
 * @param {string} name - The name to look up
 * @return {number|null} - The user ID if found, or null
 */
function lookupUserIdByName(name) {
  // In a real implementation, this would query the Pipedrive API or use cached data
  // For now, we'll just log that this function was called
  return null; // Placeholder, no lookup performed
}

// Add this function before the pushChangesToPipedrive function
/**
 * Properly handles address components in a data object
 * This ensures components like admin_area_level_2 are included in their parent address object
 * rather than as separate fields at the root level
 * @param {Object} data - The data object to process
 * @return {Object} The processed data object with address components properly structured
 */
function handleAddressComponents(data) {
  if (!data) return data;

  const result = JSON.parse(JSON.stringify(data)); // Deep clone to avoid modifying the original

  // Special handling for the problematic admin_area_level_2 field
  // This ensures it gets processed even with unusual naming
  for (const key in result) {
    if (key.includes("_admin_area_level_2")) {

      // Extract the field ID part
      const fieldIdMatch = key.match(/^([a-f0-9]{20,})_admin_area_level_2$/i);
      if (fieldIdMatch && fieldIdMatch[1]) {
        const fieldId = fieldIdMatch[1];

        // Initialize custom_fields if needed
        if (!result.custom_fields) {
          result.custom_fields = {};
        }

        // Ensure the parent address field exists in custom_fields
        if (
          !result.custom_fields[fieldId] ||
          typeof result.custom_fields[fieldId] !== "object"
        ) {
          // Initialize with existing value if available, otherwise empty
          result.custom_fields[fieldId] = {
            value: result[fieldId] || "",
          };
        }

        // Add the admin_area_level_2 component to the parent address
        // Convert to string to ensure proper format for Pipedrive API
        result.custom_fields[fieldId].admin_area_level_2 = String(result[key]);

        // Remove it from the root level
        delete result[key];
      }
    }
  }

  // Find any fields that follow the pattern fieldId_component
  const addressComponentKeys = Object.keys(result).filter((key) =>
    /^[a-f0-9]{20,}_[a-z_]+$/i.test(key)
  );

  // If no address components are found, return early
  if (addressComponentKeys.length === 0) {
    return result;
  }

  // Initialize custom_fields if needed
  if (!result.custom_fields) {
    result.custom_fields = {};
  }

  // First, collect all address components and organize them by parent field ID
  const addressComponents = {};

  for (const key of addressComponentKeys) {
    const parts = key.match(/^([a-f0-9]{20,})_(.+)$/i);
    if (parts && parts.length === 3) {
      const fieldId = parts[1];
      const component = parts[2];
      const value = result[key];


      // Initialize the address components for this field if needed
      if (!addressComponents[fieldId]) {
        addressComponents[fieldId] = {
          components: {},
          mainValue: result[fieldId] || "",
        };
      }

      // Store the component - convert to string to ensure proper format for Pipedrive API
      addressComponents[fieldId].components[component] = String(value);

      // Remove from root level immediately
      delete result[key];
    }
  }

  // Now create structured address objects
  for (const fieldId in addressComponents) {
    const addressData = addressComponents[fieldId];

    // Attempt to get the current address data from Pipedrive if we have an entity ID
    let currentAddressData = {};
    if (addressData.entityId) {
      currentAddressData = getCurrentAddressData(
        addressData.entityType,
        addressData.entityId,
        fieldId
      );
    }

    // Create a new address object
    let addressObj = {};

    // Priority handling: Check if we have a full address field that's being updated
    // If the main field ID exists in the data and isn't empty, prioritize it for the value property
    const hasFullAddressUpdate =
      result[fieldId] !== undefined && result[fieldId] !== "";

    if (hasFullAddressUpdate) {
      // Full address field is present and updated - prioritize it
      addressObj.value = String(result[fieldId]);

      // Include both current components and updated components
      // First add current components that aren't being updated
      if (currentAddressData && typeof currentAddressData === "object") {
        for (const component in currentAddressData) {
          if (component !== "value") {
            // Only add if not already being updated
            if (!addressData.components[component]) {
              addressObj[component] = String(currentAddressData[component]);
            }
          }
        }
      }

      // Then add the components being updated
      for (const component in addressData.components) {
        // Ensure all components are strings for Pipedrive API
        addressObj[component] = String(addressData.components[component]);
      }
    } else {
      // No full address update - construct address from components

      // Start with existing components from Pipedrive
      if (currentAddressData && typeof currentAddressData === "object") {
        // Use the existing value if available
        if (currentAddressData.value) {
          addressObj.value = String(currentAddressData.value);
        }

        // Add all existing components for preservation
        for (const component in currentAddressData) {
          if (component !== "value") {
            addressObj[component] = String(currentAddressData[component]);
          }
        }
      } else {
        // No current data - use main value if available
        addressObj.value = String(addressData.mainValue || "");
      }

      // Overwrite with modified components
      for (const component in addressData.components) {
        addressObj[component] = String(addressData.components[component]);
      }

      // If we now have components but no value, construct a new address string
      if (!addressObj.value && Object.keys(addressObj).length > 1) {
        // Construct a new address from all available components (both preserved and updated)
        let newAddress = "";

        // Start with street number and route (street name)
        if (addressObj.street_number && addressObj.route) {
          newAddress = `${addressObj.street_number} ${addressObj.route}`;
        } else if (addressObj.route) {
          newAddress = addressObj.route;
        }

        // Add city (locality)
        if (addressObj.locality) {
          if (newAddress) newAddress += `, ${addressObj.locality}`;
          else newAddress = addressObj.locality;
        }

        // Add state/province with comma
        if (addressObj.admin_area_level_1) {
          if (newAddress) newAddress += `, ${addressObj.admin_area_level_1}`;
          else newAddress = addressObj.admin_area_level_1;
        }

        // Add postal code (no comma before postal code)
        if (addressObj.postal_code) {
          if (newAddress) newAddress += ` ${addressObj.postal_code}`;
          else newAddress = addressObj.postal_code;
        }

        // Add country with comma
        if (addressObj.country) {
          if (newAddress) newAddress += `, ${addressObj.country}`;
          else newAddress = addressObj.country;
        }

        // Set the value property to the constructed address
        if (newAddress) {
          addressObj.value = newAddress;
        }
      }
    }

    // Add this address object to custom_fields
    result.custom_fields[fieldId] = addressObj;

    // For extra safety, ensure this is passed as an object, not a string
    if (
      !result.custom_fields[fieldId].value &&
      typeof result.custom_fields[fieldId] === "string"
    ) {
      result.custom_fields[fieldId] = {
        value: result.custom_fields[fieldId],
      };

      // Re-add components - convert to string to ensure proper format for Pipedrive API
      for (const component in addressData.components) {
        result.custom_fields[fieldId][component] = String(
          addressData.components[component]
        );
      }
    }

    // If we're using the full address, also remove it from the root level to avoid duplication
    if (hasFullAddressUpdate) {
      delete result[fieldId];
    }

  }

  // Final check for any address components still at root level - this is a safety check
  for (const key in result) {
    if (/^[a-f0-9]{20,}_[a-z_]+$/i.test(key)) {

      // Extract field ID and component
      const parts = key.match(/^([a-f0-9]{20,})_(.+)$/i);
      if (parts && parts.length === 3) {
        const fieldId = parts[1];
        const component = parts[2];

        // If parent exists in custom_fields, move component there
        if (result.custom_fields && result.custom_fields[fieldId]) {
          // Convert to string to ensure proper format for Pipedrive API
          result.custom_fields[fieldId][component] = String(result[key]);
          delete result[key];
        }
      }
    }
  }

  return result;
}

/**
 * Fetches the current address data for a specific entity
 * @param {string} entityType - The entity type (deals, persons, organizations)
 * @param {number|string} entityId - The ID of the entity
 * @param {string} addressFieldId - The custom field ID for the address
 * @return {Object} Object containing the current address components or empty object if not found
 */
function getCurrentAddressData(entityType, entityId, addressFieldId) {
  try {
    // Check access token
    const scriptProperties = PropertiesService.getScriptProperties();
    const accessToken = scriptProperties.getProperty("PIPEDRIVE_ACCESS_TOKEN");

    if (!accessToken) {
      return {};
    }

    // Get baseUrl based on subdomain
    const subdomain =
      scriptProperties.getProperty("PIPEDRIVE_SUBDOMAIN") ||
      DEFAULT_PIPEDRIVE_SUBDOMAIN;
    const baseUrl = `https://${subdomain}.pipedrive.com`;

    // Construct API URL based on entity type
    let apiUrl;
    switch (entityType.toLowerCase()) {
      case "deals":
        apiUrl = `${baseUrl}/api/v2/deals/${entityId}`;
        break;
      case "persons":
        apiUrl = `${baseUrl}/api/v2/persons/${entityId}`;
        break;
      case "organizations":
        apiUrl = `${baseUrl}/api/v2/organizations/${entityId}`;
        break;
      default:
    }

    // Make the request
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      muteHttpExceptions: true,
    });

    // Check response
    if (response.getResponseCode() !== 200) {
    }

    // Parse response
    const responseData = JSON.parse(response.getContentText());

    if (!responseData.success || !responseData.data) {
    }

    // Log the entire response structure to help debug this issue

    // Try different paths where the custom field might be located
    let addressField = null;

    // Check in custom_fields (API v2 structure)
    if (
      responseData.data.custom_fields &&
      responseData.data.custom_fields[addressFieldId]
    ) {
      addressField = responseData.data.custom_fields[addressFieldId];
    }
    // Check in data directly (some API endpoints might put it here)
    else if (responseData.data[addressFieldId]) {
      addressField = responseData.data[addressFieldId];
    }

    // If address field is just a string (not an object), create an object
    if (typeof addressField !== "object" || addressField === null) {
      addressField = {
        value: addressField,
      };
    }

    // Return the complete address object with all components
    return addressField;
  } catch (error) {
  }
}

/**
 * Sanitizes the payload before sending to Pipedrive API
 * Specifically handles address components that might be at the top level
 * @param {Object} payload - The payload to sanitize
 * @return {Object} The sanitized payload
 */
function sanitizePayloadForPipedrive(payload) {
  if (!payload) return payload;

  // Create a copy to avoid modifying the original
  const result = JSON.parse(JSON.stringify(payload));

  // Find any top-level fields that match the pattern of address components (fieldId_component)
  const addressComponentKeys = [];

  for (const key in result) {
    // Skip normal fields
    if (
      key === "custom_fields" ||
      key === "id" ||
      key === "title" ||
      key === "value" ||
      key === "phone" ||
      key === "email"
    ) {
      continue;
    }

    // Specifically check for admin_area_level_2 field that's causing problems
    if (key.includes("_admin_area_level_2")) {

      // Extract the field ID from the key
      const fieldId = key.split("_admin_area_level_2")[0];

      // Ensure custom_fields exists
      if (!result.custom_fields) {
        result.custom_fields = {};
      }

      // Ensure the parent address field exists
      if (!result.custom_fields[fieldId]) {
        result.custom_fields[fieldId] = {
          value: "Address",
        };
      }

      // Add the admin_area_level_2 component directly to the address object
      result.custom_fields[fieldId].admin_area_level_2 = String(result[key]);

      // Mark for removal
      addressComponentKeys.push(key);
      continue;
    }

    // Check for other address component fields (fieldId_component)
    const match = key.match(/^([a-f0-9]{20,})_([a-z_]+)$/i);
    if (match) {
      const fieldId = match[1];
      const component = match[2];

      // Ensure the custom_fields object exists
      if (!result.custom_fields) {
        result.custom_fields = {};
      }

      // Ensure the parent address field exists in custom_fields
      if (!result.custom_fields[fieldId]) {
        result.custom_fields[fieldId] = {
          value: "Address",
        };
      }

      // Add the component to the address object
      result.custom_fields[fieldId][component] = String(result[key]);

      // Mark for removal
      addressComponentKeys.push(key);
    }
  }

  // Remove the top-level address components
  addressComponentKeys.forEach((key) => {
    delete result[key];
  });

  // If we made changes, log the updated payload
  if (addressComponentKeys.length > 0) {
  }

  return result;
}

/**
 * Helper function to log object structure and methods
 * This helps debug the Pipedrive npm package structure
 * @param {Object} obj - Object to inspect
 * @param {string} label - Label for logging
 */
function logObjectStructure(obj, label) {

  if (obj === null || obj === undefined) {
    return;
  }

  try {
    // Log direct properties

    // Log methods if any
    const methods = Object.getOwnPropertyNames(
      Object.getPrototypeOf(obj)
    ).filter(
      (name) => typeof obj[name] === "function" && name !== "constructor"
    );

    if (methods.length) {
    }
  } catch (e) {
  }

}

/**
 * Saves Pipedrive settings for the active sheet
 * @param {string} apiKey - The Pipedrive API key (ignored in OAuth flow)
 * @param {string} entityType - The entity type (deals, persons, etc.)
 * @param {string} filterId - The filter ID
 * @param {string} subdomain - The Pipedrive subdomain (ignored in OAuth flow)
 * @param {string} sheetName - The name of the sheet
 * @param {boolean} enableTimestamp - Whether to enable timestamp after sync
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp = false) {
  // Check if user has permission to modify settings
  if (!PaymentService.canModifySettings()) {
    throw new Error('Only team admins can modify settings');
  }

  const scriptProperties = PropertiesService.getScriptProperties();

  // Note: In OAuth flow, we don't save API key or subdomain as they're handled by OAuth
  // These parameters are kept for backward compatibility with the UI
  
  // Save sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  const timestampEnabledKey = `TIMESTAMP_ENABLED_${sheetName}`;

  scriptProperties.setProperty(sheetFilterIdKey, filterId);
  scriptProperties.setProperty(sheetEntityTypeKey, entityType);
  scriptProperties.setProperty(timestampEnabledKey, enableTimestamp.toString());
  scriptProperties.setProperty('SHEET_NAME', sheetName);
  
}
