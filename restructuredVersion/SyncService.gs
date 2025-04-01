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
 * Checks if a sync operation is currently running
 * @return {boolean} True if a sync is running, false otherwise
 */
function isSyncRunning() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SYNC_RUNNING') === 'true';
}

/**
 * Sets the sync running status
 * @param {boolean} isRunning - Whether the sync is running
 */
function setSyncRunning(isRunning) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SYNC_RUNNING', isRunning ? 'true' : 'false');
}

/**
 * Main synchronization function that syncs data from Pipedrive to the sheet
 */
function syncFromPipedrive() {
  try {
    Logger.log("Starting syncFromPipedrive function");
    // Get active sheet info
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    Logger.log(`Active sheet: ${sheetName}`);
    
    // IMPORTANT: Detect any column shifts that may have occurred since last sync
    detectColumnShifts();
    
    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Check for two-way sync
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // Show confirmation dialog
    const ui = SpreadsheetApp.getUi();
    let confirmMessage = `This will sync data from Pipedrive to the current sheet "${sheetName}". Any existing data in this sheet will be replaced.`;
    
    if (twoWaySyncEnabled) {
      confirmMessage += `\n\nTwo-way sync is enabled for this sheet. Modified rows will be pushed to Pipedrive before pulling new data.`;
    }
    
    const response = ui.alert(
      'Confirm Sync',
      confirmMessage,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response !== ui.Button.OK) {
      Logger.log("User cancelled sync operation");
      return;
    }
    
    // Prevent multiple syncs running at once
    if (isSyncRunning()) {
      Logger.log("Sync already running, showing alert");
      ui.alert('A sync operation is already running. Please wait for it to complete.');
      return;
    }

    // Get configuration
    const entityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const filterIdKey = `FILTER_ID_${sheetName}`;
    
    const entityType = scriptProperties.getProperty(entityTypeKey);
    const filterId = scriptProperties.getProperty(filterIdKey);
    
    Logger.log(`Syncing sheet ${sheetName}, entity type: ${entityType}, filter ID: ${filterId}`);

    // Check for required settings
    if (!entityType) {
      Logger.log("No entity type set for this sheet");
      SpreadsheetApp.getUi().alert(
        'No Pipedrive entity type set for this sheet. Please configure your filter settings first.'
      );
      return;
    }

    // Show sync status dialog
    showSyncStatus(sheetName);
    
    // Mark sync as running
    setSyncRunning(true);
    
    // Start the sync process
    updateSyncStatus('1', 'active', 'Connecting to Pipedrive...', 50);
    
    // Perform sync with skip push parameter as false
    syncPipedriveDataToSheet(entityType, false, sheetName, filterId);
    
    // Show completion message
    Logger.log("Sync completed successfully");
    SpreadsheetApp.getUi().alert('Sync completed successfully!');
  } catch (error) {
    Logger.log(`Error in syncFromPipedrive: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Update sync status
    updateSyncStatus('3', 'error', `Error: ${error.message}`, 0);
    
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
function syncPipedriveDataToSheet(entityType, skipPush = false, sheetName = null, filterId = null) {
  try {
    Logger.log(`Starting syncPipedriveDataToSheet - Entity Type: ${entityType}, Skip Push: ${skipPush}, Sheet Name: ${sheetName}, Filter ID: ${filterId}`);
    
    // Get sheet name if not provided
    sheetName = sheetName || SpreadsheetApp.getActiveSheet().getName();
    
    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // If no filter ID provided, try to get from script properties
    if (!filterId) {
      const filterIdKey = `FILTER_ID_${sheetName}`;
      filterId = scriptProperties.getProperty(filterIdKey);
      Logger.log(`Using stored filter ID: ${filterId} from key ${filterIdKey}`);
    }
    
    // Show UI that we are retrieving data
    updateSyncStatus('2', 'active', 'Retrieving data from Pipedrive...', 10);
    
    // Check for two-way sync settings
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    Logger.log(`Two-way sync enabled: ${twoWaySyncEnabled}`);
    
    // Key for tracking column
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    
    // If user wants to skip push changes to Pipedrive, just sync
    if (!skipPush && twoWaySyncEnabled) {
      // Ask user if they want to push first
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Push Changes First?',
        'Do you want to push any changes to Pipedrive before syncing?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        Logger.log('User chose to push changes before syncing');
        // Push changes to Pipedrive first
        pushChangesToPipedrive(true, true); // true for scheduled sync, true for suppress warning
        
        // After pushing changes, continue with the sync
        Logger.log('Changes pushed, continuing with sync');
      }
    }
    
    // Get data from Pipedrive based on entity type
    let items = [];
    
    // Update status to show we're connecting to API
    updateSyncStatus('2', 'active', `Retrieving ${entityType} from Pipedrive...`, 20);
    
    Logger.log(`Retrieving data for entity type: ${entityType}`);
    
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
    
    Logger.log(`Retrieved ${items.length} items from Pipedrive`);
    
    // Log first item structure for debugging
    if (items.length > 0) {
      Logger.log("Sample item (first item) from retrieved data:");
      Logger.log(JSON.stringify(items[0], null, 2));
      
      // Specifically log email and phone fields if they exist
      if (items[0].email) {
        Logger.log("Email field structure:");
        Logger.log(JSON.stringify(items[0].email, null, 2));
      }
      
      if (items[0].phone) {
        Logger.log("Phone field structure:");
        Logger.log(JSON.stringify(items[0].phone, null, 2));
      }
    }
    
    // Check if we have any data
    if (items.length === 0) {
      throw new Error(`No ${entityType} found. Please check your filter settings.`);
    }
    
    // Update status to show data retrieval is complete
    updateSyncStatus('2', 'completed', `Retrieved ${items.length} ${entityType} from Pipedrive`, 100);
    
    // Get field options for handling picklists/enums
    let optionMappings = {};
    
    try {
      Logger.log("Getting field option mappings...");
      optionMappings = getFieldOptionMappingsForEntity(entityType);
      Logger.log(`Retrieved option mappings for fields: ${Object.keys(optionMappings).join(', ')}`);
      
      // Sample logging of one option mapping if available
      const sampleField = Object.keys(optionMappings)[0];
      if (sampleField) {
        Logger.log(`Sample option mapping for field ${sampleField}:`);
        Logger.log(JSON.stringify(optionMappings[sampleField], null, 2));
      }
    } catch (e) {
      Logger.log(`Error getting field options: ${e.message}`);
    }
    
    // Start writing to sheet
    updateSyncStatus('3', 'active', 'Writing data to spreadsheet...', 10);
    
    // Get the active spreadsheet and sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Get column preferences
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    const userEmail = Session.getEffectiveUser().getEmail();
    const userColumnSettingsKey = `COLUMNS_${sheetName}_${entityType}_${userEmail}`;
    
    Logger.log(`Looking for column preferences with keys:`);
    Logger.log(`- User specific: ${userColumnSettingsKey}`);
    Logger.log(`- Column settings key: ${columnSettingsKey}`);
    
    const savedColumnsJson = scriptProperties.getProperty(userColumnSettingsKey) || 
                            scriptProperties.getProperty(columnSettingsKey);
    
    if (scriptProperties.getProperty(userColumnSettingsKey)) {
      Logger.log(`Found user-specific column preferences`);
    } else if (scriptProperties.getProperty(columnSettingsKey)) {
      Logger.log(`Found generic column preferences`);
    } else {
      Logger.log(`No saved column preferences found with any key format`);
    }
    
    let columns = [];
    
    if (savedColumnsJson) {
      try {
        columns = JSON.parse(savedColumnsJson);
        Logger.log(`Retrieved ${columns.length} column preferences`);
        
        // Log the columns for debugging
        Logger.log("Column preferences:");
        columns.forEach((col, index) => {
          Logger.log(`Column ${index + 1}: ${JSON.stringify(col)}`);
        });
      } catch (e) {
        Logger.log(`Error parsing column preferences: ${e.message}`);
        throw new Error(`Invalid column preferences: ${e.message}`);
      }
    } else {
      // If no column preferences, use default columns
      Logger.log(`No column preferences found, using defaults for ${entityType}`);
      
      if (DEFAULT_COLUMNS[entityType]) {
        DEFAULT_COLUMNS[entityType].forEach(key => {
          columns.push({ key: key, name: formatColumnName(key) });
        });
      } else {
        DEFAULT_COLUMNS.COMMON.forEach(key => {
          columns.push({ key: key, name: formatColumnName(key) });
        });
      }
      
      Logger.log(`Using ${columns.length} default columns`);
    }
    
    // Create header row from column names
    const headers = columns.map(column => {
      if (typeof column === 'object' && column.customName) {
        return column.customName;
      }
      
      if (typeof column === 'object' && column.key) {
        // Special handling for email and phone fields
        if (typeof column.key === 'string' && column.key.includes('.')) {
          const parts = column.key.split('.');
          
          // Format email.work as "Email Work" and phone.mobile as "Phone Mobile"
          if ((parts[0] === 'email' || parts[0] === 'phone') && parts.length > 1) {
            const fieldType = formatColumnName(parts[0]);  // "Email" or "Phone"
            
            // Handle array index notation (e.g., email.0.value)
            if (parts[1] === '0' && parts.length > 2 && parts[2] === 'value') {
              return `${fieldType} Other`;
            }
            
            // Handle specific types (work, home, etc.)
            if (parts[1].toLowerCase() !== 'primary') {
              const labelType = formatColumnName(parts[1]);   // "Work", "Home", etc.
              return `${fieldType} ${labelType}`;
            }
          }
        }
        
        // For base email/phone fields, use capitalized name with "(Primary)"
        if (column.key === 'email' || column.key === 'phone') {
          const fieldType = formatColumnName(column.key);
          return `${fieldType} (Primary)`;
        }
        
        // Use the name if provided for other fields
        if (column.name) {
          return column.name;
        }
        
        // Default to formatted key
        return formatColumnName(column.key);
      }
      
      // Fall back to string value
      return typeof column === 'string' ? formatColumnName(column) : String(column);
    });
    
    // Ensure header names are unique
    const uniqueHeaders = [];
    const seenHeaders = new Map(); // Use Map to track count of each header
    
    headers.forEach(header => {
      if (seenHeaders.has(header)) {
        // Skip duplicate primary email/phone headers
        if (header.includes('(Primary)')) {
          return;
        }
        
        const count = seenHeaders.get(header);
        seenHeaders.set(header, count + 1);
        uniqueHeaders.push(`${header} (${count + 1})`);
      } else {
        seenHeaders.set(header, 1);
        uniqueHeaders.push(header);
      }
    });
    
    // Filter out any empty or undefined headers
    const finalHeaders = uniqueHeaders.filter(header => header && header.trim());
    
    Logger.log(`Created ${finalHeaders.length} unique headers: ${finalHeaders.join(', ')}`);
    
    // Options for writing data
    const options = {
      sheetName: sheetName,
      columns: columns,
      headerRow: finalHeaders,
      entityType: entityType,
      optionMappings: optionMappings,
      twoWaySyncEnabled: twoWaySyncEnabled
    };
    
    // Store original data for undo detection when two-way sync is enabled
    if (twoWaySyncEnabled) {
      try {
        Logger.log('Storing original Pipedrive data for undo detection');
        storeOriginalData(items, options);
      } catch (storageError) {
        Logger.log(`Error storing original data: ${storageError.message}`);
        // Continue with sync even if storage fails
      }
    }
    
    // Write data to the sheet
    writeDataToSheet(items, options);
    
    // Update sync status to completed
    updateSyncStatus('3', 'completed', 'Data successfully written to spreadsheet', 100);
    
    // Store sync timestamp
    const timestamp = new Date().toISOString();
    scriptProperties.setProperty(`LAST_SYNC_${sheetName}`, timestamp);
    
    Logger.log(`Successfully synced ${items.length} items from Pipedrive to sheet "${sheetName}"`);
    
    return true;
  } catch (error) {
    Logger.log(`Error in syncPipedriveDataToSheet: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Update sync status
    updateSyncStatus('3', 'error', `Error: ${error.message}`, 0);
    
    // Show error in UI
    SpreadsheetApp.getUi().alert(`Error syncing data: ${error.message}`);
    
    throw error;
  }
}

/**
 * Writes data to the sheet with the specified entities and options
 * @param {Array} items - Array of Pipedrive entities to write
 * @param {Object} options - Options for writing data
 */
function writeDataToSheet(items, options) {
  try {
    Logger.log(`Starting writeDataToSheet with ${items.length} items`);
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(options.sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${options.sheetName}" not found`);
    }
    
    // IMPORTANT: Clean up previous Sync Status column before writing new data
    cleanupPreviousSyncStatusColumn(sheet, options.sheetName);
    
    // Check for two-way sync
    const twoWaySyncEnabled = options.twoWaySyncEnabled || false;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${options.sheetName}`;
    let statusColumnIndex = -1;
    let configuredTrackingColumn = '';
    
    // Create a map to preserve existing status values
    const statusByIdMap = new Map();
    
    // If two-way sync is enabled, preserve existing status values
    if (twoWaySyncEnabled) {
      try {
        Logger.log('Two-way sync is enabled, preserving status column data');
        
        // Get the configured tracking column letter from properties
        configuredTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
        Logger.log(`Configured tracking column from properties: "${configuredTrackingColumn}"`);
        
        // If the sheet has data, extract status values
        if (sheet.getLastRow() > 0) {
          // Get all existing data
          const existingData = sheet.getDataRange().getValues();
          Logger.log(`Retrieved ${existingData.length} rows of existing data`);
          
          // Find headers row
          const headers = existingData[0];
          let idColumnIndex = 0; // Assuming ID is in first column
          let statusColumnIndex = -1;
          
          // Find the status column by looking for header that contains "Sync Status"
          for (let i = 0; i < headers.length; i++) {
            if (headers[i] && 
                headers[i].toString().toLowerCase().includes('sync') && 
                headers[i].toString().toLowerCase().includes('status')) {
              statusColumnIndex = i;
              Logger.log(`Found status column at index ${statusColumnIndex} with header "${headers[i]}"`);
              break;
            }
          }
          
          // If status column found, extract status values by ID
          if (statusColumnIndex !== -1) {
            // Process all data rows (skip header)
            for (let i = 1; i < existingData.length; i++) {
              const row = existingData[i];
              const id = row[idColumnIndex];
              const status = row[statusColumnIndex];
              
              // Skip rows without ID or status or timestamp rows
              if (!id || !status ||
                  ((typeof id === 'string') &&
                   (id.toLowerCase().includes('last') ||
                    id.toLowerCase().includes('synced') ||
                    id.toLowerCase().includes('updated')))) {
                continue;
              }
              
              // Only preserve meaningful status values
              if (status === 'Modified' || status === 'Synced' || status === 'Error') {
                statusByIdMap.set(id.toString(), status);
                Logger.log(`Preserved status "${status}" for ID ${id}`);
              }
            }
            
            Logger.log(`Preserved ${statusByIdMap.size} status values by ID`);
          } else {
            Logger.log('Could not find existing status column');
          }
        }
      } catch (e) {
        Logger.log(`Error preserving status values: ${e.message}`);
      }
    }
    
    // Clear the sheet but keep formatting
    sheet.clear();
    
    // Get headers from options
    const headers = options.headerRow;
    
    // If two-way sync is enabled, add Sync Status column at the correct position
    if (twoWaySyncEnabled) {
      // Check if we need to force the Sync Status column at the end
      const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${options.sheetName}`;
      const forceColumnAtEnd = scriptProperties.getProperty(twoWaySyncColumnAtEndKey) === 'true';

      // If we need to force the column at the end, or no configured column exists,
      // add the Sync Status column at the end
      if (forceColumnAtEnd || !configuredTrackingColumn) {
        Logger.log('Adding Sync Status column at the end');
        statusColumnIndex = headers.length;
        headers.push('Sync Status');

        // Clear the "force at end" flag after we've processed it
        if (forceColumnAtEnd) {
          scriptProperties.deleteProperty(twoWaySyncColumnAtEndKey);
          Logger.log('Cleared the force-at-end flag after repositioning the Sync Status column');
        }
      } else {
        // If not forcing at end, try to find existing column position
        if (configuredTrackingColumn) {
          // Find if there's already a "Sync Status" column in the existing sheet
          const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(options.sheetName);
          if (activeSheet && activeSheet.getLastRow() > 0) {
            const existingHeaders = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];

            // Search for "Sync Status" in the existing headers
            for (let i = 0; i < existingHeaders.length; i++) {
              if (existingHeaders[i] === "Sync Status") {
                // Use this exact position for our new status column
                const exactColumnLetter = columnToLetter(i + 1);
                Logger.log(`Found existing Sync Status column at position ${i + 1} (${exactColumnLetter})`);

                // Update the tracking column in script properties
                scriptProperties.setProperty(twoWaySyncTrackingColumnKey, exactColumnLetter);

                // Ensure our headers array has enough elements
                while (headers.length <= i) {
                  headers.push('');
                }

                // Place "Sync Status" at the exact same position
                statusColumnIndex = i;
                headers[statusColumnIndex] = 'Sync Status';

                break;
              }
            }
          }

          // If we didn't find an existing column, try using the configured index
          if (statusColumnIndex === -1) {
            const existingTrackingIndex = columnLetterToIndex(configuredTrackingColumn) - 1; // Convert to 0-based

            // Always use the configured position, regardless of current headerRow length
            Logger.log(`Using configured tracking column at position ${existingTrackingIndex + 1} (${configuredTrackingColumn})`);

            // Ensure our headers array has enough elements
            while (headers.length <= existingTrackingIndex) {
              headers.push('');
            }

            // Place the Sync Status column at its original position
            statusColumnIndex = existingTrackingIndex;
            headers[statusColumnIndex] = 'Sync Status';
          }
        } else {
          // No configured tracking column, add as the last column
          Logger.log('No configured tracking column, adding as last column');
          statusColumnIndex = headers.length;
          headers.push('Sync Status');
        }
      }

      if (configuredTrackingColumn && statusColumnIndex !== -1) {
        const newTrackingColumn = columnToLetter(statusColumnIndex + 1);
        if (newTrackingColumn !== configuredTrackingColumn) {
          scriptProperties.setProperty(`PREVIOUS_TRACKING_COLUMN_${options.sheetName}`, configuredTrackingColumn);
          Logger.log(`Stored previous tracking column ${configuredTrackingColumn} before moving to ${newTrackingColumn}`);
        }
      }
      
      // Save column letter for future reference
      const statusColumn = columnToLetter(statusColumnIndex + 1);
      scriptProperties.setProperty(twoWaySyncTrackingColumnKey, statusColumn);
      Logger.log(`Added/maintained Sync Status column at position ${statusColumnIndex + 1} (${statusColumn})`);
    }
    
    // Set header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    Logger.log(`Set up ${headers.length} headers in the sheet`);
    
    // Create data rows
    Logger.log(`Processing ${items.length} items to create data rows`);
    const dataRows = items.map(item => {
      // Create a row with empty values for all columns
      const row = Array(headers.length).fill('');
      
      // For each column, extract and format the value from the Pipedrive item
      options.columns.forEach((column, index) => {
        // Skip if index is beyond our header count
        if (index >= headers.length) {
          return;
        }
        
        const columnKey = typeof column === 'object' ? column.key : column;
        const value = getValueByPath(item, columnKey);
        row[index] = formatValue(value, columnKey, options.optionMappings);
      });
      
      // If two-way sync is enabled, add the status column
      if (twoWaySyncEnabled && statusColumnIndex !== -1) {
        // Get the item ID (assuming first column is ID)
        const id = row[0] ? row[0].toString() : '';
        
        // If we have a saved status for this ID, use it, otherwise use "Not modified"
        if (id && statusByIdMap.has(id)) {
          row[statusColumnIndex] = statusByIdMap.get(id);
        } else {
          row[statusColumnIndex] = 'Not modified';
        }
      }
      
      return row;
    });
    
    // Write all data rows at once
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
    
    // If two-way sync is enabled, set up data validation and formatting for the status column
    if (twoWaySyncEnabled && statusColumnIndex !== -1 && dataRows.length > 0) {
      try {
        // Clear any existing data validation from ALL cells in the sheet first
        sheet.clearDataValidations();
        
        // Convert to 1-based column
        const statusColumnPos = statusColumnIndex + 1;
        
        // Apply validation to each data row
        for (let i = 0; i < dataRows.length; i++) {
          const row = i + 2; // Data starts at row 2 (after header)
          const statusCell = sheet.getRange(row, statusColumnPos);
          
          // Create dropdown validation
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
            .build();
          statusCell.setDataValidation(rule);
        }
        
        // Style the header
        sheet.getRange(1, statusColumnPos)
          .setBackground('#E8F0FE')
          .setFontWeight('bold')
          .setNote('This column tracks changes for two-way sync with Pipedrive');
        
        // Style the status column
        const statusRange = sheet.getRange(2, statusColumnPos, dataRows.length, 1);
        statusRange
          .setBackground('#F8F9FA')
          .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
        
        // Set up conditional formatting
        const rules = sheet.getConditionalFormatRules();
        
        // Add rule for "Modified" status - red background
        const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Modified')
          .setBackground('#FCE8E6')
          .setFontColor('#D93025')
          .setRanges([statusRange])
          .build();
        rules.push(modifiedRule);
        
        // Add rule for "Synced" status - green background
        const syncedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Synced')
          .setBackground('#E6F4EA')
          .setFontColor('#137333')
          .setRanges([statusRange])
          .build();
        rules.push(syncedRule);
        
        // Add rule for "Error" status - yellow background
        const errorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Error')
          .setBackground('#FEF7E0')
          .setFontColor('#B06000')
          .setRanges([statusRange])
          .build();
        rules.push(errorRule);
        
        sheet.setConditionalFormatRules(rules);
      } catch (e) {
        Logger.log(`Error setting up status column formatting: ${e.message}`);
      }
    }
    
    // Clean up any lingering formatting from previous Sync Status columns
    if (twoWaySyncEnabled && statusColumnIndex !== -1) {
      try {
        Logger.log(`Performing aggressive column cleanup after sheet rebuild - current status column: ${statusColumnIndex}`);
        
        // The current status column letter
        const currentStatusColLetter = columnToLetter(statusColumnIndex + 1);
        
        // Clean ALL columns in the sheet except the current Sync Status column
        const lastCol = sheet.getLastColumn() + 5; // Add buffer for hidden columns
        for (let i = 0; i < lastCol; i++) {
          if (i !== statusColumnIndex) { // Skip the current status column
            try {
              const colLetter = columnToLetter(i + 1);
              Logger.log(`Checking column ${colLetter} for cleanup`);
              
              // Check if this column has a 'Sync Status' header or sync-related note
              const headerCell = sheet.getRange(1, i + 1);
              const headerValue = headerCell.getValue();
              const note = headerCell.getNote();
              
              if (headerValue === "Sync Status" || 
                  (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive')))) {
                Logger.log(`Found Sync Status indicators in column ${colLetter}, cleaning up`);
                cleanupColumnFormatting(sheet, colLetter);
              }
            } catch (e) {
              Logger.log(`Error cleaning up column ${i}: ${e.message}`);
            }
          }
        }
      } catch (e) {
        Logger.log(`Error during column cleanup: ${e.message}`);
      }
    }
    
    // Auto-resize columns to fit content
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log('Successfully wrote data to sheet');
  } catch (error) {
    Logger.log(`Error in writeDataToSheet: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error;
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
 * Shows the sync status dialog to the user
 * @param {string} sheetName - The name of the sheet being synced
 */
function showSyncStatus(sheetName) {
  try {
    // Reset any previous sync status
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('SYNC_COMPLETED', 'false');
    scriptProperties.setProperty('SYNC_ERROR', '');
    
    // Initialize status for each phase
    for (let phase = 1; phase <= 3; phase++) {
      const status = phase === 1 ? 'active' : 'pending';
      const detail = phase === 1 ? 'Connecting to Pipedrive...' : 'Waiting to start...';
      const progress = phase === 1 ? 0 : 0;
      
      scriptProperties.setProperty(`SYNC_PHASE_${phase}_STATUS`, status);
      scriptProperties.setProperty(`SYNC_PHASE_${phase}_DETAIL`, detail);
      scriptProperties.setProperty(`SYNC_PHASE_${phase}_PROGRESS`, progress.toString());
    }
    
    // Set current phase to 1
    scriptProperties.setProperty('SYNC_CURRENT_PHASE', '1');
    
    // Get entity type for the sheet
    const entityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType = scriptProperties.getProperty(entityTypeKey) || ENTITY_TYPES.DEALS;
    const entityTypeName = formatEntityTypeName(entityType);
    
    // Create the dialog
    const htmlTemplate = HtmlService.createTemplateFromFile('SyncStatus');
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.entityType = entityType;
    htmlTemplate.entityTypeName = entityTypeName;
    
    const html = htmlTemplate.evaluate()
      .setWidth(400)
      .setHeight(400)
      .setTitle('Sync Status');
    
    // Show dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Sync Status');
    
    // Return true to indicate success
    return true;
  } catch (error) {
    Logger.log(`Error showing sync status: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Show a fallback toast message instead
    SpreadsheetApp.getActiveSpreadsheet().toast('Starting sync operation...', 'Pipedrive Sync', 3);
    return false;
  }
}

/**
 * Helper function to format entity type name for display
 * @param {string} entityType - The entity type to format
 * @return {string} Formatted entity type name
 */
function formatEntityTypeName(entityType) {
  if (!entityType) return '';
  
  const entityMap = {
    'deals': 'Deals',
    'persons': 'Persons',
    'organizations': 'Organizations',
    'activities': 'Activities',
    'leads': 'Leads',
    'products': 'Products'
  };
  
  return entityMap[entityType] || entityType.charAt(0).toUpperCase() + entityType.slice(1);
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
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      if (trigger.getHandlerFunction() === 'onEdit') {
        Logger.log('onEdit trigger already exists');
        return; // Exit if trigger already exists
      }
    }
    
    // Create the trigger if it doesn't exist
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    Logger.log('onEdit trigger created');
  } catch (e) {
    Logger.log(`Error setting up onEdit trigger: ${e.message}`);
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
      if (trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('onEdit trigger deleted');
        return;
      }
    }
    
    Logger.log('No onEdit trigger found to delete');
  } catch (e) {
    Logger.log(`Error removing onEdit trigger: ${e.message}`);
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
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is not enabled, exit
    if (!twoWaySyncEnabled) {
      return;
    }

    // Get the edited range
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();
    
    // Check if the edit is in the header row - if it is, we might need to update tracking
    const headerRow = 1;
    if (row === headerRow) {
      // If someone renamed the Sync Status column, we'd handle that here
      // For now, just exit as we don't need special handling
      return;
    }

    // Find the "Sync Status" column by header name
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColIndex = -1;
    
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        syncStatusColIndex = i;
        break;
      }
    }
    
    // Exit if no Sync Status column found
    if (syncStatusColIndex === -1) {
      return;
    }
    
    // Convert to 1-based for sheet functions
    const syncStatusColPos = syncStatusColIndex + 1;
    
    // Check if the edit is in the Sync Status column itself (to avoid loops)
    if (column === syncStatusColPos) {
      return;
    }

    // Get the row content to check if it's a real data row or a timestamp/blank row
    const rowContent = sheet.getRange(row, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];

    // Check if this is a timestamp row
    const firstCell = String(rowContent[0] || "").toLowerCase();
    const isTimestampRow = firstCell.includes("last") ||
      firstCell.includes("updated") ||
      firstCell.includes("synced") ||
      firstCell.includes("date");

    // Count non-empty cells to determine if this is a data row
    const nonEmptyCells = rowContent.filter(cell => cell !== "" && cell !== null && cell !== undefined).length;

    // Skip if this is a timestamp row or has too few cells with data
    if (isTimestampRow || nonEmptyCells < 3) {
      return;
    }

    // Get the row ID from the first column
    const idColumnIndex = 0;
    const id = rowContent[idColumnIndex];

    // Skip rows without an ID (likely empty rows)
    if (!id) {
      return;
    }

    // Get the sync status cell
    const syncStatusCell = sheet.getRange(row, syncStatusColPos);
    const currentStatus = syncStatusCell.getValue();
    
    // Get the original data from Pipedrive
    const originalDataKey = `ORIGINAL_DATA_${sheetName}`;
    let originalData;
    
    try {
      // Try to get original data from document properties
      const originalDataJson = scriptProperties.getProperty(originalDataKey);
      originalData = originalDataJson ? JSON.parse(originalDataJson) : {};
    } catch (parseError) {
      Logger.log(`Error parsing original data: ${parseError.message}`);
      originalData = {};
    }
    
    // Check if we have original data for this row
    const rowKey = id.toString();
    
    // Handle first-time edit case
    if (currentStatus !== "Modified") {
      // New edit to unmodified row - store the current value before updating status
      if (!originalData[rowKey]) {
        originalData[rowKey] = {};
      }
      
      // Get the column header name for the edited column
      const headerName = headers[column - 1]; // Adjust for 0-based array
      
      if (headerName) {
        // Store the original value before it was changed
        originalData[rowKey][headerName] = e.oldValue !== undefined ? e.oldValue : null;
        
        // Save updated original data
        try {
          scriptProperties.setProperty(originalDataKey, JSON.stringify(originalData));
        } catch (saveError) {
          Logger.log(`Error saving original data: ${saveError.message}`);
        }
        
        // Mark as modified
        syncStatusCell.setValue("Modified");

        // Re-apply data validation to ensure consistent dropdown options
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
          .build();
        syncStatusCell.setDataValidation(rule);

        // Make sure the styling is consistent
        syncStatusCell.setBackground('#FCE8E6').setFontColor('#D93025');
      }
    } else {
      // Row is already modified - check if this edit reverts to original value
      
      // Get the column header name for the edited column
      const headerName = headers[column - 1]; // Adjust for 0-based array
      
      if (headerName && originalData[rowKey] && originalData[rowKey][headerName] !== undefined) {
        const originalValue = originalData[rowKey][headerName];
        const currentValue = e.value;
        
        // If new value matches original value
        if (originalValue == currentValue) { // Using non-strict comparison for different types
          // Check if all edited values in the row now match original values
          const allMatch = checkAllValuesMatchOriginal(sheet, row, headers, originalData[rowKey]);
          
          if (allMatch) {
            // All values in row match original - reset to Not modified
            syncStatusCell.setValue("Not modified");
            
            // Re-apply data validation
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
              .build();
            syncStatusCell.setDataValidation(rule);
            
            // Reset formatting
            syncStatusCell.setBackground('#F8F9FA').setFontColor('#000000');
          }
        }
      } else if (e.oldValue !== undefined && e.value !== undefined) {
        // Store the first known value as original if we don't have it yet
        if (!originalData[rowKey]) {
          originalData[rowKey] = {};
        }
        
        if (!originalData[rowKey][headerName]) {
          originalData[rowKey][headerName] = e.oldValue;
          
          // Save updated original data
          try {
            scriptProperties.setProperty(originalDataKey, JSON.stringify(originalData));
          } catch (saveError) {
            Logger.log(`Error saving original data: ${saveError.message}`);
          }
        }
      }
    }
  } catch (error) {
    // Silent fail for onEdit triggers
    Logger.log(`Error in onEdit trigger: ${error.message}`);
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
    
    // Check each column that has a stored original value
    for (const headerName in originalValues) {
      // Find the column index for this header
      const colIndex = headers.indexOf(headerName);
      if (colIndex === -1) continue; // Header not found
      
      const originalValue = originalValues[headerName];
      const currentValue = rowValues[colIndex];
      
      // If the current value doesn't match the original, return false
      if (originalValue != currentValue) { // Using non-strict comparison
        return false;
      }
    }
    
    // If we reach here, all values with stored originals match
    return true;
  } catch (error) {
    Logger.log(`Error in checkAllValuesMatchOriginal: ${error.message}`);
    return false;
  }
}

/**
 * Main function to sync deals from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncDealsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.DEALS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync persons from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncPersonsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PERSONS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync organizations from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncOrganizationsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ORGANIZATIONS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync activities from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncActivitiesFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ACTIVITIES, skipPush, sheetName, filterId);
}

/**
 * Main function to sync leads from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncLeadsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.LEADS, skipPush, sheetName, filterId);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 * @param {string} filterId - The filter ID to use
 * @param {boolean} skipPush - Whether to skip pushing changes back to Pipedrive
 * @param {string} sheetName - The name of the sheet to sync to
 */
function syncProductsFromFilter(filterId, skipPush = false, sheetName = null) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush, sheetName, filterId);
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

    // Find the "Sync Status" column by header name
    const headerRow = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
    let syncStatusColumnIndex = -1;

    // Search for "Sync Status" header
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === "Sync Status") {
        syncStatusColumnIndex = i;
        Logger.log(`Found Sync Status column at index ${syncStatusColumnIndex}`);
        
        // Update the stored tracking column letter
        const trackingColumn = columnToLetter(syncStatusColumnIndex + 1);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
        break;
      }
    }

    // Validate tracking column index
    if (syncStatusColumnIndex === -1) {
      Logger.log(`Sync Status column not found, cannot proceed with sync`);
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Sync Status Column Not Found',
          'The Sync Status column could not be found. Please enable two-way sync in the settings first.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

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
      const syncStatus = row[syncStatusColumnIndex];

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
          if (j === syncStatusColumnIndex) {
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
          activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setValue('Synced');
          
          // Add a timestamp if desired
          if (scriptProperties.getProperty('ENABLE_TIMESTAMP') === 'true') {
            const timestamp = new Date().toLocaleString();
            activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setNote(`Last sync: ${timestamp}`);
          }
        } else {
          // Update failed
          failureCount++;
          
          // Set the sync status to "Error" with a note about the error
          activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setValue('Error');
          activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setNote(
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
        activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setValue('Error');
        activeSheet.getRange(rowData.rowIndex + 1, syncStatusColumnIndex + 1).setNote(`Error: ${error.message}`);
        
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
    scriptProperties.setProperty(`TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`, now);

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

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is not enabled, just return silently
    if (!twoWaySyncEnabled) {
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
 * Cleans up formatting from previous sync status columns
 * @param {Sheet} sheet - The sheet to clean up
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

    // Store current column positions for future comparison
    // This helps when columns are deleted and the position shifts
    const currentColumnIndex = currentColumn ? columnLetterToIndex(currentColumn) : -1;
    if (currentColumnIndex > 0) {
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, (currentColumnIndex - 1).toString());
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
 * Gets team-aware column preferences
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of column keys
 */
SyncService.getTeamAwareColumnPreferences = function(entityType, sheetName) {
  try {
    // Get from script properties directly since UI module might not be available
    const scriptProperties = PropertiesService.getScriptProperties();
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    
    if (savedColumnsJson) {
      try {
        return JSON.parse(savedColumnsJson);
      } catch (parseError) {
        Logger.log(`Error parsing column preferences: ${parseError.message}`);
      }
    }
    
    return [];
  } catch (e) {
    Logger.log(`Error in SyncService.getTeamAwareColumnPreferences: ${e.message}`);
    return [];
  }
};

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
    const currentColLetter = scriptProperties.getProperty(trackingColumnKey) || '';
    const previousPosStr = scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`) || '-1';
    const previousPos = parseInt(previousPosStr, 10);

    // Find all "Sync Status" headers in the sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColumns = [];

    // Find ALL instances of "Sync Status" headers
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        syncStatusColumns.push(i);
      }
    }

    // If we have multiple "Sync Status" columns, clean up all but the rightmost one
    if (syncStatusColumns.length > 1) {
      Logger.log(`Found ${syncStatusColumns.length} Sync Status columns`);
      // Keep only the rightmost column
      const rightmostIndex = Math.max(...syncStatusColumns);

      // Clean up all other columns
      for (const colIndex of syncStatusColumns) {
        if (colIndex !== rightmostIndex) {
          const colLetter = columnToLetter(colIndex + 1);
          Logger.log(`Cleaning up duplicate Sync Status column at ${colLetter}`);
          cleanupColumnFormatting(sheet, colLetter);
        }
      }

      // Update the tracking to the rightmost column
      const rightmostColLetter = columnToLetter(rightmostIndex + 1);
      scriptProperties.setProperty(trackingColumnKey, rightmostColLetter);
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, rightmostIndex.toString());
      return; // Exit after handling duplicates
    }

    let actualSyncStatusIndex = syncStatusColumns.length > 0 ? syncStatusColumns[0] : -1;

    if (actualSyncStatusIndex >= 0) {
      const actualColLetter = columnToLetter(actualSyncStatusIndex + 1);

      // If there's a mismatch, columns might have shifted
      if (currentColLetter && actualColLetter !== currentColLetter) {
        Logger.log(`Column shift detected: was ${currentColLetter}, now ${actualColLetter}`);

        // If the actual position is less than the recorded position, columns were removed
        if (actualSyncStatusIndex < previousPos) {
          Logger.log(`Columns were likely removed (${previousPos} → ${actualSyncStatusIndex})`);

          // Clean ALL columns to be safe
          for (let i = 0; i < sheet.getLastColumn(); i++) {
            if (i !== actualSyncStatusIndex) { // Skip current Sync Status column
              cleanupColumnFormatting(sheet, columnToLetter(i + 1));
            }
          }
        }

        // Clean up all potential previous locations
        scanAndCleanupAllSyncColumns(sheet, actualColLetter);

        // Update the tracking column property
        scriptProperties.setProperty(trackingColumnKey, actualColLetter);
        scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, actualSyncStatusIndex.toString());
      }
    }
  } catch (error) {
    Logger.log(`Error in detectColumnShifts: ${error.message}`);
  }
}

/**
 * Cleans up formatting in a column
 * @param {Sheet} sheet - The sheet containing the column
 * @param {string} columnLetter - The letter of the column to clean up
 */
function cleanupColumnFormatting(sheet, columnLetter) {
  try {
    // Convert column letter to index (1-based for getRange)
    const columnIndex = columnLetterToIndex(columnLetter);
    
    // ALWAYS clean columns, even if they're beyond the current last column
    Logger.log(`Cleaning up formatting for column ${columnLetter} (position ${columnIndex})`);

    try {
      // Clean up data validations for the ENTIRE column
      const numRows = Math.max(sheet.getMaxRows(), 1000); // Use a large number to ensure all rows

      // Try to clear data validations - may fail for columns beyond the edge
      try {
        sheet.getRange(1, columnIndex, numRows, 1).clearDataValidations();
      } catch (e) {
        Logger.log(`Could not clear validations for column ${columnLetter}: ${e.message}`);
      }

      // Clear header note - this should work even for "out of bounds" columns
      try {
        sheet.getRange(1, columnIndex).clearNote();
        sheet.getRange(1, columnIndex).setNote(''); // Force clear
      } catch (e) {
        Logger.log(`Could not clear note for column ${columnLetter}: ${e.message}`);
      }

      // For columns within the sheet, we can do more thorough cleaning
      if (columnIndex <= sheet.getLastColumn()) {
        // Clear all formatting for the entire column (data rows only, not header)
        if (sheet.getLastRow() > 1) {
          try {
            // Clear formatting for data rows (row 2 and below), preserving header
            sheet.getRange(2, columnIndex, numRows - 1, 1).clear({
              formatOnly: true,
              contentsOnly: false,
              validationsOnly: true
            });
          } catch (e) {
            Logger.log(`Error clearing data rows: ${e.message}`);
          }
        }

        // Clear formatting for header separately, preserving bold
        try {
          const headerCell = sheet.getRange(1, columnIndex);
          const headerValue = headerCell.getValue();

          // Reset all formatting except bold
          headerCell.setBackground(null)
            .setBorder(null, null, null, null, null, null)
            .setFontColor(null);

          // Ensure header is bold
          headerCell.setFontWeight('bold');

        } catch (e) {
          Logger.log(`Error formatting header: ${e.message}`);
        }

        // Additionally clear specific formatting for data rows
        try {
          if (sheet.getLastRow() > 1) {
            const dataRows = sheet.getRange(2, columnIndex, Math.max(sheet.getLastRow() - 1, 1), 1);
            dataRows.setBackground(null);
            dataRows.setBorder(null, null, null, null, null, null);
            dataRows.setFontColor(null);
            dataRows.setFontWeight(null);
          }
        } catch (e) {
          Logger.log(`Error clearing specific formatting: ${e.message}`);
        }

        // Clear conditional formatting specifically for this column
        const rules = sheet.getConditionalFormatRules();
        let newRules = [];
        let removedRules = 0;

        for (const rule of rules) {
          const ranges = rule.getRanges();
          let shouldRemove = false;

          // Check if any range in this rule applies to our column
          for (const range of ranges) {
            if (range.getColumn() === columnIndex) {
              shouldRemove = true;
              break;
            }
          }

          if (!shouldRemove) {
            newRules.push(rule);
          } else {
            removedRules++;
          }
        }

        if (removedRules > 0) {
          sheet.setConditionalFormatRules(newRules);
          Logger.log(`Removed ${removedRules} conditional formatting rules from column ${columnLetter}`);
        }
      } else {
        Logger.log(`Column ${columnLetter} is beyond the sheet bounds (${sheet.getLastColumn()}), did minimal cleanup`);
      }

      Logger.log(`Completed cleanup for column ${columnLetter}`);
    } catch (innerError) {
      Logger.log(`Error during column cleanup operations: ${innerError.message}`);
    }
  } catch (error) {
    Logger.log(`Error cleaning up column ${columnLetter}: ${error.message}`);
  }
}

/**
 * Scans and cleans up all sync status columns except the current one
 * @param {Sheet} sheet - The sheet to scan
 * @param {string} currentColumnLetter - The letter of the current sync status column
 */
function scanAndCleanupAllSyncColumns(sheet, currentColumnLetter) {
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const currentColumnIndex = currentColumnLetter ? columnLetterToIndex(currentColumnLetter) - 1 : -1; // Convert to 0-based
    const columnsWithValidation = [];

    // First check for data validation rules that match our Sync Status patterns
    try {
      const lastRow = Math.max(sheet.getLastRow(), 2);
      const validations = sheet.getRange(1, 1, lastRow, headers.length).getDataValidations();

      // Check each column for validation rules
      for (let i = 0; i < headers.length; i++) {
        // Skip the current tracking column
        if (i === currentColumnIndex) {
          continue;
        }

        try {
          // Check validation rules in this column
          for (let j = 0; j < validations.length; j++) {
            if (validations[j][i] !== null) {
              try {
                const criteria = validations[j][i].getCriteriaType();
                const values = validations[j][i].getCriteriaValues();

                // Look for our specific validation values (Modified, Not modified, etc.)
                if (values &&
                    (values.join(',').includes('Modified') ||
                     values.join(',').includes('Not modified') ||
                     values.join(',').includes('Synced'))) {
                  columnsWithValidation.push(i);
                  Logger.log(`Found Sync Status validation at ${columnToLetter(i + 1)}`);
                  break; // Found validation in this column, no need to check more cells
                }
              } catch (e) {
                // Some validation types may not support getCriteriaValues
                Logger.log(`Error checking validation criteria: ${e.message}`);
              }
            }
          }
        } catch (error) {
          Logger.log(`Error checking column ${i} for validation: ${error.message}`);
        }
      }
    } catch (error) {
      Logger.log(`Error checking for validations: ${error.message}`);
    }

    // Check for columns with header notes that might be Sync Status columns
    for (let i = 0; i < headers.length; i++) {
      // Skip the current tracking column
      if (i === currentColumnIndex) {
        continue;
      }

      try {
        const columnPos = i + 1; // 1-based for getRange
        const headerCell = sheet.getRange(1, columnPos);
        const note = headerCell.getNote();

        if (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive'))) {
          Logger.log(`Found column with sync-related note at ${columnToLetter(columnPos)}: "${note}"`);
          const columnToClean = columnToLetter(columnPos);
          cleanupColumnFormatting(sheet, columnToClean);
        }
      } catch (error) {
        Logger.log(`Error checking column ${i} header note: ${error.message}`);
      }
    }

    // Look for columns with "Sync Status" header
    for (let i = 0; i < headers.length; i++) {
      const headerValue = headers[i];

      // Skip the current tracking column
      if (i === currentColumnIndex) {
        continue;
      }

      // If we find a "Sync Status" header in a different column, clean it up
      if (headerValue === "Sync Status") {
        const columnToClean = columnToLetter(i + 1);
        Logger.log(`Found orphaned Sync Status column at ${columnToClean}, cleaning up`);
        cleanupColumnFormatting(sheet, columnToClean);
      }
    }

    // Clean up any columns with validation that aren't the current column
    for (const columnIndex of columnsWithValidation) {
      const columnToClean = columnToLetter(columnIndex + 1);
      Logger.log(`Found column with validation at ${columnToClean}, cleaning up`);
      cleanupColumnFormatting(sheet, columnToClean);
    }

    // Scan for orphaned formatting that matches Sync Status patterns
    const dataRange = sheet.getDataRange();
    const numRows = Math.min(dataRange.getNumRows(), 10); // Check first 10 rows for formatting
    const numCols = dataRange.getNumColumns();

    // Look at ALL columns for specific background colors
    for (let col = 1; col <= numCols; col++) {
      // Skip current column (1-based)
      if (col === currentColumnIndex + 1) continue;

      let foundSyncStatusFormatting = false;

      try {
        // Check for Sync Status cell colors in first few rows
        for (let row = 2; row <= numRows; row++) {
          const cell = sheet.getRange(row, col);
          const background = cell.getBackground();

          // Check for our specific Sync Status background colors
          if (background === '#FCE8E6' || // Error/Modified 
              background === '#E6F4EA' || // Synced
              background === '#F8F9FA') { // Default
            foundSyncStatusFormatting = true;
            break;
          }
        }

        if (foundSyncStatusFormatting) {
          const colLetter = columnToLetter(col);
          Logger.log(`Found column with Sync Status formatting at ${colLetter}, cleaning up`);
          cleanupColumnFormatting(sheet, colLetter);
        }
      } catch (error) {
        Logger.log(`Error checking column ${col} for formatting: ${error.message}`);
      }
    }

    // Additionally check for specific formatting that matches Sync Status patterns
    cleanupOrphanedConditionalFormatting(sheet, currentColumnIndex);
  } catch (error) {
    Logger.log(`Error in scanAndCleanupAllSyncColumns: ${error.message}`);
  }
}

/**
 * Cleans up orphaned conditional formatting rules
 * @param {Sheet} sheet - The sheet to clean up
 * @param {number} currentColumnIndex - The index of the current Sync Status column (0-based)
 */
function cleanupOrphanedConditionalFormatting(sheet, currentColumnIndex) {
  try {
    const rules = sheet.getConditionalFormatRules();
    const newRules = [];
    let removedRules = 0;

    for (const rule of rules) {
      const ranges = rule.getRanges();
      let keepRule = true;

      // Check if this rule applies to columns other than our current one
      // and has formatting that matches our Sync Status patterns
      for (const range of ranges) {
        const column = range.getColumn();

        // Skip our current column (currentColumnIndex is 0-based, column is 1-based)
        if (column === (currentColumnIndex + 1)) {
          continue;
        }

        // Check if this rule's formatting matches our Sync Status patterns
        const bgColor = rule.getBold() || rule.getBackground();
        if (bgColor) {
          const background = rule.getBackground();
          // If background matches our Sync Status colors, this is likely an orphaned rule
          if (background === '#FCE8E6' || background === '#E6F4EA' || background === '#F8F9FA') {
            keepRule = false;
            Logger.log(`Found orphaned conditional formatting at column ${columnToLetter(column)}`);
            break;
          }
        }
      }

      if (keepRule) {
        newRules.push(rule);
      } else {
        removedRules++;
      }
    }

    if (removedRules > 0) {
      sheet.setConditionalFormatRules(newRules);
      Logger.log(`Removed ${removedRules} orphaned conditional formatting rules`);
    }
  } catch (error) {
    Logger.log(`Error cleaning up orphaned conditional formatting: ${error.message}`);
  }
}

/**
 * Converts a column index to letter (e.g., 1 -> A, 27 -> AA)
 * @param {number} columnIndex - The 1-based column index
 * @return {string} The column letter
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
 * Converts a column letter to index (e.g., A -> 1, AA -> 27)
 * @param {string} columnLetter - The column letter
 * @return {number} The 1-based column index
 */
function columnLetterToIndex(columnLetter) {
  let column = 0;
  const length = columnLetter.length;
  
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  
  return column;
}

/**
 * Writes data to the Google Sheet with columns matching the filter
 */
SyncService.writeDataToSheet = function(items, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get sheet name from options or properties
  const sheetName = options.sheetName ||
    PropertiesService.getScriptProperties().getProperty('SHEET_NAME') ||
    DEFAULT_SHEET_NAME;

  let sheet = ss.getSheetByName(sheetName);

  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // IMPORTANT: Clean up previous sync status column before writing new data
  SyncService.cleanupPreviousSyncStatusColumn(sheet, sheetName);

  // Check if two-way sync is enabled for this sheet
  const scriptProperties = PropertiesService.getScriptProperties();
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
  const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
  const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

  // For preservation of status data - key is entity ID, value is status
  let statusByIdMap = new Map();
  let configuredTrackingColumn = '';
  let statusColumnIndex = -1;

  // ... rest of existing writeDataToSheet code ...

  // CRITICAL: Clean up any lingering formatting from previous Sync Status columns
  // This needs to happen AFTER the sheet is rebuilt
  if (twoWaySyncEnabled && statusColumnIndex !== -1) {
    try {
      Logger.log(`Performing aggressive column cleanup after sheet rebuild - current status column: ${statusColumnIndex}`);

      // The current status column letter
      const currentStatusColLetter = SyncService.columnToLetter(statusColumnIndex + 1);

      // Clean ALL columns in the sheet except the current Sync Status column
      const lastCol = sheet.getLastColumn() + 5; // Add buffer for hidden columns
      for (let i = 0; i < lastCol; i++) {
        if (i !== statusColumnIndex) { // Skip the current status column
          try {
            const colLetter = SyncService.columnToLetter(i + 1);
            Logger.log(`Checking column ${colLetter} for cleanup`);

            // Check if this column has a 'Sync Status' header or sync-related note
            const headerCell = sheet.getRange(1, i + 1);
            const headerValue = headerCell.getValue();
            const note = headerCell.getNote();

            if (headerValue === "Sync Status" ||
              (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive')))) {
              Logger.log(`Found Sync Status indicators in column ${colLetter}, cleaning up`);
              SyncService.cleanupColumnFormatting(sheet, colLetter);
            }
          } catch (e) {
            Logger.log(`Error checking column ${colLetter}: ${e.message}`);
          }
        }
      }
    } catch (error) {
      Logger.log(`Error in aggressive column cleanup: ${error.message}`);
    }
  }
};

/**
 * Cleans up previous Sync Status column formatting
 * @param {Sheet} sheet - The sheet containing the column
 * @param {string} sheetName - The name of the sheet
 */
SyncService.cleanupPreviousSyncStatusColumn = function(sheet, sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const previousColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const currentColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const previousColumn = scriptProperties.getProperty(previousColumnKey) || '';

    // Clean up the specifically tracked previous column
    if (previousColumn && previousColumn !== currentColumn) {
      SyncService.cleanupColumnFormatting(sheet, previousColumn);
      scriptProperties.deleteProperty(previousColumnKey);
    }

    // NEW: Store current column positions for future comparison
    // This helps when columns are deleted and the position shifts
    const currentColumnIndex = currentColumn ? SyncService.columnLetterToIndex(currentColumn) : -1;
    if (currentColumnIndex >= 0) {
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, currentColumnIndex.toString());
    }

    // IMPORTANT: Scan ALL columns for "Sync Status" headers and validation patterns
    SyncService.scanAndCleanupAllSyncColumns(sheet, currentColumn);
  } catch (error) {
    Logger.log(`Error in cleanupPreviousSyncStatusColumn: ${error.message}`);
  }
};

// Export functions to the SyncService namespace
SyncService.syncFromPipedrive = syncFromPipedrive;
SyncService.syncPipedriveDataToSheet = syncPipedriveDataToSheet;
SyncService.isSyncRunning = isSyncRunning;
SyncService.setSyncRunning = setSyncRunning;
SyncService.updateSyncStatus = updateSyncStatus;
SyncService.showSyncStatus = showSyncStatus;
SyncService.getSyncStatus = getSyncStatus;
SyncService.formatEntityTypeName = formatEntityTypeName;
SyncService.cleanupColumnFormatting = cleanupColumnFormatting;
SyncService.scanAndCleanupAllSyncColumns = scanAndCleanupAllSyncColumns;
SyncService.columnToLetter = columnToLetter;
SyncService.columnLetterToIndex = columnLetterToIndex;

/**
 * Saves the two-way sync settings for a sheet and sets up the tracking column
 * @param {boolean} enableTwoWaySync Whether to enable two-way sync
 * @param {string} trackingColumn The column letter to use for tracking changes
 */
function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    
    // Store the previous tracking column if it exists (for cleanup purposes)
    const previousTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';

    // Save the enabled setting
    scriptProperties.setProperty(twoWaySyncEnabledKey, enableTwoWaySync.toString());
    
    // If enabling two-way sync
    if (enableTwoWaySync) {
      // Set up the onEdit trigger
      setupOnEditTrigger();

      // Determine which column to use for tracking
      let trackingColumnIndex;
      if (trackingColumn) {
        // Convert column letter to index (1-based)
        trackingColumnIndex = columnLetterToIndex(trackingColumn);
      } else {
        // Use the last column + 1 (to add a new column at the end)
        trackingColumnIndex = activeSheet.getLastColumn() + 1;
      }

      // Set up the tracking column header
      const headerRow = 1; // Assuming first row is header
      const trackingHeader = "Sync Status";

      // Check if there's already a Sync Status column
      const headers = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
      let existingSyncStatusCol = -1;
      
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] === trackingHeader) {
          existingSyncStatusCol = i + 1; // Convert to 1-based
          break;
        }
      }
      
      // If a Sync Status column already exists, use that instead
      if (existingSyncStatusCol !== -1) {
        trackingColumnIndex = existingSyncStatusCol;
        
        // Update the tracking column property
        trackingColumn = columnToLetter(trackingColumnIndex);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
        
        Logger.log(`Found existing Sync Status column at ${trackingColumn} (${trackingColumnIndex})`);
      } else {
        // No existing column, create a new one
        activeSheet.getRange(headerRow, trackingColumnIndex).setValue(trackingHeader);
        
        // Update the tracking column property
        trackingColumn = columnToLetter(trackingColumnIndex);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
        
        Logger.log(`Created new Sync Status column at ${trackingColumn} (${trackingColumnIndex})`);
      }

      // Style the header cell
      const headerCell = activeSheet.getRange(headerRow, trackingColumnIndex);
      headerCell
        .setBackground('#E8F0FE')
        .setFontWeight('bold')
        .setNote('This column tracks changes for two-way sync with Pipedrive');

      // Add data validation for all cells in the column
      const lastRow = Math.max(activeSheet.getLastRow(), 2);
      if (lastRow > 1) {
        const dataRange = activeSheet.getRange(2, trackingColumnIndex, lastRow - 1, 1);
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
          .build();
        dataRange.setDataValidation(rule);

        // Set default values for empty cells
        const values = dataRange.getValues();
        const newValues = values.map(row => {
          return [row[0] || 'Not modified'];
        });
        dataRange.setValues(newValues);

        // Style the status column
        dataRange
          .setBackground('#F8F9FA')
          .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

        // Set up conditional formatting
        const rules = activeSheet.getConditionalFormatRules();

        // Add rule for "Modified" status - red background
        const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Modified')
          .setBackground('#FCE8E6')
          .setFontColor('#D93025')
          .setRanges([dataRange])
          .build();
        rules.push(modifiedRule);

        // Add rule for "Synced" status - green background
        const syncedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Synced')
          .setBackground('#E6F4EA')
          .setFontColor('#137333')
          .setRanges([dataRange])
          .build();
        rules.push(syncedRule);

        // Add rule for "Error" status - yellow background
        const errorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Error')
          .setBackground('#FEF7E0')
          .setFontColor('#B06000')
          .setRanges([dataRange])
          .build();
        rules.push(errorRule);

        activeSheet.setConditionalFormatRules(rules);
      }
      
      // Store the current column position for comparison during cleanup
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${activeSheetName}`, (trackingColumnIndex - 1).toString());
      
      // If this is a different column than before, store the previous one for cleanup
      if (previousTrackingColumn && previousTrackingColumn !== trackingColumn) {
        scriptProperties.setProperty(`PREVIOUS_TRACKING_COLUMN_${activeSheetName}`, previousTrackingColumn);
      }
    } else {
      // If disabling two-way sync, remove the trigger
      removeOnEditTrigger();

      // Clean up the tracking column if it exists
      if (previousTrackingColumn) {
        cleanupColumnFormatting(activeSheet, previousTrackingColumn);
      }
    }

    // Return success
    return true;
  } catch (error) {
    Logger.log(`Error in saveTwoWaySyncSettings: ${error.message}`);
    return false;
  }
}

/**
 * Stores original data from Pipedrive for later comparison
 * @param {Array} items - The data items from Pipedrive
 * @param {Object} options - Options including sheet name
 */
function storeOriginalData(items, options) {
  try {
    if (!items || !options || !options.sheetName) return;
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const originalDataKey = `ORIGINAL_DATA_${options.sheetName}`;
    
    // Create a map of original values keyed by record ID
    const originalData = {};
    
    items.forEach(item => {
      // Get the item ID (should be the first column)
      const id = item.id;
      if (!id) return;
      
      // Extract values for this item
      const rowData = {};
      
      // Extract values for all configured columns
      options.columns.forEach((column, index) => {
        const columnKey = typeof column === 'object' ? column.key : column;
        const columnName = options.headerRow[index];
        
        if (columnName) {
          const value = getValueByPath(item, columnKey);
          rowData[columnName] = formatValue(value, columnKey, options.optionMappings);
        }
      });
      
      // Store this row's data
      originalData[id.toString()] = rowData;
    });
    
    // Store the original data in script properties
    scriptProperties.setProperty(originalDataKey, JSON.stringify(originalData));
    Logger.log(`Stored original data for ${items.length} records in sheet ${options.sheetName}`);
  } catch (error) {
    Logger.log(`Error in storeOriginalData: ${error.message}`);
  }
}

// Export the onEdit function to the global scope for the trigger to work correctly
this.onEdit = onEdit;