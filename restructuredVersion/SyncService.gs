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
    // First, remove any existing onEdit triggers to avoid duplicates
    removeOnEditTrigger();
    
    // Then create a new trigger
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
        if ((now - lockData.timestamp) < 5000) {
          Logger.log(`Exiting due to active lock: ${currentLock}`);
          return;
        }
        
        // Lock is old, we can override it
        Logger.log(`Override old lock from ${lockData.timestamp}`);
      }
      
      // Set a new lock
      scriptProperties.setProperty(lockKey, JSON.stringify({
        id: executionId,
        timestamp: new Date().getTime(),
        row: row,
        col: column
      }));
    } catch (lockError) {
      Logger.log(`Error setting lock: ${lockError.message}`);
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
      releaseLock(executionId, lockKey);
      return;
    }
    
    // Convert to 1-based for sheet functions
    const syncStatusColPos = syncStatusColIndex + 1;
    
    // Check if the edit is in the Sync Status column itself (to avoid loops)
    if (column === syncStatusColPos) {
      releaseLock(executionId, lockKey);
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
      cellState = cellStateJson ? JSON.parse(cellStateJson) : { status: null, lastChanged: 0, originalValues: {} };
    } catch (parseError) {
      Logger.log(`Error parsing cell state: ${parseError.message}`);
      cellState = { status: null, lastChanged: 0, originalValues: {} };
    }
    
    // Get the current time
    const now = new Date().getTime();
    
    // Check for recent changes to prevent toggling
    if (cellState.lastChanged && (now - cellState.lastChanged) < 5000 && cellState.status === currentStatus) {
      Logger.log(`Cell was recently changed to "${currentStatus}", skipping update`);
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
      Logger.log(`Error parsing original data: ${parseError.message}`);
      originalData = {};
    }
    
    // Check if we have original data for this row
    const rowKey = id.toString();
    
    // Log for debugging
    Logger.log(`onEdit triggered - Row: ${row}, Column: ${column}, Status: ${currentStatus}`);
    Logger.log(`Row ID: ${id}, Cell Value: ${e.value}, Old Value: ${e.oldValue}`);
    
    // Get the column header name for the edited column
    const headerName = headers[column - 1]; // Adjust for 0-based array
    
    // Enhanced debug logging
    Logger.log(`UNDO_DEBUG: Cell edit in ${sheetName} - Header: "${headerName}", Row: ${row}`);
    Logger.log(`UNDO_DEBUG: Current status: "${currentStatus}"`);
    
    // Handle first-time edit case (current status is not Modified)
    if (currentStatus !== "Modified") {
      // Store the original value before marking as modified
      if (!originalData[rowKey]) {
        originalData[rowKey] = {};
      }
      
      if (headerName) {
        // Store the original value
        originalData[rowKey][headerName] = e.oldValue !== undefined ? e.oldValue : null;
        Logger.log(`UNDO_DEBUG: First edit - storing original value: "${e.oldValue}" for ${headerName}`);
        
        // Save updated original data
        try {
          scriptProperties.setProperty(originalDataKey, JSON.stringify(originalData));
        } catch (saveError) {
          Logger.log(`Error saving original data: ${saveError.message}`);
        }
        
        // Mark as modified (with special prevention of change-back)
        syncStatusCell.setValue("Modified");
        Logger.log(`Changed status to Modified for row ${row}, column ${column}, header ${headerName}`);
        
        // Save new cell state to prevent toggling back
        cellState.status = "Modified";
        cellState.lastChanged = now;
        cellState.originalValues[headerName] = e.oldValue;
        
        try {
          scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
        } catch (saveError) {
          Logger.log(`Error saving cell state: ${saveError.message}`);
        }

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
      if (headerName && originalData[rowKey] && originalData[rowKey][headerName] !== undefined) {
        const originalValue = originalData[rowKey][headerName];
        const currentValue = e.value;
        
        Logger.log(`UNDO_DEBUG: === COMPARING VALUES FOR FIELD: ${headerName} ===`);
        Logger.log(`UNDO_DEBUG: Original value: ${JSON.stringify(originalValue)} (type: ${typeof originalValue})`);
        Logger.log(`UNDO_DEBUG: Current value: ${JSON.stringify(currentValue)} (type: ${typeof currentValue})`);
        
        // Improved equality check - try to normalize values for comparison regardless of type
        let originalString = originalValue === null || originalValue === undefined ? '' : String(originalValue).trim();
        let currentString = currentValue === null || currentValue === undefined ? '' : String(currentValue).trim();
        
        // Special handling for email fields - normalize domains for comparison
        if (headerName.toLowerCase().includes('email')) {
          // Apply email normalization rules
          if (originalString.includes('@')) {
            const origParts = originalString.split('@');
            const origUsername = origParts[0].toLowerCase();
            let origDomain = origParts[1].toLowerCase();
            
            // Fix common domain typos
            if (origDomain === 'gmail.comm') origDomain = 'gmail.com';
            if (origDomain === 'gmail.con') origDomain = 'gmail.com';
            if (origDomain === 'gmial.com') origDomain = 'gmail.com';
            if (origDomain === 'hotmail.comm') origDomain = 'hotmail.com';
            if (origDomain === 'hotmail.con') origDomain = 'hotmail.com';
            if (origDomain === 'yahoo.comm') origDomain = 'yahoo.com';
            if (origDomain === 'yahoo.con') origDomain = 'yahoo.com';
            
            // Reassemble normalized email
            originalString = origUsername + '@' + origDomain;
          }
          
          if (currentString.includes('@')) {
            const currParts = currentString.split('@');
            const currUsername = currParts[0].toLowerCase();
            let currDomain = currParts[1].toLowerCase();
            
            // Fix common domain typos
            if (currDomain === 'gmail.comm') currDomain = 'gmail.com';
            if (currDomain === 'gmail.con') currDomain = 'gmail.com';
            if (currDomain === 'gmial.com') currDomain = 'gmail.com';
            if (currDomain === 'hotmail.comm') currDomain = 'hotmail.com';
            if (currDomain === 'hotmail.con') currDomain = 'hotmail.com';
            if (currDomain === 'yahoo.comm') currDomain = 'yahoo.com';
            if (currDomain === 'yahoo.con') currDomain = 'yahoo.com';
            
            // Reassemble normalized email
            currentString = currUsername + '@' + currDomain;
          }
          
          Logger.log(`UNDO_DEBUG: Normalized emails for comparison - Original: "${originalString}", Current: "${currentString}"`);
        }
        // Special handling for name fields - normalize common typos
        else if (headerName.toLowerCase().includes('name')) {
          // Check for common name typos like extra letter at the end
          if (originalString.length > 0 && currentString.length > 0) {
            // Check if one string is the same as the other with an extra character at the end
            if (originalString.length === currentString.length + 1) {
              if (originalString.startsWith(currentString)) {
                Logger.log(`UNDO_DEBUG: Name has extra char at end of original: "${originalString}" vs "${currentString}"`);
                originalString = currentString;
              }
            } 
            else if (currentString.length === originalString.length + 1) {
              if (currentString.startsWith(originalString)) {
                Logger.log(`UNDO_DEBUG: Name has extra char at end of current: "${currentString}" vs "${originalString}"`);
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
                Logger.log(`UNDO_DEBUG: Name has character mismatch near end: "${originalString}" vs "${currentString}"`);
                // Normalize by taking the shorter version up to the differing character
                const normalizedName = originalString.substring(0, diffIndex);
                originalString = normalizedName;
                currentString = normalizedName;
              }
            }
          }
          
          Logger.log(`UNDO_DEBUG: Normalized names for comparison - Original: "${originalString}", Current: "${currentString}"`);
        }
        
        // For numeric values, try to normalize scientific notation and number formats
        if (!isNaN(parseFloat(originalString)) && !isNaN(parseFloat(currentString))) {
          // Convert both to numbers and back to strings for comparison
          try {
            const origNum = parseFloat(originalString);
            const currNum = parseFloat(currentString);
            
            // If both are integers, compare as integers
            if (Math.floor(origNum) === origNum && Math.floor(currNum) === currNum) {
              originalString = Math.floor(origNum).toString();
              currentString = Math.floor(currNum).toString();
              Logger.log(`UNDO_DEBUG: Normalized as integers: "${originalString}" vs "${currentString}"`);
            } else {
              // Compare with fixed decimal places for floating point numbers
              originalString = origNum.toString();
              currentString = currNum.toString();
              Logger.log(`UNDO_DEBUG: Normalized as floats: "${originalString}" vs "${currentString}"`);
            }
          } catch (numError) {
            Logger.log(`UNDO_DEBUG: Error normalizing numbers: ${numError.message}`);
          }
        }
        
        // Check if this is a structural field with complex nested structure
        if (originalValue && typeof originalValue === 'object' && originalValue.__isStructural) {
          Logger.log(`DEBUG: Found structural field with key ${originalValue.__key}`);
          
          // Simple direct comparison before complex checks
          if (originalString === currentString) {
            Logger.log(`UNDO_DEBUG: Direct string comparison match for structural field: "${originalString}" = "${currentString}"`);
            
            // Check if all edited values in the row now match original values
            Logger.log(`UNDO_DEBUG: Checking if all fields in row match original values`);
            const allMatch = checkAllValuesMatchOriginal(sheet, row, headers, originalData[rowKey]);
            
            Logger.log(`UNDO_DEBUG: All values match original: ${allMatch}`);
            
            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");
              Logger.log(`UNDO_DEBUG: Reset to Not modified for row ${row} - all values match original`);
              
              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true;  // Special flag to indicate this is an undo operation
              
              try {
                scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
              } catch (saveError) {
                Logger.log(`Error saving cell state: ${saveError.message}`);
              }
              
              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(noChangeLockKey, JSON.stringify({
                  timestamp: now,
                  expiry: now + 10000, // 10 seconds
                  status: "Not modified"
                }));
              } catch (lockError) {
                Logger.log(`Error setting no-change lock: ${lockError.message}`);
              }
              
              // Re-apply data validation
              const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
                .build();
              syncStatusCell.setDataValidation(rule);
              
              // Reset formatting
              syncStatusCell.setBackground('#F8F9FA').setFontColor('#000000');
            }
            
            return;
          }
          
          // Create a data object that mimics Pipedrive structure
          const dataObj = { id: id };
          
          // Try to reconstruct the structure based on the __key
          const key = originalValue.__key;
          const parts = key.split('.');
          const structureType = parts[0];
          
          if (['phone', 'email'].includes(structureType)) {
            // Handle phone/email fields
            dataObj[structureType] = [];
            
            // If it's a label-based path (e.g., phone.mobile)
            if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
              Logger.log(`DEBUG: Processing labeled ${structureType} field with label ${parts[1]}`);
              dataObj[structureType].push({
                label: parts[1],
                value: currentValue
              });
            } 
            // If it's an array index path (e.g., phone.0.value)
            else if (parts.length === 3 && parts[2] === 'value') {
              const idx = parseInt(parts[1]);
              Logger.log(`DEBUG: Processing indexed ${structureType} field at position ${idx}`);
              while (dataObj[structureType].length <= idx) {
                dataObj[structureType].push({});
              }
              dataObj[structureType][idx].value = currentValue;
            }
          }
          // Custom fields
          else if (structureType === 'custom_fields') {
            dataObj.custom_fields = {};
            
            if (parts.length === 2) {
              // Simple custom field
              Logger.log(`DEBUG: Processing simple custom field ${parts[1]}`);
              dataObj.custom_fields[parts[1]] = currentValue;
            } 
            else if (parts.length > 2) {
              // Nested custom field like address or currency
              Logger.log(`DEBUG: Processing complex custom field ${parts[1]}.${parts[2]}`);
              dataObj.custom_fields[parts[1]] = {};
              
              // Handle complex types
              if (parts[2] === 'formatted_address') {
                dataObj.custom_fields[parts[1]].formatted_address = currentValue;
              } 
              else if (parts[2] === 'currency') {
                dataObj.custom_fields[parts[1]].currency = currentValue;
              }
              else {
                dataObj.custom_fields[parts[1]][parts[2]] = currentValue;
              }
            }
          } else {
            // Other nested fields not covered above
            Logger.log(`DEBUG: Processing general nested field with key: ${key}`);
            
            // Build a generic nested structure
            let current = dataObj;
            for (let i = 0; i < parts.length - 1; i++) {
              if (current[parts[i]] === undefined) {
                if (!isNaN(parseInt(parts[i+1]))) {
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
          const normalizedOriginal = originalValue.__normalized || '';
          const normalizedCurrent = getNormalizedFieldValue(dataObj, key);
          
          Logger.log(`DEBUG: Structural comparison - Original: "${normalizedOriginal}", Current: "${normalizedCurrent}"`);
          
          // Check if values match
          const valuesMatch = normalizedOriginal === normalizedCurrent;
          Logger.log(`DEBUG: Structural values match: ${valuesMatch}`);
          
          // If values match, check all fields
          if (valuesMatch) {
            // Check if all edited values in the row now match original values
            Logger.log(`DEBUG: Checking if all fields in row match original values`);
            const allMatch = checkAllValuesMatchOriginal(sheet, row, headers, originalData[rowKey]);
            
            Logger.log(`DEBUG: All values match original: ${allMatch}`);
            
            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");
              Logger.log(`DEBUG: Reset to Not modified for row ${row} - all values match original`);
              
              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true;  // Special flag to indicate this is an undo operation
              
              try {
                scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
              } catch (saveError) {
                Logger.log(`Error saving cell state: ${saveError.message}`);
              }
              
              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(noChangeLockKey, JSON.stringify({
                  timestamp: now,
                  expiry: now + 10000, // 10 seconds
                  status: "Not modified"
                }));
              } catch (lockError) {
                Logger.log(`Error setting no-change lock: ${lockError.message}`);
              }
              
              // Re-apply data validation
              const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
                .build();
              syncStatusCell.setDataValidation(rule);
              
              // Reset formatting
              syncStatusCell.setBackground('#F8F9FA').setFontColor('#000000');
            }
          }
        } else {
          // This is a regular field, not a structural field
          
          // Special handling for null/empty values
          if ((originalValue === null || originalValue === "") && 
              (currentValue === null || currentValue === "")) {
            Logger.log(`DEBUG: Both values are empty, treating as match`);
          }
          
          // Simple direct comparison before complex checks
          if (originalString === currentString) {
            Logger.log(`UNDO_DEBUG: Direct string comparison match for regular field: "${originalString}" = "${currentString}"`);
            
            // Check if all edited values in the row now match original values
            Logger.log(`UNDO_DEBUG: Checking if all fields in row match original values`);
            const allMatch = checkAllValuesMatchOriginal(sheet, row, headers, originalData[rowKey]);
            
            Logger.log(`UNDO_DEBUG: All values match original: ${allMatch}`);
            
            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");
              Logger.log(`UNDO_DEBUG: Reset to Not modified for row ${row} - all values match original`);
              
              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true;  // Special flag to indicate this is an undo operation
              
              try {
                scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
              } catch (saveError) {
                Logger.log(`Error saving cell state: ${saveError.message}`);
              }
              
              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(noChangeLockKey, JSON.stringify({
                  timestamp: now,
                  expiry: now + 10000, // 10 seconds
                  status: "Not modified"
                }));
              } catch (lockError) {
                Logger.log(`Error setting no-change lock: ${lockError.message}`);
              }
              
              // Re-apply data validation
              const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
                .build();
              syncStatusCell.setDataValidation(rule);
              
              // Reset formatting
              syncStatusCell.setBackground('#F8F9FA').setFontColor('#000000');
            }
            
            return;
          }
          
          // Create a data object that mimics Pipedrive structure
          const dataObj = { id: id };
          
          // Populate the field being edited
          if (headerName.includes('.')) {
            // Handle nested structure
            const parts = headerName.split('.');
            Logger.log(`DEBUG: Building nested structure with parts: ${parts}`);
            
            if (['phone', 'email'].includes(parts[0])) {
              // Handle phone/email fields
              dataObj[parts[0]] = [];
              
              // If it's a label-based path (e.g., phone.mobile)
              if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
                Logger.log(`DEBUG: Adding label-based ${parts[0]} field with label ${parts[1]}`);
                dataObj[parts[0]].push({
                  label: parts[1],
                  value: currentValue
                });
              } 
              // If it's an array index path (e.g., phone.0.value)
              else if (parts.length === 3 && parts[2] === 'value') {
                const idx = parseInt(parts[1]);
                Logger.log(`DEBUG: Adding array-index ${parts[0]} field at index ${idx}`);
                while (dataObj[parts[0]].length <= idx) {
                  dataObj[parts[0]].push({});
                }
                dataObj[parts[0]][idx].value = currentValue;
              }
            }
            // Custom fields
            else if (parts[0] === 'custom_fields') {
              Logger.log(`DEBUG: Adding custom_fields structure`);
              dataObj.custom_fields = {};
              
              if (parts.length === 2) {
                // Simple custom field
                Logger.log(`DEBUG: Adding simple custom field ${parts[1]}`);
                dataObj.custom_fields[parts[1]] = currentValue;
              } 
              else if (parts.length > 2) {
                // Nested custom field like address or currency
                Logger.log(`DEBUG: Adding complex custom field ${parts[1]} with subfield ${parts[2]}`);
                dataObj.custom_fields[parts[1]] = {};
                
                // Handle complex types
                if (parts[2] === 'formatted_address') {
                  dataObj.custom_fields[parts[1]].formatted_address = currentValue;
                } 
                else if (parts[2] === 'currency') {
                  dataObj.custom_fields[parts[1]].currency = currentValue;
                }
                else {
                  dataObj.custom_fields[parts[1]][parts[2]] = currentValue;
                }
              }
            } else {
              // Other nested fields not covered above
              Logger.log(`DEBUG: Unhandled nested field type: ${parts[0]}`);
              
              // Build a generic nested structure
              let current = dataObj;
              for (let i = 0; i < parts.length - 1; i++) {
                if (!current[parts[i]]) {
                  current[parts[i]] = {};
                }
                current = current[parts[i]];
              }
              current[parts[parts.length - 1]] = currentValue;
              
              Logger.log(`DEBUG: Created generic nested structure: ${JSON.stringify(dataObj)}`);
            }
          } else {
            // Regular top-level field
            Logger.log(`DEBUG: Adding top-level field ${headerName}`);
            dataObj[headerName] = currentValue;
          }
          
          // Dump the constructed data object
          Logger.log(`DEBUG: Constructed data object: ${JSON.stringify(dataObj)}`);
          
          // Use the generalized field value normalization for comparison
          const normalizedOriginal = getNormalizedFieldValue({ [headerName]: originalValue }, headerName);
          const normalizedCurrent = getNormalizedFieldValue(dataObj, headerName);
          
          Logger.log(`DEBUG: Original type: ${typeof originalValue}, Current type: ${typeof currentValue}`);
          Logger.log(`DEBUG: Normalized Original: "${normalizedOriginal}"`);
          Logger.log(`DEBUG: Normalized Current: "${normalizedCurrent}"`);
          
          // Check if values match
          const valuesMatch = normalizedOriginal === normalizedCurrent;
          Logger.log(`DEBUG: Values match: ${valuesMatch}`);
          
          // If values match, check all fields
          if (valuesMatch) {
            // Check if all edited values in the row now match original values
            Logger.log(`DEBUG: Checking if all fields in row match original values`);
            const allMatch = checkAllValuesMatchOriginal(sheet, row, headers, originalData[rowKey]);
            
            Logger.log(`DEBUG: All values match original: ${allMatch}`);
            
            if (allMatch) {
              // All values in row match original - reset to Not modified
              syncStatusCell.setValue("Not modified");
              Logger.log(`DEBUG: Reset to Not modified for row ${row} - all values match original`);
              
              // Save new cell state with strong protection against toggling back
              cellState.status = "Not modified";
              cellState.lastChanged = now;
              cellState.isUndone = true;  // Special flag to indicate this is an undo operation
              
              try {
                scriptProperties.setProperty(cellStateKey, JSON.stringify(cellState));
              } catch (saveError) {
                Logger.log(`Error saving cell state: ${saveError.message}`);
              }
              
              // Create a temporary lock to prevent changes for 10 seconds
              const noChangeLockKey = `NO_CHANGE_LOCK_${sheetName}_${row}`;
              try {
                scriptProperties.setProperty(noChangeLockKey, JSON.stringify({
                  timestamp: now,
                  expiry: now + 10000, // 10 seconds
                  status: "Not modified"
                }));
              } catch (lockError) {
                Logger.log(`Error setting no-change lock: ${lockError.message}`);
              }
              
              // Re-apply data validation
              const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
                .build();
              syncStatusCell.setDataValidation(rule);

              // Reset formatting
              syncStatusCell.setBackground('#F8F9FA').setFontColor('#000000');
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
          Logger.log(`Stored original value "${e.oldValue}" for ${rowKey}.${headerName}`);
          
          // Save updated original data
          try {
            scriptProperties.setProperty(originalDataKey, JSON.stringify(originalData));
          } catch (saveError) {
            Logger.log(`Error saving original data: ${saveError.message}`);
          }
        }
      }
    }
    
    // Release the lock at the end
    releaseLock(executionId, lockKey);
  } catch (error) {
    // Silent fail for onEdit triggers
    Logger.log(`Error in onEdit trigger: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
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
        Logger.log(`Released lock: ${executionId}`);
      }
    }
  } catch (error) {
    Logger.log(`Error releasing lock: ${error.message}`);
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
      Logger.log('No original values stored to compare against');
      return false;
    }
    
    // Get current values for the entire row
    const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    
    // Get the first value (ID) to use for retrieving the original data
    const id = rowValues[0];
    
    // Create a data object that mimics Pipedrive structure for nested field handling
    const dataObj = { id: id };
    
    // Populate the data object with values from the row
    headers.forEach((header, index) => {
      if (index < rowValues.length) {
        // Use dot notation to create nested objects
        if (header.includes('.')) {
          const parts = header.split('.');
          
          // Common nested structures to handle specially
          if (['phone', 'email'].includes(parts[0])) {
            // Handle phone/email specially
            if (!dataObj[parts[0]]) {
              dataObj[parts[0]] = [];
            }
            
            // If it's a label-based path (e.g., phone.mobile)
            if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
              dataObj[parts[0]].push({
                label: parts[1],
                value: rowValues[index]
              });
            } 
            // If it's an array index path (e.g., phone.0.value)
            else if (parts.length === 3 && parts[2] === 'value') {
              const idx = parseInt(parts[1]);
              while (dataObj[parts[0]].length <= idx) {
                dataObj[parts[0]].push({});
              }
              dataObj[parts[0]][idx].value = rowValues[index];
            }
          }
          // Custom fields
          else if (parts[0] === 'custom_fields') {
            if (!dataObj.custom_fields) {
              dataObj.custom_fields = {};
            }
            
            if (parts.length === 2) {
              // Simple custom field
              dataObj.custom_fields[parts[1]] = rowValues[index];
            } 
            else if (parts.length > 2) {
              // Nested custom field
              if (!dataObj.custom_fields[parts[1]]) {
                dataObj.custom_fields[parts[1]] = {};
              }
              // Handle complex types like address
              if (parts[2] === 'formatted_address') {
                dataObj.custom_fields[parts[1]].formatted_address = rowValues[index];
              } 
              else if (parts[2] === 'currency') {
                dataObj.custom_fields[parts[1]].currency = rowValues[index];
              }
              else {
                dataObj.custom_fields[parts[1]][parts[2]] = rowValues[index];
              }
            }
          } else {
            // Other nested paths - build the structure
            let current = dataObj;
            for (let i = 0; i < parts.length - 1; i++) {
              if (current[parts[i]] === undefined) {
                // If part is numeric, create an array
                if (!isNaN(parseInt(parts[i+1]))) {
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
    Logger.log(`Checking ${Object.keys(originalValues).length} fields for original value match`);
    
    // Check each column that has a stored original value
    for (const headerName in originalValues) {
      // Find the column index for this header
      const colIndex = headers.indexOf(headerName);
      if (colIndex === -1) {
        Logger.log(`Header ${headerName} not found in current headers`);
        continue; // Header not found
      }
      
      const originalValue = originalValues[headerName];
      const currentValue = rowValues[colIndex];
      
      // Special handling for null/empty values
      if ((originalValue === null || originalValue === "") && 
          (currentValue === null || currentValue === "")) {
        Logger.log(`Both values are empty for ${headerName}, treating as match`);
        continue; // Both empty, consider a match
      }
      
      // NEW: Check if this is a structural field with complex nested structure
      if (originalValue && typeof originalValue === 'object' && originalValue.__isStructural) {
        Logger.log(`Found structural field ${headerName} with key ${originalValue.__key}`);
        
        // Use the pre-computed normalized value for comparison
        const normalizedOriginal = originalValue.__normalized || '';
        const normalizedCurrent = getNormalizedFieldValue(dataObj, originalValue.__key);
        
        Logger.log(`Structural field comparison for ${headerName}: Original="${normalizedOriginal}", Current="${normalizedCurrent}"`);
        
        // If the normalized values don't match, return false
        if (normalizedOriginal !== normalizedCurrent) {
          Logger.log(`Structural field mismatch found for ${headerName}`);
          return false;
        }
        
        // Skip to the next field
        continue;
      }
      
      // Use the generalized field value normalization for regular fields
      const normalizedOriginal = getNormalizedFieldValue({ [headerName]: originalValue }, headerName);
      const normalizedCurrent = getNormalizedFieldValue(dataObj, headerName);
      
      Logger.log(`Field comparison for ${headerName}: Original="${normalizedOriginal}", Current="${normalizedCurrent}"`);
      
      // If the normalized values don't match, return false
      if (normalizedOriginal !== normalizedCurrent) {
        Logger.log(`Mismatch found for ${headerName}`);
        return false;
      }
    }
    
    // If we reach here, all values with stored originals match
    Logger.log('All values match original values');
    return true;
  } catch (error) {
    Logger.log(`Error in checkAllValuesMatchOriginal: ${error.message}`);
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
      return '';
    }
    
    // Handle array/object (like Pipedrive phone fields)
    if (typeof value === 'object') {
      // If it's an array of phone objects (Pipedrive format)
      if (Array.isArray(value) && value.length > 0) {
        if (value[0] && value[0].value) {
          // Use the first phone number's value
          value = value[0].value;
        } else if (typeof value[0] === 'string') {
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
    Logger.log(`Error normalizing phone number: ${e.message}`);
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
      return '';
    }
    
    // If it's a number or scientific notation, convert to a regular number string
    if (typeof value === 'number' || 
        (typeof value === 'string' && value.includes('E'))) {
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
    return strValue.replace(/\D/g, '');
  } catch (e) {
    Logger.log(`Error in normalizeDigitsOnly: ${e.message}`);
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
    if (!data || !key) return '';
    
    // If key is already a value, just normalize it
    if (typeof key !== 'string') {
      return normalizePhoneNumber(key);
    }
    
    // Handle different path formats for phone numbers
    
    // Case 1: Direct phone field (e.g., "phone")
    if (key === 'phone' && data.phone) {
      return normalizePhoneNumber(data.phone);
    }
    
    // Case 2: Specific label format (e.g., "phone.mobile")
    if (key.startsWith('phone.') && key.split('.').length === 2) {
      const label = key.split('.')[1];
      
      // Handle the array of phone objects with labels
      if (Array.isArray(data.phone)) {
        // Try to find a phone with the matching label
        const match = data.phone.find(p => 
          p && p.label && p.label.toLowerCase() === label.toLowerCase()
        );
        
        if (match && match.value) {
          return normalizePhoneNumber(match.value);
        }
        
        // If not found but we were looking for primary, try to find primary flag
        if (label === 'primary') {
          const primary = data.phone.find(p => p && p.primary);
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
    if (key.startsWith('phone.') && key.includes('.value')) {
      const parts = key.split('.');
      const index = parseInt(parts[1]);
      
      if (!isNaN(index) && Array.isArray(data.phone) && 
          data.phone.length > index && data.phone[index]) {
        return normalizePhoneNumber(data.phone[index].value);
      }
    }
    
    // Case 4: Use getValueByPath as a fallback
    // This handles other complex nested paths
    try {
      const value = getValueByPath(data, key);
      return normalizePhoneNumber(value);
    } catch (e) {
      Logger.log(`Error getting phone value by path: ${e.message}`);
    }
    
    // If all else fails, return empty string
    return '';
  } catch (e) {
    Logger.log(`Error in getPhoneNumberFromField: ${e.message}`);
    return '';
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
        .setRanges([statusRange])
        .build();
      newRules.push(modifiedRule);

      // Create conditional format for "Synced" status
      const syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA')  // Light green background
        .setFontColor('#137333')   // Green text
        .setRanges([statusRange])
        .build();
      newRules.push(syncedRule);

      // Create conditional format for "Error" status
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Error')
        .setBackground('#FEF7E0')
        .setFontColor('#B06000')
        .setRanges([statusRange])
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
    
    Logger.log(`STORE_DEBUG: Storing original data for ${items.length} items`);
    
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
          let rawValue;
          try {
            // Get the raw value from the item using the column key
            rawValue = getValueByPath(item, columnKey);
            
            // For certain field types, we need to preserve the complex structure
            const isStructuralField = 
              columnKey.startsWith('phone.') || 
              columnKey.startsWith('email.') || 
              columnKey.startsWith('custom_fields.') ||
              columnKey.includes('.value');
            
            if (isStructuralField) {
              // For structural fields, store an object with the key and original structure
              // This helps with complex nested fields during comparison
              Logger.log(`STORE_DEBUG: Storing structural field ${columnName} with key ${columnKey}`);
              
              // Create a subobject for handling in onEdit
              const structureType = columnKey.split('.')[0];
              
              if (structureType === 'phone' || structureType === 'email') {
                // Store the whole array/object for phone/email fields
                if (item[structureType]) {
                  rowData[columnName] = {
                    __isStructural: true,
                    __key: columnKey,
                    __value: item[structureType],
                    __normalized: getNormalizedFieldValue(item, columnKey)
                  };
                  Logger.log(`STORE_DEBUG: Stored complex ${structureType} structure for column ${columnName}`);
                } else {
                  rowData[columnName] = formatValue(rawValue, columnKey, options.optionMappings);
                }
              } else if (structureType === 'custom_fields') {
                // Store custom field structure
                if (item.custom_fields) {
                  const parts = columnKey.split('.');
                  const fieldKey = parts[1];
                  
                  rowData[columnName] = {
                    __isStructural: true,
                    __key: columnKey,
                    __value: item.custom_fields[fieldKey],
                    __normalized: getNormalizedFieldValue(item, columnKey)
                  };
                  Logger.log(`STORE_DEBUG: Stored complex custom_field structure for column ${columnName}`);
                } else {
                  rowData[columnName] = formatValue(rawValue, columnKey, options.optionMappings);
                }
              } else {
                // Handle other nested structures
                try {
                  rowData[columnName] = {
                    __isStructural: true,
                    __key: columnKey,
                    __value: rawValue,
                    __normalized: getNormalizedFieldValue(item, columnKey)
                  };
                  Logger.log(`STORE_DEBUG: Stored complex nested structure for column ${columnName}`);
                } catch (e) {
                  rowData[columnName] = formatValue(rawValue, columnKey, options.optionMappings);
                }
              }
            } else {
              // For regular fields, just store the formatted value
              rowData[columnName] = formatValue(rawValue, columnKey, options.optionMappings);
            }
          } catch (valueError) {
            Logger.log(`STORE_DEBUG: Error getting value for column ${columnName}: ${valueError.message}`);
            rowData[columnName] = null;
          }
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

// Export the trigger functions to the SyncService namespace
SyncService.setupOnEditTrigger = setupOnEditTrigger;
SyncService.removeOnEditTrigger = removeOnEditTrigger;

/**
 * Gets the normalized value from a field regardless of its structure
 * This is crucial for handling different ways data is structured in Pipedrive
 * @param {Object} data - The data object containing the field
 * @param {string} key - The key or path to the field
 * @return {string} The normalized value for comparison
 */
function getNormalizedFieldValue(data, key) {
  try {
    Logger.log(`FIELD_DEBUG: Starting field normalization for key "${key}"`);
    
    if (!data || !key) {
      Logger.log(`FIELD_DEBUG: Empty data or key`);
      return '';
    }
    
    // If key is already a value, normalize it based on type
    if (typeof key !== 'string') {
      Logger.log(`FIELD_DEBUG: Key is a value of type ${typeof key}, normalizing directly`);
      return normalizeValueByType(key);
    }
    
    // Dump what we received for detailed debugging
    try {
      Logger.log(`FIELD_DEBUG: Data object keys: ${Object.keys(data)}`);
      
      // If the key has a period, log what we find at each level
      if (key.includes('.')) {
        const parts = key.split('.');
        let partData = data;
        
        for (let i = 0; i < parts.length; i++) {
          const part = parts[i];
          if (partData && typeof partData === 'object') {
            Logger.log(`FIELD_DEBUG: At level ${i}, part "${part}", found keys: ${Object.keys(partData)}`);
            if (partData[part] !== undefined) {
              Logger.log(`FIELD_DEBUG: Found value at level ${i}: ${typeof partData[part]}`);
              partData = partData[part];
            } else {
              Logger.log(`FIELD_DEBUG: No value found at level ${i}`);
              break;
            }
          } else {
            Logger.log(`FIELD_DEBUG: At level ${i}, partData is not an object`);
            break;
          }
        }
      } else if (data[key] !== undefined) {
        Logger.log(`FIELD_DEBUG: Direct key lookup found type: ${typeof data[key]}`);
      }
    } catch (debugError) {
      Logger.log(`FIELD_DEBUG: Error in debug logging: ${debugError.message}`);
    }
    
    // First try to get the value using standard path lookup
    let value;
    try {
      value = getValueByPath(data, key);
      Logger.log(`FIELD_DEBUG: getValueByPath returned: ${JSON.stringify(value)}`);
    } catch (e) {
      Logger.log(`FIELD_DEBUG: Error getting value by path: ${e.message}`);
      // Continue with specialized handling
    }
    
    // Special case for phone fields - we want to handle them specially to ignore formatting
    // Check if this is a phone field based on the key name
    const isPhoneField = 
      key === 'phone' || 
      key.startsWith('phone.') ||
      (typeof key === 'string' && 
        (key.toLowerCase().includes('phone') || 
         key.toLowerCase().includes('mobile') || 
         key.toLowerCase().includes('cell')));
    
    if (isPhoneField) {
      // For phone fields, use phone number normalization regardless of path format
      let phoneValue;
      if (key.includes('.')) {
        // It's a nested path, use the getPhoneNumberFromField function
        phoneValue = getPhoneNumberFromField(data, key);
      } else {
        // It's a simple field, just normalize the value
        phoneValue = normalizePhoneNumber(value || data[key]);
      }
      Logger.log(`FIELD_DEBUG: Phone field handling returned: "${phoneValue}"`);
      return phoneValue;
    }
    
    // Check if this is an email field
    const isEmailField = 
      key === 'email' || 
      key.startsWith('email.') ||
      (typeof key === 'string' && key.toLowerCase().includes('email'));
    
    if (isEmailField) {
      // Special handling for email fields - always lowercase and trim
      let emailValue;
      if (key.includes('.')) {
        // Handle nested email paths
        if (key.startsWith('email.')) {
          // Try to extract using standard email handling
          const parts = key.split('.');
          const label = parts[1];
          
          if (Array.isArray(data.email)) {
            // Try to find email with matching label
            const match = data.email.find(e => 
              e && e.label && e.label.toLowerCase() === label.toLowerCase()
            );
            
            if (match && match.value) {
              emailValue = match.value;
            } else if (label === 'primary' || label === 'main') {
              // Try to find primary email
              const primary = data.email.find(e => e && e.primary);
              if (primary && primary.value) {
                emailValue = primary.value;
              }
            }
            
            // Fallback to first email if available
            if (!emailValue && data.email.length > 0) {
              if (data.email[0].value) {
                emailValue = data.email[0].value;
              } else if (typeof data.email[0] === 'string') {
                emailValue = data.email[0];
              }
            }
          }
        } else {
          // Try to extract using general path lookup
          emailValue = value;
        }
      } else {
        // Simple email field
        emailValue = value;
        
        // Handle array of email objects
        if (Array.isArray(emailValue) && emailValue.length > 0) {
          if (emailValue[0].value) {
            // Find primary or use first
            const primary = emailValue.find(e => e && e.primary);
            emailValue = primary ? primary.value : emailValue[0].value;
          } else if (typeof emailValue[0] === 'string') {
            emailValue = emailValue[0];
          }
        } else if (emailValue && typeof emailValue === 'object' && emailValue.value) {
          emailValue = emailValue.value;
        }
      }
      
      // Ensure we have a string
      emailValue = String(emailValue || '');
      
      // Enhanced email normalization:
      // 1. Lowercase
      // 2. Trim whitespace
      // 3. Handle common email domain typos
      let normalizedEmail = emailValue.toLowerCase().trim();
      
      // Extract the domain part for better comparison
      if (normalizedEmail.includes('@')) {
        const parts = normalizedEmail.split('@');
        const username = parts[0];
        let domain = parts[1];
        
        // Fix common domain typos
        if (domain === 'gmail.comm') domain = 'gmail.com';
        if (domain === 'gmail.con') domain = 'gmail.com';
        if (domain === 'gmial.com') domain = 'gmail.com';
        if (domain === 'hotmail.comm') domain = 'hotmail.com';
        if (domain === 'hotmail.con') domain = 'hotmail.com';
        if (domain === 'yahoo.comm') domain = 'yahoo.com';
        if (domain === 'yahoo.con') domain = 'yahoo.com';
        
        // Reassemble the normalized email
        normalizedEmail = username + '@' + domain;
      }
      
      Logger.log(`FIELD_DEBUG: Email field enhanced normalization from "${emailValue}" to "${normalizedEmail}"`);
      return normalizedEmail;
    }
    
    // Add special handling for name fields with common typos
    const isNameField = 
      key === 'name' || 
      key.toLowerCase().includes('name') || 
      key.toLowerCase() === 'person' || 
      key.toLowerCase() === 'contact';
    
    if (isNameField) {
      Logger.log(`FIELD_DEBUG: Handling name field normalization for key "${key}"`);
      let nameValue = getValueByPath(data, key);
      
      // Ensure we have a string
      nameValue = nameValue !== null && nameValue !== undefined ? String(nameValue) : '';
      
      // If it's an object with a name property, extract it
      if (typeof nameValue === 'object' && nameValue !== null && nameValue.name) {
        nameValue = nameValue.name;
      }
      
      // Trim the name
      let normalizedName = nameValue.trim();
      
      // Handle common name typos like extra character at end
      if (normalizedName.length > 3) {
        // Check for common pattern where a name might have an extra character at the end
        // For example "Simpson" vs "Simpsonm" or "John" vs "Johnn"
        const lastChar = normalizedName.charAt(normalizedName.length - 1);
        const secondLastChar = normalizedName.charAt(normalizedName.length - 2);
        
        // If the last character is the same as the second-to-last character,
        // it might be a typo (double letter at end)
        if (lastChar === secondLastChar) {
          normalizedName = normalizedName.substring(0, normalizedName.length - 1);
          Logger.log(`FIELD_DEBUG: Fixed double character typo from "${nameValue}" to "${normalizedName}"`);
        }
        // If the last character is 'm' and it's a common name ending with "son",
        // it might be a typo like "Simpson" vs "Simpsonm"
        else if (lastChar === 'm' && normalizedName.toLowerCase().includes('son')) {
          normalizedName = normalizedName.substring(0, normalizedName.length - 1);
          Logger.log(`FIELD_DEBUG: Fixed common 'm' typo from "${nameValue}" to "${normalizedName}"`);
        }
        // Check for other common typos based on patterns observed in your data
        else if ((lastChar === 'n' || lastChar === 'm') && 
                 normalizedName.toLowerCase().endsWith('mann') || 
                 normalizedName.toLowerCase().endsWith('manm') ||
                 normalizedName.toLowerCase().endsWith('sonn') ||
                 normalizedName.toLowerCase().endsWith('sonm')) {
          // Normalize common surname endings
          normalizedName = normalizedName.substring(0, normalizedName.length - 1);
          Logger.log(`FIELD_DEBUG: Fixed common name ending typo from "${nameValue}" to "${normalizedName}"`);
        }
      }
      
      Logger.log(`FIELD_DEBUG: Name field normalized from "${nameValue}" to "${normalizedName}"`);
      return normalizedName;
    }
    
    // CASE 3: CUSTOM FIELDS
    if (key.startsWith('custom_fields.')) {
      const parts = key.split('.');
      const fieldKey = parts[1];
      Logger.log(`FIELD_DEBUG: Handling custom field: "${fieldKey}"`);
      
      if (data.custom_fields && data.custom_fields[fieldKey] !== undefined) {
        const customValue = data.custom_fields[fieldKey];
        Logger.log(`FIELD_DEBUG: Custom field value type: ${typeof customValue}, ${Array.isArray(customValue) ? 'array' : 'not array'}`);
        
        // Object with value and currency (money field)
        if (customValue && typeof customValue === 'object' && customValue.value !== undefined) {
          if (customValue.currency !== undefined) {
            // Money field
            const result = `${customValue.value}_${customValue.currency}`;
            Logger.log(`FIELD_DEBUG: Currency field, normalized to: "${result}"`);
            return result;
          } else if (customValue.formatted_address !== undefined) {
            // Address field
            Logger.log(`FIELD_DEBUG: Address field, normalized to: "${customValue.formatted_address}"`);
            return customValue.formatted_address;
          } else {
            // Other field with value property
            const result = normalizeValueByType(customValue.value);
            Logger.log(`FIELD_DEBUG: Generic field with value property, normalized to: "${result}"`);
            return result;
          }
        }
        
        // Multi-select fields (array of option IDs)
        if (Array.isArray(customValue)) {
          const result = customValue.join(',');
          Logger.log(`FIELD_DEBUG: Array custom field, normalized to: "${result}"`);
          return result;
        }
        
        // Regular value
        const result = normalizeValueByType(customValue);
        Logger.log(`FIELD_DEBUG: Basic custom field, normalized to: "${result}"`);
        return result;
      }
    }
    
    // CASE 4: PERSON, ORG, DEAL references
    const entityFields = ['person_id', 'org_id', 'deal_id', 'organization', 'person', 'deal', 'creator', 'owner'];
    if (entityFields.includes(key) || 
        key.endsWith('_name') || 
        key.endsWith('_id') ||
        (value && typeof value === 'object' && value.name !== undefined)) {
      
      Logger.log(`FIELD_DEBUG: Handling entity reference field`);
      
      // If it's a direct entity reference with name
      if (value && typeof value === 'object') {
        if (value.name) {
          Logger.log(`FIELD_DEBUG: Entity with name, using: "${value.name}"`);
          return value.name;
        } else if (value.id) {
          Logger.log(`FIELD_DEBUG: Entity with ID only, using: "${value.id.toString()}"`);
          return value.id.toString();
        }
      }
    }
    
    // CASE 5: General nested hierarchies not covered in specific cases above
    if (key.includes('.') && !key.startsWith('phone.') && !key.startsWith('email.') && !key.startsWith('custom_fields.')) {
      Logger.log(`FIELD_DEBUG: Handling general nested field with key: "${key}"`);
      
      // Try to get the value using the specialized method first
      try {
        if (value !== undefined) {
          const result = normalizeValueByType(value);
          Logger.log(`FIELD_DEBUG: General nested field normalized to: "${result}"`);
          return result;
        }
      } catch (e) {
        Logger.log(`FIELD_DEBUG: Error normalizing nested field: ${e.message}`);
      }
      
      // Try a different approach - manual traversal
      try {
        const parts = key.split('.');
        let current = data;
        
        for (let i = 0; i < parts.length; i++) {
          if (current === undefined || current === null) {
            Logger.log(`FIELD_DEBUG: Path traversal failed at part ${i}`);
            break;
          }
          
          // Handle array indices in the path
          if (!isNaN(parseInt(parts[i])) && Array.isArray(current)) {
            current = current[parseInt(parts[i])];
          } else {
            current = current[parts[i]];
          }
          
          Logger.log(`FIELD_DEBUG: Traversed to path part ${i} (${parts[i]}), got ${typeof current}`);
        }
        
        if (current !== undefined) {
          const result = normalizeValueByType(current);
          Logger.log(`FIELD_DEBUG: Manual traversal result: "${result}"`);
          return result;
        }
      } catch (e) {
        Logger.log(`FIELD_DEBUG: Error in manual path traversal: ${e.message}`);
      }
    }
    
    // For all other cases, normalize the value we got from getValueByPath
    const result = normalizeValueByType(value);
    Logger.log(`FIELD_DEBUG: Using standard normalization, result: "${result}"`);
    return result;
  } catch (e) {
    Logger.log(`FIELD_DEBUG: Error in getNormalizedFieldValue for ${key}: ${e.message}`);
    return '';
  }
}

/**
 * Normalizes a value based on its type for consistent comparison
 * @param {*} value - The value to normalize
 * @return {string} Normalized value
 */
function normalizeValueByType(value) {
  // Handle null/undefined
  if (value === null || value === undefined) {
    return '';
  }
  
  // Handle arrays
  if (Array.isArray(value)) {
    // If array of objects with value property
    if (value.length > 0 && typeof value[0] === 'object' && value[0] !== null) {
      if (value[0].value !== undefined) {
        // Array of objects with value properties
        return value.map(item => normalizeValueByType(item.value)).join(',');
      } else if (value[0].name !== undefined) {
        // Array of objects with names
        return value.map(item => item.name).join(',');
      }
    }
    // Simple array - join with commas
    return value.join(',');
  }
  
  // Handle objects
  if (typeof value === 'object') {
    // Object with value property
    if (value.value !== undefined) {
      return normalizeValueByType(value.value);
    }
    // Object with name property (common for entity references)
    if (value.name !== undefined) {
      return value.name;
    }
    // Object with id property
    if (value.id !== undefined) {
      return value.id.toString();
    }
    // Last resort - stringify
    return JSON.stringify(value);
  }
  
  // Handle dates
  if (value instanceof Date) {
    return value.toISOString();
  }
  
  // Handle booleans
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false';
  }
  
  // Handle numbers and scientific notation
  if (typeof value === 'number' || 
      (typeof value === 'string' && value.includes('E'))) {
    try {
      // Try to parse as number to handle scientific notation
      const numValue = Number(value);
      if (!isNaN(numValue)) {
        // For integers, use fixed notation with no decimals
        if (Number.isInteger(numValue)) {
          return numValue.toFixed(0);
        }
        // Preserve exact decimal places for better comparisons
        if (typeof value === 'string' && value.includes('.')) {
          // Get the decimal places from the original string representation
          const parts = value.split('.');
          if (parts.length > 1) {
            const decimalPlaces = parts[1].replace(/E.*$/, '').length;
            return numValue.toFixed(decimalPlaces);
          }
        }
        // For decimal numbers, convert to string directly ensuring consistent formatting
        return numValue.toString();
      }
    } catch (e) {
      // If parsing fails, continue with string handling
      Logger.log(`Error handling numeric value: ${e.message}`);
    }
  }
  
  // For ordinary strings
  if (typeof value === 'string') {
    // Check if it looks like a phone number
    if (/^[\d\+\-\(\)\s\.]+$/.test(value)) {
      // It only contains digits, plus signs, parentheses, spaces, etc.
      return normalizePhoneNumber(value);
    }
    
    // Check if it looks like an email
    if (value.includes('@') && value.includes('.')) {
      return value.toLowerCase().trim();
    }
    
    // Regular string
    return value;
  }
  
  // Numbers and other values
  return value.toString();
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
      Logger.log('No original values stored to compare against');
      return false;
    }
    
    // Get current values for the entire row
    const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    
    // Get the first value (ID) to use for retrieving the original data
    const id = rowValues[0];
    
    // Create a data object that mimics Pipedrive structure for nested field handling
    const dataObj = { id: id };
    
    // Populate the data object with values from the row
    headers.forEach((header, index) => {
      if (index < rowValues.length) {
        // Use dot notation to create nested objects
        if (header.includes('.')) {
          const parts = header.split('.');
          
          // Common nested structures to handle specially
          if (['phone', 'email'].includes(parts[0])) {
            // Handle phone/email specially
            if (!dataObj[parts[0]]) {
              dataObj[parts[0]] = [];
            }
            
            // If it's a label-based path (e.g., phone.mobile)
            if (parts.length === 2 && isNaN(parseInt(parts[1]))) {
              dataObj[parts[0]].push({
                label: parts[1],
                value: rowValues[index]
              });
            } 
            // If it's an array index path (e.g., phone.0.value)
            else if (parts.length === 3 && parts[2] === 'value') {
              const idx = parseInt(parts[1]);
              while (dataObj[parts[0]].length <= idx) {
                dataObj[parts[0]].push({});
              }
              dataObj[parts[0]][idx].value = rowValues[index];
            }
          }
          // Custom fields
          else if (parts[0] === 'custom_fields') {
            if (!dataObj.custom_fields) {
              dataObj.custom_fields = {};
            }
            
            if (parts.length === 2) {
              // Simple custom field
              dataObj.custom_fields[parts[1]] = rowValues[index];
            } 
            else if (parts.length > 2) {
              // Nested custom field
              if (!dataObj.custom_fields[parts[1]]) {
                dataObj.custom_fields[parts[1]] = {};
              }
              // Handle complex types like address
              if (parts[2] === 'formatted_address') {
                dataObj.custom_fields[parts[1]].formatted_address = rowValues[index];
              } 
              else if (parts[2] === 'currency') {
                dataObj.custom_fields[parts[1]].currency = rowValues[index];
              }
              else {
                dataObj.custom_fields[parts[1]][parts[2]] = rowValues[index];
              }
            }
          } else {
            // Other nested paths - build the structure
            let current = dataObj;
            for (let i = 0; i < parts.length - 1; i++) {
              if (current[parts[i]] === undefined) {
                // If part is numeric, create an array
                if (!isNaN(parseInt(parts[i+1]))) {
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
    Logger.log(`Checking ${Object.keys(originalValues).length} fields for original value match`);
    
    // Check each column that has a stored original value
    for (const headerName in originalValues) {
      // Find the column index for this header
      const colIndex = headers.indexOf(headerName);
      if (colIndex === -1) {
        Logger.log(`Header ${headerName} not found in current headers`);
        continue; // Header not found
      }
      
      const originalValue = originalValues[headerName];
      const currentValue = rowValues[colIndex];
      
      // Special handling for null/empty values
      if ((originalValue === null || originalValue === "") && 
          (currentValue === null || currentValue === "")) {
        Logger.log(`Both values are empty for ${headerName}, treating as match`);
        continue; // Both empty, consider a match
      }
      
      // NEW: Check if this is a structural field with complex nested structure
      if (originalValue && typeof originalValue === 'object' && originalValue.__isStructural) {
        Logger.log(`Found structural field ${headerName} with key ${originalValue.__key}`);
        
        // Use the pre-computed normalized value for comparison
        const normalizedOriginal = originalValue.__normalized || '';
        const normalizedCurrent = getNormalizedFieldValue(dataObj, originalValue.__key);
        
        Logger.log(`Structural field comparison for ${headerName}: Original="${normalizedOriginal}", Current="${normalizedCurrent}"`);
        
        // If the normalized values don't match, return false
        if (normalizedOriginal !== normalizedCurrent) {
          Logger.log(`Structural field mismatch found for ${headerName}`);
          return false;
        }
        
        // Skip to the next field
        continue;
      }
      
      // Use the generalized field value normalization for regular fields
      const normalizedOriginal = getNormalizedFieldValue({ [headerName]: originalValue }, headerName);
      const normalizedCurrent = getNormalizedFieldValue(dataObj, headerName);
      
      Logger.log(`Field comparison for ${headerName}: Original="${normalizedOriginal}", Current="${normalizedCurrent}"`);
      
      // If the normalized values don't match, return false
      if (normalizedOriginal !== normalizedCurrent) {
        Logger.log(`Mismatch found for ${headerName}`);
        return false;
      }
    }
    
    // If we reach here, all values with stored originals match
    Logger.log('All values match original values');
    return true;
  } catch (error) {
    Logger.log(`Error in checkAllValuesMatchOriginal: ${error.message}`);
    return false;
  }
}