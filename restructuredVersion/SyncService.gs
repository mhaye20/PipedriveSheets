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
    
    // IMPORTANT: If two-way sync is enabled, make sure we thoroughly clean up previous Sync Status columns
    if (twoWaySyncEnabled) {
      try {
        // Get the current Sync Status column position
        const currentStatusColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey);
        
        // Get the sheet object
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        
        if (sheet && currentStatusColumn) {
          Logger.log(`Running comprehensive cleanup of previous Sync Status columns. Current column: ${currentStatusColumn}`);
          cleanupPreviousSyncStatusColumn(sheet, currentStatusColumn);
          
          // Also perform a complete cell-by-cell scan for any sync status validation that might have been missed
          scanAllCellsForSyncStatusValidation(sheet);
        }
      } catch (cleanupError) {
        Logger.log(`Error during final Sync Status column cleanup: ${cleanupError.message}`);
        // Continue with sync even if cleanup has issues
      }
    }
    
    // Update sync status to completed
    updateSyncStatus('3', 'completed', 'Data successfully written to spreadsheet', 100);
    
    // Store sync timestamp
    const timestamp = new Date().toISOString();
    scriptProperties.setProperty(`LAST_SYNC_${sheetName}`, timestamp);
    
    Logger.log(`Successfully synced ${items.length} items from Pipedrive to sheet "${sheetName}"`);
    
    // Remove the duplicate update sync status call
    // updateSyncStatus('3', 'completed', `Successfully synced ${items.length} ${entityType} from Pipedrive`, 100);
    setSyncRunning(false);
    
    // Check if we need to recreate triggers after column changes
    checkAndRecreateTriggers();
    
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
    const previousColumnPosition = scriptProperties.getProperty(previousSyncColumnKey);
    
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
      Logger.log(`Cleared sheet ${sheetName} to ensure clean sync`);
    }
    
    // CRITICAL: Debug what headers we're getting from Pipedrive
    Logger.log(`Incoming Pipedrive headers (${options.headerRow ? options.headerRow.length : 0}): ${JSON.stringify(options.headerRow)}`);
    
    // Get headers from options - ALWAYS make a copy to avoid modifying the original
    const headers = options.headerRow ? [...options.headerRow] : [];
    Logger.log(`Working with headers (${headers.length}): ${JSON.stringify(headers)}`);
    
    // Check if two-way sync is enabled
    const twoWaySyncEnabled = options.twoWaySyncEnabled || false;
    
    // Key for tracking column
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    
    // Get current tracking column
    const currentSyncColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey);
    
    // Store previous position if we have a current one
    if (twoWaySyncEnabled && currentSyncColumn) {
      scriptProperties.setProperty(previousSyncColumnKey, currentSyncColumn);
      Logger.log(`Stored previous sync column position: ${currentSyncColumn}`);
    }
    
    // For preservation of status data - key is entity ID, value is status
    let statusByIdMap = new Map();
    let statusColumnIndex = -1;
    
    // FIRST: Write ALL Pipedrive headers to the sheet (before adding Sync Status)
    // This ensures ALL column headers from Pipedrive are properly set
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // NOW: If two-way sync is enabled, handle Sync Status column
    if (twoWaySyncEnabled) {
      // The Sync Status will be added AFTER all Pipedrive columns
      const statusHeader = 'Sync Status';
      const lastColumn = headers.length;
      
      // Add Sync Status column at the end
      statusColumnIndex = lastColumn;
      sheet.getRange(1, lastColumn + 1).setValue(statusHeader);
      sheet.getRange(1, lastColumn + 1).setFontWeight('bold');
      
      Logger.log(`Added Sync Status column at position: ${lastColumn + 1}`);
      
      // Store status column position for tracking
      const statusColumnLetter = columnToLetter(lastColumn + 1);
      scriptProperties.setProperty(twoWaySyncTrackingColumnKey, statusColumnLetter);
      Logger.log(`Set Sync Status column tracking to position ${statusColumnLetter}`);
      
      // Clean up any previous Sync Status columns
      if (previousColumnPosition && previousColumnPosition !== statusColumnLetter) {
        cleanupPreviousSyncStatusColumn(sheet, statusColumnLetter);
      }
      
      // Get current status values to preserve them
      try {
        // Get stored status data for this sheet
        const statusDataKey = `SYNC_STATUS_DATA_${sheetName}`;
        const storedStatusData = scriptProperties.getProperty(statusDataKey);
        
        if (storedStatusData) {
          try {
            const parsedStatus = JSON.parse(storedStatusData);
            
            // Populate the statusByIdMap from stored data
            if (parsedStatus && typeof parsedStatus === 'object') {
              Object.keys(parsedStatus).forEach(id => {
                statusByIdMap.set(id, parsedStatus[id]);
              });
              
              Logger.log(`Loaded ${statusByIdMap.size} stored status values`);
            }
          } catch (parseError) {
            Logger.log(`Error parsing stored status data: ${parseError.message}`);
          }
        }
      } catch (e) {
        Logger.log(`Error loading stored status values: ${e.message}`);
      }
    }
    
    // Process items and write data rows
    const dataRows = items.map(item => {
      // Create a row with empty values for all columns plus status column if needed
      const totalColumns = twoWaySyncEnabled ? headers.length + 1 : headers.length;
      const row = Array(totalColumns).fill('');
      
      // For each Pipedrive column, extract and format the value
      options.columns.forEach((column, index) => {
        // Skip if index is beyond our header count
        if (index >= headers.length) {
          return;
        }
        
        const columnKey = typeof column === 'object' ? column.key : column;
        const value = getValueByPath(item, columnKey);
        row[index] = formatValue(value, columnKey, options.optionMappings);
      });
      
      // If two-way sync is enabled, add the status value in the last column
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
    
    // Calculate the total number of columns we need to write
    const totalColumns = twoWaySyncEnabled ? headers.length + 1 : headers.length;
    
    // Write all data rows at once
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, totalColumns).setValues(dataRows);
    }
    
    // If two-way sync is enabled, set up data validation and formatting for the status column
    if (twoWaySyncEnabled && statusColumnIndex !== -1) {
      try {
        // Status column is always going to be at the end
        const statusColumnPosition = headers.length + 1;
        
        // Style the header cell
        const headerCell = sheet.getRange(1, statusColumnPosition);
        headerCell
          .setValue('Sync Status')
          .setBackground('#E8F0FE')
          .setFontWeight('bold')
          .setNote('This column tracks changes for two-way sync with Pipedrive');
        
        // Add data validation for all cells in the column
        const lastRow = Math.max(sheet.getLastRow(), 2);
        if (lastRow > 1) {
          const dataRange = sheet.getRange(2, statusColumnPosition, lastRow - 1, 1);
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
            .build();
          dataRange.setDataValidation(rule);
          
          // Set default values for empty cells - force all cells to have default values
          const values = dataRange.getValues();
          const newValues = values.map(row => {
            const currentValue = row[0];
            if (!currentValue || 
                (currentValue !== 'Modified' && 
                 currentValue !== 'Not modified' && 
                 currentValue !== 'Synced' && 
                 currentValue !== 'Error')) {
              return ['Not modified'];
            }
            return [currentValue];
          });
          dataRange.setValues(newValues);
          
          // Style the status column
          dataRange
            .setBackground('#F8F9FA')
            .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
          
          // Set up conditional formatting for the status column
          const rules = sheet.getConditionalFormatRules();
          
          // Rule for "Modified" status - red background
          const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Modified')
            .setBackground('#FCE8E6')
            .setFontColor('#D93025')
            .setRanges([dataRange])
            .build();
          
          // Rule for "Synced" status - green background
          const syncedRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Synced')
            .setBackground('#E6F4EA')
            .setFontColor('#137333')
            .setRanges([dataRange])
            .build();
          
          // Rule for "Error" status - orange background
          const errorRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Error')
            .setBackground('#FEF7E0')
            .setFontColor('#EA8600')
            .setRanges([dataRange])
            .build();
          
          // Apply all rules
          rules.push(modifiedRule);
          rules.push(syncedRule);
          rules.push(errorRule);
          sheet.setConditionalFormatRules(rules);
          
          // Save current status values for future reference
          try {
            const statusData = {};
            
            // Save only rows with data in case of ID changes
            for (let i = 0; i < dataRows.length; i++) {
              const id = dataRows[i][0].toString();
              const status = dataRows[i][statusColumnIndex];
              
              if (id && status) {
                statusData[id] = status;
              }
            }
            
            // Store the status data in script properties
            const statusDataKey = `SYNC_STATUS_DATA_${sheetName}`;
            scriptProperties.setProperty(statusDataKey, JSON.stringify(statusData));
            Logger.log(`Saved ${Object.keys(statusData).length} status values for future reference`);
          } catch (storageError) {
            Logger.log(`Error saving status data: ${storageError.message}`);
          }
        }
      } catch (e) {
        Logger.log(`Error setting up status column: ${e.message}`);
      }
    }
    
    // Final cleanup: do a thorough scan for any remaining old Sync Status columns
    if (twoWaySyncEnabled) {
      // Get the current status column letter for the final cleanup
      const finalStatusColumnLetter = columnToLetter(headers.length + 1);
      Logger.log(`Running final cleanup scan. Current Sync Status column: ${finalStatusColumnLetter}`);
      
      // This will find and clean any other columns with Sync Status headers
      cleanupPreviousSyncStatusColumn(sheet, finalStatusColumnLetter);
    }
    
    // Auto-resize columns to fit content
    sheet.autoResizeColumns(1, totalColumns);
    
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
 * @param {boolean} isCurrentColumn - Whether this is the current active Sync Status column
 */
function cleanupColumnFormatting(sheet, columnLetter, isCurrentColumn = false) {
  try {
    Logger.log(`Cleaning up column ${columnLetter} (isCurrentColumn: ${isCurrentColumn})`);
    
    // Convert column letter to index
    const columnIndex = letterToColumn(columnLetter);
    
    // Get the data range to know the last row
    const lastRow = Math.max(sheet.getLastRow(), 2);
    
    // Clean the header first - only if it's not the current column
    if (!isCurrentColumn) {
          const headerCell = sheet.getRange(1, columnIndex);
          const headerValue = headerCell.getValue();
      const note = headerCell.getNote();
      
      Logger.log(`Checking header in column ${columnLetter}: "${headerValue}"`);
      
      // Check if this column has 'Sync Status' header or sync-related note
      if (headerValue === "Sync Status" || 
          headerValue === "Sync Status (hidden)" || 
          headerValue === "Status" ||
          (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive')))) {
        
        Logger.log(`Clearing Sync Status header in column ${columnLetter}`);
        headerCell.setValue('');
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
        const values = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
        const newValues = [];
        let cleanedCount = 0;
        
        // Clear only cells containing specific sync status values
        for (let i = 0; i < values.length; i++) {
          const value = values[i][0];
          
          // Check if the value is a known sync status value
          if (value === "Modified" || 
              value === "Not modified" || 
              value === "Synced" || 
              value === "Error") {
            newValues.push(['']); // Clear known status values
            cleanedCount++;
          } else {
            newValues.push([value]); // Keep other values
          }
        }
        
        // Set the cleaned values back to the sheet
        sheet.getRange(2, columnIndex + 1, values.length, 1).setValues(newValues);
        Logger.log(`Cleared ${cleanedCount} sync status values in column ${columnLetter}`);
        
        // Remove conditional formatting for this column
        removeConditionalFormattingForColumn(sheet, columnIndex);
      }
    }
    
    Logger.log(`Cleanup of column ${columnLetter} complete`);
  } catch (error) {
    Logger.log(`Error in cleanupColumnFormatting: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

// Function to identify and clean up previous Sync Status columns
function cleanupPreviousSyncStatusColumn(sheet, currentSyncColumn) {
  try {
    Logger.log(`Looking for previous Sync Status columns to clean up (current: ${currentSyncColumn})`);
    
    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetName = sheet.getName();
    const previousSyncColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const previousSyncColumn = scriptProperties.getProperty(previousSyncColumnKey);
    
    // IMPORTANT: We do NOT want to override column headers when Pipedrive adds new fields
    // If the previous column is converted to a Pipedrive data column, we need to keep it
    
    // Clean up the known previous column first if it exists and is different from current
    if (previousSyncColumn && previousSyncColumn !== currentSyncColumn) {
      Logger.log(`Previous Sync Status column found at ${previousSyncColumn}`);
      
      try {
        // Convert letter to column index
        const previousColumnIndex = letterToColumn(previousSyncColumn);
        
        // Clean up all Sync Status-specific formatting and validation from this column
        // but do NOT clear the header cell - let the main sync function handle headers
        
        // First, clear any sync-specific formatting and validation in the data cells
        if (sheet.getLastRow() > 1) {
          const dataRange = sheet.getRange(2, previousColumnIndex, Math.max(sheet.getLastRow() - 1, 1), 1);
          
          // Clear all formatting and validation from data cells
          dataRange.clearFormat();
          dataRange.clearDataValidations();
          
          // Check for and clear status-specific values ONLY
          const values = dataRange.getValues();
          const newValues = values.map(row => {
            const value = String(row[0]).trim();
            
            // Only clear if it's one of the specific status values
            if (value === "Modified" || 
                value === "Not modified" || 
                value === "Synced" || 
                value === "Error") {
              return [""];
            }
            return [value]; // Keep any other values
          });
          
          // Write the cleaned values back
          dataRange.setValues(newValues);
          Logger.log(`Cleaned status values from previous column ${previousSyncColumn}`);
        }
        
        // Remove any sync-specific formatting or notes from the header
        // but KEEP the header cell itself for Pipedrive data
        const headerCell = sheet.getRange(1, previousColumnIndex);
        headerCell.clearFormat();
        headerCell.clearNote();
        // Do NOT call setValue() - let the main sync function set the header
        
        Logger.log(`Cleaned formatting from previous Sync Status column ${previousSyncColumn}`);
      } catch (e) {
        Logger.log(`Error cleaning previous column ${previousSyncColumn}: ${e.message}`);
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
          (headerNote && (headerNote.includes('sync') || headerNote.includes('track') || headerNote.includes('Pipedrive')));
        
        // Also check for sync status values in the data cells
        let hasSyncStatusValues = false;
        if (sheet.getLastRow() > 1) {
          // Sample a few cells to check for status values
          const sampleSize = Math.min(10, sheet.getLastRow() - 1);
          const sampleRange = sheet.getRange(2, i, sampleSize, 1);
          const sampleValues = sampleRange.getValues();
          
          hasSyncStatusValues = sampleValues.some(row => {
            const value = String(row[0]).trim();
            return value === "Modified" || 
                   value === "Not modified" || 
                   value === "Synced" || 
                   value === "Error";
          });
        }
        
        // If this column has sync status indicators, clean it
        if (isSyncStatusHeader || hasSyncStatusValues) {
          Logger.log(`Found additional Sync Status column at ${colLetter}, cleaning up...`);
          
          // Clean any sync-specific formatting and validation but preserve the header cell
          if (sheet.getLastRow() > 1) {
            const dataRange = sheet.getRange(2, i, Math.max(sheet.getLastRow() - 1, 1), 1);
            
            // Clear all formatting and validation
            dataRange.clearFormat();
            dataRange.clearDataValidations();
            
            // Only clear specific status values
            const values = dataRange.getValues();
            const newValues = values.map(row => {
              const value = String(row[0]).trim();
              if (value === "Modified" || 
                  value === "Not modified" || 
                  value === "Synced" || 
                  value === "Error") {
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
          
          Logger.log(`Cleaned sync status formatting from column ${colLetter}`);
        }
      } catch (e) {
        Logger.log(`Error checking column ${colLetter}: ${e.message}`);
      }
    }
    
    // Clear the previous column tracking since we've cleaned it up
    scriptProperties.deleteProperty(previousSyncColumnKey);
    Logger.log(`Cleanup of previous Sync Status columns complete`);
  } catch (error) {
    Logger.log(`Error in cleanupPreviousSyncStatusColumn: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
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
    Logger.log(`Removed conditional formatting rules for column ${columnToLetter(columnIndex + 1)}`);
  } catch (error) {
    Logger.log(`Error removing conditional formatting: ${error.message}`);
  }
}

// Utility function to convert a column letter to a column index (1-based)
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;
  
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  
  return column;
}

// Utility function to convert a column index to a letter (A, B, C, ..., Z, AA, AB, ...)
function columnToLetter(column) {
  let temp;
  let letter = '';
  
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  
  return letter;
}