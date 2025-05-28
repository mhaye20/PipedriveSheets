/**
 * TwoWaySyncSettingsUI
 * 
 * This module handles the UI components and functionality for two-way sync settings:
 * - Displaying the settings dialog
 * - Managing sync settings
 * - Handling settings persistence
 */

var TwoWaySyncSettingsUI = TwoWaySyncSettingsUI || {};

/**
 * Reference to sync trigger functions for two-way sync
 * These are defined in SyncService.gs
 */
// Access trigger functions when needed rather than storing as variables
// This ensures we always get the most up-to-date function implementation

/**
 * Shows the two-way sync settings dialog
 */
TwoWaySyncSettingsUI.showTwoWaySyncSettings = function() {
  try {
    // Check if user has access to two-way sync feature
    if (!PaymentService.hasFeatureAccess('two_way_sync')) {
      SpreadsheetApp.getUi().alert(
        'Two-way sync is only available on Pro and Team plans. Please upgrade to enable this feature.'
      );
      PaymentService.showUpgradeDialog();
      return;
    }
    
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get current two-way sync settings from properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

    // Get the current settings
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    const trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const lastSync = scriptProperties.getProperty(twoWaySyncLastSyncKey) || 'Never';

    // Get sheet-specific entity type
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    // Check if user is in read-only mode (team member, not admin)
    let isReadOnly = false;
    try {
      if (typeof TeamAccess !== 'undefined' && typeof TeamAccess.getUserRole === 'function') {
        const userRole = TeamAccess.getUserRole();
        isReadOnly = (userRole === 'member');
      }
    } catch (e) {
      Logger.log(`Error checking user role: ${e.message}`);
      isReadOnly = false;
    }
    
    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
    
    // Pass data to template
    template.data = {
      sheetName: activeSheetName,
      twoWaySyncEnabled: twoWaySyncEnabled,
      trackingColumn: trackingColumn,
      lastSync: lastSync,
      entityType: entityType
    };
    
    // Pass read-only status to template
    template.isReadOnly = isReadOnly;
    
    // Make include function available to the template
    template.include = function(filename) {
      if (filename === 'TwoWaySyncSettingsUI_Styles') {
        return TwoWaySyncSettingsUI.getStyles();
      } else if (filename === 'TwoWaySyncSettingsUI_Scripts') {
        return TwoWaySyncSettingsUI.getScripts();
      }
      return '';
    };
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(600)
      .setHeight(700)
      .setTitle('Two-Way Sync Settings');
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
  } catch (error) {
    Logger.log(`Error showing two-way sync settings: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to show settings: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

/**
 * Handles two-way sync settings when column preferences are saved
 * @param {string} sheetName - The name of the sheet
 */
TwoWaySyncSettingsUI.handleColumnPreferencesChange = function(sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // When columns are changed and two-way sync is enabled, handle tracking column
    if (twoWaySyncEnabled) {
      Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Adjusting sync column.`);
      
      // When columns are changed, delete the tracking column property to force repositioning
      const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
      scriptProperties.deleteProperty(twoWaySyncTrackingColumnKey);
      
      // Add a flag to indicate that the Sync Status column should be repositioned at the end
      const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${sheetName}`;
      scriptProperties.setProperty(twoWaySyncColumnAtEndKey, 'true');
      
      // Add a flag to ensure the onEdit trigger is recreated after column reorganization
      const twoWaySyncRecreateTriggersKey = `TWOWAY_SYNC_RECREATE_TRIGGERS_${sheetName}`;
      scriptProperties.setProperty(twoWaySyncRecreateTriggersKey, 'true');
      
      Logger.log(`Removed tracking column property for sheet "${sheetName}" to ensure correct positioning on next sync.`);
    }
  } catch (e) {
    Logger.log(`Error in handleColumnPreferencesChange: ${e.message}`);
    throw e;
  }
};

/**
 * Saves the two-way sync settings for a sheet and sets up the tracking column
 * @param {boolean} enableTwoWaySync Whether to enable two-way sync
 * @param {string} trackingColumn The column letter to use for tracking changes
 */
function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    // Check if user is in read-only mode (team member, not admin)
    let isReadOnly = false;
    try {
      if (typeof TeamAccess !== 'undefined' && typeof TeamAccess.getUserRole === 'function') {
        const userRole = TeamAccess.getUserRole();
        isReadOnly = (userRole === 'member');
      }
    } catch (e) {
      Logger.log(`Error checking user role: ${e.message}`);
      isReadOnly = false;
    }
    
    if (isReadOnly) {
      throw new Error('You do not have permission to modify these settings. Only team admins can change two-way sync settings.');
    }
    
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

    // Get the previous tracking column and position
    const previousTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const previousPosStr = scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${activeSheetName}`) || '-1';
    const previousPos = parseInt(previousPosStr, 10);
    const currentPos = trackingColumn ? columnLetterToIndex(trackingColumn) : -1;

    // If the position has changed, store the previous column for cleanup
    if (previousTrackingColumn && previousTrackingColumn !== trackingColumn) {
      scriptProperties.setProperty(`PREVIOUS_TRACKING_COLUMN_${activeSheetName}`, previousTrackingColumn);

      // NEW: Also track when columns have been removed (causing a left shift)
      if (previousPos >= 0 && currentPos >= 0 && currentPos < previousPos) {
        Logger.log(`Detected column removal: Sync Status moved left from ${previousPos} to ${currentPos}`);

        // Check all columns between previous and current positions (inclusive)
        // Important: Don't just check columns in between, check ALL columns from 0 to max(previousPos)
        const maxPos = Math.max(previousPos + 3, activeSheet.getLastColumn()); // Add buffer
        for (let i = 0; i <= maxPos; i++) {
          const colLetter = columnToLetter(i);
          if (colLetter !== trackingColumn) {
            // Look for sync status indicators in this column
            try {
              const headerCell = activeSheet.getRange(1, i + 1);  // i is 0-based, getRange is 1-based
              const headerValue = headerCell.getValue();
              const note = headerCell.getNote();

              // Extra check for Sync Status indicators
              if (headerValue === "Sync Status" ||
                (note && (note.includes('sync') || note.includes('track')))) {
                cleanupColumnFormatting(activeSheet, colLetter);
              }
            } catch (e) {
              Logger.log(`Error checking column ${colLetter}: ${e.message}`);
            }
          }
        }
      }
    }

    // Save settings to properties
    scriptProperties.setProperty(twoWaySyncEnabledKey, enableTwoWaySync.toString());
    scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);

    // Clean up previous Sync Status column formatting
    cleanupPreviousSyncStatusColumn(activeSheet, activeSheetName);

    // If enabling two-way sync, set up the tracking column
    if (enableTwoWaySync) {
      // Set up the onEdit trigger to track changes - use direct call to avoid variable reference issues
      if (typeof SyncService !== 'undefined' && typeof SyncService.setupOnEditTrigger === 'function') {
        SyncService.setupOnEditTrigger();
      } else {
        // Fallback to using the global function if available
        if (typeof setupOnEditTrigger === 'function') {
          setupOnEditTrigger();
        } else {
          Logger.log('Warning: setupOnEditTrigger function not found');
        }
      }
      
      // If no tracking column specified, use last column
      let columnIndex;
      if (!trackingColumn) {
        columnIndex = activeSheet.getLastColumn() + 1;
        trackingColumn = columnToLetter(columnIndex);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
      } else {
        columnIndex = columnLetterToIndex(trackingColumn);
      }

      // Set up the tracking column
      const headerCell = activeSheet.getRange(1, columnIndex);
      headerCell.setValue('Sync Status')
        .setBackground('#E8F0FE')
        .setFontWeight('bold')
        .setNote('This column tracks changes for two-way sync with Pipedrive');

      // Style the entire status column with a light background and border
      const fullStatusColumn = activeSheet.getRange(1, columnIndex, Math.max(activeSheet.getLastRow(), 2), 1);
      fullStatusColumn.setBackground('#F8F9FA') // Light gray background
        .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

      // Initialize all rows with "Not modified" status
      if (activeSheet.getLastRow() > 1) {
        // Get all data to identify which rows should have status
        const allData = activeSheet.getDataRange().getValues();
        const statusValues = activeSheet.getRange(2, columnIndex, activeSheet.getLastRow() - 1, 1).getValues();
        const newStatusValues = [];

        // Process each row (starting from row 2)
        for (let i = 1; i < allData.length; i++) {
          const row = allData[i];
          const firstCell = row[0] ? row[0].toString().toLowerCase() : '';
          const isEmpty = row.every(cell => cell === '' || cell === null || cell === undefined);

          // Skip setting status for:
          // 1. Rows where first cell contains "last" or "sync" (metadata rows)
          // 2. Empty rows
          if (firstCell.includes('last') ||
            firstCell.includes('sync') ||
            firstCell.includes('update') ||
            isEmpty) {
            newStatusValues.push(['']); // Keep empty
          } else {
            // Only set "Not modified" for actual data rows that are empty
            const currentStatus = statusValues[i - 1][0];
            newStatusValues.push([
              currentStatus === '' || currentStatus === null || currentStatus === undefined
                ? 'Not modified'
                : currentStatus
            ]);
          }
        }

        // Set all values at once
        if (newStatusValues.length > 0) {
          activeSheet.getRange(2, columnIndex, newStatusValues.length, 1).setValues(newStatusValues);
        }

        // Add data validation for status values
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
          .build();

        // Apply validation to each data row
        for (let i = 0; i < newStatusValues.length; i++) {
          if (newStatusValues[i][0] !== '') { // Only add validation to rows with status
            activeSheet.getRange(i + 2, columnIndex).setDataValidation(rule);
          }
        }

        // Set up conditional formatting
        const rules = activeSheet.getConditionalFormatRules();
        const statusRange = activeSheet.getRange(2, columnIndex, activeSheet.getLastRow() - 1, 1);

        // Create rules for each status
        const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Modified')
          .setBackground('#FCE8E6')
          .setFontColor('#D93025')
          .setRanges([statusRange])
          .build();

        const syncedRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Synced')
          .setBackground('#E6F4EA')
          .setFontColor('#137333')
          .setRanges([statusRange])
          .build();

        const errorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Error')
          .setBackground('#FCE8E6')
          .setFontColor('#D93025')
          .setBold(true)
          .setRanges([statusRange])
          .build();

        rules.push(modifiedRule, syncedRule, errorRule);
        activeSheet.setConditionalFormatRules(rules);
      }
    }

    // Update last sync time
    const now = new Date().toISOString();
    scriptProperties.setProperty(twoWaySyncLastSyncKey, now);

    // Clean up onEdit trigger if two-way sync is disabled
    if (!enableTwoWaySync) {
      // Use direct call to avoid variable reference issues
      if (typeof SyncService !== 'undefined' && typeof SyncService.removeOnEditTrigger === 'function') {
        SyncService.removeOnEditTrigger();
      } else {
        // Fallback to using the global function if available
        if (typeof removeOnEditTrigger === 'function') {
          removeOnEditTrigger();
        } else {
          Logger.log('Warning: removeOnEditTrigger function not found');
        }
      }
    }

    return true;
  } catch (error) {
    Logger.log(`Error saving two-way sync settings: ${error.message}`);
    throw error;
  }
}

/**
 * Get the styles for the two-way sync settings UI
 */
TwoWaySyncSettingsUI.getStyles = function() {
  return `<style>
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
    
    h3 {
      font-size: 18px;
      font-weight: 500;
      margin-bottom: 16px;
      color: var(--text-dark);
    }
    
    h4 {
      font-size: 16px;
      font-weight: 500;
      margin-bottom: 8px;
      color: var(--text-dark);
    }
    
    .form-container {
      max-width: 100%;
    }
    
    .sheet-info {
      background-color: var(--bg-light);
      padding: 12px 16px;
      border-radius: 8px;
      margin-bottom: 20px;
      border-left: 4px solid var(--primary-color);
      display: flex;
      align-items: center;
    }
    
    .sheet-info svg {
      margin-right: 12px;
      fill: var(--primary-color);
    }
    
    .info-alert {
      background-color: #FFF8E1;
      padding: 12px 16px;
      border-radius: 8px;
      margin-bottom: 20px;
      border-left: 4px solid var(--warning-color);
    }
    
    .info-alert h4 {
      color: var(--text-dark);
      font-size: 14px;
      margin-bottom: 4px;
    }
    
    .view-only-alert {
      background-color: #e8eaf6;
      padding: 12px 16px;
      border-radius: 8px;
      margin-bottom: 20px;
      border-left: 4px solid #3f51b5;
    }
    
    .view-only-alert h4 {
      color: #3f51b5;
      font-size: 14px;
      margin-bottom: 4px;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .switch-container {
      display: flex;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .switch {
      position: relative;
      display: inline-block;
      width: 40px;
      height: 20px;
      margin-right: 12px;
    }
    
    .switch input {
      opacity: 0;
      width: 0;
      height: 0;
    }
    
    .slider {
      position: absolute;
      cursor: pointer;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: #ccc;
      transition: .4s;
      border-radius: 20px;
    }
    
    .slider:before {
      position: absolute;
      content: "";
      height: 16px;
      width: 16px;
      left: 2px;
      bottom: 2px;
      background-color: white;
      transition: .4s;
      border-radius: 50%;
    }
    
    input:checked + .slider {
      background-color: var(--success-color);
    }
    
    input:focus + .slider {
      box-shadow: 0 0 1px var(--success-color);
    }
    
    input:checked + .slider:before {
      transform: translateX(20px);
    }
    
    label {
      display: block;
      font-weight: 500;
      margin-bottom: 8px;
      color: var(--text-dark);
    }
    
    input, select {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid var(--border-color);
      border-radius: 4px;
      font-size: 14px;
      transition: var(--transition);
    }
    
    input:focus, select:focus {
      outline: none;
      border-color: var(--primary-color);
      box-shadow: 0 0 0 2px rgba(66, 133, 244, 0.2);
    }
    
    input:disabled, select:disabled {
      background-color: #f5f5f5;
      color: #999;
      cursor: not-allowed;
    }
    
    .switch input:disabled + .slider {
      background-color: #ccc;
      cursor: not-allowed;
    }
    
    .tooltip {
      display: block;
      font-size: 12px;
      color: var(--text-light);
      margin-top: 4px;
    }
    
    .button-container {
      display: flex;
      justify-content: flex-end;
      margin-top: 24px;
    }
    
    .button-primary {
      background-color: var(--primary-color);
      color: white;
      border: none;
      padding: 10px 24px;
      border-radius: 4px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: var(--transition);
    }
    
    .button-primary:hover {
      background-color: var(--primary-dark);
      box-shadow: var(--shadow-hover);
    }
    
    .button-primary:disabled {
      background-color: #ccc;
      cursor: not-allowed;
      opacity: 0.6;
    }
    
    .button-primary:disabled:hover {
      background-color: #ccc;
      box-shadow: none;
    }
    
    .button-secondary {
      background-color: transparent;
      color: var(--primary-color);
      border: 1px solid var(--primary-color);
      padding: 9px 16px;
      margin-right: 12px;
      border-radius: 4px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: var(--transition);
    }
    
    .button-secondary:hover {
      background-color: rgba(66, 133, 244, 0.04);
    }
    
    .status {
      margin-top: 20px;
      padding: 12px 16px;
      border-radius: 4px;
      font-size: 14px;
      display: none;
    }
    
    .status.success {
      background-color: #e6f4ea;
      color: var(--success-color);
      display: block;
    }
    
    .status.error {
      background-color: #fce8e6;
      color: var(--error-color);
      display: block;
    }
    
    .spinner {
      display: inline-block;
      width: 20px;
      height: 20px;
      margin-right: 8px;
      border: 2px solid rgba(255, 255, 255, 0.3);
      border-radius: 50%;
      border-top-color: #fff;
      animation: spin 0.8s linear infinite;
      vertical-align: middle;
    }
    
    @keyframes spin {
      to {
        transform: rotate(360deg);
      }
    }
    
    .button-primary.loading {
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: wait;
      opacity: 0.9;
    }
    
    .last-sync {
      font-size: 12px;
      color: var(--text-light);
      margin-top: 8px;
    }
    
    .feature-table {
      width: 100%;
      border-collapse: collapse;
      margin: 12px 0 20px;
      background-color: white;
    }
    
    .feature-table th,
    .feature-table td {
      padding: 12px;
      text-align: left;
      border: 1px solid var(--border-color);
      line-height: 1.4;
    }
    
    .feature-table th {
      background-color: var(--bg-light);
      font-weight: 500;
      color: var(--text-dark);
    }
    
    .feature-table td {
      vertical-align: top;
    }
    
    .feature-table tr:hover td {
      background-color: var(--bg-light);
    }
    
    .hidden {
      display: none !important;
    }
  </style>`;
};

/**
 * Get the scripts for the two-way sync settings UI
 */
TwoWaySyncSettingsUI.getScripts = function() {
  return `<script>
    // Initialize read-only mode
    let isReadOnly = <?= isReadOnly ? 'true' : 'false' ?>;
    window.isReadOnly = isReadOnly;
    
    // Initialize
    document.addEventListener('DOMContentLoaded', function() {
      // Hide any loading indicators
      if (document.getElementById('loading-indicator')) {
        document.getElementById('loading-indicator').classList.add('hidden');
      }
      
      // Set up event listeners
      document.getElementById('cancelBtn').addEventListener('click', closeDialog);
      document.getElementById('saveBtn').addEventListener('click', saveSettings);
    });
    
    // Close the dialog
    function closeDialog() {
      google.script.host.close();
    }
    
    // Show a status message
    function showStatus(type, message) {
      const statusEl = document.getElementById('status');
      statusEl.className = 'status ' + type;
      statusEl.textContent = message;
    }
    
    // Save settings
    function saveSettings() {
      // Check if in read-only mode
      if (window.isReadOnly || <?= isReadOnly ? 'true' : 'false' ?>) {
        showStatus('error', 'You do not have permission to modify settings. Only team admins can change these settings.');
        return;
      }
      
      const enableTwoWaySync = document.getElementById('enableTwoWaySync').checked;
      const trackingColumn = document.getElementById('trackingColumn').value.trim();
      
      // Show loading state
      const saveBtn = document.getElementById('saveBtn');
      const saveSpinner = document.getElementById('saveSpinner');
      saveBtn.classList.add('loading');
      saveSpinner.style.display = 'inline-block';
      saveBtn.disabled = true;
      
      // Save settings via the server-side function
      google.script.run
        .withSuccessHandler(function() {
          // Hide loading state
          saveBtn.classList.remove('loading');
          saveSpinner.style.display = 'none';
          saveBtn.disabled = false;
          
          showStatus('success', 'Two-way sync settings saved successfully!');
          setTimeout(closeDialog, 1500);
        })
        .withFailureHandler(function(error) {
          // Hide loading state
          saveBtn.classList.remove('loading');
          saveSpinner.style.display = 'none';
          saveBtn.disabled = false;
          
          showStatus('error', 'Error: ' + error.message);
        })
        .saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn);
    }
  </script>`;
};

/**
 * Check and recreate the onEdit trigger after column changes
 * This function is called after syncing data
 * @param {string} sheetName - The name of the sheet that was synced
 */
function checkAndRecreateTriggers(sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // Check if we need to recreate triggers
    const twoWaySyncRecreateTriggersKey = `TWOWAY_SYNC_RECREATE_TRIGGERS_${sheetName}`;
    const shouldRecreateTriggers = scriptProperties.getProperty(twoWaySyncRecreateTriggersKey) === 'true';
    
    if (twoWaySyncEnabled && shouldRecreateTriggers) {
      Logger.log(`Recreating onEdit trigger for sheet "${sheetName}" after column changes`);
      
      // Remove the existing trigger first
      if (typeof SyncService !== 'undefined' && typeof SyncService.removeOnEditTrigger === 'function') {
        SyncService.removeOnEditTrigger();
      } else if (typeof removeOnEditTrigger === 'function') {
        removeOnEditTrigger();
      }
      
      // Create a new trigger
      if (typeof SyncService !== 'undefined' && typeof SyncService.setupOnEditTrigger === 'function') {
        SyncService.setupOnEditTrigger();
      } else if (typeof setupOnEditTrigger === 'function') {
        setupOnEditTrigger();
      }
      
      // Clear the flag
      scriptProperties.deleteProperty(twoWaySyncRecreateTriggersKey);
      
      Logger.log(`Successfully recreated onEdit trigger for sheet "${sheetName}"`);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in checkAndRecreateTriggers: ${e.message}`);
    return false;
  }
}

// Export functions to be globally accessible
this.showTwoWaySyncSettings = TwoWaySyncSettingsUI.showTwoWaySyncSettings;
this.handleColumnPreferencesChange = TwoWaySyncSettingsUI.handleColumnPreferencesChange;
this.checkAndRecreateTriggers = checkAndRecreateTriggers;

/**
 * Creates a sync trigger based on user preferences
 */
function createSyncTrigger(triggerData) {
  try {
    // Configure the trigger based on frequency
    switch (triggerData.frequency) {
      case 'hourly':
        const hourlyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .everyHours(triggerData.hourlyInterval)
          .create();

        // Store sheet name and frequency in trigger properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty(
          `TRIGGER_${hourlyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        scriptProperties.setProperty(
          `TRIGGER_${hourlyTrigger.getUniqueId()}_FREQUENCY`,
          'hourly'
        );

        return {
          success: true,
          triggerId: hourlyTrigger.getUniqueId(),
          triggerInfo: {
            id: hourlyTrigger.getUniqueId(),
            type: 'Hourly',
            description: `Every ${triggerData.hourlyInterval} hour${triggerData.hourlyInterval > 1 ? 's' : ''} for sheet "${triggerData.sheetName}"`
          }
        };

      case 'daily':
        const dailyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .atHour(triggerData.hour)
          .nearMinute(triggerData.minute)
          .everyDays(1)
          .create();

        // Store sheet name and frequency in trigger properties
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${dailyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${dailyTrigger.getUniqueId()}_FREQUENCY`,
          'daily'
        );

        // Format time string
        const hour12 = triggerData.hour % 12 === 0 ? 12 : triggerData.hour % 12;
        const ampm = triggerData.hour < 12 ? 'AM' : 'PM';
        const timeStr = `${hour12}:${triggerData.minute < 10 ? '0' + triggerData.minute : triggerData.minute} ${ampm}`;

        return {
          success: true,
          triggerId: dailyTrigger.getUniqueId(),
          triggerInfo: {
            id: dailyTrigger.getUniqueId(),
            type: 'Daily',
            description: `Every day at ${timeStr} for sheet "${triggerData.sheetName}"`
          }
        };

      case 'weekly':
        // For weekly triggers, we need to create a separate trigger for each selected day
        const weekDayTriggers = [];
        const weekDayDetails = [];
        
        // Convert numeric day values to ScriptApp.WeekDay enum values
        const weekDayMap = {
          1: ScriptApp.WeekDay.MONDAY,
          2: ScriptApp.WeekDay.TUESDAY,
          3: ScriptApp.WeekDay.WEDNESDAY,
          4: ScriptApp.WeekDay.THURSDAY,
          5: ScriptApp.WeekDay.FRIDAY,
          6: ScriptApp.WeekDay.SATURDAY,
          7: ScriptApp.WeekDay.SUNDAY
        };

        const weekDayNames = {
          1: 'Monday',
          2: 'Tuesday',
          3: 'Wednesday',
          4: 'Thursday',
          5: 'Friday',
          6: 'Saturday',
          7: 'Sunday'
        };
        
        triggerData.weekDays.forEach(day => {
          const weekDay = weekDayMap[day];
          if (!weekDay) {
            throw new Error(`Invalid week day value: ${day}`);
          }
          
          const dayTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
            .timeBased()
            .onWeekDay(weekDay)
            .atHour(triggerData.hour)
            .nearMinute(triggerData.minute)
            .create();

          // Store sheet name and frequency in trigger properties
          PropertiesService.getScriptProperties().setProperty(
            `TRIGGER_${dayTrigger.getUniqueId()}_SHEET`,
            triggerData.sheetName
          );
          PropertiesService.getScriptProperties().setProperty(
            `TRIGGER_${dayTrigger.getUniqueId()}_FREQUENCY`,
            'weekly'
          );

          weekDayTriggers.push(dayTrigger.getUniqueId());
          weekDayDetails.push(weekDayNames[day]);
        });

        // Format time string
        const weekHour12 = triggerData.hour % 12 === 0 ? 12 : triggerData.hour % 12;
        const weekAmpm = triggerData.hour < 12 ? 'AM' : 'PM';
        const weekTimeStr = `${weekHour12}:${triggerData.minute < 10 ? '0' + triggerData.minute : triggerData.minute} ${weekAmpm}`;

        return {
          success: true,
          triggerId: weekDayTriggers.join(','), // Return a comma-separated list of trigger IDs
          triggerInfo: {
            id: weekDayTriggers.join(','),
            type: 'Weekly',
            description: `Every week on ${weekDayDetails.join(', ')} at ${weekTimeStr} for sheet "${triggerData.sheetName}"`
          }
        };

      case 'monthly':
        const monthlyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .onMonthDay(triggerData.monthDay)
          .atHour(triggerData.hour)
          .nearMinute(triggerData.minute)
          .create();

        // Store sheet name and frequency in trigger properties
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${monthlyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${monthlyTrigger.getUniqueId()}_FREQUENCY`,
          'monthly'
        );

        // Format time string
        const monthHour12 = triggerData.hour % 12 === 0 ? 12 : triggerData.hour % 12;
        const monthAmpm = triggerData.hour < 12 ? 'AM' : 'PM';
        const monthTimeStr = `${monthHour12}:${triggerData.minute < 10 ? '0' + triggerData.minute : triggerData.minute} ${monthAmpm}`;

        // Format day with proper suffix
        let daySuffix = "th";
        if (triggerData.monthDay % 10 === 1 && triggerData.monthDay !== 11) daySuffix = "st";
        else if (triggerData.monthDay % 10 === 2 && triggerData.monthDay !== 12) daySuffix = "nd";
        else if (triggerData.monthDay % 10 === 3 && triggerData.monthDay !== 13) daySuffix = "rd";

        return {
          success: true,
          triggerId: monthlyTrigger.getUniqueId(),
          triggerInfo: {
            id: monthlyTrigger.getUniqueId(),
            type: 'Monthly',
            description: `Every month on the ${triggerData.monthDay}${daySuffix} at ${monthTimeStr} for sheet "${triggerData.sheetName}"`
          }
        };

      default:
        return {
          success: false,
          error: 'Invalid frequency selected'
        };
    }

  } catch (error) {
    Logger.log('Error creating trigger: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Syncs a specific sheet when triggered by a time-based trigger
 */
function syncSheetFromTrigger(event) {
  try {
    // Get the trigger ID from the event
    const triggerId = event.triggerUid;

    // Get the sheet name associated with this trigger
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetName = scriptProperties.getProperty(`TRIGGER_${triggerId}_SHEET`);

    if (!sheetName) {
      Logger.log('No sheet name found for trigger: ' + triggerId);
      return;
    }

    // Get the spreadsheet and find the sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found in the spreadsheet.`);
      return;
    }

    // Activate the sheet
    sheet.activate();

    // Set it as the current sheet for the sync operation
    scriptProperties.setProperty('SHEET_NAME', sheetName);

    // Check if two-way sync is enabled for this sheet
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is enabled, push changes to Pipedrive first
    if (twoWaySyncEnabled) {
      Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Pushing changes to Pipedrive before pulling new data.`);
      pushChangesToPipedrive(true); // Pass true to indicate this is a scheduled sync
    }

    // Get the entity type for this sheet
    const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Run the appropriate sync function based on entity type
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        syncDealsFromFilter();
        break;
      case ENTITY_TYPES.PERSONS:
        syncPersonsFromFilter();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        syncOrganizationsFromFilter();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        syncActivitiesFromFilter();
        break;
      case ENTITY_TYPES.LEADS:
        syncLeadsFromFilter();
        break;
      case ENTITY_TYPES.PRODUCTS:
        syncProductsFromFilter();
        break;
      default:
        Logger.log('Unknown entity type: ' + entityType);
        break;
    }

    // Log the completion
    Logger.log(`Scheduled sync completed for sheet "${sheetName}" with entity type "${entityType}"`);

  } catch (error) {
    Logger.log('Error in scheduled sync: ' + error.message);
  }
}

/**
 * Deletes a trigger by its unique ID
 */
function deleteTrigger(triggerId) {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    const scriptProperties = PropertiesService.getScriptProperties();

    // Handle comma-separated list of trigger IDs (for weekly triggers)
    const triggerIds = triggerId.split(',');
    let allDeleted = true;

    triggerIds.forEach(id => {
      let found = false;
      // Find and delete the trigger with this ID
      for (let i = 0; i < allTriggers.length; i++) {
        if (allTriggers[i].getUniqueId() === id) {
          try {
            ScriptApp.deleteTrigger(allTriggers[i]);
            found = true;
          } catch (e) {
            Logger.log(`Error deleting trigger ${id}: ${e.message}`);
            allDeleted = false;
          }
          break;
        }
      }

      // Always clean up the properties, even if trigger wasn't found
      // (it might have been deleted by some other process)
      try {
        // Get the sheet name first for complete cleanup
        const sheetName = scriptProperties.getProperty(`TRIGGER_${id}_SHEET`);
        
        // Delete the main properties
        scriptProperties.deleteProperty(`TRIGGER_${id}_SHEET`);
        scriptProperties.deleteProperty(`TRIGGER_${id}_FREQUENCY`);
        
        // Delete any other properties that might be related to this trigger
        const allProps = scriptProperties.getProperties();
        Object.keys(allProps).forEach(key => {
          if (key.includes(id)) {
            scriptProperties.deleteProperty(key);
          }
        });
        
        Logger.log(`Cleaned up properties for trigger ID: ${id}`);
      } catch (e) {
        Logger.log(`Error cleaning up properties for trigger ${id}: ${e.message}`);
      }
    });

    return { 
      success: allDeleted,
      error: allDeleted ? null : "Some triggers could not be deleted. Stale references have been cleaned up."
    };
  } catch (error) {
    Logger.log('Error deleting trigger: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
} 