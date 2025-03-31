/**
 * TwoWaySyncSettingsUI
 * 
 * This module handles the UI components and functionality for two-way sync settings:
 * - Displaying the settings dialog
 * - Managing sync settings
 * - Handling settings persistence
 */

const TwoWaySyncSettingsUI = {
  getStyles() {
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
      </style>
    `;
  },

  getScripts() {
    return `
      <script>
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
      </script>
    `;
  },

  showTwoWaySyncSettings() {
    const template = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
    
    // Get the active sheet name
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    
    // Get current settings
    const scriptProperties = PropertiesService.getScriptProperties();
    const settingsKey = `TWOWAY_SYNC_SETTINGS_${sheetName}`;
    const settingsJson = scriptProperties.getProperty(settingsKey) || '{}';
    const settings = JSON.parse(settingsJson);
    
    // Pass data to template
    template.sheetName = sheetName;
    template.settings = settings;
    template.entityType = 'deals'; // TODO: Make this dynamic based on sheet type
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(600)
      .setHeight(700)
      .setTitle('Two-Way Sync Settings');
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
  },

  saveTwoWaySyncSettings(enabled, trackingColumn) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    
    // Save settings
    const settings = {
      enabled: enabled,
      trackingColumn: trackingColumn,
      lastSync: new Date().toISOString()
    };
    
    const scriptProperties = PropertiesService.getScriptProperties();
    const settingsKey = `TWOWAY_SYNC_SETTINGS_${sheetName}`;
    scriptProperties.setProperty(settingsKey, JSON.stringify(settings));
    
    // Set up tracking column if enabled
    if (enabled) {
      this.setupTrackingColumn(sheet, trackingColumn);
    }
  },

  setupTrackingColumn(sheet, trackingColumn) {
    // If no tracking column specified, use last column
    if (!trackingColumn) {
      trackingColumn = sheet.getLastColumn();
    } else {
      // Convert letter to column number
      trackingColumn = this.columnToNumber(trackingColumn);
    }
    
    // Set header
    sheet.getRange(1, trackingColumn).setValue('Last Modified');
    
    // Format as date
    const range = sheet.getRange(2, trackingColumn, sheet.getLastRow() - 1);
    range.setNumberFormat('yyyy-MM-dd HH:mm:ss');
  },

  columnToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result *= 26;
      result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return result;
  }
}; 