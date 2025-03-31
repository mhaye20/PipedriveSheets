/**
 * TriggerManagerUI
 * 
 * This module handles the UI components and functionality for trigger management:
 * - Displaying the trigger manager dialog
 * - Creating and deleting triggers
 * - Managing trigger schedules
 */

const TriggerManagerUI = {
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
          margin-top: 24px;
          margin-bottom: 12px;
          color: var(--text-dark);
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
        
        .two-way-sync-notice {
          background-color: #e6f4ea;
          padding: 12px 16px;
          border-radius: 8px;
          margin-bottom: 20px;
          border-left: 4px solid var(--success-color);
          display: flex;
          align-items: flex-start;
        }
        
        .two-way-sync-notice svg {
          margin-right: 12px;
          flex-shrink: 0;
          margin-top: 2px;
          fill: var(--success-color);
        }
        
        .two-way-sync-notice p {
          margin-top: 4px;
          font-size: 13px;
        }
        
        .form-group {
          margin-bottom: 16px;
        }
        
        label {
          display: block;
          font-weight: 500;
          margin-bottom: 8px;
          color: var(--text-dark);
        }
        
        select, input {
          width: 100%;
          padding: 10px 12px;
          border: 1px solid var(--border-color);
          border-radius: 4px;
          font-size: 14px;
          transition: var(--transition);
        }
        
        select:focus, input:focus {
          outline: none;
          border-color: var(--primary-color);
          box-shadow: 0 0 0 2px rgba(66,133,244,0.2);
        }
        
        .help-text {
          font-size: 12px;
          color: var(--text-light);
          margin-top: 6px;
        }
        
        .time-inputs {
          display: flex;
          gap: 12px;
          margin-top: 8px;
        }
        
        .time-inputs label {
          margin-bottom: 4px;
          font-weight: normal;
        }
        
        .day-selection {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
          margin-top: 8px;
        }
        
        .day-button {
          padding: 6px 12px;
          border: 1px solid var(--border-color);
          border-radius: 16px;
          background: white;
          cursor: pointer;
          transition: var(--transition);
        }
        
        .day-button.selected {
          background-color: var(--primary-color);
          color: white;
          border-color: var(--primary-color);
        }
        
        .button-container {
          display: flex;
          justify-content: flex-end;
          margin-top: 24px;
          gap: 12px;
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
        
        .delete-trigger {
          padding: 6px 12px;
          background-color: transparent;
          color: var(--error-color);
          border: 1px solid var(--error-color);
          border-radius: 4px;
          font-size: 12px;
        }
        
        .delete-trigger:hover {
          background-color: rgba(219,68,55,0.08);
        }
        
        .loading {
          display: none;
          align-items: center;
          margin-right: 8px;
        }
        
        .mini-loading {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 24px;
          height: 24px;
          vertical-align: middle;
        }
        
        .loader {
          display: inline-block;
          width: 20px;
          height: 20px;
          border: 3px solid rgba(255,255,255,0.3);
          border-radius: 50%;
          border-top-color: white;
          animation: spin 1s ease-in-out infinite;
        }
        
        .mini-loader {
          display: inline-block;
          width: 16px;
          height: 16px;
          border: 2px solid rgba(219,68,55,0.3);
          border-radius: 50%;
          border-top-color: var(--error-color);
          animation: spin 1s ease-in-out infinite;
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
        
        .triggers-table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 8px;
        }
        
        .triggers-table th, .triggers-table td {
          padding: 10px;
          text-align: left;
          border-bottom: 1px solid var(--border-color);
        }
        
        .triggers-table th {
          background-color: var(--bg-light);
          font-weight: 500;
        }
        
        .no-triggers {
          padding: 16px;
          background-color: var(--bg-light);
          border-radius: 8px;
          text-align: center;
          color: var(--text-light);
          margin-top: 16px;
        }
        
        .hidden {
          display: none;
        }
        
        .fade-out {
          opacity: 0;
          transition: opacity 0.5s ease-out;
        }
        
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
      </style>
    `;
  },

  getScripts() {
    return `
      <script>
        // Form visibility based on frequency selection
        function updateFormVisibility() {
          const frequency = document.getElementById('frequency').value;
          
          // Hide all frequency-specific groups first
          document.getElementById('hourlyGroup').style.display = 'none';
          document.getElementById('weeklyGroup').style.display = 'none';
          document.getElementById('monthlyGroup').style.display = 'none';
          document.getElementById('timeGroup').style.display = 'block';
          
          // Show the relevant group based on selection
          if (frequency === 'hourly') {
            document.getElementById('hourlyGroup').style.display = 'block';
            document.getElementById('timeGroup').style.display = 'none';
          } else if (frequency === 'weekly') {
            document.getElementById('weeklyGroup').style.display = 'block';
          } else if (frequency === 'monthly') {
            document.getElementById('monthlyGroup').style.display = 'block';
          }
        }
        
        // Toggle day selection in weekly view
        function toggleDay(button) {
          button.classList.toggle('selected');
        }
        
        // Get selected days from weekly view
        function getSelectedDays() {
          const buttons = document.querySelectorAll('.day-button.selected');
          return Array.from(buttons).map(btn => parseInt(btn.dataset.day));
        }
        
        // Create a new trigger
        function createTrigger() {
          const frequency = document.getElementById('frequency').value;
          const hour = parseInt(document.getElementById('hour').value);
          const minute = parseInt(document.getElementById('minute').value);
          const sheetName = document.getElementById('sheetName').value;
          
          let triggerData = {
            frequency: frequency,
            hour: hour,
            minute: minute,
            sheetName: sheetName
          };
          
          // Add frequency-specific data
          if (frequency === 'hourly') {
            triggerData.hourlyInterval = parseInt(document.getElementById('hourlyInterval').value);
          } else if (frequency === 'weekly') {
            const selectedDays = getSelectedDays();
            if (selectedDays.length === 0) {
              showStatus('error', 'Please select at least one day of the week');
              return;
            }
            triggerData.weekDays = selectedDays;
          } else if (frequency === 'monthly') {
            triggerData.monthDay = parseInt(document.getElementById('monthDay').value);
          }
          
          // Show loading spinner
          document.getElementById('saveLoading').style.display = 'flex';
          document.getElementById('saveBtn').disabled = true;
          document.getElementById('cancelBtn').disabled = true;
          
          // Create the trigger
          google.script.run
            .withSuccessHandler(function(result) {
              document.getElementById('saveLoading').style.display = 'none';
              document.getElementById('saveBtn').disabled = false;
              document.getElementById('cancelBtn').disabled = false;
              
              if (result.success) {
                showStatus('success', 'Sync schedule created successfully!');
                
                // Reload after short delay to show the new trigger
                setTimeout(function() {
                  google.script.run.showTriggerManager();
                }, 1500);
              } else {
                showStatus('error', 'Error: ' + result.error);
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('saveLoading').style.display = 'none';
              document.getElementById('saveBtn').disabled = false;
              document.getElementById('cancelBtn').disabled = false;
              showStatus('error', 'Error: ' + error.message);
            })
            .createSyncTrigger(triggerData);
        }
        
        // Delete a trigger
        function deleteTrigger(triggerId) {
          if (confirm('Are you sure you want to delete this sync schedule?')) {
            // Show loading spinner and disable button
            const loadingElement = document.getElementById('remove-loading-' + triggerId);
            const buttonElement = document.getElementById('remove-btn-' + triggerId);
            
            if (loadingElement && buttonElement) {
              loadingElement.style.display = 'inline-flex';
              buttonElement.style.display = 'none';
            }
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  // Animation for removing the row
                  const row = document.getElementById('trigger-row-' + triggerId);
                  if (row) {
                    row.classList.add('fade-out');
                    setTimeout(() => {
                      // Reload to show updated triggers after animation
                      google.script.run.showTriggerManager();
                    }, 500);
                  } else {
                    // Fallback if row not found
                    google.script.run.showTriggerManager();
                  }
                } else {
                  // Show error and reset loading state
                  if (loadingElement && buttonElement) {
                    loadingElement.style.display = 'none';
                    buttonElement.style.display = 'inline-block';
                  }
                  showStatus('error', 'Error: ' + result.error);
                }
              })
              .withFailureHandler(function(error) {
                // Show error and reset loading state
                if (loadingElement && buttonElement) {
                  loadingElement.style.display = 'none';
                  buttonElement.style.display = 'inline-block';
                }
                showStatus('error', 'Error: ' + error.message);
              })
              .deleteTrigger(triggerId);
          }
        }
        
        // Show status message
        function showStatus(type, message) {
          const indicator = document.getElementById('statusIndicator');
          indicator.className = 'indicator ' + type;
          indicator.textContent = message;
          indicator.style.display = 'block';
          
          // Auto-hide success messages after a delay
          if (type === 'success') {
            setTimeout(function() {
              indicator.style.display = 'none';
            }, 3000);
          }
        }
        
        // Initial form setup
        updateFormVisibility();
      </script>
    `;
  },

  showTriggerManager() {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // Get current triggers for this sheet
    const currentTriggers = getTriggersForSheet(activeSheetName);
    
    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('TriggerManager');
    
    // Pass data to template
    template.sheetName = activeSheetName;
    template.twoWaySyncEnabled = twoWaySyncEnabled;
    template.currentTriggers = currentTriggers;
    
    // Make include function available to the template
    template.include = include;
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(500)
      .setHeight(650)
      .setTitle('Schedule Automatic Sync');
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Automatic Sync');
  }
}; 