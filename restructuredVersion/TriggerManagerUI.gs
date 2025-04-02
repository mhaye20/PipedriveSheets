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
                  // Close and reopen the dialog to ensure a fresh reload
                  google.script.host.close();
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
                    // Add fade-out class for animation
                    row.classList.add('fade-out');
                    
                    // After animation completes, remove the row from DOM
                    // and then reload the dialog to show updated list
                    setTimeout(() => {
                      // First remove the row completely from DOM
                      if (row.parentNode) {
                        row.parentNode.removeChild(row);
                      }
                      
                      // If this was the last row, show the "no triggers" message
                      const remainingRows = document.querySelectorAll('.triggers-table tbody tr');
                      if (remainingRows.length === 0) {
                        const table = document.querySelector('.existing-triggers');
                        if (table) {
                          table.style.display = 'none';
                          
                          // Show "no triggers" message
                          const noTriggersDiv = document.createElement('div');
                          noTriggersDiv.className = 'no-triggers';
                          noTriggersDiv.innerHTML = '<p>No automatic sync schedules are set up for this sheet.</p>';
                          table.parentNode.insertBefore(noTriggersDiv, table.nextSibling);
                        }
                      }
                      
                      // Then refresh the whole dialog to ensure we see the latest state
                      google.script.host.close();
                      google.script.run.showTriggerManager();
                    }, 500);
                  } else {
                    // Fallback if row not found - just reload the dialog
                    google.script.host.close();
                    google.script.run.showTriggerManager();
                  }
                } else {
                  // Show error and reset loading state
                  if (loadingElement && buttonElement) {
                    loadingElement.style.display = 'none';
                    buttonElement.style.display = 'inline-block';
                  }
                  showStatus('error', result.error || 'Error deleting trigger');
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
        document.addEventListener('DOMContentLoaded', function() {
          updateFormVisibility();
        });
      </script>
    `;
  },

  showTriggerManager() {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get the entity type for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    // Check if two-way sync is enabled for this sheet
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
    
    // Get current triggers for this sheet
    const currentTriggers = getTriggersForSheet(activeSheetName);
    
    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('TriggerManager');
    
    // Pass data to template
    template.sheetName = activeSheetName;
    template.entityType = entityType;
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
  },

  getTriggersForSheet(sheetName) {
    try {
      const allTriggers = ScriptApp.getProjectTriggers();
      const scriptProperties = PropertiesService.getScriptProperties();

      // Filter to get only valid triggers for this sheet
      const validTriggers = allTriggers.filter(trigger => {
        // Only time-based triggers that run syncSheetFromTrigger
        if (trigger.getHandlerFunction() === 'syncSheetFromTrigger' &&
          trigger.getEventType() === ScriptApp.EventType.CLOCK) {
          // Check if this trigger is for the specified sheet
          const triggerId = trigger.getUniqueId();
          const triggerSheet = scriptProperties.getProperty(`TRIGGER_${triggerId}_SHEET`);
          return triggerSheet === sheetName;
        }
        return false;
      });
      
      // Cleanup: Remove any stale trigger references from properties
      // This ensures we don't have "ghost" triggers showing up after deletion
      const allTriggerIds = new Set(allTriggers.map(t => t.getUniqueId()));
      
      // Get all property keys
      const allProps = scriptProperties.getProperties();
      
      // Find keys that look like they're related to triggers for this sheet
      Object.keys(allProps).forEach(key => {
        if (key.startsWith('TRIGGER_') && key.endsWith('_SHEET')) {
          const triggerId = key.replace('TRIGGER_', '').replace('_SHEET', '');
          if (!allTriggerIds.has(triggerId) && allProps[key] === sheetName) {
            // This is a stale trigger reference - delete it
            scriptProperties.deleteProperty(key);
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_FREQUENCY`);
            Logger.log(`Cleaned up stale trigger reference: ${triggerId}`);
          }
        }
      });
      
      // Map triggers to the format expected by the UI
      return validTriggers.map(trigger => {
        const info = this.getTriggerInfo(trigger);
        return {
          id: trigger.getUniqueId(),
          type: info.type,
          description: info.description
        };
      });
    } catch (e) {
      Logger.log(`Error in getTriggersForSheet: ${e.message}`);
      return [];
    }
  },

  getTriggerInfo(trigger) {
    try {
      // Check if it's a time-based trigger
      if (trigger.getEventType() !== ScriptApp.EventType.CLOCK) {
        return { type: 'Unknown', description: 'Not a time-based trigger' };
      }

      // Get the trigger ID
      const triggerId = trigger.getUniqueId();
      const scriptProperties = PropertiesService.getScriptProperties();

      // Get the sheet name
      const sheetName = scriptProperties.getProperty(`TRIGGER_${triggerId}_SHEET`);
      const sheetInfo = sheetName ? ` for sheet "${sheetName}"` : '';

      // First check if we have the frequency stored as a property
      const storedFrequency = scriptProperties.getProperty(`TRIGGER_${triggerId}_FREQUENCY`);

      if (storedFrequency) {
        // We have the stored frequency, now format description based on it
        switch (storedFrequency) {
          case 'hourly':
            let hourInterval = 1;
            try {
              hourInterval = trigger.getHours() || 1;
            } catch (e) { }

            return {
              type: 'Hourly',
              description: `Every ${hourInterval} hour${hourInterval > 1 ? 's' : ''}${sheetInfo}`
            };

          case 'daily':
            let timeStr = '';
            try {
              const atHour = trigger.getAtHour();
              const atMinute = trigger.getNearMinute();

              if (atHour !== null && atMinute !== null) {
                const hour12 = atHour % 12 === 0 ? 12 : atHour % 12;
                const ampm = atHour < 12 ? 'AM' : 'PM';
                timeStr = ` at ${hour12}:${atMinute < 10 ? '0' + atMinute : atMinute} ${ampm}`;
              }
            } catch (e) { }

            return {
              type: 'Daily',
              description: `Every day${timeStr}${sheetInfo}`
            };

          case 'weekly':
            let dayInfo = '';
            try {
              const weekDay = trigger.getWeekDay();
              if (weekDay) {
                const weekDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
                dayInfo = ` on ${weekDays[weekDay - 1] || 'a weekday'}`;
              }
            } catch (e) { }

            return {
              type: 'Weekly',
              description: `Every week${dayInfo}${sheetInfo}`
            };

          case 'monthly':
            let dayOfMonth = '';
            try {
              const monthDay = trigger.getMonthDay();
              if (monthDay) {
                dayOfMonth = ` on day ${monthDay}`;
              }
            } catch (e) { }

            return {
              type: 'Monthly',
              description: `Every month${dayOfMonth}${sheetInfo}`
            };

          default:
            return {
              type: capitalizeFirstLetter(storedFrequency),
              description: `${capitalizeFirstLetter(storedFrequency)} sync${sheetInfo}`
            };
        }
      }

      // If we don't have stored frequency, fall back to generic type
      if (sheetName) {
        return {
          type: 'Automatic',
          description: `Sync for sheet "${sheetName}"`
        };
      }

      return {
        type: 'Scheduled',
        description: 'Automatic sync'
      };
    } catch (e) {
      Logger.log(`Error in getTriggerInfo: ${e.message}`);
      return {
        type: 'Scheduled',
        description: 'Automatic sync'
      };
    }
  }
};

// Helper function to capitalize first letter
function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

// Export functions to be globally accessible
this.getTriggersForSheet = TriggerManagerUI.getTriggersForSheet;
this.getTriggerInfo = TriggerManagerUI.getTriggerInfo;
this.createSyncTrigger = createSyncTrigger;
this.deleteTrigger = deleteTrigger;

/**
 * Shows the trigger manager dialog - global function callable from client-side
 * This is the main entry point for showing the trigger manager UI.
 */
function showTriggerManager() {
  return TriggerManagerUI.showTriggerManager();
} 