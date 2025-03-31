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
        let selectedDays = [];
        let isLoading = false;

        function onLoad() {
          // Get current sheet name and two-way sync status
          const sheetName = <?= JSON.stringify(sheetName) ?>;
          const twoWaySyncEnabled = <?= JSON.stringify(twoWaySyncEnabled) ?>;
          const currentTriggers = <?= JSON.stringify(currentTriggers) ?>;

          // Update UI with current sheet info
          document.getElementById('sheetNameDisplay').textContent = sheetName;
          
          // Show two-way sync notice if enabled
          if (twoWaySyncEnabled) {
            document.getElementById('twoWaySyncNotice').style.display = 'flex';
          }

          // Display existing triggers
          displayExistingTriggers(currentTriggers);

          // Set up event listeners
          document.getElementById('frequency').addEventListener('change', handleFrequencyChange);
          document.querySelectorAll('.day-button').forEach(btn => {
            btn.addEventListener('click', () => toggleDaySelection(btn));
          });
        }

        function handleFrequencyChange() {
          const frequency = document.getElementById('frequency').value;
          const timeInputs = document.getElementById('timeInputs');
          const daySelection = document.getElementById('daySelection');
          const monthDayInput = document.getElementById('monthDayInput');

          // Hide all inputs first
          timeInputs.style.display = 'none';
          daySelection.style.display = 'none';
          monthDayInput.style.display = 'none';

          // Show relevant inputs based on frequency
          switch (frequency) {
            case 'hourly':
              break;
            case 'daily':
            case 'weekly':
              timeInputs.style.display = 'flex';
              if (frequency === 'weekly') {
                daySelection.style.display = 'flex';
              }
              break;
            case 'monthly':
              timeInputs.style.display = 'flex';
              monthDayInput.style.display = 'block';
              break;
          }
        }

        function toggleDaySelection(button) {
          const day = parseInt(button.dataset.day);
          const index = selectedDays.indexOf(day);
          
          if (index === -1) {
            selectedDays.push(day);
            button.classList.add('selected');
          } else {
            selectedDays.splice(index, 1);
            button.classList.remove('selected');
          }
        }

        function createSchedule() {
          if (isLoading) return;
          isLoading = true;

          const frequency = document.getElementById('frequency').value;
          const hour = parseInt(document.getElementById('hour').value) || 0;
          const minute = parseInt(document.getElementById('minute').value) || 0;
          const monthDay = parseInt(document.getElementById('monthDay').value) || 1;
          const hourlyInterval = parseInt(document.getElementById('hourlyInterval').value) || 1;

          // Show loading state
          const button = document.getElementById('createButton');
          button.disabled = true;
          document.getElementById('buttonLoading').style.display = 'inline-flex';

          // Prepare trigger data
          const triggerData = {
            frequency: frequency,
            sheetName: <?= JSON.stringify(sheetName) ?>,
            hour: hour,
            minute: minute,
            monthDay: monthDay,
            hourlyInterval: hourlyInterval,
            weekDays: selectedDays
          };

          // Call server-side function
          google.script.run
            .withSuccessHandler(onScheduleCreated)
            .withFailureHandler(onScheduleError)
            .createSyncTrigger(triggerData);
        }

        function onScheduleCreated(result) {
          isLoading = false;
          const button = document.getElementById('createButton');
          button.disabled = false;
          document.getElementById('buttonLoading').style.display = 'none';

          if (result.success) {
            // Show success message
            showIndicator('success', 'Sync schedule created successfully!');
            // Reload the page to show new trigger
            setTimeout(() => google.script.run.showTriggerManager(), 2000);
          } else {
            showIndicator('error', result.error || 'Failed to create sync schedule.');
          }
        }

        function onScheduleError(error) {
          isLoading = false;
          const button = document.getElementById('createButton');
          button.disabled = false;
          document.getElementById('buttonLoading').style.display = 'none';
          showIndicator('error', error.message || 'Failed to create sync schedule.');
        }

        function displayExistingTriggers(triggers) {
          const container = document.getElementById('existingTriggers');
          container.innerHTML = '';

          if (!triggers || triggers.length === 0) {
            container.innerHTML = '<p class="help-text">No sync schedules configured.</p>';
            return;
          }

          triggers.forEach(trigger => {
            const triggerEl = document.createElement('div');
            triggerEl.className = 'trigger-item';
            triggerEl.innerHTML = \`
              <div class="trigger-info">
                <strong>\${trigger.type}</strong>
                <span>\${trigger.description}</span>
              </div>
              <button class="delete-trigger" onclick="deleteTrigger('\${trigger.id}')">
                <span class="mini-loading" style="display: none;">
                  <span class="mini-loader"></span>
                </span>
                <span class="button-text">Delete</span>
              </button>
            \`;
            container.appendChild(triggerEl);
          });
        }

        function deleteTrigger(triggerId) {
          const button = event.target.closest('.delete-trigger');
          if (!button || button.disabled) return;

          button.disabled = true;
          button.querySelector('.mini-loading').style.display = 'inline-flex';
          button.querySelector('.button-text').style.display = 'none';

          google.script.run
            .withSuccessHandler(() => {
              showIndicator('success', 'Sync schedule deleted successfully!');
              setTimeout(() => google.script.run.showTriggerManager(), 2000);
            })
            .withFailureHandler(error => {
              button.disabled = false;
              button.querySelector('.mini-loading').style.display = 'none';
              button.querySelector('.button-text').style.display = 'inline';
              showIndicator('error', error.message || 'Failed to delete sync schedule.');
            })
            .deleteSyncTrigger(triggerId);
        }

        function showIndicator(type, message) {
          const indicator = document.getElementById('indicator');
          indicator.className = \`indicator \${type}\`;
          indicator.textContent = message;
          indicator.style.display = 'block';

          setTimeout(() => {
            indicator.style.display = 'none';
          }, 5000);
        }
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
      .setWidth(600)
      .setHeight(700)
      .setTitle('Schedule Automatic Sync');
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Automatic Sync');
  }
}; 