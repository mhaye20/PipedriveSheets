/**
 * TriggerManagerUI
 * 
 * This module handles the UI components and functionality for trigger management:
 * - Displaying the trigger manager dialog
 * - Creating and deleting triggers
 * - Managing trigger schedules
 */

const TriggerManagerUI = {
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