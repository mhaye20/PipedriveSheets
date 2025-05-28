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
    // Check if user has access to scheduled sync feature
    if (!PaymentService.hasFeatureAccess('scheduled_sync')) {
      SpreadsheetApp.getUi().alert(
        'Scheduled sync is only available on Pro and Team plans. Please upgrade to enable automatic syncing.'
      );
      PaymentService.showUpgradeDialog();
      return;
    }
    
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
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_WEEKDAY`);
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_HOUR`);
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MINUTE`);
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MONTHDAY`);
            scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MINUTES_INTERVAL`);
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
          case 'minutes':
            // Get stored minutes interval
            const storedMinutesInterval = scriptProperties.getProperty(`TRIGGER_${triggerId}_MINUTES_INTERVAL`);
            const minutesInterval = storedMinutesInterval ? parseInt(storedMinutesInterval) : 5;
            
            return {
              type: 'Minutes',
              description: `Every ${minutesInterval} minute${minutesInterval > 1 ? 's' : ''}${sheetInfo}`
            };
            
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
            
            // Get stored time information
            const storedDailyHour = scriptProperties.getProperty(`TRIGGER_${triggerId}_HOUR`);
            const storedDailyMinute = scriptProperties.getProperty(`TRIGGER_${triggerId}_MINUTE`);
            
            if (storedDailyHour !== null && storedDailyMinute !== null) {
              const hour = parseInt(storedDailyHour);
              const minute = parseInt(storedDailyMinute);
              const hour12 = hour % 12 === 0 ? 12 : hour % 12;
              const ampm = hour < 12 ? 'AM' : 'PM';
              timeStr = ` at ${hour12}:${minute < 10 ? '0' + minute : minute} ${ampm}`;
            }

            return {
              type: 'Daily',
              description: `Every day${timeStr}${sheetInfo}`
            };

          case 'weekly':
            let dayInfo = '';
            let weeklyTimeStr = '';
            
            // Get stored weekday information
            const storedWeekDay = scriptProperties.getProperty(`TRIGGER_${triggerId}_WEEKDAY`);
            const storedHour = scriptProperties.getProperty(`TRIGGER_${triggerId}_HOUR`);
            const storedMinute = scriptProperties.getProperty(`TRIGGER_${triggerId}_MINUTE`);
            
            if (storedWeekDay) {
              const weekDays = ['', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
              const dayNum = parseInt(storedWeekDay);
              const dayName = weekDays[dayNum] || '';
              if (dayName) {
                dayInfo = ` on ${dayName}`;
              }
            }
            
            if (storedHour !== null && storedMinute !== null) {
              const hour = parseInt(storedHour);
              const minute = parseInt(storedMinute);
              if (!isNaN(hour) && !isNaN(minute)) {
                const hour12 = hour % 12 === 0 ? 12 : hour % 12;
                const ampm = hour < 12 ? 'AM' : 'PM';
                weeklyTimeStr = ` at ${hour12}:${minute < 10 ? '0' + minute : minute} ${ampm}`;
              }
            }

            return {
              type: 'Weekly',
              description: `Every week${dayInfo}${weeklyTimeStr}${sheetInfo}`
            };

          case 'monthly':
            let dayOfMonth = '';
            let monthlyTimeStr = '';
            
            // Get stored monthly day and time information
            const storedMonthDay = scriptProperties.getProperty(`TRIGGER_${triggerId}_MONTHDAY`);
            const storedMonthlyHour = scriptProperties.getProperty(`TRIGGER_${triggerId}_HOUR`);
            const storedMonthlyMinute = scriptProperties.getProperty(`TRIGGER_${triggerId}_MINUTE`);
            
            if (storedMonthDay) {
              dayOfMonth = ` on day ${storedMonthDay}`;
            }
            
            if (storedMonthlyHour !== null && storedMonthlyMinute !== null) {
              const hour = parseInt(storedMonthlyHour);
              const minute = parseInt(storedMonthlyMinute);
              const hour12 = hour % 12 === 0 ? 12 : hour % 12;
              const ampm = hour < 12 ? 'AM' : 'PM';
              monthlyTimeStr = ` at ${hour12}:${minute < 10 ? '0' + minute : minute} ${ampm}`;
            }

            return {
              type: 'Monthly',
              description: `Every month${dayOfMonth}${monthlyTimeStr}${sheetInfo}`
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
      return {
        type: 'Scheduled',
        description: 'Automatic sync'
      };
    }
  },
  
  createSyncTrigger(triggerData) {
    try {
      
      // Validate required data
      if (!triggerData || !triggerData.sheetName || !triggerData.frequency) {
        return { success: false, error: 'Missing required trigger data' };
      }
      
      // Create trigger based on frequency
      let trigger;
      const frequency = triggerData.frequency;
      
      // Set up trigger builder
      const builder = ScriptApp.newTrigger('syncSheetFromTrigger')
        .timeBased();
      
      if (frequency === 'minutes') {
        // Minutes trigger
        const minutesInterval = parseInt(triggerData.minutesInterval) || 5;
        trigger = builder.everyMinutes(minutesInterval).create();
        
        // Save minutes interval info after trigger is created
        if (trigger) {
          const triggerId = trigger.getUniqueId();
          saveSheetInfoForTrigger(triggerId, triggerData.sheetName, frequency, {
            minutesInterval: minutesInterval
          });
          
          // Get trigger info for UI update
          const triggerInfo = this.getTriggerInfo(trigger);
          
          // Log team activity if user is in a team
          if (typeof logTeamActivity === 'function') {
            logTeamActivity('trigger', `created ${frequency} sync trigger`, {
              sheetName: triggerData.sheetName,
              frequency: frequency,
              interval: triggerData.minutesInterval || triggerData.hourlyInterval || null
            });
          }
          
          return {
            success: true,
            message: 'Trigger created successfully',
            triggerId: triggerId,
            triggerInfo: {
              id: triggerId,
              type: triggerInfo.type,
              description: triggerInfo.description
            }
          };
        }
      } else if (frequency === 'hourly') {
        // Hourly trigger
        const hourlyInterval = parseInt(triggerData.hourlyInterval) || 1;
        trigger = builder.everyHours(hourlyInterval).create();
      } else if (frequency === 'daily') {
        // Daily trigger
        const hour = parseInt(triggerData.hour) || 8;
        const minute = parseInt(triggerData.minute) || 0;
        trigger = builder.atHour(hour).nearMinute(minute).everyDays(1).create();
      } else if (frequency === 'weekly') {
        // Weekly trigger - create multiple triggers if multiple days selected
        const weekDays = triggerData.weekDays || [];
        const hour = parseInt(triggerData.hour) || 8;
        const minute = parseInt(triggerData.minute) || 0;
        
        
        if (weekDays.length === 0) {
          // Default to Monday if no days selected
          trigger = builder.atHour(hour).nearMinute(minute).onWeekDay(ScriptApp.WeekDay.MONDAY).create();
          
          // Save info with Monday as the day
          const triggerId = trigger.getUniqueId();
          saveSheetInfoForTrigger(triggerId, triggerData.sheetName, frequency, {
            weekDay: 1, // Monday
            hour: hour,
            minute: minute
          });
        } else {
          // For weekly triggers with multiple days, we'll create separate triggers
          const allCreatedTriggers = [];
          
          // Create triggers for each selected day
          for (let i = 0; i < weekDays.length; i++) {
            const dayNum = parseInt(weekDays[i]);
            const weekDay = getWeekDayFromNumber(dayNum);
            
            let newTrigger;
            if (i === 0) {
              // Use the main builder for the first trigger
              trigger = builder.atHour(hour).nearMinute(minute).onWeekDay(weekDay).create();
              newTrigger = trigger;
            } else {
              // Create additional triggers
              newTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
                .timeBased()
                .atHour(hour)
                .nearMinute(minute)
                .onWeekDay(weekDay)
                .create();
            }
            
            // Save info for each trigger including day and time
            const triggerId = newTrigger.getUniqueId();
            saveSheetInfoForTrigger(triggerId, triggerData.sheetName, frequency, {
              weekDay: dayNum,
              hour: hour,
              minute: minute
            });
            
            allCreatedTriggers.push({
              trigger: newTrigger,
              dayNum: dayNum
            });
          }
          
          // Force a full refresh after creating multiple triggers
          if (weekDays.length > 1) {
            // Add a small delay to ensure properties are saved
            Utilities.sleep(500);
            
            // Return all trigger IDs as an array
            const allTriggerIds = allCreatedTriggers.map(t => t.trigger.getUniqueId());
            // Return success but without specific trigger info to force a refresh
            return {
              success: true,
              message: 'Multiple weekly triggers created successfully',
              forceRefresh: true,
              triggerId: allTriggerIds.join(','),
              triggerInfo: {
                id: allTriggerIds.join(','),
                type: 'Weekly',
                description: `Multiple weekly sync schedules created`
              }
            };
          }
        }
      } else if (frequency === 'monthly') {
        // Monthly trigger
        const monthDay = parseInt(triggerData.monthDay) || 1;
        const hour = parseInt(triggerData.hour) || 8;
        const minute = parseInt(triggerData.minute) || 0;
        trigger = builder.atHour(hour).nearMinute(minute).onMonthDay(monthDay).create();
      } else {
        return { success: false, error: 'Invalid frequency specified' };
      }
      
      // Save sheet name and frequency in properties for this trigger
      if (trigger) {
        const triggerId = trigger.getUniqueId();
        
        // Prepare additional info based on frequency
        let additionalInfo = null;
        if (frequency === 'daily') {
          additionalInfo = {
            hour: parseInt(triggerData.hour) || 8,
            minute: parseInt(triggerData.minute) || 0
          };
        } else if (frequency === 'monthly') {
          additionalInfo = {
            monthDay: parseInt(triggerData.monthDay) || 1,
            hour: parseInt(triggerData.hour) || 8,
            minute: parseInt(triggerData.minute) || 0
          };
        } else if (frequency === 'weekly' && (!triggerData.weekDays || triggerData.weekDays.length === 0)) {
          // Default Monday case
          additionalInfo = {
            weekDay: 1, // Monday
            hour: parseInt(triggerData.hour) || 8,
            minute: parseInt(triggerData.minute) || 0
          };
        }
        
        saveSheetInfoForTrigger(triggerId, triggerData.sheetName, frequency, additionalInfo);
        
        // Get trigger info for UI update
        const triggerInfo = this.getTriggerInfo(trigger);
        
        // Log team activity if user is in a team
        if (typeof logTeamActivity === 'function') {
          logTeamActivity('trigger', `created ${frequency} sync trigger`, {
            sheetName: triggerData.sheetName,
            frequency: frequency
          });
        }
        
        return {
          success: true,
          message: 'Trigger created successfully',
          triggerId: triggerId,
          triggerInfo: {
            id: triggerId,
            type: triggerInfo.type,
            description: triggerInfo.description
          }
        };
      } else {
        return { success: false, error: 'Failed to create trigger' };
      }
    } catch (e) {
      return { success: false, error: e.message };
    }
  },
  
  deleteTrigger(triggerId) {
    try {
      
      // Find the trigger by ID
      const allTriggers = ScriptApp.getProjectTriggers();
      let triggerToDelete = null;
      
      for (const trigger of allTriggers) {
        if (trigger.getUniqueId() === triggerId) {
          triggerToDelete = trigger;
          break;
        }
      }
      
      if (!triggerToDelete) {
        return { success: false, error: 'Trigger not found' };
      }
      
      // Delete the trigger
      ScriptApp.deleteTrigger(triggerToDelete);
      
      // Delete properties related to this trigger
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_SHEET`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_FREQUENCY`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_WEEKDAY`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_HOUR`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MINUTE`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MONTHDAY`);
      scriptProperties.deleteProperty(`TRIGGER_${triggerId}_MINUTES_INTERVAL`);
      
      // Log team activity if user is in a team
      if (typeof logTeamActivity === 'function') {
        logTeamActivity('trigger', 'deleted sync trigger', {
          triggerId: triggerId
        });
      }
      
      return { success: true, message: 'Trigger deleted successfully' };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }
};

// Helper function to get ScriptApp.WeekDay from number
function getWeekDayFromNumber(dayNum) {
  switch (parseInt(dayNum)) {
    case 1: return ScriptApp.WeekDay.MONDAY;
    case 2: return ScriptApp.WeekDay.TUESDAY;
    case 3: return ScriptApp.WeekDay.WEDNESDAY;
    case 4: return ScriptApp.WeekDay.THURSDAY;
    case 5: return ScriptApp.WeekDay.FRIDAY;
    case 6: return ScriptApp.WeekDay.SATURDAY;
    case 7: return ScriptApp.WeekDay.SUNDAY;
    default: return ScriptApp.WeekDay.MONDAY;
  }
}

// Helper function to save sheet info for a trigger
function saveSheetInfoForTrigger(triggerId, sheetName, frequency, additionalInfo) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Build properties object to set all at once
  const props = {};
  props[`TRIGGER_${triggerId}_SHEET`] = sheetName;
  props[`TRIGGER_${triggerId}_FREQUENCY`] = frequency;
  
  // Add additional info if provided
  if (additionalInfo) {
    if (additionalInfo.weekDay !== undefined) {
      props[`TRIGGER_${triggerId}_WEEKDAY`] = String(additionalInfo.weekDay);
    }
    if (additionalInfo.hour !== undefined) {
      props[`TRIGGER_${triggerId}_HOUR`] = String(additionalInfo.hour);
    }
    if (additionalInfo.minute !== undefined) {
      props[`TRIGGER_${triggerId}_MINUTE`] = String(additionalInfo.minute);
    }
    if (additionalInfo.monthDay !== undefined) {
      props[`TRIGGER_${triggerId}_MONTHDAY`] = String(additionalInfo.monthDay);
    }
    if (additionalInfo.minutesInterval !== undefined) {
      props[`TRIGGER_${triggerId}_MINUTES_INTERVAL`] = String(additionalInfo.minutesInterval);
    }
  }
  
  // Set all properties at once
  scriptProperties.setProperties(props);
  
}

// Helper function to capitalize first letter
function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

// Export functions to be globally accessible
this.getTriggersForSheet = TriggerManagerUI.getTriggersForSheet;
this.getTriggerInfo = TriggerManagerUI.getTriggerInfo; 
this.createSyncTrigger = TriggerManagerUI.createSyncTrigger;
this.deleteTrigger = TriggerManagerUI.deleteTrigger;

/**
 * Shows the trigger manager dialog - global function callable from client-side
 * This is the main entry point for showing the trigger manager UI.
 */
function showTriggerManager() {
  return TriggerManagerUI.showTriggerManager();
}

/**
 * Creates a sync trigger - global function callable from client-side
 * This is the newer version that properly saves weekday/time info
 */
function createSyncTriggerV2(triggerData) {
  return TriggerManagerUI.createSyncTrigger(triggerData);
}

/**
 * Deletes a trigger - global function callable from client-side
 */
function deleteTrigger(triggerId) {
  return TriggerManagerUI.deleteTrigger(triggerId);
} 