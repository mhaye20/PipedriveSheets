/**
 * User Interface
 * 
 * This module handles all UI-related functions:
 * - Showing dialogs and sidebars
 * - Building UI components
 * - Managing user interactions
 */

// Create UI namespace if it doesn't exist
var UI = UI || {};

/**
 * Shows settings dialog where users can configure filter ID and entity type
 */
function showSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();
  
  // Get current settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Check if we're authenticated
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  if (!accessToken) {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Not Connected',
      'You need to connect to Pipedrive before configuring settings. Connect now?',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      showAuthorizationDialog();
    }
    return;
  }
  
  const savedSubdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  
  // Get sheet-specific settings using the sheet name
  const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  
  const savedFilterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const savedEntityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Create a template from the HTML file
  const template = HtmlService.createTemplateFromFile('SettingsDialog');
  
  // Pass data to the template
  template.activeSheetName = activeSheetName;
  template.savedSubdomain = savedSubdomain;
  template.savedFilterId = savedFilterId;
  template.savedEntityType = savedEntityType;
  
  // Create the HTML output from the template
  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setHeight(520)
    .setTitle(`Pipedrive Settings for "${activeSheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
}

/**
 * Shows the column selector dialog
 */
function showColumnSelector() {
  try {
    // Direct the function to the new implementation
    showColumnSelectorUI();
  } catch (e) {
    Logger.log(`Error in showColumnSelector: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the sync status dialog
 * @param {string} sheetName - Sheet name
 */
function showSyncStatus(sheetName) {
  try {
    // Get the current entity type for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
    
    // Initialize sync status properties
    scriptProperties.setProperty('SYNC_PHASE_1_STATUS', 'active');
    scriptProperties.setProperty('SYNC_PHASE_1_DETAIL', 'Connecting to Pipedrive...');
    scriptProperties.setProperty('SYNC_PHASE_1_PROGRESS', '50');
    
    scriptProperties.setProperty('SYNC_PHASE_2_STATUS', 'pending');
    scriptProperties.setProperty('SYNC_PHASE_2_DETAIL', 'Waiting to start...');
    scriptProperties.setProperty('SYNC_PHASE_2_PROGRESS', '0');
    
    scriptProperties.setProperty('SYNC_PHASE_3_STATUS', 'pending');
    scriptProperties.setProperty('SYNC_PHASE_3_DETAIL', 'Waiting to start...');
    scriptProperties.setProperty('SYNC_PHASE_3_PROGRESS', '0');
    
    scriptProperties.setProperty('SYNC_CURRENT_PHASE', '1');
    scriptProperties.setProperty('SYNC_COMPLETED', 'false');
    scriptProperties.setProperty('SYNC_ERROR', '');
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('SyncStatus');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.entityType = entityType;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Syncing Data')
      .setWidth(400)
      .setHeight(300);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Syncing Data');
  } catch (e) {
    Logger.log(`Error in showSyncStatus: ${e.message}`);
    // Just log the error but don't throw, as this is a UI enhancement
  }
}

/**
 * Includes an HTML file with the content of a script file.
 * @param {string} filename - The name of the file to include without the extension
 * @return {string} The content to be included
 */
function include(filename) {
  return HtmlService.createHtmlOutput(getFileContent_(filename)).getContent();
}

/**
 * Gets the content of a script file.
 * @param {string} filename - The name of the file without the extension
 * @return {string} The content of the file
 * @private
 */
function getFileContent_(filename) {
  if (filename === 'TeamManagerUI') {
    return `<style>${TeamManagerUI.getStyles()}</style><script>${TeamManagerUI.getScripts()}</script>`;
  } else if (filename === 'TriggerManagerUI') {
    return `<style>${TriggerManagerUI.getStyles()}</style><script>${TriggerManagerUI.getScripts()}</script>`;
  } else if (filename === 'TwoWaySyncSettingsUI') {
    return `<style>${TwoWaySyncSettingsUI.getStyles()}</style><script>${TwoWaySyncSettingsUI.getScripts()}</script>`;
  }
  return '';
}

/**
 * Shows the team management UI.
 * @param {boolean} joinOnly - Whether to show only the join team section.
 */
function showTeamManager(joinOnly = false) {
  try {
    // Get the active user's email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to retrieve your email address. Please make sure you are logged in.');
    }

    // Get team data
    const teamAccess = new TeamAccess();
    const hasTeam = teamAccess.isUserInTeam(userEmail);
    let teamName = '';
    let teamId = '';
    let teamMembers = [];
    let userRole = '';

    if (hasTeam) {
      const teamData = teamAccess.getUserTeamData(userEmail);
      teamName = teamData.name;
      teamId = teamData.id;
      teamMembers = teamAccess.getTeamMembers(teamId);
      userRole = teamData.role;
    }

    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('TeamManager');
    
    // Set template variables
    template.userEmail = userEmail;
    template.hasTeam = hasTeam;
    template.teamName = teamName;
    template.teamId = teamId;
    template.teamMembers = teamMembers;
    template.userRole = userRole;
    template.initialTab = joinOnly ? 'join' : (hasTeam ? 'manage' : 'create');
    
    // Make include function available to the template
    template.include = include;
    
    // Evaluate the template
    const htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(hasTeam ? 600 : 400)
      .setTitle(hasTeam ? 'Team Management' : 'Team Access')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Team Access');
  } catch (error) {
    Logger.log(`Error in showTeamManager: ${error.message}`);
    showError('An error occurred while loading the team management interface: ' + error.message);
  }
}

/**
 * Shows an error message to the user in a dialog
 * @param {string} message - The error message to display
 */
function showError(message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', message, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log('Error showing error dialog: ' + error.message);
  }
}

/**
 * Process a template string with PHP-like syntax using a data object
 * @param {string} template - Template string with PHP-like syntax
 * @param {Object} data - Data object containing values to substitute
 * @return {string} - Processed template
 */
function processTemplate(template, data) {
  let processedTemplate = template;
  
  // Handle conditionals
  processedTemplate = processIfElse(processedTemplate, data);
  
  // Handle foreach loops
  processedTemplate = processForEach(processedTemplate, data);
  
  // Process simple variable substitutions like <?= variable ?>
  processedTemplate = processedTemplate.replace(/\<\?=\s*([^?>]+?)\s*\?>/g, (match, variable) => {
    try {
      // Handle nested properties like member.email or complex expressions
      const value = evalInContext(variable, data) || '';
      return value;
    } catch (e) {
      Logger.log('Error processing variable: ' + e.message);
      return '';
    }
  });
  
  return processedTemplate;
}

/**
 * Process if/else statements in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
function processIfElse(template, data) {
  let result = template;
  
  // Handle if statements: <?php if (condition): ?>content<?php endif; ?>
  const IF_PATTERN = /\<\?php\s+if\s+\((.+?)\)\s*:\s*\?>([\s\S]*?)(?:\<\?php\s+else\s*:\s*\?>([\s\S]*?))?\<\?php\s+endif;\s*\?>/g;
  
  result = result.replace(IF_PATTERN, (match, condition, ifContent, elseContent = '') => {
    try {
      // Convert PHP-like condition to JavaScript
      let jsCondition = condition
        .replace(/\!\=/g, '!==')
        .replace(/\=\=/g, '===')
        .replace(/\!([\w\.]+)/g, '!$1');
      
      // Evaluate the condition in the context of the data object
      const conditionResult = evalInContext(jsCondition, data);
      
      return conditionResult ? ifContent : elseContent;
    } catch (e) {
      Logger.log('Error processing if condition: ' + e.message);
      return '';
    }
  });
  
  return result;
}

/**
 * Process foreach loops in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
function processForEach(template, data) {
  let result = template;
  
  // Handle foreach loops: <?php foreach ($items as $item): ?>content<?php endforeach; ?>
  const FOREACH_PATTERN = /\<\?php\s+foreach\s+\(\$(\w+)\s+as\s+\$(\w+)\)\s*:\s*\?>([\s\S]*?)\<\?php\s+endforeach;\s*\?>/g;
  
  result = result.replace(FOREACH_PATTERN, (match, collection, item, content) => {
    try {
      const items = data[collection];
      if (!Array.isArray(items) || items.length === 0) {
        return '';
      }
      
      return items.map(itemData => {
        // Create a context with the item for nested variable replacement
        const itemContext = Object.assign({}, data, { [item]: itemData });
        
        // Replace item variables within the loop content
        let itemContent = content;
        
        // Process variables like <?= item.property ?>
        itemContent = itemContent.replace(/\<\?=\s*(\$?)([\w\.]+)\s*\?>/g, (m, dollar, varName) => {
          try {
            // If it's a loop item variable like $member or $member.property
            if (varName.startsWith(item + '.')) {
              const propPath = varName.substring(item.length + 1);
              return evalPropertyPath(itemData, propPath) || '';
            } 
            // If it's just $member (the whole item)
            else if (varName === item) {
              return itemData || '';
            } 
            // Other variables from the parent context
            else {
              return evalInContext(varName, itemContext) || '';
            }
          } catch (e) {
            Logger.log('Error in foreach variable replacement: ' + e.message);
            return '';
          }
        });
        
        return itemContent;
      }).join('');
    } catch (e) {
      Logger.log('Error processing foreach: ' + e.message);
      return '';
    }
  });
  
  return result;
}

/**
 * Evaluate a JavaScript expression in the context of a data object
 * @param {string} expr - The expression to evaluate
 * @param {Object} context - The context object
 * @return {*} - The result of the evaluation
 */
function evalInContext(expr, context) {
  try {
    // Handle simple variable access first
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(expr)) {
      return context[expr];
    }
    
    // Handle nested properties with dot notation
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*(\.[a-zA-Z_$][a-zA-Z0-9_$]*)+$/.test(expr)) {
      return evalPropertyPath(context, expr);
    }
    
    // Handle comparisons and more complex expressions
    // Create a safe function to evaluate in the context
    const keys = Object.keys(context);
    const values = keys.map(key => context[key]);
    const evaluator = new Function(...keys, `return ${expr};`);
    return evaluator(...values);
  } catch (e) {
    Logger.log('Error evaluating expression: ' + e.message);
    return null;
  }
}

/**
 * Evaluate a property path on an object (e.g. "user.profile.name")
 * @param {Object} obj - The object to evaluate on
 * @param {string} path - The property path
 * @return {*} - The value at the path or undefined
 */
function evalPropertyPath(obj, path) {
  try {
    return path.split('.').reduce((o, p) => o && o[p], obj);
  } catch (e) {
    Logger.log('Error evaluating property path: ' + e.message);
    return undefined;
  }
}

/**
 * Shows the two-way sync settings dialog
 */
function showTwoWaySyncSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();
  
  // Get current settings
  const scriptProperties = PropertiesService.getScriptProperties();
  const settingsKey = `TWOWAY_SYNC_SETTINGS_${activeSheetName}`;
  const settingsJson = scriptProperties.getProperty(settingsKey) || '{}';
  const settings = JSON.parse(settingsJson);
  
  // Get sheet-specific entity type
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Create the HTML template
  const template = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
  
  // Pass data to template
  template.sheetName = activeSheetName;
  template.settings = settings;
  template.entityType = entityType;
  
  // Make include function available to the template
  template.include = include;
  
  // Create and show dialog
  const html = template.evaluate()
    .setWidth(600)
    .setHeight(700)
    .setTitle('Two-Way Sync Settings');
      
  SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
}

function showTriggerManager() {
  TriggerManagerUI.showTriggerManager();
}

/**
 * Gets all triggers configured for a specific sheet
 * @param {string} sheetName - The name of the sheet to get triggers for
 * @return {Array} Array of trigger objects with id, type, and description
 */
function getTriggersForSheet(sheetName) {
  try {
    // Get all triggers for the spreadsheet
    const triggers = ScriptApp.getProjectTriggers();
    
    // Filter triggers for syncFromPipedrive function and this sheet
    return triggers
      .filter(trigger => {
        // Only include sync triggers
        if (trigger.getHandlerFunction() !== 'syncFromPipedrive') {
          return false;
        }
        
        // Get trigger event source (the sheet)
        const triggerSource = trigger.getTriggerSource();
        if (triggerSource === null) {
          return false;
        }
        
        // For time-based triggers, check if they're for this sheet
        if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
          // We store the sheet name in the trigger's unique ID
          const triggerId = trigger.getUniqueId();
          return triggerId.includes(sheetName);
        }
        
        return false;
      })
      .map(trigger => {
        const info = getTriggerInfo(trigger);
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
}

/**
 * Gets human-readable information about a trigger
 * @param {Trigger} trigger - The trigger to get info for
 * @return {Object} Object with type and description
 */
function getTriggerInfo(trigger) {
  try {
    const eventType = trigger.getEventType();
    
    if (eventType === ScriptApp.EventType.CLOCK) {
      const atTime = trigger.getAtHour() !== null;
      const everyHours = !atTime;
      
      if (atTime) {
        // Daily/weekly/monthly trigger
        const hour = trigger.getAtHour();
        const minute = trigger.getAtMinute() || 0;
        const weekDay = trigger.getWeekDay();
        const monthDay = trigger.getMonthDay();
        
        const time = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
        
        if (weekDay !== null) {
          // Weekly trigger
          const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          return {
            type: 'Weekly',
            description: `Every ${days[weekDay - 1]} at ${time}`
          };
        } else if (monthDay !== null) {
          // Monthly trigger
          const suffix = ['st', 'nd', 'rd'][monthDay - 1] || 'th';
          return {
            type: 'Monthly',
            description: `On the ${monthDay}${suffix} at ${time}`
          };
        } else {
          // Daily trigger
          return {
            type: 'Daily',
            description: `Every day at ${time}`
          };
        }
      } else if (everyHours) {
        // Hourly trigger
        const hours = trigger.getAtHour() || 1;
        return {
          type: 'Hourly',
          description: `Every ${hours} hour${hours > 1 ? 's' : ''}`
        };
      }
    }
    
    return {
      type: 'Unknown',
      description: 'Unknown trigger type'
    };
  } catch (e) {
    Logger.log(`Error in getTriggerInfo: ${e.message}`);
    return {
      type: 'Error',
      description: 'Error getting trigger info'
    };
  }
}

/**
 * Shows the column selector UI
 */
function showColumnSelectorUI() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get sheet-specific settings
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;

    const filterId = scriptProperties.getProperty(sheetFilterIdKey);
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Check if we can connect to Pipedrive
    SpreadsheetApp.getActiveSpreadsheet().toast(`Connecting to Pipedrive to retrieve ${entityType} data...`);

    // Get sample data based on entity type (1 item only)
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
      case ENTITY_TYPES.PRODUCTS:
        sampleData = getProductsWithFilter(filterId, 1);
        break;
    }

    if (!sampleData || sampleData.length === 0) {
      throw new Error(`No ${entityType} data found. Please check your filter settings.`);
    }

    const sampleItem = sampleData[0];
    const availableColumns = [];

    // Get field mappings for this entity type
    const fieldMap = getCustomFieldMappings(entityType);

    // Extract fields from sample data
    function extractFields(obj, parentKey = '', parentName = '') {
      for (const key in obj) {
        if (key.startsWith('_') || typeof obj[key] === 'function') continue;

        const fullKey = parentKey ? `${parentKey}.${key}` : key;
        const displayName = parentName ? `${parentName} - ${key}` : (fieldMap[key] || formatColumnName(key));

        if (typeof obj[key] === 'object' && obj[key] !== null) {
          // Add parent field
          availableColumns.push({
            key: fullKey,
            name: displayName,
            isNested: !!parentKey,
            parentKey: parentKey || null
          });

          // Recursively extract nested fields
          extractFields(obj[key], fullKey, displayName);
        } else {
          availableColumns.push({
            key: fullKey,
            name: displayName,
            isNested: !!parentKey,
            parentKey: parentKey || null
          });
        }
      }
    }

    // Build available columns list
    extractFields(sampleItem);

    // Get previously saved column preferences
    const columnSettingsKey = `COLUMNS_${activeSheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let selectedColumns = [];

    if (savedColumnsJson) {
      try {
        selectedColumns = JSON.parse(savedColumnsJson);
        selectedColumns = selectedColumns.filter(col =>
          availableColumns.some(availCol => availCol.key === col.key)
        );
      } catch (e) {
        Logger.log('Error parsing saved columns: ' + e.message);
        selectedColumns = [];
      }
    }

    // Create the data script that will be injected into the HTML template
    const dataScript = `<script>
      const availableColumns = ${JSON.stringify(availableColumns)};
      const selectedColumns = ${JSON.stringify(selectedColumns)};
      const entityType = "${entityType}";
      const sheetName = "${activeSheetName}";
      const entityTypeName = "${formatEntityTypeName(entityType)}";
    </script>`;

    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('ColumnSelector');
    
    // Pass data to template
    template.dataScript = dataScript;
    template.entityTypeName = formatEntityTypeName(entityType);
    template.sheetName = activeSheetName;
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(800)
      .setHeight(600)
      .setTitle(`Select Columns for ${entityType} in "${activeSheetName}"`);
      
    SpreadsheetApp.getUi().showModalDialog(html, `Select Columns for ${entityType} in "${activeSheetName}"`);
  } catch (error) {
    Logger.log('Error in showColumnSelectorUI: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
    throw error;
  }
}

/**
 * Formats an entity type name for display
 * @param {string} entityType - The entity type to format
 * @return {string} The formatted entity type name
 */
function formatEntityTypeName(entityType) {
  if (!entityType) return '';
  
  // Remove any prefix/suffix and convert to title case
  const name = entityType.replace(/^ENTITY_TYPES\./, '').toLowerCase();
  return name.charAt(0).toUpperCase() + name.slice(1);
}

/**
 * Shows the help and about dialog
 */
function showHelp() {
  try {
    // Get current version from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const currentVersion = scriptProperties.getProperty('APP_VERSION') || '1.0.0';

    // Get version history if available
    const versionHistoryJson = scriptProperties.getProperty('VERSION_HISTORY');
    let versionHistory = [];
    
    if (versionHistoryJson) {
      try {
        versionHistory = JSON.parse(versionHistoryJson);
      } catch (e) {
        Logger.log('Error parsing version history: ' + e.message);
      }
    }

    // Create the HTML template
    const template = HtmlService.createTemplateFromFile('Help');
    
    // Pass data to template
    template.currentVersion = currentVersion;
    template.versionHistory = versionHistory;
    
    // Create and show dialog
    const html = template.evaluate()
      .setWidth(600)
      .setHeight(400)
      .setTitle('Help & About');
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Help & About');
  } catch (error) {
    Logger.log('Error in showHelp: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error showing help: ' + error.message);
  }
}

/**
 * Saves column preferences to script properties
 * @param {Array} columns - Array of column objects with key and name properties
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
function saveColumnPreferences(columns, entityType, sheetName) {
  try {
    Logger.log(`Saving column preferences for ${entityType} in sheet "${sheetName}"`);
    
    // Get current user email for user-specific preferences
    const userEmail = Session.getActiveUser().getEmail();
    
    // Store the full column objects to preserve names and other properties
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Store columns based on both entity type and sheet name for sheet-specific preferences
    const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    scriptProperties.setProperty(columnSettingsKey, JSON.stringify(columns));
    
    // Also store in user-specific property
    if (userEmail) {
      const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
      scriptProperties.setProperty(userColumnSettingsKey, JSON.stringify(columns));
      Logger.log(`Saved user-specific column preferences with key: ${userColumnSettingsKey}`);
    }
    
    // Check if two-way sync is enabled for this sheet
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
      
      Logger.log(`Removed tracking column property for sheet "${sheetName}" to ensure correct positioning on next sync.`);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveColumnPreferences: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    throw e;
  }
}