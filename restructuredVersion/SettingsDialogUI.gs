/**
 * Settings Dialog UI Helper
 * 
 * This module handles the settings dialog UI:
 * - Displaying the settings dialog
 * - Managing Pipedrive settings
 * - Handling settings persistence
 */

var SettingsDialogUI = SettingsDialogUI || {};

/**
 * Shows settings dialog where users can configure filter ID and entity type
 */
SettingsDialogUI.showSettings = function() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get current settings from properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
    
    // Only show auth dialog if we have no access token at all
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
  } catch (error) {
    Logger.log('Error showing settings: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error showing settings: ' + error.message);
  }
};

/**
 * Shows the help and about dialog
 */
SettingsDialogUI.showHelp = function() {
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
};

// Export functions to be globally accessible
this.showSettings = SettingsDialogUI.showSettings;
this.showHelp = SettingsDialogUI.showHelp; 