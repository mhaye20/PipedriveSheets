/**
 * Quick Setup Functions for PipedriveSheets
 * Provides a streamlined onboarding experience
 */

/**
 * Quick setup wizard for new users
 */
function showQuickSetup() {
  const template = HtmlService.createTemplateFromFile('QuickSetupDialog');
  
  // Check current status
  const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const hasAccess = userEmail && (checkAnyUserAccess(userEmail) || hasVerifiedTeamAccess());
  const currentPlan = PaymentService.getCurrentPlan();
  
  template.userEmail = userEmail;
  template.hasAccess = hasAccess;
  template.currentPlan = currentPlan;
  
  const html = template.evaluate()
    .setWidth(500)
    .setHeight(600);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'PipedriveSheets Setup');
}

/**
 * Complete setup and refresh menu
 */
function completeSetup() {
  try {
    // Initialize menu
    initializePipedriveMenu();
    
    // Mark setup as complete
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('SETUP_COMPLETE', 'true');
    
    return true;
  } catch (error) {
    Logger.log('Error in completeSetup: ' + error.message);
    return false;
  }
}

/**
 * Check if user needs setup
 */
function needsSetup() {
  const userProperties = PropertiesService.getUserProperties();
  const setupComplete = userProperties.getProperty('SETUP_COMPLETE');
  const hasApiKey = PropertiesService.getUserProperties().getProperty('API_KEY');
  
  return !setupComplete || !hasApiKey;
}