/**
 * Pipedrive to Google Sheets Integration
 * Version: 1.0.0
 * 
 * This add-on connects to Pipedrive API and fetches data based on filters
 * to populate a Google Sheet with the requested fields. It allows for two-way
 * synchronization, team collaboration, and scheduled updates.
 * 
 * @author Your Name
 * @license Apache License 2.0
 */

/**
 * Main entry point
 * 
 * This module contains:
 * - onOpen trigger
 * - Menu creation and initialization
 * - Main entry point functions
 */

/**
 * Constants
 */
const API_KEY = ''; // Default empty, user will provide
const FILTER_ID = ''; // Default filter ID
const DEFAULT_PIPEDRIVE_SUBDOMAIN = '';
const DEFAULT_SHEET_NAME = 'Pipedrive Data';
const PIPEDRIVE_API_URL_PREFIX = 'https://';
const PIPEDRIVE_API_URL_SUFFIX = '.pipedrive.com/api/v2';
const ENTITY_TYPES = {
  DEALS: 'deals',
  PERSONS: 'persons',
  ORGANIZATIONS: 'organizations',
  ACTIVITIES: 'activities',
  LEADS: 'leads',
  PRODUCTS: 'products'
};

/**
 * Cache for verified users
 */
let VERIFIED_USERS = {};

/**
 * Field definitions cache
 */
const fieldDefinitionsCache = {};

/**
 * Runs when the sheet is opened
 * Creates the Pipedrive menu and initializes the application
 */
function onOpen() {
  try {
    // Create a custom menu in Google Sheets
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Pipedrive', [
      {name: 'Initialize Pipedrive', functionName: 'initializePipedriveMenu'},
    ]);
  } catch (e) {
    Logger.log(`Error in onOpen: ${e.message}`);
  }
}

/**
 * Initializes the Pipedrive menu
 * Always the first function that runs when user clicks the menu
 */
function initializePipedriveMenu() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to determine your email address. Please ensure you are signed in.');
    }
    
    Logger.log('Initializing Pipedrive menu for user: ' + userEmail);
    
    // First check if the user is the script installer/owner
    let isScriptOwner = false;
    try {
      const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
      isScriptOwner = (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED);
      if (isScriptOwner) {
        Logger.log('User is the script owner/installer with full authorization');
      }
    } catch (e) {
      Logger.log('Error checking script owner: ' + e.message);
    }
    
    // Check if user is in any team - getting a direct team check
    // instead of using checkAnyUserAccess which might return true for script owners
    // even if they don't have a team yet
    const userTeam = getUserTeam(userEmail);
    const hasTeamAccess = userTeam !== null;
    
    // If user is the script owner and doesn't have a team yet, create one automatically
    if (isScriptOwner && !hasTeamAccess) {
      Logger.log('Script owner detected without a team. Auto-creating team for ' + userEmail);
      const teamName = 'My Pipedrive Team';
      const result = createTeam(teamName);
      
      if (result.success) {
        Logger.log('Successfully created team for script owner: ' + result.teamId);
        // Team created successfully, now create the full menu
        createPipedriveMenu();
        
        // Show success message
        const html = HtmlService.createHtmlOutput(`
          <p>Welcome to Pipedrive for Sheets!</p>
          <p>A team has been automatically created for you with ID: <strong>${result.teamId}</strong></p>
          <p>Share this Team ID with your colleagues to let them join your team.</p>
          <script>
            setTimeout(function() {
              google.script.host.close();
            }, 5000);
          </script>
        `)
        .setWidth(400)
        .setHeight(150);
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Team Created');
        
        return true;
      } else {
        Logger.log('Failed to create team for script owner: ' + result.message);
      }
    }
    
    // After potential team creation, check team access again
    const finalTeamAccess = getUserTeam(userEmail) !== null;
    
    // If user has team access, show regular menu
    if (finalTeamAccess) {
      // User has team access, create the full menu
      createPipedriveMenu();
      
      // Show success message
      const html = HtmlService.createHtmlOutput(`
        <p>Pipedrive menu has been initialized successfully.</p>
        <script>
          setTimeout(function() {
            google.script.host.close();
          }, 1500);
        </script>
      `)
      .setWidth(300)
      .setHeight(80);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Pipedrive Ready');
      
      return true;
    } else {
      // User is not in a team, show join team dialog
      Logger.log('User not in any team, showing join team request');
      createBasicMenu();
      showTeamJoinRequest();
      return false;
    }
  } catch (e) {
    Logger.log('Error in initializePipedriveMenu: ' + e.message);
    
    // Show error message
    SpreadsheetApp.getUi().alert(
      'Initialization Error',
      'Failed to initialize Pipedrive menu: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Fallback to basic menu on error
    try {
      createBasicMenu();
    } catch (menuError) {
      Logger.log('Error creating basic menu: ' + menuError.message);
    }
    
    return false;
  }
}

/**
 * Clears all team data for testing purposes
 * CAUTION: This will remove all teams and team memberships
 */
function clearTeamDataForTesting() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get confirmation from user
    const response = ui.alert(
      'Clear All Team Data',
      'This will permanently delete ALL team data. This action cannot be undone.\n\n' + 
      'Are you sure you want to proceed?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'No data was deleted.', ui.ButtonSet.OK);
      return;
    }
    
    // Clear the team data
    PropertiesService.getDocumentProperties().deleteProperty('TEAMS_DATA');
    PropertiesService.getDocumentProperties().deleteProperty('EMAIL_TO_TEAM_MAP');
    
    // Clear verified users
    PropertiesService.getDocumentProperties().deleteProperty('VERIFIED_USER_IDS');
    PropertiesService.getDocumentProperties().deleteProperty('VERIFIED_TEAM_USERS');
    
    Logger.log('All team data has been cleared');
    
    // Show confirmation
    ui.alert(
      'Data Cleared',
      'All team data has been successfully deleted. Please reload the page to see the changes.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Error in clearTeamDataForTesting: ' + e.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to clear team data: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Creates a basic menu with limited options
 */
function createBasicMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Pipedrive');
  
  menu.addItem('Settings', 'showSettings')
      .addItem('Join Team', 'showTeamJoinRequest')
      .addItem('Force Reauthorize', 'forceReauthorize')
      .addItem('Help & About', 'showHelp')
      .addSeparator()
      .addItem('🧪 Test First Installation', 'testFirstInstallation')
      .addItem('🧨 Clear Team Data', 'clearTeamDataForTesting');
      
  menu.addToUi();
}

/**
 * Creates the full Pipedrive menu with all options
 */
function createPipedriveMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Pipedrive');
  
  menu.addItem('🔄 Sync Data', 'syncFromPipedrive')
      .addItem('⬆️ Push Changes to Pipedrive', 'pushChangesToPipedrive')
      .addSeparator()
      .addItem('📊 Select Columns', 'showColumnSelector')
      .addItem('⚙️ Filter Settings', 'showSettings')
      .addItem('🔁 Two-Way Sync Settings', 'showTwoWaySyncSettings')
      .addItem('👥 Team Management', 'showTeamManager')
      .addSeparator()
      .addItem('⏱️ Schedule Sync', 'showTriggerManager')
      .addSeparator()
      .addItem('Force Reauthorize', 'forceReauthorize')
      .addItem('ℹ️ Help & About', 'showHelp');
      
  menu.addToUi();
}

/**
 * Checks if a user has any type of access
 * @param {string} userEmail The email address to check
 * @returns {boolean} True if the user has access, false otherwise
 */
function checkAnyUserAccess(userEmail) {
  try {
    if (!userEmail) return false;
    
    // Normalize email for case-insensitive comparison
    const normalizedEmail = userEmail.toLowerCase();
    
    // Check if we've already verified this user in this session
    if (VERIFIED_USERS[normalizedEmail]) {
      return true;
    }
    
    // Try fast path using email-to-team map
    const emailMapStr = PropertiesService.getDocumentProperties().getProperty('EMAIL_TO_TEAM_MAP');
    if (emailMapStr) {
      try {
        const emailMap = JSON.parse(emailMapStr);
        if (emailMap[normalizedEmail]) {
          // Save to verified users cache
          VERIFIED_USERS[normalizedEmail] = true;
          return true;
        }
      } catch (e) {
        Logger.log('Error parsing email map: ' + e.message);
      }
    }
    
    // Slower path - check if user is in any team
    if (isUserInTeam(userEmail)) {
      // Save to verified users cache
      VERIFIED_USERS[normalizedEmail] = true;
      return true;
    }
    
    Logger.log('User ' + userEmail + ' does not have access');
    return false;
  } catch (e) {
    Logger.log('Error in checkAnyUserAccess: ' + e.message);
    return false;
  }
}

/**
 * Forces a check of team membership for a user
 * @param {string} userEmail Email address of the user
 * @returns {boolean} True if the user is in a team, false otherwise
 */
function forceTeamMembershipCheck(userEmail) {
  try {
    if (!userEmail) return false;
    
    // Try fast path using email-to-team map
    const emailMapStr = PropertiesService.getDocumentProperties().getProperty('EMAIL_TO_TEAM_MAP');
    if (emailMapStr) {
      try {
        const emailMap = JSON.parse(emailMapStr);
        if (emailMap[userEmail.toLowerCase()]) {
          // User is in a team - create the full menu
          createPipedriveMenu();
          
          // Add to verified users cache
          VERIFIED_USERS[userEmail.toLowerCase()] = true;
          
          return true;
        }
      } catch (e) {
        Logger.log('Error parsing email map: ' + e.message);
      }
    }
    
    // Try checking if user is in any team directly
    if (isUserInTeam(userEmail)) {
      // User is in a team - create the full menu
      createPipedriveMenu();
      
      // Add to verified users cache
      VERIFIED_USERS[userEmail.toLowerCase()] = true;
      
      return true;
    }
    
    // User is not in a team - show the basic menu and team join dialog
    createBasicMenu();
    showTeamJoinRequest();
    
    return false;
  } catch (e) {
    Logger.log('Error in forceTeamMembershipCheck: ' + e.message);
    createBasicMenu();
    return false;
  }
}

/**
 * Checks if the user has verified team access
 * @return {boolean} True if the user has verified access, false otherwise
 */
function hasVerifiedTeamAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return false;
    }
    
    return checkAnyUserAccess(userEmail);
  } catch (e) {
    Logger.log(`Error in hasVerifiedTeamAccess: ${e.message}`);
    return false;
  }
}

/**
 * Preloads verified users for faster access checks
 */
function preloadVerifiedUsers() {
  try {
    // Clear the cache
    VERIFIED_USERS = {};
    
    // Get the email-to-team map
    const docProps = PropertiesService.getDocumentProperties();
    const emailMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
    
    if (emailMapStr) {
      const emailMap = JSON.parse(emailMapStr);
      
      // Mark all users in the map as verified
      for (const email in emailMap) {
        VERIFIED_USERS[email.toLowerCase()] = true;
      }
      
      Logger.log('Teams data preloaded successfully');
    }
  } catch (e) {
    Logger.log(`Error in preloadVerifiedUsers: ${e.message}`);
  }
}

/**
 * Verifies team access and refreshes the menu accordingly
 */
function verifyTeamAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return false;
    }
    
    // Force a fresh check by ignoring the cached result
    const normalizedEmail = userEmail.toLowerCase();
    delete VERIFIED_USERS[normalizedEmail];
    
    // Check access again
    if (checkAnyUserAccess(userEmail)) {
      // User has access, refresh the menu
      createPipedriveMenu();
      return true;
    } else {
      // User doesn't have access, show the basic menu
      createBasicMenu();
      return false;
    }
  } catch (e) {
    Logger.log(`Error in verifyTeamAccess: ${e.message}`);
    return false;
  }
}

/**
 * Refreshes the menu after team verification
 */
function refreshMenuAfterVerification() {
  try {
    if (hasVerifiedTeamAccess()) {
      createPipedriveMenu();
    } else {
      createBasicMenu();
    }
  } catch (e) {
    Logger.log(`Error in refreshMenuAfterVerification: ${e.message}`);
  }
}

/**
 * Force a reauthorization of the script with all required permissions
 */
function forceReauthorize() {
  try {
    // First try to get user info to verify authentication
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log(`Current user: ${userEmail}`);
    
    // Create HTML with a button that explicitly requests the permission
    const html = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
          text-align: center;
        }
        h3 {
          color: #4285F4;
        }
        p {
          margin: 15px 0;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 10px 20px;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        }
        button:hover {
          background-color: #3367D6;
        }
        .success {
          color: #0f9d58;
          font-weight: bold;
        }
        .error {
          color: #d32f2f;
          font-weight: bold;
        }
      </style>
      <h3>Script Authorization</h3>
      <p>This add-on requires additional permissions to function correctly.</p>
      <p>Click the button below to grant the necessary permissions:</p>
      <button id="authorize-button">Authorize Add-on</button>
      <p id="status"></p>
      
      <script>
        // Authorization button handler
        document.getElementById('authorize-button').addEventListener('click', function() {
          document.getElementById('status').textContent = 'Requesting permission...';
          
          // Request permissions by calling a function that uses them
          google.script.run
            .withSuccessHandler(function(result) {
              document.getElementById('status').className = 'success';
              document.getElementById('status').textContent = 'Authorization successful! Please refresh the page.';
              
              // Disable the button
              document.getElementById('authorize-button').disabled = true;
            })
            .withFailureHandler(function(error) {
              document.getElementById('status').className = 'error';
              document.getElementById('status').textContent = 'Error: ' + error.message;
            })
            .testUIPermission();
        });
      </script>
    `)
    .setWidth(400)
    .setHeight(300);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Authorization Required');
  } catch (e) {
    Logger.log(`Error in forceReauthorize: ${e.message}`);
    
    // Show error to the user
    SpreadsheetApp.getUi().alert(
      'Authorization Error',
      'Could not show authorization dialog: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Tests UI permission by showing a small dialog
 * Used for authorization flow
 */
function testUIPermission() {
  // Show a small dialog to test UI permission
  const ui = SpreadsheetApp.getUi();
  const testHtml = HtmlService.createHtmlOutput('<p>Permissions test successful!</p>')
    .setWidth(200)
    .setHeight(50);
  
  ui.showModalDialog(testHtml, 'Permission Test');
  return true;
}

/**
 * Refreshes the menu after joining a team
 * This function is called after successful team join/create operations
 */
function refreshMenuAfterJoin() {
  try {
    // Force a fresh verification
    verifyTeamAccess();
    
    // Show confirmation
    const html = HtmlService.createHtmlOutput(`
      <p>Team access verified. Initializing Pipedrive menu...</p>
      <script>
        setTimeout(function() {
          window.top.location.reload();
        }, 1500);
      </script>
    `)
    .setWidth(300)
    .setHeight(80);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Access Granted');
  } catch (e) {
    Logger.log(`Error in refreshMenuAfterJoin: ${e.message}`);
  }
}

/**
 * Fixes the menu after joining a team
 * @returns {boolean} True if successful, false otherwise
 */
function fixMenuAfterJoin() {
  try {
    // First reapply team verification
    verifyTeamAccess();
    
    // Create the Pipedrive menu
    createPipedriveMenu();
    
    // Create a toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Team access granted successfully! Menu has been updated.',
      'Success',
      5
    );
    
    return true;
  } catch (e) {
    Logger.log('Error in fixMenuAfterJoin: ' + e.message);
    return false;
  }
}

/**
 * Test function to directly show the join team dialog
 * This can be used for testing the team join flow
 */
function testShowJoinTeamDialog() {
  try {
    // Directly show the join team dialog
    showTeamJoinRequest();
    return true;
  } catch (e) {
    Logger.log('Error in testShowJoinTeamDialog: ' + e.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to show join team dialog: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * Test function to simulate first installation and auto-team creation
 * Use this to test the team creation flow for new installers
 */
function testFirstInstallation() {
  try {
    // Get the current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to determine your email address');
    }
    
    // Get current team data
    const currentTeamData = getUserTeam(userEmail);
    
    // If user is already in a team, ask for confirmation to proceed
    if (currentTeamData) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Existing Team Found',
        `You are already a member of team "${currentTeamData.name}" with ID: ${currentTeamData.teamId}.\n\n` +
        'Do you want to test first installation anyway? This will NOT remove you from your current team.',
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        ui.alert('Test Cancelled', 'First installation test was cancelled.', ui.ButtonSet.OK);
        return;
      }
    }
    
    // Force script owner check to return true for testing
    const html = HtmlService.createHtmlOutput(`
      <p>Testing first installation process...</p>
      <p>Email: ${userEmail}</p>
      <div id="status">Running initialization...</div>
      
      <script>
        // Simulate first-time installer flow
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('status').innerHTML = 
              'Initialization complete.<br><br>' +
              (result.teamId ? 
                'Team created successfully with ID: <strong>' + result.teamId + '</strong>' :
                'No new team was created. See logs for details.');
          })
          .withFailureHandler(function(error) {
            document.getElementById('status').innerHTML = 
              'Error: ' + error.message;
          })
          .testCreateTeamForOwner();
      </script>
    `)
    .setWidth(400)
    .setHeight(200);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Testing First Installation');
  } catch (e) {
    Logger.log('Error in testFirstInstallation: ' + e.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to run test: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function for testing team creation for owner
 * @return {Object} Result of team creation
 */
function testCreateTeamForOwner() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to determine user email');
    }
    
    Logger.log('Creating test team for user: ' + userEmail);
    
    // For testing purposes, we'll create a team directly regardless of current status
    // This simulates the behavior that happens during first installation
    
    // Generate a team ID
    const teamId = Utilities.getUuid();
    
    // Get or initialize teams data
    const teamsData = getTeamsData() || {};
    
    // Create the new team entry
    teamsData[teamId] = {
      name: 'Test Owner Team',
      createdBy: userEmail,
      createdAt: new Date().toISOString(),
      adminEmails: [userEmail],
      memberEmails: [userEmail],
      settings: {
        shareFilters: true,
        shareColumns: true
      }
    };
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      
      // Return success with team information
      return {
        success: true,
        teamId: teamId,
        message: 'Test team created successfully'
      };
    } else {
      return { 
        success: false, 
        message: 'Failed to save team data'
      };
    }
  } catch (e) {
    Logger.log('Error in testCreateTeamForOwner: ' + e.message);
    throw e;
  }
} 