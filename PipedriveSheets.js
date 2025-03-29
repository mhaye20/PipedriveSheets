/**
 * Pipedrive to Google Sheets Integration
 * 
 * This script connects to Pipedrive API and fetches data based on filters
 * to populate a Google Sheet with the requested fields.
 */

// Configuration constants (can be overridden in the script properties)
const API_KEY = '57b67dc4677482d649c23bd469888994c5ffeb3e';
const FILTER_ID = ''; // Default filter ID
const DEFAULT_PIPEDRIVE_SUBDOMAIN = 'api';
const DEFAULT_SHEET_NAME = 'PDexport'; // Default sheet name
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
 * This function runs when a user opens any spreadsheet with your add-on
 */
function onOpen() {
  try {
    Logger.log(`onOpen running, creating initial menu`);
    const ui = SpreadsheetApp.getUi();
    
    // Always create a basic menu first, regardless of verification
    const menu = ui.createMenu('Pipedrive');
    
    // Add a verification check item that will then create the full menu
    menu.addItem('Initialize Pipedrive Menu', 'initializePipedriveMenu');
    menu.addToUi();
    
    // Always try to detect column shifts
    detectColumnShifts();
  } catch (e) {
    Logger.log(`Error in onOpen: ${e.message}`);
    try {
      // Maintain your existing error handling
      detectColumnShifts();
    } catch (shiftError) {
      Logger.log(`Error in detectColumnShifts: ${shiftError.message}`);
    }
  }
}

/**
 * This function will be called when the user clicks "Initialize Pipedrive Menu"
 * At this point, we'll definitely have their email and can do proper verification
 */
function initializePipedriveMenu() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log(`Initializing Pipedrive menu for user: ${userEmail}`);
    
    // Preload verified users
    preloadVerifiedUsers();
    
    // Now perform access checks
    if (checkAnyUserAccess(userEmail) || 
        hasVerifiedTeamAccess() || 
        forceTeamMembershipCheck(userEmail)) {
      
      // Replace the menu with the full Pipedrive menu
      createPipedriveMenu();
      
      // Show a brief confirmation that initialization was successful
      const html = HtmlService.createHtmlOutput(`
        <p>Pipedrive menu has been successfully initialized!</p>
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
      // Show the team access request dialog
      verifyTeamAccess();
      return false;
    }
  } catch (e) {
    Logger.log(`Error in initializePipedriveMenu: ${e.message}`);
    
    // Show error to user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'An error occurred while initializing Pipedrive menu: ' + e.message, ui.ButtonSet.OK);
    return false;
  }
}

// Keep your existing createPipedriveMenu function as is:
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
          .addItem('ℹ️ Help & About', 'showHelp');
      
  menu.addToUi();
}

function checkAnyUserAccess(userEmail) {
  try {
    if (!userEmail) return false;
    
    // APPROACH 1: Direct ownership check - keep this
    try {
      const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
      if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED) {
        Logger.log(`${userEmail} is the script owner, granting access`);
        return true;
      }
    } catch (e) {
      Logger.log(`Error checking script owner: ${e.message}`);
    }
    
    // APPROACH 2: Check email map in document properties
    try {
      const docProps = PropertiesService.getDocumentProperties();
      const emailMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
      
      if (emailMapStr) {
        const emailMap = JSON.parse(emailMapStr);
        
        // Simplified check - just use lowercase consistently
        if (emailMap[userEmail.toLowerCase()]) {
          Logger.log(`${userEmail} found in email map`);
          return true;
        }
      }
    } catch (mapError) {
      Logger.log(`Error checking email map: ${mapError.message}`);
    }
    
    // APPROACH 3: Direct check of teams data
    try {
      if (isUserInTeam(userEmail)) {
        Logger.log(`${userEmail} found in team members list`);
        return true;
      }
    } catch (teamError) {
      Logger.log(`Error checking teams data: ${teamError.message}`);
    }
    
    Logger.log(`${userEmail} not found in any access list`);
    return false;
  } catch (e) {
    Logger.log(`Error in checkAnyUserAccess: ${e.message}`);
    return false;
  }
}

function forceTeamMembershipCheck(userEmail) {
  try {
    if (!userEmail) return false;
    
    Logger.log(`Force checking team membership for: ${userEmail}`);
    
    // Try getting from document properties directly
    const docProps = PropertiesService.getDocumentProperties();
    const emailToTeamMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
    
    if (emailToTeamMapStr) {
      const emailToTeamMap = JSON.parse(emailToTeamMapStr);
      Logger.log(`Current email map: ${JSON.stringify(emailToTeamMap)}`);
      
      if (emailToTeamMap[userEmail.toLowerCase()]) {
        Logger.log(`User found in email map, creating Pipedrive menu`);
        return true;
      }
    }
    
    // Direct check of teams data
    const teamsData = getTeamsData();
    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      const memberEmails = team.memberEmails || [];
      
      // Case-insensitive check
      for (let i = 0; i < memberEmails.length; i++) {
        if (memberEmails[i].toLowerCase() === userEmail.toLowerCase()) {
          Logger.log(`User found in teams data`);
          // Update the map since it seems out of date
          updateEmailToTeamMap();
          return true;
        }
      }
    }
    
    Logger.log(`User not found in any team membership lists`);
    return false;
  } catch (e) {
    Logger.log(`Error in forceTeamMembershipCheck: ${e.message}`);
    return false;
  }
}

/**
 * Checks if a user has verified team access with improved persistence
 */
function hasVerifiedTeamAccess() {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return false;
    
    // First check if the user installed the add-on (always give access to installer)
    try {
      const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
      if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED) {
        Logger.log(`User ${userEmail} is the script owner/installer, granting full access`);
        return true;
      }
    } catch (e) {
      Logger.log(`Error checking if user is script owner: ${e.message}`);
    }
    
    // Directly check if user is in a team
    if (isUserInTeam(userEmail)) {
      Logger.log(`User ${userEmail} found in teams data`);
      return true;
    }
    
    return false;
  } catch (e) {
    Logger.log(`Error in hasVerifiedTeamAccess: ${e.message}`);
    return false;
  }
}

/**
 * Preloads verified user data to ensure faster access checks
 * This should be called at the start of onOpen
 */
function preloadVerifiedUsers() {
  try {
    // Basic check to ensure teams data is available
    const teamsData = getTeamsData();
    if (teamsData) {
      Logger.log('Teams data preloaded successfully');
    }
    return true;
  } catch (e) {
    Logger.log(`Error in preloadVerifiedUsers: ${e.message}`);
    return false;
  }
}

/**
 * Shows a dialog for users to verify their team membership
 */
function verifyTeamAccess() {
  try {
    const ui = SpreadsheetApp.getUi();
    const userEmail = Session.getActiveUser().getEmail();
    
    // Check if user is already in a team
    if (isUserInTeam(userEmail)) {
      // Add user to verified list
      try {
        const verifiedUsersStr = PropertiesService.getDocumentProperties().getProperty('VERIFIED_USER_IDS') || '[]';
        const verifiedUsers = JSON.parse(verifiedUsersStr);
        
        // Only add if not already in list
        if (!verifiedUsers.includes(userEmail.toLowerCase())) {
          verifiedUsers.push(userEmail.toLowerCase());
          PropertiesService.getDocumentProperties().setProperty('VERIFIED_USER_IDS', JSON.stringify(verifiedUsers));
          // Update the email-to-team map to ensure persistence between sessions
          updateEmailToTeamMap();
          // Also store in cache for faster lookup
          Logger.log(`Added ${userEmail} to verified users list`);
        }
      } catch (e) {
        Logger.log(`Error adding to verified users: ${e.message}`);
      }
      
      // Show success with forced reload
      const html = HtmlService.createHtmlOutput(`
        <p>Your team access has been verified!</p>
        <p>Please click the button below to reload the page and see the Pipedrive menu.</p>
        <script>
          function forceReload() {
            google.script.run.withSuccessHandler(function(result) {
              google.script.host.close();
              window.top.location.reload(true);
            }).refreshMenuAfterVerification();
          }
        </script>
        <div style="text-align: center; margin-top: 20px;">
          <button 
            style="padding: 10px 20px; background: #4285F4; color: white; border: none; border-radius: 4px; cursor: pointer;"
            onclick="forceReload()">Reload Page</button>
        </div>
      `)
      .setWidth(350)
      .setHeight(180);
      
      ui.showModalDialog(html, 'Team Access Verified');
      return;
    }
    
    // Show team join dialog for non-members
    showTeamJoinRequest();
  } catch (e) {
    Logger.log(`Error in verifyTeamAccess: ${e.message}`);
    ui.alert('Error', 'An error occurred: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Refreshes the menu after verification
 */
function refreshMenuAfterVerification() {
  try {
    // Clear existing menus by creating an entirely new menu
    const ui = SpreadsheetApp.getUi();
    
    // Force update the email-to-team map
    updateEmailToTeamMap();
    
    // Create the full menu directly instead of calling onOpen
    createPipedriveMenu();
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Team access verified successfully!',
      'Success',
      5
    );
    return true;
  } catch (e) {
    Logger.log(`Error refreshing menu: ${e.message}`);
    return false;
  }
}

/**
 * Checks if the user is a team member and verifies their status
 */
function checkAndVerifyTeamMembership() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Unable to determine your email. Please make sure you are signed in.' };
    }
    
    // Check if user is in a team
    if (isUserInTeam(userEmail)) {
      // Mark as verified in both UserProperties and DocumentProperties for persistence
      try {
        // Update user properties
        const userProps = PropertiesService.getUserProperties();
        userProps.setProperty('VERIFIED_TEAM_ACCESS', 'true');
        
        // Update document properties for master list of verified users
        const docProps = PropertiesService.getDocumentProperties();
        const verifiedUsersJson = docProps.getProperty('VERIFIED_TEAM_USERS');
        const verifiedUsers = verifiedUsersJson ? JSON.parse(verifiedUsersJson) : [];
        
        if (!verifiedUsers.includes(userEmail)) {
          verifiedUsers.push(userEmail);
          docProps.setProperty('VERIFIED_TEAM_USERS', JSON.stringify(verifiedUsers));
        }
        
        Logger.log(`User ${userEmail} verified successfully`);
      } catch (e) {
        Logger.log(`Error setting verification properties: ${e.message}`);
      }
      
      return { success: true };
    } else {
      return { success: false, error: 'You are not a member of any team. Please join a team first.' };
    }
  } catch (e) {
    Logger.log(`Error verifying team membership: ${e.message}`);
    return { success: false, error: e.message };
  }
}

// Add this function to handle join requests
function showTeamJoinRequest() {
  // Simply call showTeamManager with a joinOnly flag
  showTeamManager(true);
}

// Add this at the top of your file to cache field definitions
const fieldDefinitionsCache = {};

function isMultiOptionField(fieldKey, entityType) {
  try {
    // Get field definitions (cached if possible)
    let fieldDefinitions = [];
    const cacheKey = `fields_${entityType}`;

    if (!fieldDefinitionsCache[cacheKey]) {
      switch (entityType) {
        case ENTITY_TYPES.PERSONS:
          fieldDefinitions = getPersonFields();
          break;
        case ENTITY_TYPES.DEALS:
          fieldDefinitions = getDealFields();
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          fieldDefinitions = getOrganizationFields();
          break;
        case ENTITY_TYPES.PRODUCTS:
          fieldDefinitions = getProductFields();
          break;
        default:
          return false;
      }
      fieldDefinitionsCache[cacheKey] = fieldDefinitions;
      Logger.log(`Successfully retrieved ${entityType} fields using v1 API`);
    } else {
      fieldDefinitions = fieldDefinitionsCache[cacheKey];
    }

    // Find the field definition
    const fieldDef = fieldDefinitions.find(f => f.key === fieldKey);
    if (fieldDef) {
      return fieldDef.field_type === 'set' ||
        (fieldDef.options && fieldDef.multiple);
    }
  } catch (e) {
    Logger.log(`Error checking field type: ${e.message}`);
  }

  return false;
}

/**
 * Gets all teams data from document properties
 * @returns {Object} An object containing all teams data
 */
function getTeamsData() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const teamsDataJson = documentProperties.getProperty('TEAMS_DATA');
  let teamsData = {};

  if (teamsDataJson) {
    try {
      teamsData = JSON.parse(teamsDataJson);
    } catch (e) {
      Logger.log('Error parsing teams data: ' + e.message);
      teamsData = {};
    }
  }

  return teamsData;
}

/**
 * Gets the team that a user belongs to
 * @param {string} userEmail The user's email address
 * @param {Object} teamsData Optional teams data object (to avoid repeated lookups)
 * @returns {Object|null} The team object or null if user doesn't belong to a team
 */
function getUserTeam(userEmail, teamsData = null) {
  try {
    // Try to get team ID from user properties first (faster and persists across refreshes)
    const userProperties = PropertiesService.getUserProperties();
    const userTeamId = userProperties.getProperty('USER_TEAM_ID');

    // If we have a cached team ID, verify it's still valid
    if (userTeamId) {
      // Load teams data if not provided
      if (!teamsData) {
        teamsData = getTeamsData();
      }

      // Check if the team exists and user is still a member
      if (teamsData[userTeamId] &&
        teamsData[userTeamId].memberEmails &&
        teamsData[userTeamId].memberEmails.includes(userEmail)) {
        // Add the ID to the team object
        const team = teamsData[userTeamId];
        team.id = userTeamId;
        Logger.log(`Found user ${userEmail} in team ${userTeamId} via user properties`);
        return team;
      }
    }

    // Fall back to searching all teams
    if (!teamsData) {
      teamsData = getTeamsData();
    }

    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      if (team.memberEmails.includes(userEmail)) {
        // Add the ID to the team object
        team.id = teamId;

        // Update the user's cached team ID
        userProperties.setProperty('USER_TEAM_ID', teamId);
        Logger.log(`Found user ${userEmail} in team ${teamId} via full search, updated user properties`);

        return team;
      }
    }

    // No team found, clear any stale user property
    if (userTeamId) {
      userProperties.deleteProperty('USER_TEAM_ID');
      Logger.log(`Cleared stale team ID for user ${userEmail}`);
    }

    return null;
  } catch (e) {
    Logger.log(`Error in getUserTeam: ${e.message}`);
    return null;
  }
}

/**
 * Saves the teams data to document properties
 * @param {Object} teamsData The teams data to save
 * @returns {boolean} Whether the save was successful
 */
function saveTeamsData(teamsData) {
  try {
    const documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('TEAMS_DATA', JSON.stringify(teamsData));
    return true;
  } catch (e) {
    Logger.log('Error saving teams data: ' + e.message);
    return false;
  }
}

/**
 * Creates a new team with the current user as admin
 * @param {string} teamName The name of the team to create
 * @returns {Object} Result object with success/error information
 */
function createTeam(teamName) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Check if user is already in a team
    const existingTeam = getUserTeam(userEmail, teamsData);
    if (existingTeam) {
      return { success: false, error: 'You are already a member of a team. Please leave your current team before creating a new one.' };
    }

    // Generate a unique team ID (using timestamp and random string)
    const teamId = new Date().getTime().toString(36) + Math.random().toString(36).substr(2, 5);

    // Create the new team
    teamsData[teamId] = {
      name: teamName,
      adminEmails: [userEmail],
      memberEmails: [userEmail],
      dateCreated: new Date().toISOString(),
      shareFilters: true,
      shareColumns: true,
    };

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      return { success: true, teamId: teamId };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error creating team: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Adds a user to an existing team
 * @param {string} teamId The ID of the team to join
 * @returns {Object} Result object with success/error information
 */
function joinTeam(teamId) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data from properties
    const teamsData = getTeamsData();

    // Check if user is already in a team in the main data store
    const existingTeam = getUserTeam(userEmail, teamsData);
    if (existingTeam) {
      return { success: false, error: 'You are already a member of a team. Please leave your current team before joining another one.' };
    }

    // Check if the team exists
    if (!teamsData[teamId]) {
      return { success: false, error: 'Team not found. Please check the team ID and try again.' };
    }

    // Check if the team is at maximum capacity
    if (teamsData[teamId].memberEmails.length >= MAX_TEAM_MEMBERS) {
      return { success: false, error: `This team has reached the maximum of ${MAX_TEAM_MEMBERS} members.` };
    }

    // Add the user to the team
    teamsData[teamId].memberEmails.push(userEmail);

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error joining team: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

function fixMenuAfterJoin() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Add debug logging
    Logger.log(`Fixing menu after join for user: ${userEmail}`);
    
    // Show success dialog with reload button
    const html = HtmlService.createHtmlOutput(`
      <p>You have successfully joined the team.</p>
      <p>Please reload the page to see the Pipedrive menu.</p>
      <script>
        function reloadPage() {
          google.script.host.close();
          window.top.location.reload();
        }
      </script>
      <button onclick="reloadPage()">Reload Page</button>
    `)
    .setWidth(350)
    .setHeight(180);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Team Joined');
    return true;
  } catch (e) {
    Logger.log(`Error fixing menu after join: ${e.message}`);
    return false;
  }
}

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
          .addItem('ℹ️ Help & About', 'showHelp');
      
  menu.addToUi();
}

/**
 * Creates a flattened map of user emails to team IDs for quick lookup
 * This is stored in document properties for reliability across refreshes
 */
function updateEmailToTeamMap() {
  try {
    const teamsData = getTeamsData();
    const emailToTeamMap = {};
    
    // Build the map
    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      const memberEmails = team.memberEmails || [];
      
      // Add each member to the map
      for (let i = 0; i < memberEmails.length; i++) {
        // Store in lowercase for consistent lookup
        emailToTeamMap[memberEmails[i].toLowerCase()] = teamId;
      }
    }
    
    // Save the map to document properties only
    PropertiesService.getDocumentProperties().setProperty('EMAIL_TO_TEAM_MAP', JSON.stringify(emailToTeamMap));
    Logger.log('Email-to-team map updated successfully');
    return true;
  } catch (e) {
    Logger.log(`Error updating email-to-team map: ${e.message}`);
    return false;
  }
}

/**
 * Removes the current user from their team
 * @returns {Object} Result object with success/error information
 */
function leaveTeam() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    const teamId = userTeam.id;

    // If user is the only admin, and there are other members, don't allow leaving
    if (userTeam.adminEmails.includes(userEmail) &&
      userTeam.adminEmails.length === 1 &&
      userTeam.memberEmails.length > 1) {
      return {
        success: false,
        error: 'You are the only admin of this team. Please promote another member to admin before leaving.'
      };
    }

    // If user is admin and the only member, delete the team
    if (userTeam.adminEmails.includes(userEmail) && userTeam.memberEmails.length === 1) {
      delete teamsData[teamId];
    } else {
      // Remove user from the team
      teamsData[teamId].memberEmails = teamsData[teamId].memberEmails.filter(email => email !== userEmail);

      // If user is an admin, remove from admin list too
      if (userTeam.adminEmails.includes(userEmail)) {
        teamsData[teamId].adminEmails = teamsData[teamId].adminEmails.filter(email => email !== userEmail);
      }
    }

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      // Update the email to team mapping
      updateEmailToTeamMap();
      
      // Clear user properties
      try {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.deleteProperty('USER_TEAM_ID');
      } catch (e) {
        Logger.log(`Error clearing user property: ${e.message}`);
      }
      
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error leaving team: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Invites a user to the current user's team
 * @param {string} email The email of the user to invite
 * @returns {Object} Result object with success/error information
 */
function inviteMember(email) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    // Check if the user is an admin
    if (!userTeam.adminEmails.includes(userEmail)) {
      return { success: false, error: 'Only team admins can invite members.' };
    }

    // Check if the team is at maximum capacity
    if (userTeam.memberEmails.length >= MAX_TEAM_MEMBERS) {
      return { success: false, error: `Your team has reached the maximum of ${MAX_TEAM_MEMBERS} members.` };
    }

    // Check if the user is already a member
    if (userTeam.memberEmails.includes(email)) {
      return { success: false, error: 'This user is already a member of your team.' };
    }

    // Check if the user is already in another team
    for (const tid in teamsData) {
      if (teamsData[tid].memberEmails.includes(email)) {
        return { success: false, error: 'This user is already a member of another team.' };
      }
    }

    // Add the user to the team
    teamsData[userTeam.id].memberEmails.push(email);

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error inviting member: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Promotes a team member to admin
 * @param {string} email The email of the member to promote
 * @returns {Object} Result object with success/error information
 */
function promoteTeamMember(email) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    // Check if the user is an admin
    if (!userTeam.adminEmails.includes(userEmail)) {
      return { success: false, error: 'Only team admins can promote members.' };
    }

    // Check if the member is in the team
    if (!userTeam.memberEmails.includes(email)) {
      return { success: false, error: 'This user is not a member of your team.' };
    }

    // Check if the member is already an admin
    if (userTeam.adminEmails.includes(email)) {
      return { success: false, error: 'This user is already an admin.' };
    }

    // Promote the member to admin
    teamsData[userTeam.id].adminEmails.push(email);

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error promoting team member: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Demotes a team admin to regular member
 * @param {string} email The email of the admin to demote
 * @returns {Object} Result object with success/error information
 */
function demoteTeamMember(email) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    // Check if the user is an admin
    if (!userTeam.adminEmails.includes(userEmail)) {
      return { success: false, error: 'Only team admins can demote other admins.' };
    }

    // Check if the admin is in the team
    if (!userTeam.memberEmails.includes(email)) {
      return { success: false, error: 'This user is not a member of your team.' };
    }

    // Check if the member is an admin
    if (!userTeam.adminEmails.includes(email)) {
      return { success: false, error: 'This user is not an admin.' };
    }

    // Check if this would remove the last admin
    if (userTeam.adminEmails.length <= 1) {
      return { success: false, error: 'Cannot demote the last admin of the team.' };
    }

    // Demote the admin to regular member
    teamsData[userTeam.id].adminEmails = userTeam.adminEmails.filter(a => a !== email);

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error demoting team admin: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Removes a team member (admin function)
 * @param {string} email The email of the member to remove
 * @returns {Object} Result object with success/error information
 */
function removeTeamMember(email) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    if (!adminEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the admin's team
    const adminTeam = getUserTeam(adminEmail, teamsData);
    if (!adminTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    // Check if user is an admin
    if (!adminTeam.adminEmails.includes(adminEmail)) {
      return { success: false, error: 'Only team admins can remove members.' };
    }

    const teamId = adminTeam.id;

    // Check if the member exists in the team
    if (!adminTeam.memberEmails.includes(email)) {
      return { success: false, error: 'This user is not a member of your team.' };
    }

    // Remove the member from the team
    teamsData[teamId].memberEmails = teamsData[teamId].memberEmails.filter(e => e !== email);

    // If the member was an admin, remove from admin list too
    if (adminTeam.adminEmails.includes(email)) {
      teamsData[teamId].adminEmails = teamsData[teamId].adminEmails.filter(e => e !== email);
    }

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      // Update the email to team mapping to reflect the removal
      updateEmailToTeamMap();

      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error removing team member: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Saves team sharing settings
 * @param {boolean} shareFilters Whether to share filters with team members
 * @param {boolean} shareColumns Whether to share column preferences with team members
 * @returns {Object} Result object with success/error information
 */
function saveTeamSettings(shareFilters, shareColumns) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data
    const teamsData = getTeamsData();

    // Find the user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, error: 'You are not a member of any team.' };
    }

    // Check if the user is an admin
    if (!userTeam.adminEmails.includes(userEmail)) {
      return { success: false, error: 'Only team admins can change team settings.' };
    }

    // Update the team settings
    teamsData[userTeam.id].shareFilters = shareFilters;
    teamsData[userTeam.id].shareColumns = shareColumns;

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      return { success: true };
    } else {
      return { success: false, error: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error saving team settings: ' + e.message);
    return { success: false, error: 'An error occurred: ' + e.message };
  }
}

/**
 * Modified function to get saved column preferences for an entity type
 * Modified to consider team sharing settings
 * @param {string} entityType The entity type (deals, persons, etc)
 * @param {string} sheetName The sheet name
 * @returns {Array} Array of column configurations
 */
function getTeamAwareColumnPreferences(entityType, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const userEmail = Session.getActiveUser().getEmail();

  // Get the user's team and check sharing settings
  const userTeam = getUserTeam(userEmail);

  // First try to get user-specific column preferences
  const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
  let savedColumnsJson = scriptProperties.getProperty(userColumnSettingsKey);

  // If the user is part of a team and column sharing is enabled, check team preferences
  if ((!savedColumnsJson || savedColumnsJson === '[]') && userTeam && userTeam.shareColumns) {
    // Look for team members' configurations
    for (const teamMemberEmail of userTeam.memberEmails) {
      if (teamMemberEmail === userEmail) continue; // Skip the current user

      const teamMemberColumnSettingsKey = `COLUMNS_${teamMemberEmail}_${sheetName}_${entityType}`;
      const teamMemberColumnsJson = scriptProperties.getProperty(teamMemberColumnSettingsKey);

      if (teamMemberColumnsJson && teamMemberColumnsJson !== '[]') {
        savedColumnsJson = teamMemberColumnsJson;
        Logger.log(`Using team member ${teamMemberEmail}'s column preferences for ${entityType}`);
        break;
      }
    }
  }

  // If still no team column preferences, fall back to the global setting
  if (!savedColumnsJson || savedColumnsJson === '[]') {
    const globalColumnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
    savedColumnsJson = scriptProperties.getProperty(globalColumnSettingsKey);
  }

  let selectedColumns = [];
  if (savedColumnsJson) {
    try {
      selectedColumns = JSON.parse(savedColumnsJson);
    } catch (e) {
      Logger.log('Error parsing saved columns: ' + e.message);
      selectedColumns = [];
    }
  }

  return selectedColumns;
}

/**
 * Modified function to save column preferences for an entity type
 * Modified to use team-aware column storage
 * @param {Array} columns Array of column configurations
 * @param {string} entityType The entity type (deals, persons, etc)
 * @param {string} sheetName The sheet name
 */
function saveTeamAwareColumnPreferences(columns, entityType, sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const userEmail = Session.getActiveUser().getEmail();

  // Store columns in user-specific property
  const userColumnSettingsKey = `COLUMNS_${userEmail}_${sheetName}_${entityType}`;
  scriptProperties.setProperty(userColumnSettingsKey, JSON.stringify(columns));

  // Also store in global property for backward compatibility
  const columnSettingsKey = `COLUMNS_${sheetName}_${entityType}`;
  scriptProperties.setProperty(columnSettingsKey, JSON.stringify(columns));

  // Check if two-way sync is enabled for this sheet
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
  const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

  // When columns are changed and two-way sync is enabled, handle tracking column
  if (twoWaySyncEnabled) {
    Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Checking if we need to adjust sync column.`);

    // When columns are changed, delete the tracking column property to force repositioning
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    scriptProperties.deleteProperty(twoWaySyncTrackingColumnKey);

    // Add a flag to indicate that the Sync Status column should be repositioned at the end
    const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${sheetName}`;
    scriptProperties.setProperty(twoWaySyncColumnAtEndKey, 'true');

    Logger.log(`Removed tracking column property for sheet "${sheetName}" to ensure correct positioning on next sync.`);
  }
}

/**
 * Modified function to get Pipedrive filters, including team members' filters
 * @returns {Object} An object containing filter data
 */
function getTeamAwarePipedriveFilters() {
  try {
    // Get API key from properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      return { success: false, error: 'API key not configured' };
    }

    // Get the user's email
    const userEmail = Session.getActiveUser().getEmail();

    // Get the user's team
    const userTeam = getUserTeam(userEmail);

    // Initialize result
    let filters = {};

    // Get filters using the current user's API key
    const url = `https://${subdomain}.pipedrive.com/v1/filters?api_token=${apiKey}`;

    Logger.log(`Fetching filters from: ${url}`);

    // Add muteHttpExceptions to get more detailed error information
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Filter API response code: ${responseCode}`);

    if (responseCode !== 200) {
      return {
        success: false,
        error: `Server returned status code ${responseCode}: ${responseText.substring(0, 100)}`
      };
    }

    // Parse the response
    const result = JSON.parse(responseText);

    if (!result.success) {
      return {
        success: false,
        error: result.error || 'Failed to retrieve filters'
      };
    }

    // Group filters by type
    filters = result.data.reduce((acc, filter) => {
      const type = filter.type || 'other';
      if (!acc[type]) {
        acc[type] = [];
      }

      // Flag this as the user's own filter
      filter.isOwn = true;

      acc[type].push(filter);
      return acc;
    }, {});

    // If the user is part of a team and filter sharing is enabled, get team members' filters
    if (userTeam && userTeam.shareFilters) {
      // Store user's filter IDs to avoid duplicates
      const userFilterIds = new Set();
      for (const type in filters) {
        filters[type].forEach(filter => userFilterIds.add(filter.id));
      }

      // For each team member, store their filters in the document properties
      for (const memberEmail of userTeam.memberEmails) {
        if (memberEmail === userEmail) continue; // Skip the current user

        // Get the team member's filters from document properties
        const memberFiltersKey = `TEAM_MEMBER_FILTERS_${memberEmail}`;
        const memberFiltersJson = scriptProperties.getProperty(memberFiltersKey);

        if (memberFiltersJson) {
          try {
            const memberFilters = JSON.parse(memberFiltersJson);

            // Group member's filters by type and add to results
            for (const filter of memberFilters) {
              // Skip if this filter ID is already in the user's filters
              if (userFilterIds.has(filter.id)) continue;

              const type = filter.type || 'other';
              if (!filters[type]) {
                filters[type] = [];
              }

              // Flag this as a team member's filter
              filter.isTeamFilter = true;
              filter.ownerEmail = memberEmail;

              filters[type].push(filter);
            }
          } catch (e) {
            Logger.log(`Error parsing ${memberEmail}'s filters: ${e.message}`);
          }
        }
      }
    }

    return {
      success: true,
      data: filters
    };
  } catch (e) {
    Logger.log(`Error getting Pipedrive filters: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Stores the user's Pipedrive filters for team sharing
 * This should be called whenever filters are fetched
 * @param {Array} filters Array of filter objects from Pipedrive API
 */
function storeUserFilters(filters) {
  try {
    if (!filters || !Array.isArray(filters)) return;

    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return;

    const scriptProperties = PropertiesService.getScriptProperties();
    const memberFiltersKey = `TEAM_MEMBER_FILTERS_${userEmail}`;

    scriptProperties.setProperty(memberFiltersKey, JSON.stringify(filters));
    Logger.log(`Stored ${filters.length} filters for user ${userEmail}`);
  } catch (e) {
    Logger.log(`Error storing user filters: ${e.message}`);
  }
}

/**
 * Main entry point for syncing data from Pipedrive
 */
function syncFromPipedrive() {
  // Show a loading message
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();

  detectColumnShifts();

  // Get the script properties
  const scriptProperties = PropertiesService.getScriptProperties();

  // Check if two-way sync is enabled for this sheet
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
  const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
  const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

  // Get the current entity type and the last synced entity type
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  const currentEntityType = scriptProperties.getProperty(sheetEntityTypeKey);
  const lastEntityTypeKey = `LAST_ENTITY_TYPE_${activeSheetName}`;
  const lastEntityType = scriptProperties.getProperty(lastEntityTypeKey);

  // Check if entity type has changed
  if (currentEntityType !== lastEntityType && currentEntityType && twoWaySyncEnabled) {
    // Entity type has changed - clear the Sync Status column letter
    Logger.log(`Entity type changed from ${lastEntityType || 'none'} to ${currentEntityType}. Clearing tracking column.`);
    scriptProperties.deleteProperty(twoWaySyncTrackingColumnKey);

    // Add a flag to indicate that the Sync Status column should be repositioned
    const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${activeSheetName}`;
    scriptProperties.setProperty(twoWaySyncColumnAtEndKey, 'true');

    // Store the new entity type as the last synced entity type
    scriptProperties.setProperty(lastEntityTypeKey, currentEntityType);
  }

  // First show a confirmation dialog, including info about pushing changes if two-way sync is enabled
  let confirmMessage = `This will sync data from Pipedrive to the current sheet "${activeSheetName}". Any existing data in this sheet will be replaced.`;

  if (twoWaySyncEnabled) {
    confirmMessage += `\n\nTwo-way sync is enabled for this sheet. Modified rows will be pushed to Pipedrive before pulling new data.`;
  }

  confirmMessage += `\n\nDo you want to continue?`;

  const confirmation = ui.alert(
    'Sync Pipedrive Data',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );

  if (confirmation === ui.Button.NO) {
    return;
  }

  // If two-way sync is enabled, push changes to Pipedrive first
  if (twoWaySyncEnabled) {
    // Show a message that we're pushing changes
    spreadsheet.toast('Two-way sync enabled. Pushing modified rows to Pipedrive first...', 'Syncing', 5);

    try {
      // Call pushChangesToPipedrive() with true for isScheduledSync to suppress duplicate UI messages
      pushChangesToPipedrive(false, true);
    } catch (error) {
      // Log the error but continue with the sync
      Logger.log(`Error pushing changes: ${error.message}`);
      spreadsheet.toast(`Warning: Error pushing changes to Pipedrive: ${error.message}`, 'Sync Warning', 10);
    }
  }

  // Show a sync status UI
  showSyncStatus(activeSheetName);

  // Set the active sheet as the current sheet for this operation
  scriptProperties.setProperty('SHEET_NAME', activeSheetName);

  try {
    switch (currentEntityType) {
      case ENTITY_TYPES.DEALS:
        syncDealsFromFilter(true);
        break;
      case ENTITY_TYPES.PERSONS:
        syncPersonsFromFilter(true);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        syncOrganizationsFromFilter(true);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        syncActivitiesFromFilter(true);
        break;
      case ENTITY_TYPES.LEADS:
        syncLeadsFromFilter(true);
        break;
      case ENTITY_TYPES.PRODUCTS:
        syncProductsFromFilter(true);
        break;
      default:
        spreadsheet.toast('Unknown entity type. Please check settings.', 'Sync Error', 10);
        break;
    }

    // After successful sync, update the last entity type
    scriptProperties.setProperty(lastEntityTypeKey, currentEntityType);
  } catch (error) {
    // If there's an error, show it
    spreadsheet.toast('Error: ' + error.message, 'Sync Error', 10);
    Logger.log('Sync error: ' + error.message);
  }
}

// Add this helper function to your code
function refreshMenuAfterJoin() {
  try {
    const ui = SpreadsheetApp.getUi();
    const userEmail = Session.getActiveUser().getEmail();
    const teamsData = getTeamsData();
    const userTeam = getUserTeam(userEmail, teamsData);

    // Show debug info
    Logger.log(`Menu refresh - User: ${userEmail}, Has team: ${userTeam !== null}`);
    if (userTeam) {
      Logger.log(`Team ID: ${userTeam.id}, Members: ${userTeam.memberEmails.length}`);
    }

    // Force proper menu creation
    onOpen();

    // Show confirmation toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      userTeam ? 'Full Pipedrive menu activated.' : 'Limited menu shown - no team membership found.',
      'Menu Refreshed',
      5
    );
  } catch (e) {
    Logger.log(`Error in refreshMenuAfterJoin: ${e.message}`);
  }
}

/**
 * Team management functionality for Pipedrive Sheets
 * This extension allows users to create teams, share filters and column preferences
 * with team members, and set admin permissions.
 */

// Team-related constants
const MAX_TEAM_MEMBERS = 5;
const TEAM_ADMIN_ROLE = 'admin';
const TEAM_MEMBER_ROLE = 'member';

/**
 * Shows the team management UI where users can create/join teams and manage team members
 * @param {boolean} joinOnly - If true, only show the join team section
 */
function showTeamManager(joinOnly = false) {
  // Get the active user's email
  const userEmail = Session.getActiveUser().getEmail();

  // Get current team settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const documentProperties = PropertiesService.getDocumentProperties();

  // Get team data
  const teamsData = getTeamsData();
  const userTeam = getUserTeam(userEmail, teamsData);
  const isAdmin = userTeam && userTeam.adminEmails.includes(userEmail);

  // Create HTML content for the team manager dialog
  let htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
        --section-bg: #f8f9fa;
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

      /* Loading spinner animation */
@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

.loading-spinner {
  display: inline-block;
  width: 12px;
  height: 12px;
  border: 2px solid rgba(255,255,255,0.5);
  border-radius: 50%;
  border-top-color: white;
  animation: spin 0.8s linear infinite;
  margin-right: 5px;
  vertical-align: middle;
  position: relative;
  top: -1px;
}

.button-text {
  display: inline-block;
  vertical-align: middle;
}

.button-loading {
  background-color: #6B9AE7 !important;
  cursor: wait !important;
  opacity: 0.85;
  transition: all 0.2s ease;
}
      
      .section-title {
        font-size: 15px;
        font-weight: 500;
        color: var(--text-dark);
        margin-bottom: 12px;
        padding-bottom: 6px;
        border-bottom: 1px solid var(--border-color);
      }
      
      .form-container {
        max-width: 100%;
      }
      
      .user-info {
        background-color: var(--bg-light);
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 16px;
        border-left: 4px solid var(--primary-color);
        display: flex;
        align-items: center;
        font-size: 13px;
      }
      
      .user-info svg {
        margin-right: 12px;
        fill: var(--primary-color);
      }
      
      .section {
        background-color: var(--section-bg);
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
        border: 1px solid var(--border-color);
      }
      
      .form-row {
        display: flex;
        gap: 16px;
        margin-bottom: 16px;
      }
      
      .form-group {
        margin-bottom: 16px;
        flex: 1;
      }
      
      .form-group:last-child {
        margin-bottom: 0;
      }
      
      label {
        display: block;
        font-weight: 500;
        margin-bottom: 6px;
        color: var(--text-dark);
        font-size: 13px;
      }
      
      input, select {
        width: 100%;
        padding: 8px 12px;
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
        font-size: 11px;
        color: var(--text-light);
        margin-top: 4px;
      }
      
      .button-container {
        display: flex;
        justify-content: flex-end;
        margin-top: 20px;
      }
      
      .button-primary {
        background-color: var(--primary-color);
        color: white;
        border: none;
        padding: 8px 20px;
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
        padding: 7px 14px;
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
      
      .button-danger {
        background-color: var(--error-color);
        color: white;
      }
      
      .button-danger:hover {
        background-color: #c53929;
      }
      
      .toggle-switch {
        position: relative;
        display: inline-block;
        width: 40px;
        height: 22px;
        margin-right: 10px;
        vertical-align: middle;
      }
      
      .toggle-switch input {
        opacity: 0;
        width: 0;
        height: 0;
      }
      
      .toggle-slider {
        position: absolute;
        cursor: pointer;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background-color: #ccc;
        transition: .3s;
        border-radius: 34px;
      }
      
      .toggle-slider:before {
        position: absolute;
        content: "";
        height: 18px;
        width: 18px;
        left: 2px;
        bottom: 2px;
        background-color: white;
        transition: .3s;
        border-radius: 50%;
      }
      
      input:checked + .toggle-slider {
        background-color: var(--primary-color);
      }
      
      input:focus + .toggle-slider {
        box-shadow: 0 0 1px var(--primary-color);
      }
      
      input:checked + .toggle-slider:before {
        transform: translateX(18px);
      }
      
      .toggle-wrapper {
        display: flex;
        align-items: center;
        margin-bottom: 8px;
      }
      
      .toggle-label {
        font-weight: normal;
        cursor: pointer;
      }
      
      .tab-container {
        margin-bottom: 20px;
      }
      
      .tabs {
        display: flex;
        border-bottom: 1px solid var(--border-color);
      }
      
      .tab {
        padding: 8px 16px;
        cursor: pointer;
        font-weight: 500;
        color: var(--text-light);
        border-bottom: 2px solid transparent;
        transition: var(--transition);
      }
      
      .tab:hover {
        color: var(--primary-color);
      }
      
      .tab.active {
        color: var(--primary-color);
        border-bottom-color: var(--primary-color);
      }
      
      .tab-content {
        padding: 16px 0;
        display: none;
      }
      
      .tab-content.active {
        display: block;
      }
      
      .team-card {
        border: 1px solid var(--border-color);
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
        background-color: white;
        transition: var(--transition);
      }
      
      .team-card:hover {
        box-shadow: var(--shadow);
      }
      
      .team-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
      }
      
      .team-name {
        font-size: 16px;
        font-weight: 500;
      }

      button {
  padding: 6px 12px;
  border-radius: 4px;
  border: none;
  background-color: var(--primary-color);
  color: white;
  cursor: pointer;
  font-weight: 500;
  transition: background-color 0.2s ease;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 32px;
}

button:hover {
  background-color: var(--primary-dark);
}

.remove-member {
  background-color: var(--error-color);
  min-width: 80px;
}

.remove-member:hover {
  background-color: #c53929;
  color: white !important;
}
      
      .badge {
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 12px;
        background-color: var(--bg-light);
        color: var(--text-light);
      }
      
      .badge.admin {
        background-color: #e8f0fe;
        color: var(--primary-color);
      }
      
      .team-members {
        margin-top: 12px;
      }
      
      .team-member {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 6px 0;
        border-bottom: 1px solid var(--border-color);
      }
      
      .team-member:last-child {
        border-bottom: none;
      }
      
      .member-info {
        display: flex;
        align-items: center;
      }
      
      .member-avatar {
        width: 24px;
        height: 24px;
        border-radius: 50%;
        background-color: var(--primary-color);
        color: white;
        text-align: center;
        line-height: 24px;
        margin-right: 10px;
        font-size: 12px;
      }
      
      .member-email {
        font-size: 13px;
      }
      
      .member-role {
        font-size: 11px;
        color: var(--text-light);
      }
      
      .member-actions button {
        background: none;
        border: none;
        cursor: pointer;
        color: var(--text-light);
        transition: var(--transition);
      }
      
      .member-actions button:hover {
        color: var(--primary-color);
      }
      
      .empty-state {
        text-align: center;
        padding: 20px;
        color: var(--text-light);
      }
      
      .status {
        margin-top: 16px;
        padding: 10px 14px;
        border-radius: 4px;
        font-size: 13px;
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
        width: 16px;
        height: 16px;
        margin-right: 8px;
        border: 2px solid rgba(255, 255, 255, 0.3);
        border-radius: 50%;
        border-top-color: #fff;
        animation: spin 0.8s linear infinite;
      }
      
      @keyframes spin {
        to {
          transform: rotate(360deg);
        }
      }
    </style>`;

  // Add user info section if not in join-only mode
  if (!joinOnly) {
    htmlContent += `
    <div class="user-info">
      <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
        <path d="M0 0h24v24H0z" fill="none"/>
        <path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/>
      </svg>
      <div>Logged in as: <strong>${userEmail}</strong></div>
    </div>`;
  }

  if (joinOnly) {
    // Only show join team section for non-team members
    htmlContent += `
    <div class="section">
      <div class="section-title">Join a Pipedrive Team</div>
      <div class="form-group">
        <label for="joinTeamId">Team ID</label>
        <input type="text" id="joinTeamId" placeholder="Enter the team ID provided by the team admin" />
        <span class="tooltip">Ask your team admin for the Team ID</span>
      </div>
      <div class="button-container">
        <button id="joinTeamButton" class="button-primary">Join Team</button>
      </div>
    </div>
    <div id="status" class="status"></div>`;
  } else {
    // Show full team management UI with tabs
    htmlContent += `
    <div class="tab-container">
      <div class="tabs">
        <div class="tab active" data-tab="my-team">My Team</div>
        ${isAdmin ? '<div class="tab" data-tab="create-team">Create Team</div>' : ''}
        <div class="tab" data-tab="join-team">Join Team</div>
        ${isAdmin ? '<div class="tab" data-tab="admin">Admin Settings</div>' : ''}
      </div>
      
      <div class="tab-content active" id="my-team">
        ${userTeam ? `
          <div class="team-card">
            <div class="team-header">
              <div class="team-name">${userTeam.name}</div>
              ${userTeam.adminEmails.includes(userEmail) ? '<span class="badge admin">Admin</span>' : '<span class="badge">Member</span>'}
            </div>
            <div>Team ID: <span id="teamId">${userTeam.id}</span> <button id="copyTeamId" class="button-secondary" style="padding: 2px 6px; margin: 0 0 0 5px; font-size: 11px;">Copy</button></div>
            <div class="team-members">
              <h4>Team Members (${userTeam.memberEmails.length}/${MAX_TEAM_MEMBERS})</h4>
              ${userTeam.memberEmails.map(email => `
                <div class="team-member">
                  <div class="member-info">
                    <div class="member-avatar">${email.charAt(0).toUpperCase()}</div>
                    <div>
                      <div class="member-email">${email}</div>
                      <div class="member-role">${userTeam.adminEmails.includes(email) ? 'Admin' : 'Member'}</div>
                    </div>
                  </div>
                  ${userTeam.adminEmails.includes(userEmail) && email !== userEmail ? `
                    <div class="member-actions">
                      ${!userTeam.adminEmails.includes(email) ? `
                        <button class="promote-member" data-email="${email}">Make Admin</button>
                      ` : userTeam.adminEmails.length > 1 ? `
                        <button class="demote-member" data-email="${email}">Remove Admin</button>
                      ` : ''}
                      <button class="remove-member" data-email="${email}">
  <span class="loading-spinner" style="display:none;"></span>
  <span class="button-text">Remove</span>
</button>
                    </div>
                  ` : ''}
                </div>
              `).join('')}
            </div>
            <div class="button-container">
              ${userTeam.adminEmails.includes(userEmail) ? `
                <button id="inviteButton" class="button-secondary">Invite Member</button>
                <button id="leaveTeamButton" class="button-secondary button-danger">Delete Team</button>
              ` : `
                <button id="leaveTeamButton" class="button-secondary button-danger">Leave Team</button>
              `}
            </div>
          </div>
        ` : `
          <div class="empty-state">
            <p>You're not currently part of any team.</p>
            <p>${isAdmin ? 'Create a new team or join an existing one.' : 'Join an existing team to collaborate with others.'}</p>
          </div>
        `}
      </div>
      
      ${isAdmin ? `
      <div class="tab-content" id="create-team">
        <div class="section">
          <div class="section-title">Create a New Team</div>
          <div class="form-group">
            <label for="teamName">Team Name</label>
            <input type="text" id="teamName" placeholder="Enter a name for your team" />
          </div>
          <div class="button-container">
            <button id="createTeamButton" class="button-primary">Create Team</button>
          </div>
        </div>
      </div>
      ` : ''}
      
      <div class="tab-content" id="join-team">
        <div class="section">
          <div class="section-title">Join an Existing Team</div>
          <div class="form-group">
            <label for="joinTeamId">Team ID</label>
            <input type="text" id="joinTeamId" placeholder="Enter the team ID provided by the team admin" />
            <span class="tooltip">Ask your team admin for the Team ID</span>
          </div>
          <div class="button-container">
            <button id="joinTeamButton" class="button-primary">Join Team</button>
          </div>
        </div>
      </div>
      
      ${isAdmin ? `
        <div class="tab-content" id="admin">
          <div class="section">
            <div class="section-title">Team Sharing Settings</div>
            <div class="form-group">
              <div class="toggle-wrapper">
                <label class="toggle-switch">
                  <input type="checkbox" id="shareFilters" ${userTeam.shareFilters ? 'checked' : ''} />
                  <span class="toggle-slider"></span>
                </label>
                <label for="shareFilters" class="toggle-label">Share Pipedrive Filters with team members</label>
              </div>
              <span class="tooltip">When enabled, all team members will see filters saved by other team members</span>
            </div>
            
            <div class="form-group">
              <div class="toggle-wrapper">
                <label class="toggle-switch">
                  <input type="checkbox" id="shareColumns" ${userTeam.shareColumns ? 'checked' : ''} />
                  <span class="toggle-slider"></span>
                </label>
                <label for="shareColumns" class="toggle-label">Share Column Preferences with team members</label>
              </div>
              <span class="tooltip">When enabled, all team members will use the same column configurations</span>
            </div>
            
            <div class="button-container">
              <button id="saveAdminSettingsButton" class="button-primary">Save Settings</button>
            </div>
          </div>
        </div>
      ` : ''}
    </div>
    
    <div id="status" class="status"></div>`;
  }

  // JavaScript for the UI - different based on mode
  if (joinOnly) {
    htmlContent += `
    <script>
      // Join team
      document.getElementById('joinTeamButton').addEventListener('click', function() {
        const teamId = document.getElementById('joinTeamId').value.trim();
        
        if (!teamId) {
          showStatus('error', 'Please enter a team ID');
          return;
        }
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> Joining...';
        
        google.script.run
  .withSuccessHandler(function(result) {
    if (result.success) {
      showStatus('success', 'Joined team successfully!');
      
      // Fix the button
      document.getElementById('joinTeamButton').disabled = false;
      document.getElementById('joinTeamButton').innerHTML = 'Join Team';
      
      // Show clear instructions about refreshing
      const statusDiv = document.getElementById('status');
      const refreshText = document.createElement('p');
      refreshText.innerHTML = '<br><strong>⚠️ IMPORTANT:</strong> You must refresh the page to see the full menu. <br>Click the button below to refresh now.';
      refreshText.style.color = '#d14836';
      statusDiv.appendChild(refreshText);
      
      const refreshBtn = document.createElement('button');
      refreshBtn.innerText = '🔄 Refresh Page Now';
      refreshBtn.style.marginTop = '10px';
      refreshBtn.style.padding = '8px 16px';
      refreshBtn.style.backgroundColor = '#4285f4';
      refreshBtn.style.color = 'white';
      refreshBtn.style.border = 'none';
      refreshBtn.style.borderRadius = '4px';
      refreshBtn.style.cursor = 'pointer';
      refreshBtn.onclick = function() {
        window.top.location.reload();
      };
      statusDiv.appendChild(refreshBtn);
      
      // Close the dialog after 10 seconds
      setTimeout(function() {
        google.script.host.close();
      }, 10000);
    } else {
      showStatus('error', result.error || 'Failed to join team');
      document.getElementById('joinTeamButton').disabled = false;
      document.getElementById('joinTeamButton').innerHTML = 'Join Team';
    }
  })
  .withFailureHandler(function(error) {
    showStatus('error', error.message || 'Failed to join team');
    document.getElementById('joinTeamButton').disabled = false;
    document.getElementById('joinTeamButton').innerHTML = 'Join Team';
  })
  .joinTeam(teamId);
      });
      
      // Helper function to show status messages
      function showStatus(type, message) {
        const statusElement = document.getElementById('status');
        statusElement.className = 'status ' + type;
        statusElement.textContent = message;
        statusElement.style.display = 'block';
        
        // Auto hide after 5 seconds
        setTimeout(() => {
          statusElement.style.display = 'none';
        }, 5000);
      }
    </script>`;
  } else {
    // Full JavaScript for the complete team management UI
    htmlContent += `
    <script>
      // Show active tab
      document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', function() {
          // Remove active class from all tabs
          document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
          document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
          
          // Add active class to clicked tab
          this.classList.add('active');
          document.getElementById(this.dataset.tab).classList.add('active');
        });
      });
      
      // Copy team ID
      document.getElementById('copyTeamId')?.addEventListener('click', function() {
        const teamId = document.getElementById('teamId').textContent;
        navigator.clipboard.writeText(teamId)
          .then(() => showStatus('success', 'Team ID copied to clipboard!'))
          .catch(() => showStatus('error', 'Failed to copy team ID'));
      });
      
      // Success handler function to show message and reopen dialog
      function handleSuccess(message) {
        showStatus('success', message);
        
        // Close after showing the success message and reopen the dialog
        setTimeout(() => {
          // Tell the server to reopen the dialog
          google.script.run.reopenTeamManager();
          // Close this dialog
          google.script.host.close();
        }, 1500);
      }
      
      ${isAdmin ? `
      // Create team
      document.getElementById('createTeamButton')?.addEventListener('click', function() {
        const teamName = document.getElementById('teamName').value.trim();
        
        if (!teamName) {
          showStatus('error', 'Please enter a team name');
          return;
        }
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> Creating...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              handleSuccess('Team created successfully!');
            } else {
              showStatus('error', result.error || 'Failed to create team');
              document.getElementById('createTeamButton').disabled = false;
              document.getElementById('createTeamButton').textContent = 'Create Team';
            }
          })
          .withFailureHandler(function(error) {
            showStatus('error', error.message || 'Failed to create team');
            document.getElementById('createTeamButton').disabled = false;
            document.getElementById('createTeamButton').textContent = 'Create Team';
          })
          .createTeam(teamName);
      });
      ` : ''}
      
      // Join team
      document.getElementById('joinTeamButton').addEventListener('click', function() {
        const teamId = document.getElementById('joinTeamId').value.trim();
        
        if (!teamId) {
          showStatus('error', 'Please enter a team ID');
          return;
        }
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> Joining...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              handleSuccess('Joined team successfully!');
            } else {
              showStatus('error', result.error || 'Failed to join team');
              document.getElementById('joinTeamButton').disabled = false;
              document.getElementById('joinTeamButton').textContent = 'Join Team';
            }
          })
          .withFailureHandler(function(error) {
            showStatus('error', error.message || 'Failed to join team');
            document.getElementById('joinTeamButton').disabled = false;
            document.getElementById('joinTeamButton').textContent = 'Join Team';
          })
          .joinTeam(teamId);
      });
      
      // Leave team
      document.getElementById('leaveTeamButton')?.addEventListener('click', function() {
        const isAdmin = ${userTeam && userTeam.adminEmails.includes(userEmail) ? 'true' : 'false'};
        const actionText = isAdmin ? 'delete the team' : 'leave the team';
        
        if (!confirm('Are you sure you want to ' + actionText + '? ${isAdmin ? 'This will remove all team members and cannot be undone.' : ''}')) {
          return;
        }
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> ' + (isAdmin ? 'Deleting...' : 'Leaving...');
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              handleSuccess(isAdmin ? 'Team deleted successfully!' : 'Left team successfully!');
            } else {
              showStatus('error', result.error || 'Failed to ' + actionText);
              document.getElementById('leaveTeamButton').disabled = false;
              document.getElementById('leaveTeamButton').textContent = isAdmin ? 'Delete Team' : 'Leave Team';
            }
          })
          .withFailureHandler(function(error) {
            showStatus('error', error.message || 'Failed to ' + actionText);
            document.getElementById('leaveTeamButton').disabled = false;
            document.getElementById('leaveTeamButton').textContent = isAdmin ? 'Delete Team' : 'Leave Team';
          })
          .leaveTeam();
      });
      
      // Save admin settings
      document.getElementById('saveAdminSettingsButton')?.addEventListener('click', function() {
        const shareFilters = document.getElementById('shareFilters').checked;
        const shareColumns = document.getElementById('shareColumns').checked;
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> Saving...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              handleSuccess('Team settings saved successfully!');
            } else {
              showStatus('error', result.error || 'Failed to save team settings');
              document.getElementById('saveAdminSettingsButton').disabled = false;
              document.getElementById('saveAdminSettingsButton').textContent = 'Save Settings';
            }
          })
          .withFailureHandler(function(error) {
            showStatus('error', error.message || 'Failed to save team settings');
            document.getElementById('saveAdminSettingsButton').disabled = false;
            document.getElementById('saveAdminSettingsButton').textContent = 'Save Settings';
          })
          .saveTeamSettings(shareFilters, shareColumns);
      });
      
      // Invite member button
      document.getElementById('inviteButton')?.addEventListener('click', function() {
        const email = prompt('Enter the email address of the person you want to invite:');
        if (!email) return;
        
        if (!validateEmail(email)) {
          showStatus('error', 'Please enter a valid email address');
          return;
        }
        
        this.disabled = true;
        this.innerHTML = '<span class="spinner"></span> Inviting...';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              handleSuccess('Member invited successfully!');
            } else {
              showStatus('error', result.error || 'Failed to invite member');
              document.getElementById('inviteButton').disabled = false;
              document.getElementById('inviteButton').textContent = 'Invite Member';
            }
          })
          .withFailureHandler(function(error) {
            showStatus('error', error.message || 'Failed to invite member');
            document.getElementById('inviteButton').disabled = false;
            document.getElementById('inviteButton').textContent = 'Invite Member';
          })
          .inviteMember(email);
      });
      
      // Member promotion/demotion/removal
      document.querySelectorAll('.promote-member').forEach(button => {
        button.addEventListener('click', function() {
          const email = this.dataset.email;
          
          if (confirm('Make ' + email + ' an admin? Admins can manage team members and settings.')) {
            this.disabled = true;
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  handleSuccess('Member promoted to admin successfully!');
                } else {
                  showStatus('error', result.error || 'Failed to promote member');
                  document.querySelectorAll('.promote-member').forEach(b => b.disabled = false);
                }
              })
              .withFailureHandler(function(error) {
                showStatus('error', error.message || 'Failed to promote member');
                document.querySelectorAll('.promote-member').forEach(b => b.disabled = false);
              })
              .promoteTeamMember(email);
          }
        });
      });
      
      document.querySelectorAll('.demote-member').forEach(button => {
        button.addEventListener('click', function() {
          const email = this.dataset.email;
          
          if (confirm('Remove admin rights from ' + email + '?')) {
            this.disabled = true;
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  handleSuccess('Admin rights removed successfully!');
                } else {
                  showStatus('error', result.error || 'Failed to remove admin rights');
                  document.querySelectorAll('.demote-member').forEach(b => b.disabled = false);
                }
              })
              .withFailureHandler(function(error) {
                showStatus('error', error.message || 'Failed to remove admin rights');
                document.querySelectorAll('.demote-member').forEach(b => b.disabled = false);
              })
              .demoteTeamMember(email);
          }
        });
      });
      
      document.querySelectorAll('.remove-member').forEach(button => {
  button.addEventListener('click', function() {
    const email = this.dataset.email;
    
    if (confirm('Remove ' + email + ' from the team?')) {
      // Show loading state
      this.disabled = true;
      this.classList.add('button-loading');
      const spinner = this.querySelector('.loading-spinner');
      const buttonText = this.querySelector('.button-text');
      spinner.style.display = 'inline-block';
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            handleSuccess('Member removed successfully!');
          } else {
            showStatus('error', result.error || 'Failed to remove member');
            // Reset all remove buttons
            document.querySelectorAll('.remove-member').forEach(b => {
              b.disabled = false;
              b.classList.remove('button-loading');
              b.querySelector('.loading-spinner').style.display = 'none';
            });
          }
        })
        .withFailureHandler(function(error) {
          showStatus('error', error.message || 'Failed to remove member');
          // Reset all remove buttons
          document.querySelectorAll('.remove-member').forEach(b => {
            b.disabled = false;
            b.classList.remove('button-loading');
            b.querySelector('.loading-spinner').style.display = 'none';
          });
        })
        .removeTeamMember(email);
    }
  });
});
      
      // Helper functions
      function showStatus(type, message) {
        const statusElement = document.getElementById('status');
        statusElement.className = 'status ' + type;
        statusElement.textContent = message;
        statusElement.style.display = 'block';
        
        // Auto hide after 5 seconds
        setTimeout(() => {
          statusElement.style.display = 'none';
        }, 5000);
      }
      
      function validateEmail(email) {
        const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return re.test(email);
      }
    </script>`;
  }

  // Create and show the HTML output
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(joinOnly ? 400 : 500)
    .setHeight(joinOnly ? 280 : 600)
    .setTitle(joinOnly ? 'Join a Pipedrive Team' : 'Team Management');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, joinOnly ? 'Join a Pipedrive Team' : 'Team Management');
}

/**
 * Helper function to reopen the team manager dialog
 * Called by client-side JavaScript after a successful operation
 */
function reopenTeamManager() {
  showTeamManager();
}

/**
 * Shows settings dialog where users can configure API key, filter ID, and entity type
 */
function showSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();

  // Get current settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const savedApiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || '';
  const savedSubdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

  // Get sheet-specific settings using the sheet name
  const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;

  const savedFilterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const savedEntityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  const savedTimestampEnabled = scriptProperties.getProperty(`TIMESTAMP_ENABLED_${activeSheetName}`) || 'false';

  // Attempt to get filters if we have API key and subdomain
  let filtersHtml = '';
  let hasFilters = false;

  if (savedApiKey && savedSubdomain) {
    try {
      const filtersResult = getPipedriveFilters();

      if (filtersResult.success) {
        hasFilters = true;
        const filtersByType = filtersResult.data;

        // Generate filter options grouped by type
        for (const type in filtersByType) {
          const filters = filtersByType[type];

          if (filters.length > 0) {
            filtersHtml += `
              <optgroup>
            `;

            filters.forEach(filter => {
              filtersHtml += `
                <option style="margin-left: 10px;" value="${filter.id}" data-type="${type}" ${savedFilterId === filter.id.toString() ? 'selected' : ''}>${filter.name} (ID: ${filter.id})</option>
              `;
            });

            filtersHtml += `</optgroup>`;
          }
        }
      }
    } catch (e) {
      // Silently fail - we'll just show the manual filter ID input
      Logger.log('Error fetching filters: ' + e.message);
    }
  }

  // Create HTML content for the settings dialog
  const htmlOutput = HtmlService.createHtmlOutput(`
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
    --section-bg: #f8f9fa;
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
  
  .section-title {
    font-size: 15px;
    font-weight: 500;
    color: var(--text-dark);
    margin-bottom: 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--border-color);
  }
  
  .form-container {
    max-width: 100%;
  }
  
  .sheet-info {
    background-color: var(--bg-light);
    padding: 10px 14px;
    border-radius: 6px;
    margin-bottom: 16px;
    border-left: 4px solid var(--primary-color);
    display: flex;
    align-items: center;
    font-size: 13px;
  }
  
  .sheet-info svg {
    margin-right: 12px;
    fill: var(--primary-color);
  }
  
  .section {
    background-color: var(--section-bg);
    border-radius: 8px;
    padding: 16px;
    margin-bottom: 16px;
    border: 1px solid var(--border-color);
  }
  
  .form-row {
    display: flex;
    gap: 16px;
    margin-bottom: 16px;
  }
  
  .form-group {
    margin-bottom: 16px;
    flex: 1;
  }
  
  .form-group:last-child {
    margin-bottom: 0;
  }
  
  label {
    display: block;
    font-weight: 500;
    margin-bottom: 6px;
    color: var(--text-dark);
    font-size: 13px;
  }
  
  input, select {
    width: 100%;
    padding: 8px 12px;
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
    font-size: 11px;
    color: var(--text-light);
    margin-top: 4px;
  }
  
  .button-container {
    display: flex;
    justify-content: flex-end;
    margin-top: 20px;
  }
  
  .button-primary {
    background-color: var(--primary-color);
    color: white;
    border: none;
    padding: 8px 20px;
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
    padding: 7px 14px;
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
  
  /* Toggle switch styles */
  .toggle-switch {
    position: relative;
    display: inline-block;
    width: 66px;
    height: 22px;
    margin-right: 10px;
    vertical-align: middle;
  }
  
  .toggle-switch input {
    opacity: 0;
    width: 0;
    height: 0;
  }
  
  .toggle-slider {
    position: absolute;
    cursor: pointer;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: #ccc;
    transition: .3s;
    border-radius: 34px;
  }
  
  .toggle-slider:before {
    position: absolute;
    content: "";
    height: 18px;
    width: 18px;
    left: 2px;
    bottom: 2px;
    background-color: white;
    transition: .3s;
    border-radius: 50%;
  }
  
  input:checked + .toggle-slider {
    background-color: var(--primary-color);
  }
  
  input:focus + .toggle-slider {
    box-shadow: 0 0 1px var(--primary-color);
  }
  
  input:checked + .toggle-slider:before {
    transform: translateX(24px);
  }
  
  .toggle-wrapper {
    display: flex;
    align-items: center;
    margin-bottom: 8px;
  }
  
  .toggle-label {
    font-weight: normal;
    cursor: pointer;
  }
  
  .filter-section {
    margin-bottom: 0;
  }
  
  .filter-selector {
    margin-bottom: 8px;
  }
  
  .filter-search {
  margin-bottom: 10px;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  padding: 8px 12px 8px 36px;
  font-size: 13px;
  width: 100%;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%23757575' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'%3E%3C/circle%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'%3E%3C/line%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: 10px center;
  background-size: 16px;
}

.filter-search:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 0 2px rgba(66, 133, 244, 0.1);
}
  
  #filterSelector {
  height: 160px;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  background-color: white;
  width: 100%;
  font-size: 13px;
  overflow-y: auto;
  scrollbar-width: thin;
}

#filterSelector:focus {
  border-color: var(--primary-color);
  box-shadow: 0 0 0 2px rgba(66, 133, 244, 0.1);
  outline: none;
}

#filterSelector option {
  padding: 8px 12px;
  border-bottom: 1px solid #f0f0f0;
}

#filterSelector option:hover {
  background-color: var(--bg-light);
}

#filterSelector optgroup {
  padding: 6px;
  font-weight: 500;
  background-color: #f5f5f5;
  color: var(--text-dark);
}
  
  .filter-count {
  margin-bottom: 6px;
  font-size: 11px;
  color: var(--text-light);
  background-color: var(--bg-light);
  padding: 2px 8px;
  border-radius: 4px;
  display: inline-block;
}
  
  .filter-toggle {
  display: inline-block;
  color: var(--primary-color);
  font-size: 12px;
  cursor: pointer;
  margin-top: 10px;
  text-decoration: none;
  font-weight: 500;
}

.filter-toggle:hover {
  text-decoration: underline;
}
  
  .filter-input-container {
  margin-top: 10px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.filter-input {
  flex: 1;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  padding: 8px 12px;
  font-size: 13px;
}

.filter-input:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 0 2px rgba(66, 133, 244, 0.1);
}
  
  .filter-button {
  width: 36px;
  height: 36px;
  display: flex;
  align-items: center;
  justify-content: center;
  background-color: var(--primary-color);
  border: none;
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.2s ease;
}

.filter-button:hover {
  background-color: var(--primary-dark);
  box-shadow: var(--shadow-hover);
}

.filter-button svg {
  width: 18px;
  height: 18px;
  fill: white;
}
  
  .status {
    margin-top: 16px;
    padding: 10px 14px;
    border-radius: 4px;
    font-size: 13px;
    display: none;
  }
  
  .status.success {
    background-color: #e6f4ea;
    color: var(--success-color);
    display: block;
  }
  
  .hidden {
    display: none;
  }
</style>

<div class="sheet-info">
  <svg width="16" height="16" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M0 0h24v24H0z" fill="none"/>
    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
  </svg>
  <div>Configuring settings for sheet "<strong>${activeSheetName}</strong>"</div>
</div>

<div class="form-container">
  <!-- Pipedrive API Connection Section -->
  <div class="section">
    <div class="section-title">Pipedrive API Connection</div>
    <div class="form-row">
      <div class="form-group">
        <label for="apiKey">API Key</label>
        <input type="text" id="apiKey" value="${savedApiKey}" placeholder="Enter your Pipedrive API Key" />
        <span class="tooltip">Find in Pipedrive: Settings > Personal > API</span>
      </div>
      <div class="form-group">
        <label for="subdomain">Subdomain</label>
        <input type="text" id="subdomain" value="${savedSubdomain}" placeholder="Enter Pipedrive subdomain" />
        <span class="tooltip">E.g., yourcompany.pipedrive.com (enter only 'yourcompany')</span>
      </div>
    </div>
  </div>
  
  <!-- Data Settings Section -->
  <div class="section">
    <div class="section-title">Data Settings</div>
    <div class="form-row">
      <div class="form-group">
        <label for="entityType">Data Type</label>
        <select id="entityType">
          <option value="deals" ${savedEntityType === 'deals' ? 'selected' : ''}>Deals</option>
          <option value="persons" ${savedEntityType === 'persons' ? 'selected' : ''}>Persons</option>
          <option value="organizations" ${savedEntityType === 'organizations' ? 'selected' : ''}>Organizations</option>
          <option value="activities" ${savedEntityType === 'activities' ? 'selected' : ''}>Activities</option>
          <option value="leads" ${savedEntityType === 'leads' ? 'selected' : ''}>Leads</option>
          <option value="products" ${savedEntityType === 'products' ? 'selected' : ''}>Products</option>
        </select>
        <span class="tooltip">Select the type of data to sync from Pipedrive</span>
      </div>
      <div class="form-group">
        <label>Sync Timestamp</label>
        <div class="toggle-wrapper">
          <label class="toggle-switch">
            <input type="checkbox" id="enableTimestamp" ${savedTimestampEnabled === 'true' ? 'checked' : ''} />
            <span class="toggle-slider"></span>
          </label>
          <label for="enableTimestamp" class="toggle-label">Add timestamp row after each sync</label>
        </div>
        <span class="tooltip">Adds date/time of the last sync at the bottom of the sheet</span>
      </div>
    </div>
  </div>
  
  <!-- Filter Selection Section -->
  <div class="section">
    <div class="section-title">Filter Selection</div>
    <div class="form-group filter-section">
      ${hasFilters ? `
        <div class="filter-selector" id="filterSelectorContainer">
          <input type="text" id="filterSearch" class="filter-search" placeholder="Search filters..." />
          <div class="filter-count" id="filterCount"></div>
          <select id="filterSelector" size="6">
            <option value="">-- Select a filter --</option>
            ${filtersHtml}
          </select>
          <span class="tooltip">Choose a filter from your Pipedrive account</span>
          <span class="filter-toggle" id="showManualFilterInput">Enter filter ID manually instead</span>
        </div>
        
        <div class="filter-input-container hidden" id="manualFilterContainer">
          <input type="text" id="filterId" value="${savedFilterId}" placeholder="Enter Pipedrive filter ID" class="filter-input" />
          <button type="button" id="findFiltersBtn" class="filter-button" title="Find your filters">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor">
              <path d="M0 0h24v24H0V0z" fill="none"/>
              <path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
            </svg>
          </button>
          <span class="filter-toggle hidden" id="showFilterSelector">Choose from filter list instead</span>
        </div>
      ` : `
        <div class="filter-input-container" id="manualFilterContainer">
          <input type="text" id="filterId" value="${savedFilterId}" placeholder="Enter Pipedrive filter ID" class="filter-input" />
          <button type="button" id="findFiltersBtn" class="filter-button" title="Find your filters">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor">
              <path d="M0 0h24v24H0V0z" fill="none"/>
              <path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
            </svg>
          </button>
        </div>
      `}
      <span id="filterTypeWarning" class="note hidden">Warning: Selected filter type doesn't match the data type.</span>
    </div>
  </div>
  
  <div id="status" class="status"></div>
  
  <div class="button-container">
    <button type="button" id="cancelBtn" class="button-secondary">Cancel</button>
    <button type="button" id="saveBtn" class="button-primary"><span id="saveSpinner" class="spinner" style="display:none;"></span>Save Settings</button>
  </div>
</div>

<script>
  // Initialize
  let originalFilterOptions = [];
  document.addEventListener('DOMContentLoaded', function() {
    if (document.getElementById('filterSelector')) {
      originalFilterOptions = Array.from(document.getElementById('filterSelector').querySelectorAll('option[data-type]'));
      }
        
        // Set up event listeners
        document.getElementById('cancelBtn').addEventListener('click', closeDialog);
        document.getElementById('saveBtn').addEventListener('click', saveSettings);
        
        if (document.getElementById('findFiltersBtn')) {
          document.getElementById('findFiltersBtn').addEventListener('click', findFilters);
        }
        
        if (document.getElementById('showManualFilterInput')) {
          document.getElementById('showManualFilterInput').addEventListener('click', function() {
            document.getElementById('filterSelectorContainer').classList.add('hidden');
            document.getElementById('manualFilterContainer').classList.remove('hidden');
            document.getElementById('showManualFilterInput').classList.add('hidden');
            document.getElementById('showFilterSelector').classList.remove('hidden');
          });
        }
        
        if (document.getElementById('showFilterSelector')) {
          document.getElementById('showFilterSelector').addEventListener('click', function() {
            document.getElementById('filterSelectorContainer').classList.remove('hidden');
            document.getElementById('manualFilterContainer').classList.add('hidden');
            document.getElementById('showFilterSelector').classList.add('hidden');
            document.getElementById('showManualFilterInput').classList.remove('hidden');
          });
        }

        // Add highlighting functionality to make selected filter more visible
if (document.getElementById('filterSelector')) {
  const filterSelector = document.getElementById('filterSelector');
  
  // Function to highlight the selected option
  function highlightSelectedOption() {
    const selectedIndex = filterSelector.selectedIndex;
    
    if (selectedIndex > 0) {
      // Reset any previous highlighting
      Array.from(filterSelector.querySelectorAll('option')).forEach(opt => {
        opt.style.background = '';
        opt.style.color = '';
        opt.style.fontWeight = '';
      });
      
      // Add prominent styling to the selected option
      const selectedOption = filterSelector.options[selectedIndex];
      if (selectedOption) {
        selectedOption.style.background = '#4285f4';
        selectedOption.style.color = 'white';
        selectedOption.style.fontWeight = 'bold';
        
        // Try multiple scroll approaches
        try {
          // Approach 1: Use scrollIntoView
          selectedOption.scrollIntoView({ block: 'center', behavior: 'auto' });
          
          // Approach 2: Click the option to ensure it's visible
          setTimeout(() => {
            // Temporary click handler to prevent actual selection change
            const origOnChange = filterSelector.onchange;
            filterSelector.onchange = null;
            
            selectedOption.click();
            selectedOption.focus();
            
            // Restore original handler
            setTimeout(() => {
              filterSelector.onchange = origOnChange;
            }, 10);
          }, 200);
          
          // Approach 3: Focus the select element first, then the option
          setTimeout(() => {
            filterSelector.focus();
            setTimeout(() => {
              selectedOption.focus();
            }, 50);
          }, 300);
        } catch(e) {
          console.log('Error highlighting option:', e);
        }
      }
    }
  }
  
  // Run the highlighting on load with multiple timing attempts
  highlightSelectedOption();
  [100, 300, 500, 1000].forEach(delay => {
    setTimeout(highlightSelectedOption, delay);
  });
  
  // Also add a clear visual indicator text
  const selectedText = filterSelector.options[filterSelector.selectedIndex]?.text || '';
  if (selectedText) {
    const indicatorDiv = document.createElement('div');
    indicatorDiv.style.color = '#4285f4';
    indicatorDiv.style.fontWeight = 'bold';
    indicatorDiv.style.marginTop = '8px';
    indicatorDiv.innerHTML = '✓ Current filter: <strong>' + selectedText + '</strong>';
    filterSelector.parentNode.insertBefore(indicatorDiv, filterSelector.nextSibling);
  }
}

        // Scroll to selected filter in dropdown
        if (document.getElementById('filterSelector')) {
          const filterSelector = document.getElementById('filterSelector');
          
          // Function to calculate and set the proper scroll position
          function scrollToSelectedOption() {
            const selectedIndex = filterSelector.selectedIndex;
    
            if (selectedIndex > 0) {
              // Get all options including those in optgroups
              const allOptions = Array.from(filterSelector.querySelectorAll('option'));
      
              // Find the selected option's position
              if (allOptions[selectedIndex]) {
                // Calculate option height (approximately 24px for most browsers)
                const optionHeight = 24;
        
                // Calculate where to scroll based on the option's position
                // Multiply by index to get approximate position
                const scrollPosition = (selectedIndex * optionHeight) - (filterSelector.clientHeight / 2);
        
                // Set the scroll position directly
                filterSelector.scrollTop = Math.max(0, scrollPosition);
              }
            }
          }
        
          // Try immediately
          scrollToSelectedOption();
        
          // Also try after a delay to ensure UI is fully rendered
          setTimeout(scrollToSelectedOption, 100);
        
          // And once more after a longer delay as a fallback
          setTimeout(scrollToSelectedOption, 500);
        }
        
        if (document.getElementById('filterSelector')) {
          document.getElementById('filterSelector').addEventListener('change', function() {
            if (this.value) {
              document.getElementById('filterId').value = this.value;
              
              // Check if the filter type matches the entity type
              const selectedOption = this.options[this.selectedIndex];
              const filterType = selectedOption.getAttribute('data-type');
              const entityType = document.getElementById('entityType').value;
              
              const filterTypeMap = {
                'deals': 'deals',
                'people': 'persons',
                'org': 'organizations',
                'activity': 'activities',
                'leads': 'leads',
                'products': 'products'
              };
              
              const expectedEntityType = filterTypeMap[filterType] || filterType;
              
              if (expectedEntityType.toLowerCase().trim() !== entityType) {
                document.getElementById('filterTypeWarning').classList.remove('hidden');
              } else {
                document.getElementById('filterTypeWarning').classList.add('hidden');
              }
            }
          });
        }
        
        // The properly formatted version should look like this:
if (document.getElementById('entityType')) {
  // When entityType changes, rebuild the filter selector
document.getElementById('entityType').addEventListener('change', function() {
  if (document.getElementById('filterSelector')) {
    const filterSelector = document.getElementById('filterSelector');
    const selectedEntityType = this.value;
    
    // Store current value if any
    const currentValue = filterSelector.value;
    
    // Map from Pipedrive filter types to our entity types
    const filterTypeMap = {
      'deals': 'deals',
      'people': 'persons',
      'org': 'organizations',
      'activity': 'activities',
      'leads': 'leads',
      'products': 'products'
    };
    
    // Create a new select element that will replace the old one
    const newSelector = document.createElement('select');
    newSelector.id = 'filterSelector';
    newSelector.size = 8;
  
    
    // Keep track of filter types and their options
    const filtersByGroup = {};
    
    // Collect all valid options matching the entity type
    const allOptions = originalFilterOptions;
    let visibleCount = 0;
    
    allOptions.forEach(option => {
      const filterType = option.getAttribute('data-type');
      const expectedEntityType = filterTypeMap[filterType] || filterType;
      
      if (expectedEntityType === selectedEntityType) {
        // This is a matching filter - collect it by its group
        const optgroup = option.closest('optgroup');
        const groupLabel = optgroup ? optgroup.label : null;
        
        if (!filtersByGroup[groupLabel]) {
          filtersByGroup[groupLabel] = [];
        }
        
        // Create a clone of the option
        const newOption = document.createElement('option');
        newOption.value = option.value;
        newOption.textContent = option.textContent;
        newOption.setAttribute('data-type', filterType);
        if (option.value === currentValue) {
          newOption.selected = true;
        }
        
        filtersByGroup[groupLabel].push(newOption);
        visibleCount++;
      }
    });
    
    // Add groups and their options to the new selector
    for (const groupLabel in filtersByGroup) {
      if (groupLabel) {
        const group = document.createElement('optgroup');
        group.label = groupLabel;
        
        filtersByGroup[groupLabel].forEach(option => {
          group.appendChild(option);
        });
        
        newSelector.appendChild(group);
      } else {
        // Direct options without a group
        filtersByGroup[groupLabel].forEach(option => {
          newSelector.appendChild(option);
        });
      }
    }
    
    // Replace the old selector with the new one
    filterSelector.parentNode.replaceChild(newSelector, filterSelector);
    
    // Add change event listener to the new selector
    newSelector.addEventListener('change', function() {
      if (this.value) {
        document.getElementById('filterId').value = this.value;
        
        // Check if the filter type matches the entity type
        const selectedOption = this.options[this.selectedIndex];
        const filterType = selectedOption.getAttribute('data-type');
        const entityType = document.getElementById('entityType').value;
        
        const expectedEntityType = filterTypeMap[filterType] || filterType;
        
        if (expectedEntityType.toLowerCase().trim()  !== entityType) {
          document.getElementById('filterTypeWarning').classList.remove('hidden');
        } else {
          document.getElementById('filterTypeWarning').classList.add('hidden');
        }
      }
    });
    
    // Update filter count
    const countElement = document.getElementById('filterCount');
    if (countElement) {
      countElement.textContent = visibleCount + ' filter(s) available for this data type';
    }
    
    // Reset filter ID if none are valid for this type
    if (visibleCount === 0 && currentValue) {
      document.getElementById('filterId').value = '';
    }
  }
});
  
  // Trigger the change event to filter options on initial load
  document.getElementById('entityType').dispatchEvent(new Event('change'));
}
        
        // Add search functionality to filter dropdown
        if (document.getElementById('filterSearch')) {
          document.getElementById('filterSearch').addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            const filterSelector = document.getElementById('filterSelector');
            const options = filterSelector.querySelectorAll('option');
            const optgroups = filterSelector.querySelectorAll('optgroup');
            let visibleCount = 0;
            
            // First hide all optgroups
            optgroups.forEach(group => {
              group.style.display = 'none';
            });
            
            // Then show options that match the search
            options.forEach(option => {
              if (option.value === '') {
                // Always show the default "Select a filter" option
                option.style.display = '';
                return;
              }
              
              const optionText = option.textContent.toLowerCase();
              const optionId = option.value.toLowerCase();
              
              if (optionText.includes(searchTerm) || optionId.includes(searchTerm)) {
                option.style.display = '';
                // Show the parent optgroup
                const parentGroup = option.parentNode;
                if (parentGroup.tagName === 'OPTGROUP') {
                  parentGroup.style.display = '';
                }
                visibleCount++;
              } else {
                option.style.display = 'none';
              }
            });
            
            // Update filter count
            const countEl = document.getElementById('filterCount');
            if (searchTerm) {
              countEl.textContent = 'Found ' + visibleCount + ' matching filter' + (visibleCount !== 1 ? 's' : '');
            } else {
              countEl.textContent = visibleCount + ' total filters';
            }
          });
          
          // Trigger initial count
          document.getElementById('filterSearch').dispatchEvent(new Event('input'));
        }
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
        const apiKey = document.getElementById('apiKey').value.trim();
        const entityType = document.getElementById('entityType').value;
        const filterId = document.getElementById('filterId').value.trim();
        const subdomain = document.getElementById('subdomain').value.trim();
        const enableTimestamp = document.getElementById('enableTimestamp').checked;
        const sheetName = '${activeSheetName}';
        
        // Validate inputs
        if (!apiKey) {
          showStatus('error', 'API Key is required');
          return;
        }
        
        if (!filterId) {
          showStatus('error', 'Filter ID is required');
          return;
        }
        
        if (!subdomain) {
          showStatus('error', 'Subdomain is required');
          return;
        }
        
        // Store the selected filter option for visual consistency
        const filterSelector = document.getElementById('filterSelector');
        let selectedIndex = -1;
        if (filterSelector) {
          selectedIndex = filterSelector.selectedIndex;
        }
        
        // Show loading state
        const saveBtn = document.getElementById('saveBtn');
        const saveSpinner = document.getElementById('saveSpinner');
        saveBtn.classList.add('loading');
        saveSpinner.style.display = 'inline-block';
        saveBtn.disabled = true;
  
        // Disable all form controls to prevent changes during save
        const formElements = document.querySelectorAll('input, select, button');
        formElements.forEach(el => {
          if (el !== saveBtn) el.disabled = true;
        });
  
        // Save settings via the server-side function
        google.script.run
          .withSuccessHandler(function() {
            // Maintain selected filter visual state
            if (filterSelector && selectedIndex >= 0) {
              filterSelector.selectedIndex = selectedIndex;
            }
      
            // Hide loading state but keep form disabled
            saveBtn.classList.remove('loading');
            saveSpinner.style.display = 'none';
            saveBtn.disabled = false;
            
            showStatus('success', 'Settings saved successfully!');
            setTimeout(closeDialog, 1500);
          })
          .withFailureHandler(function(error) {
            // Re-enable all form controls
            formElements.forEach(el => {
              el.disabled = false;
            });
      
            // Maintain selected filter visual state
            if (filterSelector && selectedIndex >= 0) {
              filterSelector.selectedIndex = selectedIndex;
            }
      
            // Hide loading state
            saveBtn.classList.remove('loading');
            saveSpinner.style.display = 'none';
            saveBtn.disabled = false;
      
            showStatus('error', 'Error: ' + error.message);
          })
          .saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp);
      }
      
      // Open the filter finder dialog
      function findFilters() {
        // Save current settings first
        const apiKey = document.getElementById('apiKey').value.trim();
        const subdomain = document.getElementById('subdomain').value.trim();
        
        if (!apiKey || !subdomain) {
          showStatus('error', 'Please enter your API Key and Subdomain first to find filters');
          return;
        }
        
        const entityType = document.getElementById('entityType').value;
        const sheetName = '${activeSheetName}';
        
        // Save the API key and subdomain temporarily
        google.script.run
          .withSuccessHandler(function() {
            // Then open the filter finder
            google.script.run.showFilterFinder();
          })
          .withFailureHandler(function(error) {
            showStatus('error', 'Error: ' + error.message);
          })
          .saveSettings(apiKey, entityType, "", subdomain, sheetName);
      }
    </script>
  `)
    .setWidth(500)
    .setHeight(650)
    .setTitle(`Pipedrive Settings for "${activeSheetName}"`);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
}

/**
 * Shows a column selector UI for users to choose which columns to display
 */
function showColumnSelector() {
  try {
    // Get the active sheet name
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    const addedCustomFields = new Set();

    // Get API key and filter ID from properties or use defaults
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;

    // Get sheet-specific settings
    const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;

    const filterId = scriptProperties.getProperty(sheetFilterIdKey) || FILTER_ID;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Use the active sheet for the operation
    scriptProperties.setProperty('SHEET_NAME', activeSheetName);

    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      SpreadsheetApp.getUi().alert('Please configure your Pipedrive API Key in Settings first.');
      showSettings();
      return;
    }

    // Check if we can connect to Pipedrive and get a sample item
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

    if (sampleData && sampleData[0] && sampleData[0].custom_fields) {
      Logger.log('Custom fields in sample data:');
      for (const key in sampleData[0].custom_fields) {
        Logger.log(`Sample custom field: ${key}`);
      }
    }

    if (!sampleData || sampleData.length === 0) {
      SpreadsheetApp.getUi().alert(`Could not retrieve any ${entityType} from Pipedrive. Please check your API key and filter ID.`);
      return;
    }

    // Log the raw sample data for debugging
    Logger.log(`PIPEDRIVE DEBUG - Sample ${entityType} raw data:`);
    Logger.log(JSON.stringify(sampleData[0], null, 2));

    // Get all available columns from the sample data
    const sampleItem = sampleData[0];
    const availableColumns = [];

    // Get field definitions to show friendly names
    let fieldDefinitions = [];
    let customFieldMap = {}; // Add this line

    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fieldDefinitions = getDealFields();
        customFieldMap = getCustomFieldMappings(ENTITY_TYPES.DEALS);
        break;
      case ENTITY_TYPES.PERSONS:
        fieldDefinitions = getPersonFields();
        customFieldMap = getCustomFieldMappings(ENTITY_TYPES.PERSONS);
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fieldDefinitions = getOrganizationFields();
        customFieldMap = getCustomFieldMappings(ENTITY_TYPES.ORGANIZATIONS);
        break;
      case ENTITY_TYPES.ACTIVITIES:
        fieldDefinitions = getActivityFields();
        customFieldMap = getCustomFieldMappings(ENTITY_TYPES.ACTIVITIES);
        break;
      case ENTITY_TYPES.LEADS:
        fieldDefinitions = getLeadFields();
        customFieldMap = getCustomFieldMappings(ENTITY_TYPES.LEADS);
        break;
      case ENTITY_TYPES.PRODUCTS:
        fieldDefinitions = getProductFields();
        customFieldMap = getProductFieldOptionMappings();
        break;
    }

    const fieldMap = {};
    fieldDefinitions.forEach(field => {
      fieldMap[field.key] = field.name;
    });

    // Function to recursively extract fields from complex objects
    function extractFields(obj, parentPath = '', parentName = '') {
      // Scan all sample items for custom fields
      Logger.log(`Scanning all ${sampleData.length} sample items for custom fields with hash keys...`);
      const hashPattern = /^[0-9a-f]{40}$/;
      const processedHashes = new Set();

      // Look through all sample items to find hash keys
      sampleData.forEach(item => {
        for (const key in item) {
          if (hashPattern.test(key) && customFieldMap[key] && !addedCustomFields.has(key) && !processedHashes.has(key)) {
            processedHashes.add(key);
            Logger.log(`Found hash key in sample data: ${key} => ${customFieldMap[key]}`);

            // Add this custom field with its friendly name
            availableColumns.push({
              key: key,
              name: `${customFieldMap[key]}`,
              isNested: false
            });

            addedCustomFields.add(key);

            // If it's a complex object, also add fields for its properties
            const customValue = item[key];
            if (customValue && typeof customValue === 'object') {
              for (const propKey in customValue) {
                if (propKey !== 'value') { // Skip the main value property
                  availableColumns.push({
                    key: `${key}_${propKey}`,
                    name: `${customFieldMap[key]} - ${propKey}`,
                    isNested: true,
                    parentKey: key
                  });
                }
              }
            }
          }
        }
      });

      // Check for hash pattern keys at the root level
      if (parentPath === '') {  // Only check at the root level
        for (const key in obj) {
          if (hashPattern.test(key) && customFieldMap[key] && !addedCustomFields.has(key)) {
            Logger.log(`Hash pattern match: ${key} => ${customFieldMap[key]}`);

            addedCustomFields.add(key);
            // This is a custom field hash that we have a mapping for
            const friendlyName = customFieldMap[key];
            const customValue = obj[key];

            // Add entry for this custom field with friendly name
            availableColumns.push({
              key: key,
              name: `${friendlyName}`,
              isNested: false
            });

            // If it's a complex object, also add fields for its properties
            if (customValue && typeof customValue === 'object') {
              // Handle complex fields like persons, addresses, etc.
              for (const propKey in customValue) {
                if (propKey !== 'value') { // Skip the main value property
                  availableColumns.push({
                    key: `${key}_${propKey}`,
                    name: `${friendlyName} - ${propKey}`,
                    isNested: true,
                    parentKey: key
                  });
                }
              }
            }
          }
        }
      }

      // Special handling for custom_fields in API v2
      if (parentPath === 'custom_fields') {
        for (const key in obj) {

          Logger.log(`Processing custom_fields key: ${key}`);
          Logger.log(`Custom field mapping exists?: ${Boolean(customFieldMap[key])}`);
          if (customFieldMap[key]) {
            Logger.log(`Mapping found: ${key} => ${customFieldMap[key]}`);
          } else {
            Logger.log(`MISSING MAPPING: ${key} (NOT FOUND IN MAPPING)`);

            // Try with trimmed key
            const trimmedKey = key.trim();
            if (customFieldMap[trimmedKey]) {
              Logger.log(`FOUND with trimming: ${trimmedKey} => ${customFieldMap[trimmedKey]}`);
            }

            // Log all the keys in customFieldMap that start with the same 5 characters
            const prefix = key.substring(0, 5);
            const similarKeys = Object.keys(customFieldMap).filter(k => k.startsWith(prefix));
            if (similarKeys.length > 0) {
              Logger.log(`Similar keys in mapping: ${similarKeys.join(', ')}`);
            }
          }
          if (addedCustomFields.has(key)) {
            // Skip this custom field as it's already been added
            continue;
          }
          addedCustomFields.add(key);
          const customFieldValue = obj[key];
          const currentPath = `custom_fields.${key}`;

          // Add this logging
          if (!customFieldMap[key]) {
            Logger.log(`MISSING MAPPING: Custom field key ${key} has no friendly name mapping`);
          }

          // Simple value custom fields (text, date, number, single option)
          if (typeof customFieldValue !== 'object' || customFieldValue === null) {
            availableColumns.push({
              key: currentPath,
              name: `${customFieldMap[key] || key}`,
              isNested: true,
              parentKey: 'custom_fields'
            });
            Logger.log(`Added to availableColumns: ${currentPath} => ${customFieldMap[key] || '(using key fallback)'}`);

            continue;
          }

          if (customFieldMap) {
            Logger.log(`DEBUGGING: Custom field map keys (first 10): ${Object.keys(customFieldMap).slice(0, 10).join(', ')}`);
            Logger.log(`DEBUGGING: Sample mapping: ${Object.keys(customFieldMap)[0]} => ${customFieldMap[Object.keys(customFieldMap)[0]]}`);
          }

          // Handle complex custom fields (object-based)
          if (typeof customFieldValue === 'object') {
            // Currency fields
            if (customFieldValue.value !== undefined && customFieldValue.currency !== undefined) {
              availableColumns.push({
                key: currentPath,
                name: `${customFieldMap[key] || key} (Currency)`,
                isNested: true,
                parentKey: 'custom_fields'
              });
            }
            // Date/Time range fields
            else if (customFieldValue.value !== undefined && customFieldValue.until !== undefined) {
              availableColumns.push({
                key: currentPath,
                name: `${customFieldMap[key] || key} (Range)`,
                isNested: true,
                parentKey: 'custom_fields'
              });
            }
            // Address fields
            else if (customFieldValue.value !== undefined && customFieldValue.formatted_address !== undefined) {
              availableColumns.push({
                key: currentPath,
                name: `${customFieldMap[key] || key} (Address)`,
                isNested: true,
                parentKey: 'custom_fields'
              });

              // Add formatted address as a separate column option
              availableColumns.push({
                key: `${currentPath}.formatted_address`,
                name: `${customFieldMap[key] || key} (Formatted Address)`,
                isNested: true,
                parentKey: currentPath
              });
            }
            // For all other object types
            else {
              availableColumns.push({
                key: currentPath,
                name: `${customFieldMap[key] || key} (Complex)`,
                isNested: true,
                parentKey: 'custom_fields'
              });

              // Extract nested fields from complex custom field
              extractFields(customFieldValue, currentPath, `Custom Field: ${key}`);
            }
          }
        }
        return;
      }

      // Handle arrays by looking at the first item
      if (Array.isArray(obj)) {
        // If it's an array of objects with a common structure, extract from first item
        if (obj.length > 0 && typeof obj[0] === 'object' && obj[0] !== null) {
          // Special handling for multiple options fields
          if (obj[0].hasOwnProperty('label') && obj[0].hasOwnProperty('id')) {
            // Add a field for the entire multiple options array
            availableColumns.push({
              key: parentPath,
              name: parentName + ' (Multiple Options)',
              isNested: true,
              parentKey: parentPath.split('.').slice(0, -1).join('.')
            });
            return;
          }

          // For arrays of structured objects, like emails or phones
          if (obj[0].hasOwnProperty('value') && obj[0].hasOwnProperty('primary')) {
            let displayName = 'Primary ' + (parentName || 'Item');
            if (obj[0].hasOwnProperty('label')) {
              displayName = 'Primary ' + parentName + ' (' + obj[0].label + ')';
            }

            availableColumns.push({
              key: parentPath + '.0.value',
              name: displayName,
              isNested: true,
              parentKey: parentPath
            });
          } else {
            extractFields(obj[0], parentPath + '.0', parentName + ' (First Item)');
          }
        }
        return;
      }

      // Extract properties from this object
      for (const key in obj) {
        // Skip internal properties, functions, or empty objects
        if (key.startsWith('_') || typeof obj[key] === 'function') {
          continue;
        }

        const currentPath = parentPath ? parentPath + '.' + key : key;

        // For common Pipedrive objects, create shortcuts
        if (key === 'name' && parentPath && (obj.hasOwnProperty('email') || obj.hasOwnProperty('phone') || obj.hasOwnProperty('address'))) {
          // For person or organization name
          availableColumns.push({
            key: currentPath,
            name: (parentName || parentPath) + ' Name',
            isNested: true,
            parentKey: parentPath
          });
        } else if (typeof obj[key] === 'object' && obj[key] !== null) {
          // Recursively extract from nested objects
          let nestedParentName = parentName ? parentName + ' ' + formatColumnName(key) : formatColumnName(key);
          extractFields(obj[key], currentPath, nestedParentName);
        } else {
          // Simple property
          const displayName = parentName ? parentName + ' ' + formatColumnName(key) : formatColumnName(key);
          availableColumns.push({
            key: currentPath,
            name: displayName,
            isNested: parentPath ? true : false,
            parentKey: parentPath
          });
        }
      }
    }

    // Build top-level column data first
    for (const key in sampleItem) {
      // Skip internal properties or functions
      if (key.startsWith('_') || typeof sampleItem[key] === 'function') {
        continue;
      }

      const displayName = fieldMap[key] || formatColumnName(key);

      // Add the top-level column
      availableColumns.push({
        key: key,
        name: displayName,
        isNested: false
      });

      // If it's a complex object, extract nested fields
      if (typeof sampleItem[key] === 'object' && sampleItem[key] !== null) {
        extractFields(sampleItem[key], key, displayName);
      }
    }

    // Log all extracted columns for debugging
    Logger.log(`PIPEDRIVE DEBUG - All available columns for ${entityType}:`);
    Logger.log(JSON.stringify(availableColumns, null, 2));

    // Add this just before the columns are rendered to the UI
    Logger.log(`Final availableColumns length: ${availableColumns.length}`);
    Logger.log(`First 10 columns: ${availableColumns.slice(0, 10).map(c => c.name).join(', ')}`);

    // Get previously saved column preferences for this specific sheet and entity type
    const columnSettingsKey = `COLUMNS_${activeSheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let selectedColumns = [];

    if (savedColumnsJson) {
      try {
        selectedColumns = JSON.parse(savedColumnsJson);
        // Filter out any columns that no longer exist
        selectedColumns = selectedColumns.filter(col =>
          availableColumns.some(availCol => availCol.key === col.key)
        );
      } catch (e) {
        Logger.log('Error parsing saved columns: ' + e.message);
        selectedColumns = [];
      }
    }

    Logger.log('Custom field entries in availableColumns:');
    const customFieldEntries = availableColumns.filter(col => col.name.startsWith('Custom Field:'));
    customFieldEntries.forEach(entry => {
      Logger.log(`Column entry: key=${entry.key}, name=${entry.name}`);
    });

    // Show the column selector UI
    showColumnSelectorUI(availableColumns, selectedColumns, entityType, activeSheetName);
  } catch (error) {
    Logger.log('Error in column selector: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
  }
}

/**
 * Creates a mapping between custom field keys and their human-readable names
 * @param {string} entityType The entity type (deals, persons, etc.)
 * @return {Object} Mapping of custom field keys to field names
 */
function getCustomFieldMappings(entityType) {
  Logger.log(`Getting custom field mappings for ${entityType}`);

  let customFields = getCustomFieldsForEntity(entityType);
  const map = {};

  Logger.log(`Retrieved ${customFields.length} custom fields for ${entityType}`);

  customFields.forEach(field => {
    map[field.key] = field.name;
    // Log each mapping for debugging
    Logger.log(`Custom field mapping: ${field.key} => ${field.name}`);
  });
  Logger.log(`Total custom field mappings: ${Object.keys(map).length}`);

  return map;
}

/**
 * Checks if a field key is likely to be a date field
 * @param {string} fieldKey The field key to check
 * @return {boolean} True if the field is likely a date field
 */
function isDateField(fieldKey) {
  // Common date field patterns in Pipedrive
  const dateFieldPatterns = [
    'date',
    '_at',
    '_time',
    'birthday',
    'deadline',
    'start_date',
    'end_date',
    'due_date',
    'expiry'
  ];

  // Convert the field key to lowercase for case-insensitive matching
  const lowerFieldKey = fieldKey.toLowerCase();

  // Check if the field key contains any of the date patterns
  return dateFieldPatterns.some(pattern => lowerFieldKey.includes(pattern));
}

/**
 * Detects if a value might be a date and converts it to YYYY-MM-DD format for Pipedrive
 * @param {any} value The value to check and potentially convert
 * @return {any} The original value or converted date string
 */
function convertToStandardDateFormat(value) {
  // Skip if not a string or already in the correct format (YYYY-MM-DD)
  if (typeof value !== 'string' && !(value instanceof Date)) {
    return value;
  }

  if (typeof value === 'string') {
    // Skip if it's already in YYYY-MM-DD format
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return value;
    }

    // Skip if it's definitely not a date
    if (!/\d/.test(value)) {
      return value;
    }
  }

  try {
    // Try to parse the date using Date.parse
    const dateObj = new Date(value);

    // Check if it's a valid date
    if (isNaN(dateObj.getTime())) {
      return value;
    }

    // Format as YYYY-MM-DD
    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  } catch (e) {
    // If anything goes wrong, return the original value
    return value;
  }
}

/**
 * Shows the column selector UI
 */
function showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName) {
  const htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
      
      .header {
        margin-bottom: 16px;
      }
      
      h3 {
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 12px;
        color: var(--text-dark);
      }
      
      .sheet-info {
        background-color: var(--bg-light);
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 14px;
        border-left: 4px solid var(--primary-color);
        display: flex;
        align-items: center;
      }
      
      .sheet-info svg {
        margin-right: 12px;
        fill: var(--primary-color);
      }
      
      .info {
        font-size: 13px;
        color: var(--text-light);
        margin-bottom: 16px;
      }
      
      .container {
        display: flex;
        height: 400px;
        gap: 16px;
      }
      
      .column {
        width: 50%;
        display: flex;
        flex-direction: column;
      }
      
      .column-header {
        font-weight: 500;
        margin-bottom: 8px;
        padding: 0 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .column-count {
        font-size: 12px;
        color: var(--text-light);
        font-weight: normal;
      }
      
      .search {
        padding: 10px 12px;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        margin-bottom: 12px;
        font-size: 14px;
        width: 100%;
        transition: var(--transition);
      }
      
      .search:focus {
        outline: none;
        border-color: var(--primary-color);
        box-shadow: 0 0 0 2px rgba(66,133,244,0.2);
      }
      
      .scrollable {
        flex-grow: 1;
        overflow-y: auto;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        background-color: white;
      }
      
      .item {
        padding: 8px 12px;
        margin: 2px 4px;
        cursor: pointer;
        border-radius: 4px;
        transition: var(--transition);
        display: flex;
        align-items: center;
      }
      
      .item:hover {
        background-color: rgba(66,133,244,0.08);
      }
      
      .item.selected {
        background-color: #e8f0fe;
        position: relative;
      }
      
      .item.selected:hover {
        background-color: #d2e3fc;
      }
      
      .category {
        font-weight: 500;
        padding: 8px 12px;
        background-color: var(--bg-light);
        margin: 8px 4px 4px 4px;
        border-radius: 4px;
        color: var(--text-dark);
        font-size: 13px;
      }
      
      .nested {
        margin-left: 16px;
        position: relative;
      }
      
      .nested::before {
        content: '';
        position: absolute;
        left: -8px;
        top: 50%;
        width: 6px;
        height: 1px;
        background-color: var(--border-color);
      }
      
      .footer {
        margin-top: 16px;
        display: flex;
        justify-content: space-between;
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
      
      .action-btns {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      
      .drag-handle {
        display: inline-block;
        width: 12px;
        height: 20px;
        background-image: radial-gradient(circle, var(--text-light) 1px, transparent 1px);
        background-size: 3px 3px;
        background-position: center;
        background-repeat: repeat;
        margin-right: 8px;
        cursor: grab;
        opacity: 0.5;
      }
      
      .selected:hover .drag-handle {
        opacity: 0.8;
      }
      
      .column-text {
        flex-grow: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      
      .add-btn, .remove-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
      }
      
      .add-btn {
        color: var(--success-color);
        background-color: rgba(15,157,88,0.1);
      }

      .rename-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
        color: var(--warning-color);
        background-color: rgba(244,180,0,0.1);
      }

      .item:hover .rename-btn {
        opacity: 1;
      }
      
      .remove-btn {
        color: var(--error-color);
        background-color: rgba(219,68,55,0.1);
      }
      
      .item:hover .add-btn,
      .item:hover .remove-btn {
        opacity: 1;
      }
      
      .loading {
        display: none;
        align-items: center;
        margin-right: 8px;
      }
      
      .loader {
        display: inline-block;
        width: 18px;
        height: 18px;
        border: 2px solid rgba(66,133,244,0.2);
        border-radius: 50%;
        border-top-color: var(--primary-color);
        animation: spin 1s ease-in-out infinite;
      }
      
      .dragging {
        opacity: 0.4;
        box-shadow: var(--shadow-hover);
      }
      
      .over {
        border-top: 2px solid var(--primary-color);
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
      
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
    
    <div class="header">
      <h3>Column Selection</h3>
      
      <div id="statusIndicator" class="indicator"></div>
      
      <div class="sheet-info">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24">
          <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14z"/>
          <path d="M7 12h10v2H7zm0-4h10v2H7zm0 8h7v2H7z"/>
        </svg>
        <div>
          Configuring columns for <strong>${entityType}</strong> in sheet <strong>"${sheetName}"</strong>
        </div>
      </div>
      
      <p class="info">
        Select which Pipedrive data columns to display in your sheet. Drag items in the right panel to change their order.
      </p>
    </div>
    
    <div class="container">
      <div class="column">
        <input type="text" id="searchBox" class="search" placeholder="Search for columns...">
        <div class="column-header">
          Available Columns <span id="availableCount" class="column-count"></span>
        </div>
        <div id="availableList" class="scrollable">
          <!-- Available columns will be populated here by JavaScript -->
        </div>
      </div>
      
      <div class="column">
        <div class="column-header">
          Selected Columns <span id="selectedCount" class="column-count"></span>
        </div>
        <div id="selectedList" class="scrollable">
          <!-- Selected columns will be populated here by JavaScript -->
        </div>
      </div>
    </div>
    
    <div class="footer">
      <div class="action-btns">
        <button class="secondary-btn" id="helpBtn">Help & Tips</button>
      </div>
      <div class="action-btns">
        <div class="loading" id="saveLoading">
          <span class="loader"></span>
        </div>
        <button class="secondary-btn" id="cancelBtn">Cancel</button>
        <button class="primary-btn" id="saveBtn">Save & Close</button>
      </div>
    </div>

    <script>
      // Initialize data
      let availableColumns = ${JSON.stringify(availableColumns)};
      let selectedColumns = ${JSON.stringify(selectedColumns)};
      const entityType = "${entityType}";
      const sheetName = "${sheetName}";
      
      // DOM elements
      const availableList = document.getElementById('availableList');
      const selectedList = document.getElementById('selectedList');
      const searchBox = document.getElementById('searchBox');
      const availableCountEl = document.getElementById('availableCount');
      const selectedCountEl = document.getElementById('selectedCount');
      
      // Render the lists
      function renderAvailableList(searchTerm = '') {
        availableList.innerHTML = '';
        
        // Group columns by parent key or top-level
        const topLevel = [];
        const nested = {};
        let availableCount = 0;
        
        availableColumns.forEach(col => {
          if (!selectedColumns.some(selected => selected.key === col.key)) {
            availableCount++;
            if (col.name.toLowerCase().includes(searchTerm.toLowerCase())) {
              if (!col.isNested) {
                topLevel.push(col);
              } else {
                const parentKey = col.parentKey || 'unknown';
                if (!nested[parentKey]) {
                  nested[parentKey] = [];
                }
                nested[parentKey].push(col);
              }
            }
          }
        });
        
        // Update available count
        availableCountEl.textContent = '(' + availableCount + ')';
        
        // Add top-level columns first
        if (topLevel.length > 0) {
          const topLevelHeader = document.createElement('div');
          topLevelHeader.className = 'category';
          topLevelHeader.textContent = 'Main Fields';
          availableList.appendChild(topLevelHeader);
          
          topLevel.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.dataset.key = col.key;
            
            const columnText = document.createElement('span');
            columnText.className = 'column-text';
            columnText.textContent = col.name;
            item.appendChild(columnText);
            
            const addBtn = document.createElement('span');
            addBtn.className = 'add-btn';
            addBtn.innerHTML = '&plus;';
            addBtn.title = 'Add column';
            addBtn.onclick = (e) => {
              e.stopPropagation();
              addColumn(col);
            };
            item.appendChild(addBtn);
            
            item.onclick = () => addColumn(col);
            availableList.appendChild(item);
          });
        }
        
        // Then add nested columns by parent
        for (const parentKey in nested) {
          if (nested[parentKey].length > 0) {
            const parentName = availableColumns.find(col => col.key === parentKey)?.name || parentKey;
            
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'category';
            categoryHeader.textContent = parentName;
            availableList.appendChild(categoryHeader);
            
            nested[parentKey].forEach(col => {
              const item = document.createElement('div');
              item.className = 'item nested';
              item.dataset.key = col.key;
              
              const columnText = document.createElement('span');
              columnText.className = 'column-text';
              columnText.textContent = col.name;
              item.appendChild(columnText);
              
              const addBtn = document.createElement('span');
              addBtn.className = 'add-btn';
              addBtn.innerHTML = '&plus;';
              addBtn.title = 'Add column';
              addBtn.onclick = (e) => {
                e.stopPropagation();
                addColumn(col);
              };
              item.appendChild(addBtn);
              
              item.onclick = () => addColumn(col);
              availableList.appendChild(item);
            });
          }
        }
        
        // Show "no results" message if nothing matches the search
        if (availableList.children.length === 0 && searchTerm) {
          const noResults = document.createElement('div');
          noResults.style.padding = '16px';
          noResults.style.textAlign = 'center';
          noResults.style.color = 'var(--text-light)';
          noResults.textContent = 'No matching columns found';
          availableList.appendChild(noResults);
        }
      }
      
      function renderSelectedList() {
        selectedList.innerHTML = '';
        
        // Update selected count
        selectedCountEl.textContent = '(' + selectedColumns.length + ')';
        
        if (selectedColumns.length === 0) {
          const emptyState = document.createElement('div');
          emptyState.style.padding = '16px';
          emptyState.style.textAlign = 'center';
          emptyState.style.color = 'var(--text-light)';
          emptyState.innerHTML = 'No columns selected yet<br>Select columns from the left panel';
          selectedList.appendChild(emptyState);
          return;
        }
        
        selectedColumns.forEach((col, index) => {
          const item = document.createElement('div');
          item.className = 'item selected';
          item.dataset.key = col.key;
          item.dataset.index = index;
          item.draggable = true;
          
          // Add drag handle
          const dragHandle = document.createElement('span');
          dragHandle.className = 'drag-handle';
          item.appendChild(dragHandle);
          
          // Add column name
          const columnText = document.createElement('span');
          columnText.className = 'column-text';
          columnText.textContent = col.customName || col.name;
          if (col.customName) {
            columnText.textContent = col.customName;
            columnText.title = "Original field: " + col.name;
            columnText.style.fontStyle = 'italic';
          }
          item.appendChild(columnText);
          
          // Add remove button
          const removeBtn = document.createElement('span');
          removeBtn.className = 'remove-btn';
          removeBtn.innerHTML = '❌';
          removeBtn.title = 'Remove column';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeColumn(col);
          };
          item.appendChild(removeBtn);
          
          // Add rename button
          const renameBtn = document.createElement('span');
          renameBtn.className = 'rename-btn';
          renameBtn.innerHTML = '✏️';
          renameBtn.title = 'Rename column';
          renameBtn.onclick = (e) => {
            e.stopPropagation();
            renameColumn(index);
          };
          item.appendChild(renameBtn);
          
          // Set up drag events
          item.ondragstart = handleDragStart;
          item.ondragover = handleDragOver;
          item.ondrop = handleDrop;
          item.ondragend = handleDragEnd;
          
          selectedList.appendChild(item);
        });
      }
      
      // Drag and drop functionality
      let draggedItem = null;
      
      function handleDragStart(e) {
        draggedItem = this;
        this.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', this.dataset.index);
        
        // Add a small delay to make the visual change noticeable
        setTimeout(() => {
          this.style.opacity = '0.4';
        }, 0);
      }
      
      function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this.classList.add('over');
      }
      
      function handleDrop(e) {
        e.preventDefault();
        this.classList.remove('over');
        
        const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
        const toIndex = parseInt(this.dataset.index);
        
        if (fromIndex !== toIndex) {
          const item = selectedColumns[fromIndex];
          selectedColumns.splice(fromIndex, 1);
          selectedColumns.splice(toIndex, 0, item);
          renderSelectedList();
        }
      }
      
      function handleDragEnd() {
        this.classList.remove('dragging');
        document.querySelectorAll('.item').forEach(item => {
          item.classList.remove('over');
        });
      }
      
      // Column management
      function addColumn(column) {
        selectedColumns.push(column);
        renderAvailableList(searchBox.value);
        renderSelectedList();
        showStatus('success', 'Column added: ' + column.name);
      }
      
      function removeColumn(column) {
        selectedColumns = selectedColumns.filter(col => col.key !== column.key);
        renderAvailableList(searchBox.value);
        renderSelectedList();
        showStatus('success', 'Column removed: ' + column.name);
      }

      function renameColumn(columnIndex) {
        const column = selectedColumns[columnIndex];
        const currentName = column.customName || column.name;
        const newName = prompt("Enter custom header name for '" + column.name + "'", currentName);
        
        if (newName !== null) {
          // Update the column with the custom name
          selectedColumns[columnIndex].customName = newName;
          renderSelectedList();
          showStatus('success', 'Column renamed to: ' + newName);
        }
      }
      
      function showStatus(type, message) {
        const indicator = document.getElementById('statusIndicator');
        indicator.className = 'indicator ' + type;
        indicator.textContent = message;
        indicator.style.display = 'block';
        
        // Auto-hide after a delay
        setTimeout(function() {
          indicator.style.display = 'none';
        }, 2000);
      }
      
      // Event listeners
      document.getElementById('saveBtn').onclick = () => {
        if (selectedColumns.length === 0) {
          showStatus('error', 'Please select at least one column');
          return;
        }
        
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'flex';
        document.getElementById('saveBtn').disabled = true;
        document.getElementById('cancelBtn').disabled = true;
        
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('saveLoading').style.display = 'none';
            showStatus('success', 'Column preferences saved successfully!');
            
            // Close after a short delay to show the success message
            setTimeout(() => {
              google.script.host.close();
            }, 1500);
          })
          .withFailureHandler((error) => {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            document.getElementById('cancelBtn').disabled = false;
            showStatus('error', 'Error: ' + error.message);
          })
          .saveColumnPreferences(selectedColumns, entityType, sheetName);
      };
      
      document.getElementById('cancelBtn').onclick = () => {
        google.script.host.close();
      };
      
      document.getElementById('helpBtn').onclick = () => {
        const helpContent = 'Tips for selecting columns:\\n' +
                           '\\n• Search for specific columns using the search box' +
                           '\\n• Main fields are top-level Pipedrive fields' +
                           '\\n• Nested fields provide more detailed information' +
                           '\\n• Drag and drop to reorder selected columns' +
                           '\\n• The column order here determines the order in your sheet';
                           
        alert(helpContent);
      };
      
      searchBox.oninput = () => {
        renderAvailableList(searchBox.value);
      };
      
      // Initial render
      renderAvailableList();
      renderSelectedList();
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(800)
    .setHeight(600)
    .setTitle(`Select Columns for ${entityType} in "${sheetName}"`);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Select Columns for ${entityType} in "${sheetName}"`);
}

/**
 * Saves column preferences to script properties
 */
function saveColumnPreferences(columns, entityType, sheetName) {
  // Save using the team-aware function
  saveTeamAwareColumnPreferences(columns, entityType, sheetName);
}

/**
 * Saves settings to script properties
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp = false) {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Save global settings (API key and subdomain are global)
  scriptProperties.setProperty('PIPEDRIVE_API_KEY', apiKey);
  scriptProperties.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);

  // Save sheet-specific settings
  const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  const timestampEnabledKey = `TIMESTAMP_ENABLED_${sheetName}`;

  scriptProperties.setProperty(sheetFilterIdKey, filterId);
  scriptProperties.setProperty(sheetEntityTypeKey, entityType);
  scriptProperties.setProperty(timestampEnabledKey, enableTimestamp.toString());
  scriptProperties.setProperty('SHEET_NAME', sheetName);
}

/**
 * Main function to sync deals from a Pipedrive filter to the Google Sheet
 */
function syncDealsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.DEALS, skipPush);
}

/**
 * Main function to sync persons from a Pipedrive filter to the Google Sheet
 */
function syncPersonsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PERSONS, skipPush);
}

/**
 * Main function to sync organizations from a Pipedrive filter to the Google Sheet
 */
function syncOrganizationsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ORGANIZATIONS, skipPush);
}

/**
 * Main function to sync activities from a Pipedrive filter to the Google Sheet
 */
function syncActivitiesFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.ACTIVITIES, skipPush);
}

/**
 * Main function to sync leads from a Pipedrive filter to the Google Sheet
 */
function syncLeadsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.LEADS, skipPush);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 */
function syncProductsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush);
}

/**
 * Gets products using a specific filter ID
 * @param {string} filterId The Pipedrive filter ID
 * @param {number} limit The maximum number of products to fetch
 * @returns {Array} Array of products
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
}

/**
 * Gets all available product fields from Pipedrive
 * @returns {Array} Array of product fields
 */
function getProductFields() {
  // Note: Product fields use API v1 according to the documentation
  const url = getPipedriveApiUrl().replace('/v2', '/v1') + '/productFields?limit=500';
  const fields = getCustomFieldsForEntity(ENTITY_TYPES.PRODUCTS);
  return fields;
}

/**
 * Gets mappings for product custom fields
 * @returns {Object} Object mapping custom field keys to names
 */
function getProductFieldOptionMappings() {
  return getFieldOptionMappingsForEntity(ENTITY_TYPES.PRODUCTS);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 */
function syncProductsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush);
}

/**
 * Gets products using a specific filter ID
 * @param {string} filterId The Pipedrive filter ID
 * @param {number} limit The maximum number of products to fetch
 * @returns {Array} Array of products
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
}

/**
 * Gets all available product fields from Pipedrive
 * @returns {Array} Array of product fields
 */
function getProductFields() {
  // Note: Product fields use API v1 according to the documentation
  const url = getPipedriveApiUrl().replace('/v2', '/v1') + '/productFields?limit=500';
  const fields = getCustomFieldsForEntity(ENTITY_TYPES.PRODUCTS);
  return fields;
}

/**
 * Gets mappings for product custom fields
 * @returns {Object} Object mapping custom field keys to names
 */
function getProductFieldOptionMappings() {
  return getFieldOptionMappingsForEntity(ENTITY_TYPES.PRODUCTS);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 */
function syncProductsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush);
}

/**
 * Gets products using a specific filter ID
 * @param {string} filterId The Pipedrive filter ID
 * @param {number} limit The maximum number of products to fetch
 * @returns {Array} Array of products
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
}

/**
 * Generic function to sync any Pipedrive entity to a Google Sheet
 */
function syncPipedriveDataToSheet(entityType, skipPush = false) {
  try {
    // Update sync status to Phase 1 (connecting)
    updateSyncStatus(1, 'active', 'Connecting to Pipedrive...', 50);

    // Get API key and filter ID from properties or use defaults
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    const sheetName = scriptProperties.getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;

    // Check if two-way sync is enabled for this sheet
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is enabled, push changes to Pipedrive first
    if (twoWaySyncEnabled && !skipPush) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Pushing modified rows to Pipedrive first.`);

      try {
        // Call pushChangesToPipedrive() with false for isScheduledSync to show UI feedback
        pushChangesToPipedrive(true, true);
      } catch (error) {
        // Log the error but continue with the sync
        Logger.log(`Error pushing changes: ${error.message}`);
        spreadsheet.toast(`Warning: Error pushing changes to Pipedrive: ${error.message}`, 'Sync Warning', 10);
      }
    }

    // Get sheet-specific filter ID
    const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
    const filterId = scriptProperties.getProperty(sheetFilterIdKey) || FILTER_ID;

    // Check if API key is configured
    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      updateSyncStatus(1, 'error', 'Pipedrive API Key not configured. Please check settings.', 0);

      const ui = SpreadsheetApp.getUi();
      ui.alert('Pipedrive API Key Required',
        'Please configure your Pipedrive API Key in Settings first.',
        ui.ButtonSet.OK);
      showSettings();
      return;
    }

    // Check if filter ID is configured
    if (!filterId) {
      updateSyncStatus(1, 'error', 'Filter ID not configured. Please check settings.', 0);

      const ui = SpreadsheetApp.getUi();
      ui.alert('Filter ID Required',
        'Please configure a Pipedrive Filter ID in Settings first.',
        ui.ButtonSet.OK);
      showSettings();
      return;
    }

    // Mark phase 1 as complete
    updateSyncStatus(1, 'completed', 'Connected to Pipedrive', 100);

    // Get saved column preferences for this specific sheet and entity type
    let selectedColumns = getTeamAwareColumnPreferences(entityType, sheetName);

    // If no columns are selected, prompt to configure
    if (selectedColumns.length === 0) {
      updateSyncStatus(2, 'error', 'No columns selected for display. Please configure columns first.', 0);

      const ui = SpreadsheetApp.getUi();
      const response = ui.alert('No Columns Selected',
        `You need to select which columns to display for ${entityType} in "${sheetName}". Would you like to configure them now?`,
        ui.ButtonSet.YES_NO);

      if (response === ui.Button.YES) {
        showColumnSelector();
      }

      return;
    }

    // Start phase 2 (retrieving data)
    updateSyncStatus(2, 'active', `Retrieving ${entityType} data from Pipedrive...`, 10);

    // Create column mapping for the export
    const columnsToUse = selectedColumns.map(col => col.key);
    const headerRow = selectedColumns.map(col => col.customName || col.name);

    // Get option mappings for multiple option fields
    const optionMappings = getFieldOptionMappingsForEntity(entityType);

    // Get all data for the export based on entity type
    let items = [];
    try {
      switch (entityType) {
        case ENTITY_TYPES.DEALS:
          items = getDealsWithFilter(filterId, 0);
          break;
        case ENTITY_TYPES.PERSONS:
          items = getPersonsWithFilter(filterId, 0);
          break;
        case ENTITY_TYPES.ORGANIZATIONS:
          items = getOrganizationsWithFilter(filterId, 0);
          break;
        case ENTITY_TYPES.ACTIVITIES:
          items = getActivitiesWithFilter(filterId, 0);
          break;
        case ENTITY_TYPES.LEADS:
          items = getLeadsWithFilter(filterId, 0);
          break;
        case ENTITY_TYPES.PRODUCTS:
          items = getProductsWithFilter(filterId, 0);
          break;
      }
    } catch (error) {
      updateSyncStatus(2, 'error', `Error retrieving data: ${error.message}`, 0);

      // Show a more user-friendly error message based on the error
      if (error.message.includes('Filter ID') && error.message.includes('not found')) {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Filter Error',
          `The filter ID ${filterId} was not found in your Pipedrive account. Please check your filter ID in Settings and make sure it exists in your Pipedrive account.`,
          ui.ButtonSet.OK);
        showSettings();
      } else if (error.message.includes('filter type mismatch')) {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Filter Type Mismatch',
          error.message + ' Please update your settings with the correct filter ID.',
          ui.ButtonSet.OK);
        showSettings();
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, 'Sync Error', 10);
      }
      return;
    }

    if (!items || items.length === 0) {
      updateSyncStatus(2, 'warning', `No ${entityType} found with the specified filter.`, 100);
      SpreadsheetApp.getActiveSpreadsheet().toast(`No ${entityType} found with the specified filter ID: ${filterId}. The filter might be empty or not accessible.`);
      return;
    }

    // Complete phase 2
    updateSyncStatus(2, 'completed', `Retrieved ${items.length} ${entityType} from Pipedrive`, 100);

    // Start phase 3 (writing to sheet)
    updateSyncStatus(3, 'active', `Writing ${items.length} ${entityType} to spreadsheet...`, 20);

    // Prepare and write data to sheet with the selected columns
    writeDataToSheet(items, {
      columns: columnsToUse,
      headerRow: headerRow,
      optionMappings: optionMappings,
      entityType: entityType,
      sheetName: sheetName
    });

    // Complete phase 3
    updateSyncStatus(3, 'completed', `Data successfully synced to "${sheetName}"!`, 100);

    // Show success toast
    SpreadsheetApp.getActiveSpreadsheet().toast(`${entityType} successfully synced from Pipedrive to "${sheetName}"! (${items.length} items total)`);
    refreshSyncStatusStyling();
  } catch (error) {
    Logger.log('Error: ' + error.message);
    updateSyncStatus(0, 'error', error.message, 0);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.message);
  }
}

/**
 * Gets the Pipedrive API URL with the subdomain from settings
 */
function getPipedriveApiUrl() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  return PIPEDRIVE_API_URL_PREFIX + subdomain + PIPEDRIVE_API_URL_SUFFIX;
}

/**
 * Gets the field definitions to understand column mappings
 */
function getDealFields() {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

  // Use v1 API directly
  const urlV1 = `https://${subdomain}.pipedrive.com/api/v1/dealFields?api_token=${apiKey}`;

  try {
    const responseV1 = UrlFetchApp.fetch(urlV1, { muteHttpExceptions: true });
    const responseTextV1 = responseV1.getContentText();

    try {
      const responseDataV1 = JSON.parse(responseTextV1);

      if (responseDataV1.success) {
        Logger.log('Successfully retrieved deal fields using v1 API');
        return responseDataV1.data;
      } else {
        Logger.log('Failed to retrieve deal fields with v1 API: ' + responseDataV1.error);
        return [];
      }
    } catch (parseError) {
      Logger.log('Error parsing v1 API response: ' + parseError.message);
      return [];
    }
  } catch (error) {
    Logger.log('Error retrieving deal fields: ' + error.message);
    return [];
  }
}

/**
 * Builds a mapping of field option IDs to their labels for use with multiple option fields
 */
function getFieldOptionMappings() {
  // Get regular deal fields
  const dealFields = getDealFields();
  // Get custom fields (we need to make a separate request for these)
  const customFields = getCustomFields();

  // Combine all fields for processing
  const allFields = [...dealFields, ...customFields];

  const optionMappings = {};

  allFields.forEach(field => {
    // Only process fields with options (dropdown, multiple options, etc.)
    if (field.options && field.options.length > 0) {
      // Create a mapping for this field
      optionMappings[field.key] = {};

      // Map each option ID to its label
      field.options.forEach(option => {
        optionMappings[field.key][option.id] = option.label;
      });

      // Log for debugging
      Logger.log(`Field ${field.name} (${field.key}) has ${field.options.length} options: ${JSON.stringify(field.options)}`);
    }
  });

  return optionMappings;
}

/**
 * Gets custom field definitions
 */
function getCustomFields() {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;

  const url = `${getPipedriveApiUrl()}/dealFields?api_token=${apiKey}&filter_type=custom_field`;

  try {
    const response = UrlFetchApp.fetch(url);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.success) {
      // Log the custom fields for debugging
      Logger.log(`Retrieved ${responseData.data.length} custom fields`);
      return responseData.data;
    } else {
      Logger.log('Failed to retrieve custom fields: ' + responseData.error);
      return [];
    }
  } catch (error) {
    Logger.log('Error retrieving custom fields: ' + error.message);
    return [];
  }
}

/**
 * Gets the field definitions for a specific entity type
 */
function getPipedriverFields(entityType) {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

  // Convert plural entity type to singular for fields endpoints
  let fieldsEndpoint;
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      fieldsEndpoint = 'dealFields';
      break;
    case ENTITY_TYPES.PERSONS:
      fieldsEndpoint = 'personFields';
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      fieldsEndpoint = 'organizationFields';
      break;
    case ENTITY_TYPES.ACTIVITIES:
      fieldsEndpoint = 'activityFields';
      break;
    case ENTITY_TYPES.LEADS:
      fieldsEndpoint = 'leadFields';
      break;
    case ENTITY_TYPES.PRODUCTS:
      fieldsEndpoint = 'productFields';
      break;
    default:
      Logger.log(`Unknown entity type: ${entityType}`);
      return [];
  }

  // Use v1 API directly
  const urlV1 = `https://${subdomain}.pipedrive.com/api/v1/${fieldsEndpoint}?api_token=${apiKey}`;

  try {
    const responseV1 = UrlFetchApp.fetch(urlV1, { muteHttpExceptions: true });
    const responseTextV1 = responseV1.getContentText();

    try {
      const responseDataV1 = JSON.parse(responseTextV1);

      if (responseDataV1.success) {
        Logger.log(`Successfully retrieved ${entityType} fields using v1 API`);
        return responseDataV1.data;
      } else {
        Logger.log(`Failed to retrieve ${entityType} fields with v1 API: ${responseDataV1.error}`);
        return [];
      }
    } catch (parseError) {
      Logger.log(`Error parsing v1 API response for ${entityType}: ${parseError.message}`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error retrieving ${entityType} fields: ${error.message}`);
    return [];
  }
}

/**
 * Gets custom field definitions for a specific entity type
 */
function getCustomFieldsForEntity(entityType) {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

  // Convert plural entity type to singular for fields endpoints
  let fieldsEndpoint;
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      fieldsEndpoint = 'dealFields';
      break;
    case ENTITY_TYPES.PERSONS:
      fieldsEndpoint = 'personFields';
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      fieldsEndpoint = 'organizationFields';
      break;
    case ENTITY_TYPES.ACTIVITIES:
      fieldsEndpoint = 'activityFields';
      break;
    case ENTITY_TYPES.LEADS:
      fieldsEndpoint = 'leadFields';
      break;
    case ENTITY_TYPES.PRODUCTS:
      fieldsEndpoint = 'productFields';
      break;
    default:
      fieldsEndpoint = 'dealFields';
  }

  const url = `https://${subdomain}.pipedrive.com/api/v1/${fieldsEndpoint}?api_token=${apiKey}&filter_type=custom_field`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseText = response.getContentText();

    // Add logging to see what's being returned
    Logger.log(`Custom fields API response for ${entityType}: ${responseText.substring(0, 200)}...`);

    const responseData = JSON.parse(responseText);

    if (responseData && responseData.success) {
      Logger.log(`Successfully retrieved ${responseData.data.length} custom fields for ${entityType}`);
      return responseData.data;
    } else {
      Logger.log(`Failed to retrieve custom fields for ${entityType}: ${responseData ? responseData.error : 'Unknown error'}`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error retrieving custom fields for ${entityType}: ${error.message}`);
    return [];
  }
}

/**
 * Builds a mapping of field option IDs to their labels for a specific entity type
 */
function getFieldOptionMappingsForEntity(entityType) {
  // Get both standard and custom fields
  const standardFields = getPipedriverFields(entityType);
  const customFields = getCustomFieldsForEntity(entityType);

  // Combine all fields for processing
  const allFields = [...standardFields, ...customFields];

  const optionMappings = {};

  allFields.forEach(field => {
    // Only process fields with options (dropdown, multiple options, etc.)
    if (field.options && field.options.length > 0) {
      // Create a mapping for this field
      optionMappings[field.key] = {};

      // Map each option ID to its label
      field.options.forEach(option => {
        optionMappings[field.key][option.id] = option.label;
      });

      // Log for debugging
      Logger.log(`Field ${field.name} (${field.key}) has ${field.options.length} options`);
    }
  });

  return optionMappings;
}

/**
 * Gets deals using a specific filter
 */
function getDealsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.DEALS, filterId, limit);
}

/**
 * Gets persons using a specific filter
 */
function getPersonsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PERSONS, filterId, limit);
}

/**
 * Gets organizations using a specific filter
 */
function getOrganizationsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ORGANIZATIONS, filterId, limit);
}

/**
 * Gets activities using a specific filter
 */
function getActivitiesWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.ACTIVITIES, filterId, limit);
}

/**
 * Gets leads using a specific filter
 */
function getLeadsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.LEADS, filterId, limit);
}

/**
 * Gets deal fields
 */
function getDealFields() {
  return getPipedriverFields(ENTITY_TYPES.DEALS);
}

/**
 * Gets person fields
 */
function getPersonFields() {
  return getPipedriverFields(ENTITY_TYPES.PERSONS);
}

/**
 * Gets organization fields
 */
function getOrganizationFields() {
  return getPipedriverFields(ENTITY_TYPES.ORGANIZATIONS);
}

/**
 * Gets activity fields
 */
function getActivityFields() {
  return getPipedriverFields(ENTITY_TYPES.ACTIVITIES);
}

/**
 * Gets lead fields
 */
function getLeadFields() {
  return getPipedriverFields(ENTITY_TYPES.LEADS);
}

/**
 * Gets product fields
 */
function getProductFields() {
  // Note: Product fields use API v1 according to the documentation
  const url = getPipedriveApiUrl().replace('/v2', '/v1') + '/productFields?limit=500';
  const fields = getCustomFieldsForEntity(ENTITY_TYPES.PRODUCTS);
  return fields;
}

/**
 * Gets mappings for product custom fields
 * @returns {Object} Object mapping custom field keys to names
 */
function getProductFieldOptionMappings() {
  return getFieldOptionMappingsForEntity(ENTITY_TYPES.PRODUCTS);
}

/**
 * Main function to sync products from a Pipedrive filter to the Google Sheet
 */
function syncProductsFromFilter(skipPush = false) {
  syncPipedriveDataToSheet(ENTITY_TYPES.PRODUCTS, skipPush);
}

/**
 * Gets products using a specific filter ID
 * @param {string} filterId The Pipedrive filter ID
 * @param {number} limit The maximum number of products to fetch
 * @returns {Array} Array of products
 */
function getProductsWithFilter(filterId, limit = 100) {
  return getFilteredDataFromPipedrive(ENTITY_TYPES.PRODUCTS, filterId, limit);
}

/**
 * Gets data from Pipedrive using a specific filter based on entity type
 */
function getFilteredDataFromPipedrive(entityType, filterId, limit = 100) {
  // Get API key from properties or use default
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
  const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

  // Define optional fields to include based on entity type
  let includeFields = '';
  switch (entityType) {
    case ENTITY_TYPES.DEALS:
      includeFields = 'next_activity_id,last_activity_id,first_won_time,products_count,files_count,notes_count,followers_count,email_messages_count,activities_count,done_activities_count,undone_activities_count,participants_count,last_incoming_mail_time,last_outgoing_mail_time';
      break;
    case ENTITY_TYPES.PERSONS:
      includeFields = 'next_activity_id,last_activity_id,open_deals_count,related_open_deals_count,closed_deals_count,related_closed_deals_count,participant_open_deals_count,participant_closed_deals_count,email_messages_count,activities_count,done_activities_count,undone_activities_count,files_count,notes_count,followers_count,won_deals_count,related_won_deals_count,lost_deals_count,related_lost_deals_count,last_incoming_mail_time,last_outgoing_mail_time';
      break;
    case ENTITY_TYPES.ORGANIZATIONS:
      includeFields = 'next_activity_id,last_activity_id,open_deals_count,related_open_deals_count,closed_deals_count,related_closed_deals_count,email_messages_count,activities_count,done_activities_count,undone_activities_count,files_count,notes_count,followers_count,won_deals_count,related_won_deals_count,lost_deals_count,related_lost_deals_count';
      break;
    case ENTITY_TYPES.ACTIVITIES:
      includeFields = 'attendees';
      break;
  }

  // First, try to get filter details to ensure it exists (using v1 API for filters)
  const filterUrl = `https://${subdomain}.pipedrive.com/v1/filters/${filterId}?api_token=${apiKey}`;

  try {
    Logger.log(`Checking filter: ${filterUrl}`);

    const filterResponse = UrlFetchApp.fetch(filterUrl, {
      muteHttpExceptions: true
    });

    const responseCode = filterResponse.getResponseCode();
    if (responseCode !== 200) {
      Logger.log(`Filter check failed with status code ${responseCode}`);
      throw new Error(`Filter ID ${filterId} not found or not accessible.`);
    }

    const filterData = JSON.parse(filterResponse.getContentText());

    // Check if the filter exists
    if (!filterData.success || !filterData.data) {
      Logger.log(`Filter ID ${filterId} not found or not accessible. Error: ${JSON.stringify(filterData)}`);
      throw new Error(`Filter ID ${filterId} not found or not accessible.`);
    }

    // Check if filter type matches the entity type
    const filterType = filterData.data.type;
    const normalizedEntityType = entityType.toLowerCase();
    const normalizedFilterType = filterType.toLowerCase();

    // Map entity types to filter types (e.g. 'deals' filter type for 'deals' entity)
    // In Pipedrive API, filter types sometimes differ slightly from entity types
    const filterTypeMap = {
      'deals': 'deals',
      'persons': 'people',
      'organizations': 'org',
      'activities': 'activity',
      'leads': 'leads',
      'products': 'products'
    };

    const expectedFilterType = filterTypeMap[normalizedEntityType] || normalizedEntityType;

    if (normalizedFilterType !== expectedFilterType) {
      Logger.log(`Filter type mismatch. Expected: ${expectedFilterType}, Actual: ${normalizedFilterType}`);
      throw new Error(`The filter ID ${filterId} is for ${normalizedFilterType} but you're trying to use it for ${normalizedEntityType}. Please use a filter created for ${normalizedEntityType}.`);
    }

    try {
      // For all entity types, we need to use v1 API with filters due to issues with v2 API
      // Construct the base URL using v1 API with filter_id
      let url = `https://${subdomain}.pipedrive.com/v1/${entityType}?api_token=${apiKey}&filter_id=${filterId}`;

      // Add include_fields parameter if we have fields to include
      if (includeFields) {
        url += `&include_fields=${includeFields}`;
      }

      let allItems = [];
      let hasMore = true;
      let start = 0;
      const pageLimit = 100; // Pipedrive API limit per page is 100

      // Handle pagination for v1 API
      while (hasMore) {
        const paginatedUrl = `${url}&start=${start}&limit=${pageLimit}`;

        // Log the full URL for debugging
        Logger.log(`Fetching data from: ${paginatedUrl}`);

        const response = UrlFetchApp.fetch(paginatedUrl, {
          muteHttpExceptions: true
        });

        const responseCode = response.getResponseCode();
        if (responseCode !== 200) {
          Logger.log(`API request failed with status code ${responseCode}, response: ${response.getContentText().substring(0, 200)}`);
          throw new Error(`Failed to retrieve ${entityType} data from Pipedrive. Server returned status code ${responseCode}.`);
        }

        const responseText = response.getContentText();

        try {
          const responseData = JSON.parse(responseText);

          // Log the raw response for debugging
          Logger.log(`Response success: ${responseData.success}, items: ${responseData.data ? responseData.data.length : 0}`);

          if (responseData.success) {
            const items = responseData.data;
            if (items && items.length > 0) {
              allItems = allItems.concat(items);

              // If we have a specified limit (not 0), check if we've reached it
              if (limit > 0 && allItems.length >= limit) {
                allItems = allItems.slice(0, limit);
                hasMore = false;
              }
              // Check if there are more results
              else if (responseData.additional_data &&
                responseData.additional_data.pagination &&
                responseData.additional_data.pagination.more_items_in_collection) {
                hasMore = true;
                start += pageLimit;

                // Status update for large datasets
                if (allItems.length % 500 === 0) {
                  SpreadsheetApp.getActiveSpreadsheet().toast(`Retrieved ${allItems.length} ${entityType} so far...`);
                }
              } else {
                hasMore = false;
              }
            } else {
              hasMore = false;
            }
          } else {
            Logger.log(`Failed to retrieve ${entityType}: ${JSON.stringify(responseData)}`);
            throw new Error(`Failed to retrieve ${entityType}: ${responseData.error || 'Unknown error'}`);
          }
        } catch (parseError) {
          Logger.log(`JSON parse error: ${parseError.message}`);
          Logger.log(`Response text: ${responseText.substring(0, 200)}...`);
          throw new Error(`Failed to parse API response: ${parseError.message}`);
        }
      }

      // Log completion status for large datasets
      if (allItems.length > 100) {
        Logger.log(`Retrieved ${allItems.length} ${entityType} from Pipedrive filter`);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Retrieved ${allItems.length} ${entityType} from Pipedrive filter. Preparing data for the sheet...`);
      }

      return allItems;
    } catch (dataFetchError) {
      Logger.log(`Error fetching data: ${dataFetchError.message}`);
      throw dataFetchError;
    }
  } catch (error) {
    Logger.log(`Error retrieving ${entityType} with filter ${filterId}: ${error.message}`);
    throw new Error(`Error retrieving ${entityType} with filter ${filterId}: ${error.message}`);
  }
}

/**
 * Format the value for display in the spreadsheet
 */
function formatValue(value, columnPath, optionMappings = {}) {
  if (value === null || value === undefined) {
    return '';
  }

  // Handle custom fields from API v2
  // In API v2, custom fields are in a nested object
  if (columnPath.startsWith('custom_fields.')) {
    // The field key is the part after "custom_fields."
    const fieldKey = columnPath.split('.')[1];

    // If it's a multiple option field in new format (array of numbers)
    if (Array.isArray(value)) {
      // Check if we have option mappings for this field
      if (optionMappings[fieldKey]) {
        const labels = value.map(id => {
          // Return the label if we have it, otherwise just return the ID
          return optionMappings[fieldKey][id] || id;
        });
        return labels.join(', ');
      }
      return value.join(', ');
    }

    // Handle currency fields and other object-based custom fields
    if (typeof value === 'object' && value !== null) {
      if (value.value !== undefined && value.currency !== undefined) {
        return `${value.value} ${value.currency}`;
      }
      if (value.value !== undefined && value.until !== undefined) {
        return `${value.value} - ${value.until}`;
      }
      // For address fields
      if (value.value !== undefined && value.formatted_address !== undefined) {
        return value.formatted_address;
      }
      return JSON.stringify(value);
    }

    // For single option fields (now just a number)
    if (typeof value === 'number' && optionMappings[fieldKey]) {
      return optionMappings[fieldKey][value] || value;
    }
  }

  // Handle comma-separated IDs for multiple options fields (API v1 format - for backward compatibility)
  if (typeof value === 'string' && /^[0-9]+(,[0-9]+)*$/.test(value)) {
    // This looks like a comma-separated list of IDs, likely a multiple option field

    // Extract the field key from the column path (remove any array or nested indices)
    let fieldKey;

    // Special handling for custom_fields paths
    if (columnPath.startsWith('custom_fields.')) {
      // For custom fields, the field key is after "custom_fields."
      fieldKey = columnPath.split('.')[1];
    } else {
      // For regular fields, use the first part of the path
      fieldKey = columnPath.split('.')[0];
    }

    // Check if we have option mappings for this field
    if (optionMappings[fieldKey]) {
      const ids = value.split(',');
      const labels = ids.map(id => {
        // Return the label if we have it, otherwise just return the ID
        return optionMappings[fieldKey][id] || id;
      });
      return labels.join(', ');
    }
  }

  // Regular object handling
  if (typeof value === 'object') {
    // Check if it's an array of option objects (multiple options field)
    if (Array.isArray(value) && value.length > 0 && value[0] && typeof value[0] === 'object' && value[0].label) {
      // It's a multiple options field with label property, extract and join labels
      return value.map(option => option.label).join(', ');
    }
    // Check if it's a single option object
    else if (value.label !== undefined) {
      return value.label;
    }
    // Handle person/org objects
    else if (value.name !== undefined) {
      return value.name;
    }
    // Handle currency objects
    else if (value.currency !== undefined && value.value !== undefined) {
      return `${value.value} ${value.currency}`;
    }
    // For other objects, convert to JSON string
    return JSON.stringify(value);
  } else if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }

  return value.toString();
}

/**
 * Gets a value from an object by path notation
 * Handles nested properties using dot notation (e.g., "creator_user.name")
 * Supports array indexing with numeric indices
 * Special handling for custom_fields in API v2
 */
function getValueByPath(obj, path) {
  // If path is already an object with a key property, use that
  if (typeof path === 'object' && path.key) {
    path = path.key;
  }

  // Special handling for custom_fields in API v2
  if (path.startsWith('custom_fields.') && obj.custom_fields) {
    const parts = path.split('.');
    const fieldKey = parts[1];
    const nestedField = parts.length > 2 ? parts.slice(2).join('.') : null;

    // If the custom field exists
    if (obj.custom_fields[fieldKey] !== undefined) {
      const fieldValue = obj.custom_fields[fieldKey];

      // If we need to extract a nested property from the custom field
      if (nestedField) {
        return getValueByPath(fieldValue, nestedField);
      }

      // Otherwise return the field value itself
      return fieldValue;
    }

    return undefined;
  }

  // Handle simple non-nested case
  if (!path.includes('.')) {
    return obj[path];
  }

  // Handle nested paths
  const parts = path.split('.');
  let current = obj;

  for (const part of parts) {
    if (current === null || current === undefined) {
      return undefined;
    }

    // Handle array indexing
    if (!isNaN(part) && Array.isArray(current)) {
      const index = parseInt(part);
      current = current[index];
    } else {
      current = current[part];
    }
  }

  return current;
}

/**
 * Format column names for better readability
 */
function formatColumnName(name) {
  // Convert snake_case to Title Case
  return name.split('_')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

/**
 * Sets up a trigger to run the sync automatically on a schedule
 * Uncomment and modify as needed
 */
/*
function createDailyTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncDealsFromFilter') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger that runs daily
  ScriptApp.newTrigger('syncDealsFromFilter')
    .timeBased()
    .everyDays(1)
    .atHour(6) // Run at 6 AM
    .create();
}
*/

/**
 * Shows a UI with sync status information
 */
function showSyncStatus(sheetName) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

  // Reset all sync status properties to ensure clean start
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

  const htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
      }
      
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      
      body {
        font-family: 'Roboto', Arial, sans-serif;
        color: var(--text-dark);
        margin: 0;
        padding: 16px;
        font-size: 14px;
        line-height: 1.5;
      }
      
      .header {
        margin-bottom: 20px;
      }
      
      h3 {
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 8px;
      }
      
      .sync-info {
        background-color: var(--bg-light);
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 16px;
        border-left: 4px solid var(--primary-color);
        display: flex;
        align-items: center;
      }
      
      .sync-info svg {
        margin-right: 12px;
        fill: var(--primary-color);
      }
      
      #syncStatus {
        margin-top: 20px;
        padding: 16px;
        border-radius: 8px;
        background-color: var(--bg-light);
        text-align: center;
      }
      
      .phase {
        margin-top: 16px;
        display: flex;
        align-items: center;
        padding: 8px 12px;
        border-radius: 4px;
        background: white;
        border: 1px solid var(--border-color);
      }
      
      .phase-icon {
        width: 24px;
        height: 24px;
        border-radius: 50%;
        background-color: var(--bg-light);
        margin-right: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
      }
      
      .phase-icon.pending {
        color: var(--text-light);
      }
      
      .phase-icon.active {
        background-color: var(--primary-color);
        color: white;
      }
      
      .phase-icon.completed {
        background-color: var(--success-color);
        color: white;
      }
      
      .phase-icon.error {
        background-color: var(--error-color);
        color: white;
      }
      
      .phase-text {
        flex-grow: 1;
      }
      
      .phase-name {
        font-weight: 500;
      }
      
      .phase-detail {
        font-size: 12px;
        color: var(--text-light);
        margin-top: 2px;
      }
      
      .progress-container {
        height: 4px;
        background-color: var(--bg-light);
        border-radius: 2px;
        margin-top: 4px;
        overflow: hidden;
      }
      
      .progress-bar {
        height: 100%;
        width: 0%;
        background-color: var(--primary-color);
        transition: width 0.3s ease;
      }
      
      .progress-bar.completed {
        background-color: var(--success-color);
        width: 100% !important;
      }
      
      #errorMessage {
        margin-top: 16px;
        padding: 12px;
        border-radius: 4px;
        background-color: rgba(219,68,55,0.1);
        color: var(--error-color);
        border-left: 4px solid var(--error-color);
        display: none;
      }
      
      .loader {
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 2px solid rgba(66,133,244,0.2);
        border-radius: 50%;
        border-top-color: var(--primary-color);
        animation: spin 1s ease-in-out infinite;
      }
      
      .checkmark, .error-icon {
        display: inline-block;
        width: 18px;
        height: 18px;
        fill: currentColor;
        vertical-align: middle;
      }
      
      .button-container {
        margin-top: 24px;
        text-align: right;
      }
      
      button {
        padding: 10px 16px;
        border: none;
        border-radius: 4px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
      }
      
      button:focus {
        outline: none;
      }
      
      button.primary {
        background-color: var(--primary-color);
        color: white;
      }
      
      button.primary:hover {
        background-color: var(--primary-dark);
        box-shadow: 0 1px 3px rgba(60,64,67,0.3);
      }
      
      button:disabled {
        opacity: 0.5;
        cursor: not-allowed;
      }
      
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
    
    <div class="header">
      <h3>Syncing Pipedrive Data</h3>
      
      <div class="sync-info">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24">
          <path d="M12 4V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.01-.25 1.97-.7 2.8l1.46 1.46C19.54 15.03 20 13.57 20 12c0-4.42-3.58-8-8-8zm0 14c-3.31 0-6-2.69-6-6 0-1.01.25-1.97.7-2.8L5.24 7.74C4.46 8.97 4 10.43 4 12c0 4.42 3.58 8 8 8v3l4-4-4-4v3z"/>
        </svg>
        <div>
          Syncing <strong>${entityType}</strong> data to sheet <strong>"${sheetName}"</strong>
        </div>
      </div>
    </div>
    
    <div id="syncStatus">
      <div class="phase" id="phase1">
        <div class="phase-icon active" id="phase1Icon"><span class="loader"></span></div>
        <div class="phase-text">
          <div class="phase-name">Connecting to Pipedrive</div>
          <div class="phase-detail" id="phase1Detail">Connecting to Pipedrive...</div>
          <div class="progress-container">
            <div class="progress-bar" id="phase1Progress" style="width: 50%"></div>
          </div>
        </div>
      </div>
      
      <div class="phase" id="phase2">
        <div class="phase-icon pending" id="phase2Icon">2</div>
        <div class="phase-text">
          <div class="phase-name">Retrieving data from Pipedrive</div>
          <div class="phase-detail" id="phase2Detail">Waiting to start...</div>
          <div class="progress-container">
            <div class="progress-bar" id="phase2Progress"></div>
          </div>
        </div>
      </div>
      
      <div class="phase" id="phase3">
        <div class="phase-icon pending" id="phase3Icon">3</div>
        <div class="phase-text">
          <div class="phase-name">Writing data to spreadsheet</div>
          <div class="phase-detail" id="phase3Detail">Waiting to start...</div>
          <div class="progress-container">
            <div class="progress-bar" id="phase3Progress"></div>
          </div>
        </div>
      </div>
      
      <div id="errorMessage"></div>
      
      <div class="button-container">
        <button class="primary" id="closeBtn" disabled>Close</button>
      </div>
    </div>
    
    <script>
      // Function to update UI based on status
      function updatePhase(phaseNumber, status, detail, progress) {
        const phaseElement = document.getElementById('phase' + phaseNumber);
        const phaseIcon = document.getElementById('phase' + phaseNumber + 'Icon');
        const phaseDetail = document.getElementById('phase' + phaseNumber + 'Detail');
        const progressBar = document.getElementById('phase' + phaseNumber + 'Progress');
        
        // Update phase status
        phaseIcon.className = 'phase-icon ' + status;
        
        // Update icon content based on status
        if (status === 'active') {
          phaseIcon.innerHTML = '<span class="loader"></span>';
        } else if (status === 'completed') {
          phaseIcon.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="checkmark"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>';
          
          // Force progress bar to show completed
          progressBar.style.width = '100%';
          progressBar.className = 'progress-bar completed';
        } else if (status === 'error') {
          phaseIcon.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="error-icon"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/></svg>';
        } else {
          phaseIcon.textContent = phaseNumber;
        }
        
        // Update detail text
        if (detail && detail !== 'undefined') {
          phaseDetail.textContent = detail;
        }
        
        // Update progress bar for non-completed states
        if (progress !== undefined && status !== 'completed') {
          progressBar.style.width = progress + '%';
        }
      }
      
      // Function to show error message
      function showError(message) {
        if (!message) return;
        
        const errorElement = document.getElementById('errorMessage');
        errorElement.textContent = message;
        errorElement.style.display = 'block';
        
        // Enable close button
        document.getElementById('closeBtn').disabled = false;
      }
      
      // Set up close button
      document.getElementById('closeBtn').onclick = function() {
        google.script.host.close();
      };
      
      // Poll for updates every 0.75 seconds
      let pollInterval = setInterval(function() {
        google.script.run
          .withSuccessHandler(function(status) {
            if (!status) return;
            
            // Update each phase status
            if (status.phase1) {
              updatePhase(1, status.phase1.status, status.phase1.detail, status.phase1.progress);
            }
            
            if (status.phase2) {
              updatePhase(2, status.phase2.status, status.phase2.detail, status.phase2.progress);
            }
            
            if (status.phase3) {
              updatePhase(3, status.phase3.status, status.phase3.detail, status.phase3.progress);
            }
            
            // Show error message if needed
            if (status.error) {
              showError(status.error);
            }
            
            // Enable close button when completed
            if (status.completed === 'true') {
              clearInterval(pollInterval);
              document.getElementById('closeBtn').disabled = false;
            }
          })
          .withFailureHandler(function(error) {
            console.error('Error polling for status:', error);
          })
          .getSyncStatus();
      }, 750);
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(480)
    .setTitle('Syncing Pipedrive Data');

  // Show the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Syncing Pipedrive Data');
}

/**
 * Gets the current sync status for the dialog to poll
 */
function getSyncStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();

  return {
    phase1: {
      status: scriptProperties.getProperty('SYNC_PHASE_1_STATUS') || 'active',
      detail: scriptProperties.getProperty('SYNC_PHASE_1_DETAIL') || 'Connecting to Pipedrive...',
      progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_1_PROGRESS') || '0')
    },
    phase2: {
      status: scriptProperties.getProperty('SYNC_PHASE_2_STATUS') || 'pending',
      detail: scriptProperties.getProperty('SYNC_PHASE_2_DETAIL') || 'Waiting to start...',
      progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_2_PROGRESS') || '0')
    },
    phase3: {
      status: scriptProperties.getProperty('SYNC_PHASE_3_STATUS') || 'pending',
      detail: scriptProperties.getProperty('SYNC_PHASE_3_DETAIL') || 'Waiting to start...',
      progress: parseInt(scriptProperties.getProperty('SYNC_PHASE_3_PROGRESS') || '0')
    },
    currentPhase: scriptProperties.getProperty('SYNC_CURRENT_PHASE') || '1',
    completed: scriptProperties.getProperty('SYNC_COMPLETED') || 'false',
    error: scriptProperties.getProperty('SYNC_ERROR') || ''
  };
}

/**
 * Updates the sync status properties
 */
function updateSyncStatus(phase, status, detail, progress) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();

    // Ensure progress is 100% for completed phases
    if (status === 'completed') {
      progress = 100;
    }

    // Store status for the specific phase
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_STATUS`, status);
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_DETAIL`, detail || '');
    scriptProperties.setProperty(`SYNC_PHASE_${phase}_PROGRESS`, progress ? progress.toString() : '0');

    // Set current phase
    scriptProperties.setProperty('SYNC_CURRENT_PHASE', phase.toString());

    // If status is error, store the error
    if (status === 'error') {
      scriptProperties.setProperty('SYNC_ERROR', detail || 'An error occurred');
      scriptProperties.setProperty('SYNC_COMPLETED', 'true');
    }

    // If this is the final phase completion, mark as completed
    if (status === 'completed' && phase === 3) {
      scriptProperties.setProperty('SYNC_COMPLETED', 'true');
    }

    // Also show a toast message for visibility
    let toastMessage = '';
    if (phase === 1) toastMessage = 'Connecting to Pipedrive...';
    else if (phase === 2) toastMessage = 'Retrieving data from Pipedrive...';
    else if (phase === 3) toastMessage = 'Writing data to spreadsheet...';

    if (status === 'error') {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${detail}`, 'Sync Error', 5);
    } else if (status === 'completed' && phase === 3) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Sync completed successfully!', 'Sync Status', 3);
    } else if (detail) {
      SpreadsheetApp.getActiveSpreadsheet().toast(detail, toastMessage, 2);
    }
  } catch (e) {
    // If we have issues updating the status, just continue with the sync
    Logger.log('Could not update sync status: ' + e.message);
    // Still show a toast message as backup
    SpreadsheetApp.getActiveSpreadsheet().toast(detail || 'Processing...', 'Sync Status', 2);
  }
}

/**
 * Shows help and about information
 */
function showHelp() {
  const htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
    <style>
      :root {
        --primary-color: #4285f4;
        --text-dark: #202124;
        --text-light: #5f6368;
        --bg-light: #f8f9fa;
        --border-color: #dadce0;
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
        font-size: 14px;
      }
      
      h3 {
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 16px;
      }
      
      .header-icon {
        text-align: center;
        margin-bottom: 20px;
      }
      
      .header-icon svg {
        width: 64px;
        height: 64px;
      }
      
      .section {
        margin-bottom: 24px;
        padding-bottom: 16px;
        border-bottom: 1px solid var(--border-color);
      }
      
      .section-title {
        font-weight: 500;
        font-size: 15px;
        margin-bottom: 8px;
        color: var(--primary-color);
      }
      
      ul {
        padding-left: 20px;
      }
      
      li {
        margin-bottom: 8px;
      }
      
      .help-item {
        display: flex;
        margin-bottom: 12px;
      }
      
      .help-icon {
        width: 24px;
        margin-right: 12px;
        color: var(--primary-color);
        text-align: center;
      }
      
      .help-text {
        flex: 1;
      }
      
      .footer {
        font-size: 12px;
        color: var(--text-light);
        text-align: center;
        margin-top: 16px;
      }
    </style>
    
    <div class="header-icon">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" fill="#4285f4">
        <path d="M100 15c-47.5 0-86 38.5-86 86s38.5 86 86 86 86-38.5 86-86-38.5-86-86-86zm0 160c-40.8 0-74-33.2-74-74s33.2-74 74-74 74 33.2 74 74-33.2 74-74 74z"/>
        <path d="M100 55c-5.8 0-10 4.2-10 10s4.2 10 10 10 10-4.2 10-10-4.2-10-10-10zm0 80c-5.8 0-10-4.2-10-10V95c0-5.8 4.2-10 10-10s10 4.2 10 10v30c0 5.8-4.2 10-10 10z"/>
      </svg>
    </div>
    
    <h3>Pipedrive Integration Help</h3>
    
    <div class="section">
      <div class="section-title">Getting Started</div>
      <div class="help-item">
        <div class="help-icon">⚙️</div>
        <div class="help-text">
          First, configure your integration in the <strong>Settings</strong> menu. You'll need your Pipedrive API key and subdomain.
        </div>
      </div>
      <div class="help-item">
        <div class="help-icon">📊</div>
        <div class="help-text">
          Next, use <strong>Select Columns</strong> to choose which Pipedrive data fields to display in your sheet.
        </div>
      </div>
      <div class="help-item">
        <div class="help-icon">🔄</div>
        <div class="help-text">
          Finally, use <strong>Sync Data</strong> to import your Pipedrive data into the current sheet.
        </div>
      </div>
    </div>
    
    <div class="section">
      <div class="section-title">Tips & Tricks</div>
      <ul>
        <li>Create multiple sheets to track different types of Pipedrive data (deals, contacts, etc.)</li>
        <li>Use Pipedrive filters to control exactly which records are imported</li>
        <li>Customize columns to show only the data you need</li>
        <li>The sync will overwrite all data in the current sheet</li>
      </ul>
    </div>
    
    <div class="section">
      <div class="section-title">Troubleshooting</div>
      <ul>
        <li><strong>No data syncing:</strong> Check your API key and filter ID in Settings</li>
        <li><strong>Missing columns:</strong> Use the Select Columns menu to add more fields</li>
        <li><strong>Error messages:</strong> Ensure your Pipedrive account has access to the requested data</li>
      </ul>
    </div>
    
    <div class="footer">
      PipedriveSheets<br>
      Version 1.0.1
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(500)
    .setTitle('Pipedrive Integration Help');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Pipedrive Integration Help');
}

/**
 * Writes data to the Google Sheet with columns matching the filter
 */
function writeDataToSheet(items, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get sheet name from options or properties
  const sheetName = options.sheetName ||
    PropertiesService.getScriptProperties().getProperty('SHEET_NAME') ||
    DEFAULT_SHEET_NAME;

  let sheet = ss.getSheetByName(sheetName);

  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  cleanupPreviousSyncStatusColumn(sheet, sheetName);

  // Check if two-way sync is enabled for this sheet
  const scriptProperties = PropertiesService.getScriptProperties();
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
  const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
  const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

  // For preservation of status data - key is entity ID, value is status
  let statusByIdMap = new Map();
  let configuredTrackingColumn = '';

  // Step 1: If two-way sync is enabled, extract all status values by entity ID
  if (twoWaySyncEnabled) {
    try {
      Logger.log('Two-way sync is enabled, preserving status column data');

      // Get the configured tracking column letter from properties
      configuredTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
      Logger.log(`Configured tracking column from properties: "${configuredTrackingColumn}"`);

      // If the sheet has data, extract status values
      if (sheet.getLastRow() > 0) {
        // Get all existing data
        const existingData = sheet.getDataRange().getValues();
        Logger.log(`Retrieved ${existingData.length} rows of existing data`);

        // Find headers row
        const headers = existingData[0];
        let idColumnIndex = 0; // Assuming ID is in first column
        let statusColumnIndex = -1;

        // Find the status column by looking for header that contains "Sync Status"
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] &&
            headers[i].toString().toLowerCase().includes('sync') &&
            headers[i].toString().toLowerCase().includes('status')) {
            statusColumnIndex = i;
            Logger.log(`Found status column at index ${statusColumnIndex} with header "${headers[i]}"`);
            break;
          }
        }

        // If status column found, extract status values by ID
        if (statusColumnIndex !== -1) {
          // Process all data rows (skip header)
          for (let i = 1; i < existingData.length; i++) {
            const row = existingData[i];
            const id = row[idColumnIndex];
            const status = row[statusColumnIndex];

            // Skip rows without ID or status or timestamp rows
            if (!id || !status ||
              ((typeof id === 'string') &&
                (id.toLowerCase().includes('last') ||
                  id.toLowerCase().includes('synced') ||
                  id.toLowerCase().includes('updated')))) {
              continue;
            }

            // Only preserve meaningful status values
            if (status === 'Modified' || status === 'Synced' || status === 'Error') {
              statusByIdMap.set(id.toString(), status);
              Logger.log(`Preserved status "${status}" for ID ${id}`);
            }
          }

          Logger.log(`Preserved ${statusByIdMap.size} status values by ID`);
        } else {
          Logger.log('Could not find existing status column');
        }
      }
    } catch (e) {
      Logger.log(`Error preserving status data: ${e.message}`);
    }
  }

  // Clear existing data now that we've preserved what we need
  sheet.clear();

  // Use the columns and headers from the options
  const columns = options.columns;
  const headerRow = options.headerRow;
  const optionMappings = options.optionMappings || {};
  const entityType = options.entityType || ENTITY_TYPES.DEALS;

  // Create the full header row with status column if needed
  let fullHeaderRow = [];

  // Use custom header names when available
  headerRow.forEach((originalHeader, index) => {
    const column = columns[index];
    if (column && column.customName) {
      fullHeaderRow.push(column.customName);
    } else {
      fullHeaderRow.push(originalHeader);
    }
  });

  // Add status column if two-way sync is enabled
  let statusColumnIndex = -1;
  if (twoWaySyncEnabled) {
    // Check if we need to force the Sync Status column at the end
    const twoWaySyncColumnAtEndKey = `TWOWAY_SYNC_COLUMN_AT_END_${sheetName}`;
    const forceColumnAtEnd = scriptProperties.getProperty(twoWaySyncColumnAtEndKey) === 'true';

    // Get the configured tracking column letter and verify if it exists
    configuredTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';

    // If we need to force the column at the end, or no configured column exists,
    // add the Sync Status column at the end
    if (forceColumnAtEnd || !configuredTrackingColumn) {
      Logger.log('Adding Sync Status column at the end');
      statusColumnIndex = fullHeaderRow.length;
      fullHeaderRow.push('Sync Status');

      // Clear the "force at end" flag after we've processed it
      if (forceColumnAtEnd) {
        scriptProperties.deleteProperty(twoWaySyncColumnAtEndKey);
        Logger.log('Cleared the force-at-end flag after repositioning the Sync Status column');
      }
    } else {
      // If not forcing at end, try to find existing column position
      if (configuredTrackingColumn) {
        // Find if there's already a "Sync Status" column in the existing sheet
        const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (activeSheet && activeSheet.getLastRow() > 0) {
          const existingHeaders = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];

          // Search for "Sync Status" in the existing headers
          for (let i = 0; i < existingHeaders.length; i++) {
            if (existingHeaders[i] === "Sync Status") {
              // Use this exact position for our new status column
              const exactColumnLetter = columnToLetter(i);
              Logger.log(`Found existing Sync Status column at position ${i + 1} (${exactColumnLetter})`);

              // Update the tracking column in script properties
              scriptProperties.setProperty(twoWaySyncTrackingColumnKey, exactColumnLetter);

              // Ensure our headers array has enough elements
              while (fullHeaderRow.length <= i) {
                fullHeaderRow.push('');
              }

              // Place "Sync Status" at the exact same position
              statusColumnIndex = i;
              fullHeaderRow[statusColumnIndex] = 'Sync Status';

              break;
            }
          }
        }

        // If we didn't find an existing column, try using the configured index
        if (statusColumnIndex === -1) {
          const existingTrackingIndex = columnLetterToIndex(configuredTrackingColumn);

          // Always use the configured position, regardless of current headerRow length
          Logger.log(`Using configured tracking column at position ${existingTrackingIndex} (${configuredTrackingColumn})`);

          // Ensure our headers array has enough elements
          while (fullHeaderRow.length <= existingTrackingIndex) {
            fullHeaderRow.push('');
          }

          // Place the Sync Status column at its original position
          statusColumnIndex = existingTrackingIndex;
          fullHeaderRow[statusColumnIndex] = 'Sync Status';
        }
      } else {
        // No configured tracking column, add as the last column
        Logger.log('No configured tracking column, adding as last column');
        statusColumnIndex = fullHeaderRow.length;
        fullHeaderRow.push('Sync Status');
      }
    }

    if (configuredTrackingColumn && statusColumnIndex !== -1) {
      const newTrackingColumn = columnToLetter(statusColumnIndex);
      if (newTrackingColumn !== configuredTrackingColumn) {
        scriptProperties.setProperty(`PREVIOUS_TRACKING_COLUMN_${sheetName}`, configuredTrackingColumn);
        Logger.log(`Stored previous tracking column ${configuredTrackingColumn} before moving to ${newTrackingColumn}`);
      }
    }

    // Save the status column position to properties for future use
    const statusColumnLetter = columnToLetter(statusColumnIndex);
    scriptProperties.setProperty(twoWaySyncTrackingColumnKey, statusColumnLetter);
    Logger.log(`Setting status column at index ${statusColumnIndex} (column ${statusColumnLetter})`);
  }

  // Write the header row
  sheet.getRange(1, 1, 1, fullHeaderRow.length).setValues([fullHeaderRow]).setFontWeight('bold');

  // Prepare the data rows
  const dataRows = [];

  // Process each item from Pipedrive
  items.forEach(item => {
    // Create a row array with the right number of columns
    const row = new Array(fullHeaderRow.length).fill('');

    // Extract values for each column from the Pipedrive item
    columns.forEach((column, index) => {
      // Get the value using path notation
      const value = getValueByPath(item, column);

      // Format the value and add it to the row
      row[index] = formatValue(value, column, optionMappings);
    });

    // If two-way sync is enabled, add the status column
    if (twoWaySyncEnabled && statusColumnIndex !== -1) {
      // Get the item ID (assuming first column is ID)
      const id = row[0] ? row[0].toString() : '';

      // If we have a saved status for this ID, use it, otherwise use "Not modified"
      if (id && statusByIdMap.has(id)) {
        row[statusColumnIndex] = statusByIdMap.get(id);
      } else {
        row[statusColumnIndex] = 'Not modified';
      }
    }

    // Add the row to our data array
    dataRows.push(row);
  });

  // Write all data rows at once
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, fullHeaderRow.length).setValues(dataRows);
  }

  // If two-way sync is enabled, set up data validation and formatting for the status column
  if (twoWaySyncEnabled && statusColumnIndex !== -1 && dataRows.length > 0) {
    try {
      // Clear any existing data validation from ALL cells in the sheet first
      sheet.clearDataValidations();

      // Convert to 1-based column
      const statusColumnPos = statusColumnIndex + 1;

      // IMPORTANT: Apply validation EXPLICITLY to EACH data row instead of the entire range
      // This gives us precise control over which cells get validation
      for (let i = 0; i < dataRows.length; i++) {
        const row = i + 2; // Data starts at row 2 (after header)
        const statusCell = sheet.getRange(row, statusColumnPos);

        // Use the same dropdown for all cells
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
          .build();
        statusCell.setDataValidation(rule);
      }

      // Style the header to make it clear it's for sync status
      sheet.getRange(1, statusColumnPos).setBackground('#E8F0FE')
        .setFontWeight('bold')
        .setNote('This column tracks changes for two-way sync with Pipedrive');

      // Style the status column
      const statusRange = sheet.getRange(2, statusColumnPos, dataRows.length, 1);
      statusRange.setBackground('#F8F9FA')  // Light gray
        .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

      // Set up conditional formatting for the status values
      const rules = [];

      // Add rule for "Modified" status - red background
      let modifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setRanges([statusRange])
        .build();
      rules.push(modifiedRule);

      // Add rule for "Synced" status - green background
      let syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA')  // Light green background
        .setFontColor('#137333')   // Green text
        .setRanges([statusRange])
        .build();
      rules.push(syncedRule);

      // Add rule for "Error" status - red background with bold text
      let errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Error')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setBold(true)             // Bold text for errors
        .setRanges([statusRange])
        .build();
      rules.push(errorRule);

      sheet.setConditionalFormatRules(rules);

      Logger.log('Applied data validation and conditional formatting to status column');
    } catch (e) {
      Logger.log(`Error setting up status column formatting: ${e.message}`);
    }
  }

  // Check if timestamp is enabled (moved outside the else block to work for all sheets)
  const timestampEnabledKey = `TIMESTAMP_ENABLED_${sheetName}`;
  const timestampEnabled = scriptProperties.getProperty(timestampEnabledKey) === 'true';

  // Add timestamp of last sync if enabled
  if (timestampEnabled) {
    const timestampRow = sheet.getLastRow() + 2;
    sheet.getRange(timestampRow, 1).setValue('Last synced:');
    sheet.getRange(timestampRow, 2).setValue(new Date()).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    // Format the timestamp for better visibility
    sheet.getRange(timestampRow, 1, 1, 2).setFontWeight('bold');
    sheet.getRange(timestampRow, 1, 1, 2).setBackground('#f1f3f4');  // Light gray background

    // Clear any data validation from the timestamp row and spacer row
    if (twoWaySyncEnabled && statusColumnIndex !== -1) {
      // Clear the entire spacer row and timestamp row from any data validation
      const spacerRow = timestampRow - 1;
      if (spacerRow > dataRows.length + 1) { // Make sure we don't clear data rows
        sheet.getRange(spacerRow, 1, 2, sheet.getLastColumn()).setDataValidation(null);
      }
    }
  }

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  // CRITICAL: Clean up any lingering formatting from previous Sync Status columns
  // This needs to happen AFTER the sheet is rebuilt
  if (twoWaySyncEnabled && statusColumnIndex !== -1) {
    try {
      Logger.log(`Performing aggressive column cleanup after sheet rebuild - current status column: ${statusColumnIndex}`);

      // The current status column letter
      const currentStatusColLetter = columnToLetter(statusColumnIndex);

      // Clean ALL columns in the sheet except the current Sync Status column
      const lastCol = sheet.getLastColumn() + 5; // Add buffer for hidden columns
      for (let i = 0; i < lastCol; i++) {
        if (i !== statusColumnIndex) { // Skip the current status column
          try {
            const colLetter = columnToLetter(i);
            Logger.log(`Checking column ${colLetter} for cleanup`);

            // Check if this column has a 'Sync Status' header or sync-related note
            const headerCell = sheet.getRange(1, i + 1);
            const headerValue = headerCell.getValue();
            const note = headerCell.getNote();

            if (headerValue === "Sync Status" ||
              (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive')))) {
              Logger.log(`Found Sync Status indicators in column ${colLetter}, cleaning up`);
              cleanupColumnFormatting(sheet, colLetter);
            }

            // Also check for data validation in this column
            if (i < sheet.getLastColumn()) {
              try {
                // Check a sample cell for validation
                const sampleCell = sheet.getRange(2, i + 1);
                const validation = sampleCell.getDataValidation();

                if (validation) {
                  try {
                    const values = validation.getCriteriaValues();
                    if (values && values.length > 0 &&
                      (values[0].join(',').includes('Modified') ||
                        values[0].join(',').includes('Synced'))) {
                      Logger.log(`Found validation in column ${colLetter}, cleaning up`);
                      cleanupColumnFormatting(sheet, colLetter);
                    }
                  } catch (e) { }
                }
              } catch (e) { }
            }
          } catch (e) {
            // Ignore errors for individual columns
          }
        }
      }

      // Update tracking
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, statusColumnIndex.toString());

    } catch (e) {
      Logger.log(`Error during aggressive post-rebuild cleanup: ${e.message}`);
    }
  }
}

/**
 * Converts a 0-based column index to a column letter (A, B, C, ..., Z, AA, AB, etc.)
 * @param {number} columnIndex The 0-based column index
 * @return {string} The column letter
 */
function columnToLetter(columnIndex) {
  let temp, letter = '';
  columnIndex += 1; // Convert to 1-based index

  while (columnIndex > 0) {
    temp = (columnIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIndex = (columnIndex - temp - 1) / 26;
  }

  return letter;
}

/**
 * Shows the trigger management UI
 */
function showTriggerManager() {
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

  // Get existing triggers for this sheet
  const existingTriggers = getTriggersForSheet(activeSheetName);

  // Build the HTML for the list of existing triggers
  let existingTriggersHtml = '';
  if (existingTriggers.length > 0) {
    existingTriggersHtml = `
      <div class="existing-triggers">
        <h4>Existing Sync Schedules</h4>
        <table class="triggers-table">
          <thead>
            <tr>
              <th>Frequency</th>
              <th>Details</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>`;

    existingTriggers.forEach(trigger => {
      const triggerInfo = getTriggerInfo(trigger);
      const triggerId = trigger.getUniqueId();
      existingTriggersHtml += `
        <tr id="trigger-row-${triggerId}">
          <td>${triggerInfo.type}</td>
          <td>${triggerInfo.description}</td>
          <td>
            <div id="remove-loading-${triggerId}" class="mini-loading" style="display:none;">
              <span class="mini-loader"></span>
            </div>
            <button type="button" id="remove-btn-${triggerId}" class="delete-trigger" 
              onclick="deleteTrigger('${triggerId}')">
              Remove
            </button>
          </td>
        </tr>`;
    });

    existingTriggersHtml += `
          </tbody>
        </table>
      </div>`;
  } else {
    existingTriggersHtml = `
      <div class="no-triggers">
        <p>No automatic sync schedules are set up for this sheet.</p>
      </div>`;
  }

  // Create a notice about two-way sync if it's enabled
  const twoWaySyncNotice = twoWaySyncEnabled ? `
    <div class="two-way-sync-notice">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
        <path d="M18 7l-1.41-1.41-6.34 6.34 1.41 1.41L18 7zm4.24-1.41L11.66 16.17 7.48 12l-1.41 1.41L11.66 19l12-12-1.42-1.41zM.41 13.41L6 19l1.41-1.41L1.83 12 .41 13.41z"/>
      </svg>
      <div>
        <strong>Two-Way Sync is enabled for this sheet.</strong>
        <p>When scheduled sync runs, any modified rows will be pushed to Pipedrive before pulling new data.</p>
      </div>
    </div>` : '';

  const htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
    
    <div class="container">
      <h3>Schedule Automatic Sync</h3>
      
      <div id="statusIndicator" class="indicator"></div>
      
      <div class="sheet-info">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24">
          <path d="M15 1H9v2h6V1zm-4 13h2V8h-2v6zm8.03-6.61l1.42-1.42c-.43-.51-.9-.99-1.41-1.41l-1.42 1.42C16.07 4.74 14.12 4 12 4c-4.97 0-9 4.03-9 9s4.02 9 9 9 9-4.03 9-9c0-2.12-.74-4.07-1.97-5.61zM12 20c-3.87 0-7-3.13-7-7s3.13-7 7-7 7 3.13 7 7-3.13 7-7 7z"/>
        </svg>
        <div>Scheduling automatic sync for <strong>${entityType}</strong> in sheet <strong>"${activeSheetName}"</strong></div>
      </div>
      
      ${twoWaySyncNotice}
      
      ${existingTriggersHtml}
      
      <h4>Create New Schedule</h4>
      
      <form id="triggerForm">
        <div class="form-group">
          <label for="frequency">Sync Frequency</label>
          <select id="frequency" onchange="updateFormVisibility()">
            <option value="daily">Daily</option>
            <option value="weekly">Weekly</option>
            <option value="monthly">Monthly</option>
            <option value="hourly">Hourly</option>
          </select>
        </div>
        
        <div class="form-group" id="hourlyGroup">
          <label>Run every</label>
          <select id="hourlyInterval">
            <option value="1">1 hour</option>
            <option value="2">2 hours</option>
            <option value="3">3 hours</option>
            <option value="4">4 hours</option>
            <option value="6">6 hours</option>
            <option value="8">8 hours</option>
            <option value="12">12 hours</option>
          </select>
          <p class="help-text">The sync will run at regular intervals throughout the day.</p>
        </div>
        
        <div class="form-group" id="weeklyGroup" style="display:none;">
          <label>Days of Week</label>
          <div class="day-selection">
            <button type="button" class="day-button" data-day="1" onclick="toggleDay(this)">Monday</button>
            <button type="button" class="day-button" data-day="2" onclick="toggleDay(this)">Tuesday</button>
            <button type="button" class="day-button" data-day="3" onclick="toggleDay(this)">Wednesday</button>
            <button type="button" class="day-button" data-day="4" onclick="toggleDay(this)">Thursday</button>
            <button type="button" class="day-button" data-day="5" onclick="toggleDay(this)">Friday</button>
            <button type="button" class="day-button" data-day="6" onclick="toggleDay(this)">Saturday</button>
            <button type="button" class="day-button" data-day="7" onclick="toggleDay(this)">Sunday</button>
          </div>
        </div>
        
        <div class="form-group" id="monthlyGroup" style="display:none;">
          <label for="monthDay">Day of Month</label>
          <select id="monthDay">
            <option value="1">1st</option>
            <option value="2">2nd</option>
            <option value="3">3rd</option>
            <option value="4">4th</option>
            <option value="5">5th</option>
            <option value="10">10th</option>
            <option value="15">15th</option>
            <option value="20">20th</option>
            <option value="25">25th</option>
            <option value="28">28th</option>
          </select>
          <p class="help-text">The sync will run once per month on this day.</p>
        </div>
        
        <div class="form-group" id="timeGroup">
          <label>Time of Day</label>
          <div class="time-inputs">
            <div>
              <label for="hour">Hour</label>
              <select id="hour">
                <option value="0">12 AM</option>
                <option value="1">1 AM</option>
                <option value="2">2 AM</option>
                <option value="3">3 AM</option>
                <option value="4">4 AM</option>
                <option value="5">5 AM</option>
                <option value="6">6 AM</option>
                <option value="7">7 AM</option>
                <option value="8" selected>8 AM</option>
                <option value="9">9 AM</option>
                <option value="10">10 AM</option>
                <option value="11">11 AM</option>
                <option value="12">12 PM</option>
                <option value="13">1 PM</option>
                <option value="14">2 PM</option>
                <option value="15">3 PM</option>
                <option value="16">4 PM</option>
                <option value="17">5 PM</option>
                <option value="18">6 PM</option>
                <option value="19">7 PM</option>
                <option value="20">8 PM</option>
                <option value="21">9 PM</option>
                <option value="22">10 PM</option>
                <option value="23">11 PM</option>
              </select>
            </div>
            <div>
              <label for="minute">Minute</label>
              <select id="minute">
                <option value="0" selected>00</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="45">45</option>
              </select>
            </div>
          </div>
          <p class="help-text">Times are based on your Google account's timezone.</p>
        </div>
        
        <input type="hidden" id="sheetName" value="${activeSheetName}" />
        
        <div class="button-container">
          <button type="button" class="secondary-btn" id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
          <div class="loading" id="saveLoading">
            <span class="loader"></span>
          </div>
          <button type="button" class="primary-btn" id="saveBtn" onclick="createTrigger()">Create Schedule</button>
        </div>
      </form>
    </div>
    
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

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(650)
    .setTitle('Schedule Automatic Sync');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Schedule Automatic Sync');
}

/**
 * Creates a sync trigger based on user preferences
 */
function createSyncTrigger(triggerData) {
  try {
    // Configure the trigger based on frequency
    switch (triggerData.frequency) {
      case 'hourly':
        const hourlyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .everyHours(triggerData.hourlyInterval)
          .create();

        // Store sheet name and frequency in trigger properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty(
          `TRIGGER_${hourlyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        scriptProperties.setProperty(
          `TRIGGER_${hourlyTrigger.getUniqueId()}_FREQUENCY`,
          'hourly'
        );

        return {
          success: true,
          triggerId: hourlyTrigger.getUniqueId()
        };

      case 'daily':
        const dailyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .atHour(triggerData.hour)
          .nearMinute(triggerData.minute)
          .everyDays(1)
          .create();

        // Store sheet name and frequency in trigger properties
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${dailyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${dailyTrigger.getUniqueId()}_FREQUENCY`,
          'daily'
        );

        return {
          success: true,
          triggerId: dailyTrigger.getUniqueId()
        };

      case 'weekly':
        // For weekly triggers, we need to create a separate trigger for each selected day
        const weekDayTriggers = [];
        triggerData.weekDays.forEach(day => {
          const dayTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
            .timeBased()
            .onWeekDay(day)
            .atHour(triggerData.hour)
            .nearMinute(triggerData.minute)
            .create();

          // Store sheet name and frequency in trigger properties
          PropertiesService.getScriptProperties().setProperty(
            `TRIGGER_${dayTrigger.getUniqueId()}_SHEET`,
            triggerData.sheetName
          );
          PropertiesService.getScriptProperties().setProperty(
            `TRIGGER_${dayTrigger.getUniqueId()}_FREQUENCY`,
            'weekly'
          );

          weekDayTriggers.push(dayTrigger.getUniqueId());
        });

        return {
          success: true,
          triggerId: weekDayTriggers.join(',') // Return a comma-separated list of trigger IDs
        };

      case 'monthly':
        const monthlyTrigger = ScriptApp.newTrigger('syncSheetFromTrigger')
          .timeBased()
          .onMonthDay(triggerData.monthDay)
          .atHour(triggerData.hour)
          .nearMinute(triggerData.minute)
          .create();

        // Store sheet name and frequency in trigger properties
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${monthlyTrigger.getUniqueId()}_SHEET`,
          triggerData.sheetName
        );
        PropertiesService.getScriptProperties().setProperty(
          `TRIGGER_${monthlyTrigger.getUniqueId()}_FREQUENCY`,
          'monthly'
        );

        return {
          success: true,
          triggerId: monthlyTrigger.getUniqueId()
        };

      default:
        return {
          success: false,
          error: 'Invalid frequency selected'
        };
    }

  } catch (error) {
    Logger.log('Error creating trigger: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Syncs a specific sheet when triggered by a time-based trigger
 */
function syncSheetFromTrigger(event) {
  try {
    // Get the trigger ID from the event
    const triggerId = event.triggerUid;

    // Get the sheet name associated with this trigger
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetName = scriptProperties.getProperty(`TRIGGER_${triggerId}_SHEET`);

    if (!sheetName) {
      Logger.log('No sheet name found for trigger: ' + triggerId);
      return;
    }

    // Get the spreadsheet and find the sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found in the spreadsheet.`);
      return;
    }

    // Activate the sheet
    sheet.activate();

    // Set it as the current sheet for the sync operation
    scriptProperties.setProperty('SHEET_NAME', sheetName);

    // Check if two-way sync is enabled for this sheet
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is enabled, push changes to Pipedrive first
    if (twoWaySyncEnabled) {
      Logger.log(`Two-way sync is enabled for sheet "${sheetName}". Pushing changes to Pipedrive before pulling new data.`);
      pushChangesToPipedrive(true); // Pass true to indicate this is a scheduled sync
    }

    // Get the entity type for this sheet
    const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Run the appropriate sync function based on entity type
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        syncDealsFromFilter();
        break;
      case ENTITY_TYPES.PERSONS:
        syncPersonsFromFilter();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        syncOrganizationsFromFilter();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        syncActivitiesFromFilter();
        break;
      case ENTITY_TYPES.LEADS:
        syncLeadsFromFilter();
        break;
      case ENTITY_TYPES.PRODUCTS:
        syncProductsFromFilter();
        break;
      default:
        Logger.log('Unknown entity type: ' + entityType);
        break;
    }

    // Log the completion
    Logger.log(`Scheduled sync completed for sheet "${sheetName}" with entity type "${entityType}"`);

  } catch (error) {
    Logger.log('Error in scheduled sync: ' + error.message);
  }
}

/**
 * Gets all triggers for a specific sheet
 */
function getTriggersForSheet(sheetName) {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    const scriptProperties = PropertiesService.getScriptProperties();

    return allTriggers.filter(trigger => {
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
  } catch (error) {
    Logger.log('Error getting triggers: ' + error.message);
    return [];
  }
}

/**
 * Gets readable information about a trigger
 */
function getTriggerInfo(trigger) {
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

  } catch (error) {
    Logger.log('Error getting trigger info: ' + error.message);
    return {
      type: 'Scheduled',
      description: 'Automatic sync'
    };
  }
}

/**
 * Deletes a trigger by its unique ID
 */
function deleteTrigger(triggerId) {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    const scriptProperties = PropertiesService.getScriptProperties();

    // Handle comma-separated list of trigger IDs (for weekly triggers)
    const triggerIds = triggerId.split(',');

    triggerIds.forEach(id => {
      // Find and delete the trigger with this ID
      for (let i = 0; i < allTriggers.length; i++) {
        if (allTriggers[i].getUniqueId() === id) {
          ScriptApp.deleteTrigger(allTriggers[i]);

          // Also clean up the properties
          scriptProperties.deleteProperty(`TRIGGER_${id}_SHEET`);
          scriptProperties.deleteProperty(`TRIGGER_${id}_FREQUENCY`);
          break;
        }
      }
    });

    return { success: true };
  } catch (error) {
    Logger.log('Error deleting trigger: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

function initializeTeamFeatures() {
  // Check if this is the first time running
  const documentProperties = PropertiesService.getDocumentProperties();
  const isTeamFeatureInitialized = documentProperties.getProperty('TEAM_FEATURE_INITIALIZED');

  if (!isTeamFeatureInitialized) {
    // Set flag to indicate team features are initialized
    documentProperties.setProperty('TEAM_FEATURE_INITIALIZED', 'true');

    // Create initial empty team data structure if it doesn't exist
    const teamsDataJson = documentProperties.getProperty('TEAMS_DATA');
    if (!teamsDataJson) {
      documentProperties.setProperty('TEAMS_DATA', '{}');
    }

    // Log initialization
    Logger.log('Team features initialized successfully');
  }
}

/**
 * Gets all available filters from Pipedrive for the current user
 * Modified to support team sharing of filters
 */
function getPipedriveFilters() {
  try {
    // Get API key from properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY') || API_KEY;
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

    if (!apiKey || apiKey === 'YOUR_PIPEDRIVE_API_KEY') {
      return { success: false, error: 'API key not configured' };
    }

    // Get the user's email
    const userEmail = Session.getActiveUser().getEmail();

    // Get the user's team
    const userTeam = getUserTeam(userEmail);

    // Construct the v1 API URL for filters (filters not yet available in v2)
    const url = `https://${subdomain}.pipedrive.com/v1/filters?api_token=${apiKey}`;

    Logger.log(`Fetching filters from: ${url}`);

    // Add muteHttpExceptions to get more detailed error information
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Filter API response code: ${responseCode}`);

    if (responseCode !== 200) {
      return {
        success: false,
        error: `Server returned status code ${responseCode}: ${responseText.substring(0, 100)}`
      };
    }

    try {
      const responseData = JSON.parse(responseText);

      if (!responseData.success) {
        return {
          success: false,
          error: responseData.error || 'Unknown error'
        };
      }

      // Store the user's filters for team sharing
      storeUserFilters(responseData.data);

      // Group filters by type for easier selection
      let filtersByType = {};

      // Group user's own filters first
      responseData.data.forEach(filter => {
        if (!filtersByType[filter.type]) {
          filtersByType[filter.type] = [];
        }

        // Flag this as the user's own filter
        filter.isOwn = true;

        filtersByType[filter.type].push({
          id: filter.id,
          name: filter.name,
          type: filter.type,
          isOwn: true
        });
      });

      // If the user is part of a team and filter sharing is enabled, add team members' filters
      if (userTeam && userTeam.shareFilters) {
        // Keep track of filter IDs to avoid duplicates
        const userFilterIds = new Set();
        for (const type in filtersByType) {
          filtersByType[type].forEach(filter => userFilterIds.add(filter.id));
        }

        // For each team member, get their filters from document properties
        for (const memberEmail of userTeam.memberEmails) {
          if (memberEmail === userEmail) continue; // Skip the current user

          // Get the team member's filters from document properties
          const memberFiltersKey = `TEAM_MEMBER_FILTERS_${memberEmail}`;
          const memberFiltersJson = scriptProperties.getProperty(memberFiltersKey);

          if (memberFiltersJson) {
            try {
              const memberFilters = JSON.parse(memberFiltersJson);

              // Add each member's filter if it's not a duplicate
              memberFilters.forEach(filter => {
                if (!userFilterIds.has(filter.id)) {
                  if (!filtersByType[filter.type]) {
                    filtersByType[filter.type] = [];
                  }

                  // Add the team member's filter with additional metadata
                  filtersByType[filter.type].push({
                    id: filter.id,
                    name: filter.name + ' (' + memberEmail.split('@')[0] + ')',
                    type: filter.type,
                    isTeamFilter: true,
                    ownerEmail: memberEmail
                  });

                  // Mark as processed to avoid duplicates
                  userFilterIds.add(filter.id);
                }
              });
            } catch (e) {
              Logger.log(`Error parsing ${memberEmail}'s filters: ${e.message}`);
            }
          }
        }
      }

      return {
        success: true,
        data: filtersByType
      };
    } catch (parseError) {
      return { success: false, error: `JSON Parse error: ${parseError.message}` };
    }
  } catch (error) {
    Logger.log('Error fetching filters: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Shows a dialog to help users find their filters
 */
function showFilterFinder() {
  // Get available filters
  const filtersResult = getPipedriveFilters();

  if (!filtersResult.success) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Could not retrieve filters: ' + filtersResult.error + '\n\nPlease make sure your API key is correct.', ui.ButtonSet.OK);
    return;
  }

  const filtersByType = filtersResult.data;

  // Create HTML content for the dialog
  let filtersHtml = '';

  for (const type in filtersByType) {
    const filters = filtersByType[type];

    if (filters.length > 0) {
      filtersHtml += `
        <div class="filter-group">
          <table class="filter-table">
            <thead>
              <tr>
                <th>Filter Name</th>
                <th>Filter ID</th>
              </tr>
            </thead>
            <tbody>
      `;

      filters.forEach(filter => {
        filtersHtml += `
          <tr>
            <td>${filter.name}</td>
            <td><code>${filter.id}</code> <button class="copy-btn" data-id="${filter.id}">Copy</button></td>
          </tr>
        `;
      });

      filtersHtml += `
            </tbody>
          </table>
        </div>
      `;
    }
  }

  // If no filters found
  if (filtersHtml === '') {
    filtersHtml = '<p>No filters found in your Pipedrive account. Please create filters in Pipedrive first.</p>';
  }

  const htmlOutput = HtmlService.createHtmlOutput(`
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
        margin: 12px 0;
        color: var(--text-dark);
      }
      
      .filter-group {
        margin-bottom: 24px;
      }
      
      .filter-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 12px;
      }
      
      .filter-table th,
      .filter-table td {
        padding: 8px 12px;
        text-align: left;
        border: 1px solid var(--border-color);
      }
      
      .filter-table th {
        background-color: var(--bg-light);
        font-weight: 500;
      }
      
      code {
        background-color: var(--bg-light);
        padding: 2px 4px;
        border-radius: 4px;
        font-family: monospace;
        font-size: 13px;
      }
      
      .copy-btn {
        background-color: var(--primary-color);
        color: white;
        border: none;
        border-radius: 4px;
        padding: 4px 8px;
        font-size: 12px;
        cursor: pointer;
        margin-left: 8px;
        transition: var(--transition);
      }
      
      .copy-btn:hover {
        background-color: var(--primary-dark);
      }
      
      .info {
        background-color: var(--bg-light);
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 20px;
        border-left: 4px solid var(--primary-color);
      }
      
      .back-btn {
        background-color: var(--bg-light);
        color: var(--text-dark);
        border: 1px solid var(--border-color);
        border-radius: 4px;
        padding: 8px 16px;
        font-size: 14px;
        cursor: pointer;
        margin-top: 16px;
        transition: var(--transition);
      }
      
      .back-btn:hover {
        background-color: #e8eaed;
      }
    </style>
    
    <h3>Your Pipedrive Filters</h3>
    
    <div class="info">
      Find your filters below. Click the "Copy" button next to a filter ID to copy it, then use it in the settings.
    </div>
    
    ${filtersHtml}
    
    <button class="back-btn" onclick="google.script.host.close()">Close</button>
    
    <script>
      // Add event listeners to copy buttons
      document.querySelectorAll('.copy-btn').forEach(btn => {
        btn.addEventListener('click', function() {
          const filterId = this.getAttribute('data-id');
          
          // Create a temporary textarea element to copy the text
          const textarea = document.createElement('textarea');
          textarea.value = filterId;
          document.body.appendChild(textarea);
          textarea.select();
          document.execCommand('copy');
          document.body.removeChild(textarea);
          
          // Change button text temporarily
          const originalText = this.textContent;
          this.textContent = 'Copied!';
          
          // Reset button text after a delay
          setTimeout(() => {
            this.textContent = originalText;
          }, 2000);
        });
      });
    </script>
  `)
    .setWidth(600)
    .setHeight(500)
    .setTitle('Your Pipedrive Filters');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Your Pipedrive Filters');
}

/**
 * Formats a filter type for display
 */
function formatFilterType(type) {
  // Map filter types to friendly names
  const typeMap = {
    'deals': 'Deals',
    'people': 'Persons',
    'org': 'Organizations',
    'activity': 'Activities',
    'leads': 'Leads',
    'products': 'Products',
  };

  return typeMap[type] || type.charAt(0).toUpperCase() + type.slice(1);
}

/**
 * Shows the two-way sync settings UI for configuring bidirectional updates
 */
function showTwoWaySyncSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();

  // Get current two-way sync settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
  const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
  const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

  const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';
  const trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
  const lastSync = scriptProperties.getProperty(twoWaySyncLastSyncKey) || 'Never';

  // Get sheet-specific entity type
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

  // Create HTML content for the settings dialog
  const htmlOutput = HtmlService.createHtmlOutput(`
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
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
        background-color: var(--primary-color);
      }
      
      input:focus + .slider {
        box-shadow: 0 0 1px var(--primary-color);
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
        margin-bottom: 20px;
      }
      
      .feature-table th,
      .feature-table td {
        padding: 8px 12px;
        text-align: left;
        border: 1px solid var(--border-color);
      }
      
      .feature-table th {
        background-color: var(--bg-light);
        font-weight: 500;
      }
    </style>
    
    <h3>Two-Way Sync Settings</h3>
    
    <div class="sheet-info">
      <svg xmlns="http://www.w3.org/2000/svg" height="24" viewBox="0 0 24 24" width="24">
        <path d="M0 0h24v24H0z" fill="none"/>
        <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10zm0-12H5V5h14v2zM7 11h10v2H7zm0 4h7v2H7z"/>
      </svg>
      <div>Configuring two-way sync for sheet "<strong>${activeSheetName}</strong>"</div>
    </div>
    
    <div class="info-alert">
      <h4>About Two-Way Sync</h4>
      <p>Two-way sync allows you to make changes in your sheet and push them back to Pipedrive. The system will track which rows have been modified since the last sync.</p>
    </div>
    
    <div class="form-container">
      <div class="form-group">
        <div class="switch-container">
          <label class="switch">
            <input type="checkbox" id="enableTwoWaySync" ${twoWaySyncEnabled ? 'checked' : ''}>
            <span class="slider"></span>
          </label>
          <strong>Enable Two-Way Sync for this sheet</strong>
        </div>
        <span class="tooltip">When enabled, a tracking column will be added to the sheet to track modified rows.</span>
      </div>
      
      <div class="form-group">
        <label for="trackingColumn">
          Tracking Column Letter (Optional)
        </label>
        <input type="text" id="trackingColumn" value="${trackingColumn}" placeholder="e.g., Z" />
        <span class="tooltip">Specify which column to use for tracking changes. Leave empty to use the last column.</span>
        <div class="last-sync">Last sync: ${lastSync}</div>
      </div>
      
      <div class="form-group">
        <h4>Supported Fields for Bidirectional Updates</h4>
        <p>The following fields for ${entityType} can be updated in Pipedrive:</p>
        <table class="feature-table">
          <thead>
            <tr>
              <th>Field Type</th>
              <th>Supported</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Basic fields (name, value, etc.)</td>
              <td>Yes</td>
              <td>All standard fields are supported</td>
            </tr>
            <tr>
              <td>Dropdown fields</td>
              <td>Yes</td>
              <td>Must use exact option values</td>
            </tr>
            <tr>
              <td>Date fields</td>
              <td>Yes</td>
              <td>Use YYYY-MM-DD format</td>
            </tr>
            <tr>
              <td>User assignments</td>
              <td>Yes</td>
              <td>Use Pipedrive user IDs</td>
            </tr>
            <tr>
              <td>Custom fields</td>
              <td>Yes</td>
              <td>All custom field types supported</td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <div id="status" class="status"></div>
      
      <div class="button-container">
        <button type="button" id="cancelBtn" class="button-secondary">Cancel</button>
        <button type="button" id="saveBtn" class="button-primary"><span id="saveSpinner" class="spinner" style="display:none;"></span>Save Settings</button>
      </div>
    </div>
    
    <script>
      // Initialize
      document.addEventListener('DOMContentLoaded', function() {
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
  `)
    .setWidth(600)
    .setHeight(650)
    .setTitle(`Two-Way Sync Settings for "${activeSheetName}"`);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Two-Way Sync Settings for "${activeSheetName}"`);
}

/**
 * Saves the two-way sync settings for a sheet and sets up the tracking column
 * @param {boolean} enableTwoWaySync Whether to enable two-way sync
 * @param {string} trackingColumn The column letter to use for tracking changes
 */
function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Get script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

    // Save settings to properties
    scriptProperties.setProperty(twoWaySyncEnabledKey, enableTwoWaySync.toString());
    scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);

    // Store the previous tracking column if it exists
    const previousTrackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const previousPosStr = scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${activeSheetName}`) || '-1';
    const previousPos = parseInt(previousPosStr, 10);
    const currentPos = trackingColumn ? columnLetterToIndex(trackingColumn) : -1;

    if (previousTrackingColumn && previousTrackingColumn !== trackingColumn) {
      scriptProperties.setProperty(`PREVIOUS_TRACKING_COLUMN_${activeSheetName}`, previousTrackingColumn);

      // NEW: Also track when columns have been removed (causing a left shift)
      if (previousPos >= 0 && currentPos >= 0 && currentPos < previousPos) {
        Logger.log(`Detected column removal: Sync Status moved left from ${previousPos} to ${currentPos}`);

        // Check all columns between previous and current positions (inclusive)
        // Important: Don't just check columns in between, check ALL columns from 0 to max(previousPos)
        const maxPos = Math.max(previousPos + 3, activeSheet.getLastColumn()); // Add buffer
        for (let i = 0; i <= maxPos; i++) {
          const colLetter = columnToLetter(i);
          if (colLetter !== trackingColumn) {
            // Look for sync status indicators in this column
            try {
              const headerCell = activeSheet.getRange(1, i + 1);  // i is 0-based, getRange is 1-based
              const headerValue = headerCell.getValue();
              const note = headerCell.getNote();

              // Extra check for Sync Status indicators
              if (headerValue === "Sync Status" ||
                (note && (note.includes('sync') || note.includes('track')))) {
                cleanupColumnFormatting(activeSheet, colLetter);
              }
            } catch (e) {
              Logger.log(`Error checking column ${colLetter}: ${e.message}`);
            }
          }
        }
      }
    }

    // Clean up previous Sync Status column formatting
    cleanupPreviousSyncStatusColumn(activeSheet, activeSheetName);

    // If enabling two-way sync, set up the tracking column
    if (enableTwoWaySync) {
      // Determine which column to use for tracking
      let trackingColumnIndex;
      if (trackingColumn) {
        // Convert column letter to index (0-based)
        trackingColumnIndex = columnLetterToIndex(trackingColumn);
      } else {
        // Use the last column
        trackingColumnIndex = activeSheet.getLastColumn();
      }

      // Set up the tracking column header
      const headerRow = 1; // Assuming first row is header
      const trackingHeader = "Sync Status";

      // Create the tracking column if it doesn't exist
      if (trackingColumnIndex >= activeSheet.getLastColumn()) {
        // Column doesn't exist yet, add it
        activeSheet.getRange(headerRow, trackingColumnIndex + 1).setValue(trackingHeader);

        // Update the tracking column letter based on the actual position
        const actualColumnIndex = trackingColumnIndex;
        trackingColumn = columnToLetter(actualColumnIndex);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
        Logger.log(`Created tracking column at position ${trackingColumnIndex + 1} (${trackingColumn})`);
      } else {
        // Column exists, update header
        activeSheet.getRange(headerRow, trackingColumnIndex + 1).setValue(trackingHeader);

        // Verify the tracking column letter is correct
        const actualColumnIndex = trackingColumnIndex;
        trackingColumn = columnToLetter(actualColumnIndex);
        scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
        Logger.log(`Updated tracking column at position ${trackingColumnIndex + 1} (${trackingColumn})`);
      }

      // Visually style the Sync Status column to distinguish it from data columns
      // Style the header cell
      const headerCell = activeSheet.getRange(headerRow, trackingColumnIndex + 1);
      headerCell.setBackground('#E8F0FE') // Light blue background
        .setFontWeight('bold')
        .setNote('This column tracks changes for two-way sync with Pipedrive');

      // Style the entire status column with a light background and border
      const fullStatusColumn = activeSheet.getRange(1, trackingColumnIndex + 1, Math.max(activeSheet.getLastRow(), 2), 1);
      fullStatusColumn.setBackground('#F8F9FA') // Light gray background
        .setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

      // Initialize all rows with "Not modified" status
      if (activeSheet.getLastRow() > 1) {
        // Get all data to identify which rows should have status
        const allData = activeSheet.getDataRange().getValues();
        const statusValues = activeSheet.getRange(2, trackingColumnIndex + 1, activeSheet.getLastRow() - 1, 1).getValues();
        const newStatusValues = [];

        // Process each row (starting from row 2)
        for (let i = 1; i < allData.length; i++) {
          const row = allData[i];
          const firstCell = row[0] ? row[0].toString().toLowerCase() : '';
          const isEmpty = row.every(cell => cell === '' || cell === null || cell === undefined);

          // Skip setting status for:
          // 1. Rows where first cell contains "last" or "sync" (metadata rows)
          // 2. Empty rows
          if (firstCell.includes('last') ||
            firstCell.includes('sync') ||
            firstCell.includes('update') ||
            isEmpty) {
            newStatusValues.push(['']); // Keep empty
          } else {
            // Only set "Not modified" for actual data rows that are empty
            const currentStatus = statusValues[i - 1][0];
            newStatusValues.push([
              currentStatus === '' || currentStatus === null || currentStatus === undefined
                ? 'Not modified'
                : currentStatus
            ]);
          }
        }

        // Apply the statuses
        activeSheet.getRange(2, trackingColumnIndex + 1, activeSheet.getLastRow() - 1, 1).setValues(newStatusValues);
      }

      // Format the tracking column
      const dataRange = activeSheet.getDataRange();

      // Add conditional formatting for the status column
      if (trackingColumnIndex > -1) {
        const trackingRange = activeSheet.getRange(headerRow + 1, trackingColumnIndex + 1, Math.max(activeSheet.getLastRow() - headerRow, 1), 1);
        if (trackingRange.getNumRows() > 0) {
          // Get all existing conditional formatting rules
          const existingRules = activeSheet.getConditionalFormatRules();
          // Clear any existing rules for the tracking column
          const newRules = existingRules.filter(rule => {
            const ranges = rule.getRanges();
            return !ranges.some(range => range.getColumn() === trackingColumnIndex + 1);
          });

          // Create conditional format for "Modified" status
          const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Modified')
            .setBackground('#FCE8E6')  // Light red background
            .setFontColor('#D93025')   // Red text
            .setRanges([trackingRange])
            .build();
          newRules.push(modifiedRule);

          // Create conditional format for "Synced" status
          const syncedRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Synced')
            .setBackground('#E6F4EA')  // Light green background
            .setFontColor('#137333')   // Green text
            .setRanges([trackingRange])
            .build();
          newRules.push(syncedRule);

          // Create conditional format for "Error" status
          const errorRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Error')
            .setBackground('#FCE8E6')  // Light red background
            .setFontColor('#D93025')   // Red text
            .setBold(true)             // Bold text for errors
            .setRanges([trackingRange])
            .build();
          newRules.push(errorRule);

          // Apply all rules
          activeSheet.setConditionalFormatRules(newRules);
        }
      }

      // Store the current position for future reference
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${activeSheetName}`, trackingColumnIndex.toString());
    }

    // Set up the onEdit trigger to detect changes
    setupOnEditTrigger();

    // Update last sync time
    const now = new Date();
    scriptProperties.setProperty(twoWaySyncLastSyncKey, now.toISOString());
  } catch (error) {
    Logger.log(`Error saving two-way sync settings: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Converts a column letter (A, B, C, ..., Z, AA, AB, etc.) to a 0-based index
 * @param {string} columnLetter The column letter
 * @return {number} The 0-based column index
 */
function columnLetterToIndex(columnLetter) {
  if (!columnLetter) {
    Logger.log('Warning: Empty column letter passed to columnLetterToIndex');
    return -1;
  }

  columnLetter = columnLetter.toUpperCase();
  Logger.log(`Converting column letter ${columnLetter} to index`);

  let index = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    index = index * 26 + (columnLetter.charCodeAt(i) - 64);
  }

  const result = index - 1; // Convert to 0-based index
  Logger.log(`Column letter ${columnLetter} converted to index ${result}`);
  return result;
}

// Function to clean up previous Sync Status column formatting
function cleanupPreviousSyncStatusColumn(sheet, sheetName) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const previousColumnKey = `PREVIOUS_TRACKING_COLUMN_${sheetName}`;
    const currentColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    const previousColumn = scriptProperties.getProperty(previousColumnKey) || '';

    // Clean up the specifically tracked previous column
    if (previousColumn && previousColumn !== currentColumn) {
      cleanupColumnFormatting(sheet, previousColumn);
      scriptProperties.deleteProperty(previousColumnKey);
    }

    // NEW: Store current column positions for future comparison
    // This helps when columns are deleted and the position shifts
    const currentColumnIndex = currentColumn ? columnLetterToIndex(currentColumn) : -1;
    if (currentColumnIndex >= 0) {
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, currentColumnIndex.toString());
    }

    // IMPORTANT: Scan ALL columns for "Sync Status" headers and validation patterns
    scanAndCleanupAllSyncColumns(sheet, currentColumn);
  } catch (error) {
    Logger.log(`Error in cleanupPreviousSyncStatusColumn: ${error.message}`);
  }
}

// New helper function to consolidate cleanup logic
function cleanupColumnFormatting(sheet, columnLetter) {
  try {
    const columnIndex = columnLetterToIndex(columnLetter);
    const columnPos = columnIndex + 1; // 1-based for getRange

    // ALWAYS clean columns, even if they're beyond the current last column
    Logger.log(`Cleaning up formatting for column ${columnLetter} (position ${columnPos})`);

    try {
      // Clean up data validations for the ENTIRE column
      const numRows = Math.max(sheet.getMaxRows(), 1000); // Use a large number to ensure all rows

      // Try to clear data validations - may fail for columns beyond the edge
      try {
        sheet.getRange(1, columnPos, numRows, 1).clearDataValidations();
      } catch (e) {
        Logger.log(`Could not clear validations for column ${columnLetter}: ${e.message}`);
      }

      // Clear header note - this should work even for "out of bounds" columns
      try {
        sheet.getRange(1, columnPos).clearNote();
        sheet.getRange(1, columnPos).setNote(''); // Force clear
      } catch (e) {
        Logger.log(`Could not clear note for column ${columnLetter}: ${e.message}`);
      }

      // For columns within the sheet, we can do more thorough cleaning
      if (columnPos <= sheet.getLastColumn()) {
        // Clear all formatting for the entire column (data rows only, not header)
        if (sheet.getLastRow() > 1) {
          try {
            // Clear formatting for data rows (row 2 and below), preserving header
            sheet.getRange(2, columnPos, numRows - 1, 1).clear({
              formatOnly: true,
              contentsOnly: false,
              validationsOnly: true
            });
          } catch (e) {
            Logger.log(`Error clearing data rows: ${e.message}`);
          }
        }

        // Clear formatting for header separately, preserving bold
        try {
          const headerCell = sheet.getRange(1, columnPos);
          const headerValue = headerCell.getValue();

          // Reset all formatting except bold
          headerCell.setBackground(null)
            .setBorder(null, null, null, null, null, null)
            .setFontColor(null);

          // Ensure header is bold
          headerCell.setFontWeight('bold');

        } catch (e) {
          Logger.log(`Error formatting header: ${e.message}`);
        }

        // Additionally clear specific formatting for data rows
        try {
          if (sheet.getLastRow() > 1) {
            const dataRows = sheet.getRange(2, columnPos, Math.max(sheet.getLastRow() - 1, 1), 1);
            dataRows.setBackground(null);
            dataRows.setBorder(null, null, null, null, null, null);
            dataRows.setFontColor(null);
            dataRows.setFontWeight(null);
          }
        } catch (e) {
          Logger.log(`Error clearing specific formatting: ${e.message}`);
        }

        // Clear conditional formatting specifically for this column
        const rules = sheet.getConditionalFormatRules();
        let newRules = [];
        let removedRules = 0;

        for (const rule of rules) {
          const ranges = rule.getRanges();
          let shouldRemove = false;

          // Check if any range in this rule applies to our column
          for (const range of ranges) {
            if (range.getColumn() === columnPos) {
              shouldRemove = true;
              break;
            }
          }

          if (!shouldRemove) {
            newRules.push(rule);
          } else {
            removedRules++;
          }
        }

        if (removedRules > 0) {
          sheet.setConditionalFormatRules(newRules);
          Logger.log(`Removed ${removedRules} conditional formatting rules from column ${columnLetter}`);
        }
      } else {
        Logger.log(`Column ${columnLetter} is beyond the sheet bounds (${sheet.getLastColumn()}), did minimal cleanup`);
      }

      Logger.log(`Completed cleanup for column ${columnLetter}`);
    } catch (innerError) {
      Logger.log(`Error during column cleanup operations: ${innerError.message}`);
    }
  } catch (error) {
    Logger.log(`Error cleaning up column ${columnLetter}: ${error.message}`);
  }
}

// Enhanced function to scan all columns for Sync Status headers
function scanAndCleanupAllSyncColumns(sheet, currentColumnLetter) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentColumnIndex = currentColumnLetter ? columnLetterToIndex(currentColumnLetter) : -1;

  // Track all columns with data validation (which sync status columns would have)
  const lastRow = Math.max(sheet.getLastRow(), 2);
  let columnsWithValidation = [];

  if (lastRow > 1) {
    // Check ALL columns for data validation instead of just sampling
    for (let i = 0; i < headers.length; i++) {
      // Skip the current tracking column
      if (i === currentColumnIndex) {
        continue;
      }

      try {
        // Check the entire column for validation
        const columnPos = i + 1; // 1-based for getRange
        const validationColumn = sheet.getRange(2, columnPos, lastRow - 1, 1);

        // Check if the column has data validation with specific values
        const validations = validationColumn.getDataValidations();
        for (let j = 0; j < validations.length; j++) {
          if (validations[j][0] !== null) {
            try {
              const criteria = validations[j][0].getCriteriaType();
              const values = validations[j][0].getCriteriaValues();

              // Look for our specific validation values (Modified, Not modified, etc.)
              if (values &&
                (values.join(',').includes('Modified') ||
                  values.join(',').includes('Not modified') ||
                  values.join(',').includes('Synced'))) {
                columnsWithValidation.push(i);
                Logger.log(`Found Sync Status validation at ${columnToLetter(i)}`);
                break; // Found validation in this column, no need to check more cells
              }
            } catch (e) {
              // Some validation types may not support getCriteriaValues
              Logger.log(`Error checking validation criteria: ${e.message}`);
            }
          }
        }
      } catch (error) {
        Logger.log(`Error checking column ${i} for validation: ${error.message}`);
      }
    }
  }

  // Check for columns with header notes that might be Sync Status columns
  for (let i = 0; i < headers.length; i++) {
    // Skip the current tracking column
    if (i === currentColumnIndex) {
      continue;
    }

    try {
      const columnPos = i + 1; // 1-based for getRange
      const headerCell = sheet.getRange(1, columnPos);
      const note = headerCell.getNote();

      if (note && (note.includes('sync') || note.includes('track') || note.includes('Pipedrive'))) {
        Logger.log(`Found column with sync-related note at ${columnToLetter(i)}: "${note}"`);
        const columnToClean = columnToLetter(i);
        cleanupColumnFormatting(sheet, columnToClean);
      }
    } catch (error) {
      Logger.log(`Error checking column ${i} header note: ${error.message}`);
    }
  }

  // Look for columns with "Sync Status" header
  for (let i = 0; i < headers.length; i++) {
    const headerValue = headers[i];

    // Skip the current tracking column
    if (i === currentColumnIndex) {
      continue;
    }

    // If we find a "Sync Status" header in a different column, clean it up
    if (headerValue === "Sync Status") {
      const columnToClean = columnToLetter(i);
      Logger.log(`Found orphaned Sync Status column at ${columnToClean}, cleaning up`);
      cleanupColumnFormatting(sheet, columnToClean);
    }
  }

  // Then, also clean up any columns with validation that aren't the current column
  for (const columnIndex of columnsWithValidation) {
    const columnToClean = columnToLetter(columnIndex);
    Logger.log(`Found column with validation at ${columnToClean}, cleaning up`);
    cleanupColumnFormatting(sheet, columnToClean);
  }

  // Add after your existing column checks
  // Also scan for orphaned formatting that matches your Sync Status patterns
  const dataRange = sheet.getDataRange();
  const numRows = Math.min(dataRange.getNumRows(), 10); // Check first 10 rows for formatting
  const numCols = dataRange.getNumColumns();

  // Look at ALL columns for specific background colors
  for (let col = 1; col <= numCols; col++) {
    // Skip current column 
    if (col === currentColumnIndex + 1) continue;

    let foundSyncStatusFormatting = false;

    try {
      // Check for Sync Status cell colors in first few rows
      for (let row = 2; row <= numRows; row++) {
        const cell = sheet.getRange(row, col);
        const background = cell.getBackground();

        // Check for our specific Sync Status background colors
        if (background === '#FCE8E6' || // Error/Modified 
          background === '#E6F4EA' || // Synced
          background === '#F8F9FA') { // Default
          foundSyncStatusFormatting = true;
          break;
        }
      }

      if (foundSyncStatusFormatting) {
        const colLetter = columnToLetter(col - 1); // col is 1-based, columnToLetter expects 0-based
        Logger.log(`Found column with Sync Status formatting at ${colLetter}, cleaning up`);
        cleanupColumnFormatting(sheet, colLetter);
      }
    } catch (error) {
      Logger.log(`Error checking column ${col} for formatting: ${error.message}`);
    }
  }

  // Additionally check for specific formatting that matches Sync Status patterns
  cleanupOrphanedConditionalFormatting(sheet, currentColumnIndex);
}

function detectColumnShifts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    const scriptProperties = PropertiesService.getScriptProperties();

    // Get current and previous positions
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const currentColLetter = scriptProperties.getProperty(trackingColumnKey) || '';
    const previousPosStr = scriptProperties.getProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`) || '-1';
    const previousPos = parseInt(previousPosStr, 10);

    // Find all "Sync Status" headers in the sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColumns = [];

    // Find ALL instances of "Sync Status" headers
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Sync Status") {
        syncStatusColumns.push(i);
      }
    }

    // If we have multiple "Sync Status" columns, clean up all but the rightmost one
    if (syncStatusColumns.length > 1) {
      Logger.log(`Found ${syncStatusColumns.length} Sync Status columns`);
      // Keep only the rightmost column
      const rightmostIndex = Math.max(...syncStatusColumns);

      // Clean up all other columns
      for (const colIndex of syncStatusColumns) {
        if (colIndex !== rightmostIndex) {
          const colLetter = columnToLetter(colIndex);
          Logger.log(`Cleaning up duplicate Sync Status column at ${colLetter}`);
          cleanupColumnFormatting(sheet, colLetter);
        }
      }

      // Update the tracking to the rightmost column
      const rightmostColLetter = columnToLetter(rightmostIndex);
      scriptProperties.setProperty(trackingColumnKey, rightmostColLetter);
      scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, rightmostIndex.toString());
      return; // Exit after handling duplicates
    }

    let actualSyncStatusIndex = syncStatusColumns.length > 0 ? syncStatusColumns[0] : -1;

    if (actualSyncStatusIndex >= 0) {
      const actualColLetter = columnToLetter(actualSyncStatusIndex);

      // If there's a mismatch, columns might have shifted
      if (currentColLetter && actualColLetter !== currentColLetter) {
        Logger.log(`Column shift detected: was ${currentColLetter}, now ${actualColLetter}`);

        // If the actual position is less than the recorded position, columns were removed
        if (actualSyncStatusIndex < previousPos) {
          Logger.log(`Columns were likely removed (${previousPos} → ${actualSyncStatusIndex})`);

          // Clean ALL columns to be safe
          for (let i = 0; i < sheet.getLastColumn(); i++) {
            if (i !== actualSyncStatusIndex) { // Skip current Sync Status column
              cleanupColumnFormatting(sheet, columnToLetter(i));
            }
          }
        }

        // Clean up all potential previous locations
        scanAndCleanupAllSyncColumns(sheet, actualColLetter);

        // Update the tracking column property
        scriptProperties.setProperty(trackingColumnKey, actualColLetter);
        scriptProperties.setProperty(`CURRENT_SYNCSTATUS_POS_${sheetName}`, actualSyncStatusIndex.toString());
      }
    }
  } catch (error) {
    Logger.log(`Error in detectColumnShifts: ${error.message}`);
  }
}

// New helper function to clean up orphaned conditional formatting
function cleanupOrphanedConditionalFormatting(sheet, currentColumnIndex) {
  try {
    const rules = sheet.getConditionalFormatRules();
    const newRules = [];
    let removedRules = 0;

    for (const rule of rules) {
      const ranges = rule.getRanges();
      let keepRule = true;

      // Check if this rule applies to columns other than our current one
      // and has formatting that matches our Sync Status patterns
      for (const range of ranges) {
        const column = range.getColumn();

        // Skip our current column
        if (column === (currentColumnIndex + 1)) {
          continue;
        }

        // Check if this rule's formatting matches our Sync Status patterns
        const bgColor = rule.getBold() || rule.getBackground();
        if (bgColor) {
          const background = rule.getBackground();
          // If background matches our Sync Status colors, this is likely an orphaned rule
          if (background === '#FCE8E6' || background === '#E6F4EA' || background === '#F8F9FA') {
            keepRule = false;
            Logger.log(`Found orphaned conditional formatting at column ${columnToLetter(column - 1)}`);
            break;
          }
        }
      }

      if (keepRule) {
        newRules.push(rule);
      } else {
        removedRules++;
      }
    }

    if (removedRules > 0) {
      sheet.setConditionalFormatRules(newRules);
      Logger.log(`Removed ${removedRules} orphaned conditional formatting rules`);
    }
  } catch (error) {
    Logger.log(`Error cleaning up orphaned conditional formatting: ${error.message}`);
  }
}

/**
 * Sets up an onEdit trigger to detect changes in the sheet
 */
function setupOnEditTrigger() {
  // Check if the trigger already exists
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      // Trigger already exists
      return;
    }
  }

  // Create a new trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * Removes the onEdit trigger
 */
function removeOnEditTrigger() {
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Handles edits to the sheet and marks rows as modified for two-way sync
 * This function is automatically triggered when a user edits the sheet
 * @param {Object} e The edit event object
 */
function onEdit(e) {
  try {
    // Get the edited sheet
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    // If two-way sync is not enabled, exit
    if (!twoWaySyncEnabled) {
      return;
    }

    // Get the tracking column
    let trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    let trackingColumnIndex;

    if (trackingColumn) {
      // Convert column letter to index (0-based)
      trackingColumnIndex = columnLetterToIndex(trackingColumn);
    } else {
      // Use the last column
      trackingColumnIndex = sheet.getLastColumn() - 1;
    }

    // Get the edited range
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();

    // Check if the edit is in the tracking column itself (to avoid loops)
    if (column === trackingColumnIndex + 1) {
      return;
    }

    // Check if the edit is in the header row
    const headerRow = 1;
    if (row === headerRow) {
      return;
    }

    // Get the row content to check if it's a real data row or a timestamp/blank row
    const rowContent = sheet.getRange(row, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0];

    // Check if this is a timestamp row
    const firstCell = String(rowContent[0] || "").toLowerCase();
    const isTimestampRow = firstCell.includes("last") ||
      firstCell.includes("updated") ||
      firstCell.includes("synced") ||
      firstCell.includes("date");

    // Count non-empty cells to determine if this is a data row
    const nonEmptyCells = rowContent.filter(cell => cell !== "" && cell !== null && cell !== undefined).length;

    // Skip if this is a timestamp row or has too few cells with data
    if (isTimestampRow || nonEmptyCells < 3) {
      return;
    }

    // Get the row ID from the first column
    const idColumnIndex = 0;
    const id = rowContent[idColumnIndex];

    // Skip rows without an ID (likely empty rows)
    if (!id) {
      return;
    }

    // Update the tracking column to mark as modified
    const trackingRange = sheet.getRange(row, trackingColumnIndex + 1);
    const currentStatus = trackingRange.getValue();

    // Only mark as modified if it's not already marked or if it was previously synced
    if (currentStatus === "Not modified" || currentStatus === "Synced") {
      trackingRange.setValue("Modified");

      // Re-apply data validation to ensure consistent dropdown options
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
        .build();
      trackingRange.setDataValidation(rule);

      // Make sure the styling is consistent
      // This will be overridden by conditional formatting but helps with visual feedback
      trackingRange.setBackground('#FCE8E6').setFontColor('#D93025');
    }
  } catch (error) {
    // Silent fail for onEdit triggers
    Logger.log(`Error in onEdit trigger: ${error.message}`);
  }
}

// Add this function to your code
function getOptionIdByLabel(fieldKey, optionLabel, entityType) {
  try {
    // Get field definitions
    let fieldDefinitions = [];
    switch (entityType) {
      case ENTITY_TYPES.PERSONS:
        fieldDefinitions = getPersonFields();
        break;
      case ENTITY_TYPES.DEALS:
        fieldDefinitions = getDealFields();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fieldDefinitions = getOrganizationFields();
        break;
      case ENTITY_TYPES.PRODUCTS:
        fieldDefinitions = getProductFields();
        break;
      default:
        return null;
    }

    // Find the field
    const field = fieldDefinitions.find(f => f.key === fieldKey);
    if (!field || !field.options || !Array.isArray(field.options)) {
      Logger.log(`No options found for field ${fieldKey}`);
      return null;
    }

    // Find the option by label
    const option = field.options.find(o =>
      (o.label && o.label.toLowerCase() === optionLabel.toLowerCase()) ||
      (o.name && o.name.toLowerCase() === optionLabel.toLowerCase())
    );

    if (option && (option.id !== undefined || option.value !== undefined)) {
      Logger.log(`Found option ID ${option.id || option.value} for label "${optionLabel}"`);
      return option.id !== undefined ? option.id : option.value;
    }

    Logger.log(`No matching option found for label "${optionLabel}" in field ${fieldKey}`);
    return null;
  } catch (e) {
    Logger.log(`Error getting option ID: ${e.message}`);
    return null;
  }
}

/**
 * Sends modified data from the sheet back to Pipedrive
 * This function handles the bulk update functionality
 */
function pushChangesToPipedrive(isScheduledSync = false, suppressNoModifiedWarning = false) {
  detectColumnShifts();
  try {
    // Get the active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    const twoWaySyncLastSyncKey = `TWOWAY_SYNC_LAST_SYNC_${activeSheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    if (!twoWaySyncEnabled) {
      // Show an error message if two-way sync is not enabled, only for manual syncs
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Two-Way Sync Not Enabled',
          'Two-way sync is not enabled for this sheet. Please enable it in the Two-Way Sync Settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Get sheet-specific entity type
    const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
    const entityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;

    // Get API key from properties
    const apiKey = scriptProperties.getProperty('PIPEDRIVE_API_KEY');
    const subdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;

    if (!apiKey) {
      // Show an error message if API key is not set, only for manual syncs
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'API Key Not Found',
          'Please set your Pipedrive API key in the Settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Get the original column configuration that maps headers to field keys
    const columnSettingsKey = `COLUMNS_${activeSheetName}_${entityType}`;
    const savedColumnsJson = scriptProperties.getProperty(columnSettingsKey);
    let columnConfig = [];

    if (savedColumnsJson) {
      try {
        columnConfig = JSON.parse(savedColumnsJson);
        Logger.log(`Retrieved column configuration for ${entityType}`);
      } catch (e) {
        Logger.log(`Error parsing column configuration: ${e.message}`);
      }
    }

    // Create a mapping from display names to field keys
    const headerToFieldKeyMap = {};
    columnConfig.forEach(col => {
      const displayName = col.customName || col.name;
      headerToFieldKeyMap[displayName] = col.key;
    });

    // Get the tracking column
    let trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    let trackingColumnIndex;

    Logger.log(`Retrieved tracking column from properties: "${trackingColumn}"`);

    // Look for a column named "Sync Status"
    const headerRow = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
    let syncStatusColumnIndex = -1;

    // First, try to use the stored tracking column if available
    if (trackingColumn && trackingColumn.trim() !== '') {
      trackingColumnIndex = columnLetterToIndex(trackingColumn);
      // Verify the header matches "Sync Status"
      if (trackingColumnIndex >= 0 && trackingColumnIndex < headerRow.length) {
        if (headerRow[trackingColumnIndex] === "Sync Status") {
          Logger.log(`Using configured tracking column ${trackingColumn} (index: ${trackingColumnIndex})`);
          syncStatusColumnIndex = trackingColumnIndex;
        }
      }
    }

    // If not found by letter or the header doesn't match, search for "Sync Status" header
    if (syncStatusColumnIndex === -1) {
      for (let i = 0; i < headerRow.length; i++) {
        if (headerRow[i] === "Sync Status") {
          syncStatusColumnIndex = i;
          Logger.log(`Found Sync Status column at index ${syncStatusColumnIndex}`);

          // Update the stored tracking column letter
          trackingColumn = columnToLetter(syncStatusColumnIndex);
          scriptProperties.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
          break;
        }
      }
    }

    // Use the found column index
    trackingColumnIndex = syncStatusColumnIndex;

    // Validate tracking column index
    if (trackingColumnIndex < 0 || trackingColumnIndex >= activeSheet.getLastColumn()) {
      Logger.log(`Invalid tracking column index ${trackingColumnIndex}, cannot proceed with sync`);
      if (!isScheduledSync) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Sync Status Column Not Found',
          'The Sync Status column could not be found. Please check your two-way sync settings.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Double check the header of the tracking column
    const trackingHeader = activeSheet.getRange(1, trackingColumnIndex + 1).getValue();
    Logger.log(`Tracking column header: "${trackingHeader}" at column index ${trackingColumnIndex} (column ${trackingColumnIndex + 1})`);

    // Get the data range
    const dataRange = activeSheet.getDataRange();
    const values = dataRange.getValues();

    // Get column headers (first row)
    const headers = values[0];

    // Find the ID column index (usually first column)
    const idColumnIndex = 0; // Assuming ID is in column A (index 0)

    // Get field mappings based on entity type
    const fieldMappings = getFieldMappingsForEntity(entityType);

    // Track rows that need updating
    const modifiedRows = [];

    // Collect modified rows
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const syncStatus = row[trackingColumnIndex];

      // Only process rows marked as "Modified"
      if (syncStatus === 'Modified') {
        // Use let instead of const for rowId since we might need to change it
        let rowId = row[idColumnIndex];

        // Skip rows without an ID
        if (!rowId) {
          continue;
        }

        // For products, ensure we're using the correct ID field
        if (entityType === ENTITY_TYPES.PRODUCTS) {
          // If the first column contains a name instead of an ID, 
          // try to look for the ID in another column
          const idColumnName = 'ID'; // This should match your header for the ID column
          const idColumnIdx = headers.indexOf(idColumnName);
          if (idColumnIdx !== -1 && idColumnIdx !== idColumnIndex) {
            // Use the value from the specific ID column
            rowId = row[idColumnIdx]; // Using let above allows this reassignment
            Logger.log(`Using product ID ${rowId} from column ${idColumnName} instead of first column value ${row[idColumnIndex]}`);
          }
        }

        // Create an object with field values to update
        const updateData = {};

        // For API v2 custom fields
        if (!entityType.endsWith('Fields') && entityType !== ENTITY_TYPES.LEADS) {
          // Initialize custom fields container for API v2 as an object, not an array
          updateData.custom_fields = {};
        }

        // Map column values to API fields
        for (let j = 0; j < headers.length; j++) {
          // Skip the tracking column
          if (j === trackingColumnIndex) {
            continue;
          }

          const header = headers[j];
          const value = row[j];

          // Skip empty values
          if (value === '' || value === null || value === undefined) {
            continue;
          }

          // Get the field key for this header - first try the stored column config
          const fieldKey = headerToFieldKeyMap[header] || fieldMappings[header];

          // Skip headers that don't have a mapping or if it's the ID field (since ID is in the URL path)
          if (!fieldKey || fieldKey === 'id') {
            if (fieldKey === 'id') {
              Logger.log(`Skipping ID field - this is already in the URL path`);
            } else {
              Logger.log(`Skipping field "${header}" - no mapping found for Pipedrive API`);
            }
            continue;
          }

          // Special handling for products - don't include owner_id.name directly
          if (entityType === ENTITY_TYPES.PRODUCTS && fieldKey === 'owner_id.name') {
            // Skip owner_id.name because Pipedrive products API requires numeric owner_id
            Logger.log(`Skipping owner_id.name for products as it requires numeric owner_id`);
            continue;
          }

          // Handle custom fields vs standard fields
          if (fieldKey.length > 30 && /^[0-9a-f]+$/.test(fieldKey)) {
            // This is likely a custom field ID (long hex string)
            if (entityType.endsWith('Fields') || entityType === ENTITY_TYPES.LEADS) {
              // For API v1 endpoints (fields and leads), use the format "field_key"
              updateData[fieldKey] = isDateField(fieldKey) ? convertToStandardDateFormat(value) : value;
            } else {
              // Make sure custom_fields is initialized
              if (!updateData.custom_fields) {
                updateData.custom_fields = {};
              }

              // For API v2 endpoints, add to custom_fields object with field key as property
              let processedValue = value;

              // Extra logging to diagnose the issue
              Logger.log(`Processing custom field ${fieldKey} with value: ${value} (type: ${typeof value})`);

              // Check if this is a multi-option field that requires array format
              if (isMultiOptionField(fieldKey, entityType)) {
                // First get the array of label values (as you're already doing)
                let labelsArray = [];
                if (Array.isArray(value)) {
                  labelsArray = value;
                } else if (typeof value === 'string') {
                  // For string values, handle different possible formats
                  if (value.startsWith('[') && value.endsWith(']')) {
                    // Looks like JSON array string - parse it
                    try {
                      labelsArray = JSON.parse(value);
                      Logger.log(`Converted JSON string to array for multi-option field: ${JSON.stringify(labelsArray)}`);
                    } catch (e) {
                      // If parsing fails, create single-item array
                      labelsArray = [value];
                      Logger.log(`Created single-item array for multi-option field: ${JSON.stringify(labelsArray)}`);
                    }
                  } else if (value.includes(',')) {
                    // Comma-separated list - split into array
                    labelsArray = value.split(',').map(item => item.trim());
                    Logger.log(`Split comma-separated string into array for multi-option field: ${JSON.stringify(labelsArray)}`);
                  } else if (value.trim() !== '') {
                    // Single value - wrap in array
                    labelsArray = [value];
                    Logger.log(`Wrapped single value in array for multi-option field: ${JSON.stringify(labelsArray)}`);
                  } else {
                    // Empty string - use empty array
                    labelsArray = [];
                  }
                } else if (value === null || value === undefined) {
                  // Null/undefined - use empty array
                  labelsArray = [];
                } else {
                  // Other values - wrap in array
                  labelsArray = [value];
                }

                // NEW CODE: Convert labels to IDs
                const idArray = [];
                for (const label of labelsArray) {
                  if (label && typeof label === 'string') {
                    const optionId = getOptionIdByLabel(fieldKey, label, entityType);
                    if (optionId !== null) {
                      // Use the numeric ID
                      idArray.push(optionId);
                      Logger.log(`Converted option label "${label}" to ID ${optionId}`);
                    } else {
                      // If we can't find the ID, try to convert to number if possible
                      const numericLabel = Number(label);
                      if (!isNaN(numericLabel)) {
                        idArray.push(numericLabel); // If it can be converted to a number, use that
                        Logger.log(`Using numeric value ${numericLabel} for option "${label}"`);
                      } else {
                        // Last resort: use the string as-is
                        Logger.log(`Warning: Could not find ID for option "${label}" - using as-is`);
                        idArray.push(label);
                      }
                    }
                  } else if (typeof label === 'number') {
                    // If it's already a number, assume it's an ID
                    idArray.push(label);
                  } else if (label !== null && label !== undefined) {
                    // Any other non-null value
                    idArray.push(label);
                  }
                }

                Logger.log(`Final option IDs for field ${fieldKey}: ${JSON.stringify(idArray)}`);

                // CRITICAL: Use the array of IDs instead of the original array
                updateData.custom_fields[fieldKey] = idArray;
                continue; // Skip the remaining processing for this field
              }

              // For non-multi-option fields, process date values
              if (value instanceof Date) {
                // Format as YYYY-MM-DD without time component
                const year = value.getFullYear();
                const month = String(value.getMonth() + 1).padStart(2, '0');
                const day = String(value.getDate()).padStart(2, '0');
                processedValue = `${year}-${month}-${day}`;
                Logger.log(`Converted Date object to YYYY-MM-DD format: ${processedValue}`);
              }
              // For string values that look like dates
              else if (typeof value === 'string' && /\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}/.test(value)) {
                Logger.log(`Value looks like a date, converting: ${value}`);
                processedValue = convertToStandardDateFormat(value);
              }
              // Check if it's a date field based on the field key
              else if (isDateField(fieldKey)) {
                Logger.log(`Field key ${fieldKey} detected as date field, converting: ${value}`);
                processedValue = convertToStandardDateFormat(value);
              }

              // Assign to custom_fields
              updateData.custom_fields[fieldKey] = processedValue;
            }
          } else {
            // This is a standard field
            // Check if this might be a date field and convert if needed
            if (isDateField(fieldKey)) {
              updateData[fieldKey] = convertToStandardDateFormat(value);
            } else {
              updateData[fieldKey] = value;
            }
          }
        }

        // If using API v2 and there are no custom fields, remove the empty object
        if (!entityType.endsWith('Fields') && entityType !== ENTITY_TYPES.LEADS &&
          updateData.custom_fields && Object.keys(updateData.custom_fields).length === 0) {
          delete updateData.custom_fields;
        }

        // Add to the list of rows to update
        modifiedRows.push({
          id: rowId,
          rowIndex: i + 1, // Convert to 1-based index for sheet access
          data: updateData
        });
      }
    }

    // If no modified rows found, inform the user (only for manual syncs)
    if (modifiedRows.length === 0) {
      if (!isScheduledSync && !suppressNoModifiedWarning) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'No Modified Data',
          'No modified rows found to update in Pipedrive.',
          ui.ButtonSet.OK
        );
      }
      return;
    }

    // Confirm before proceeding (only for manual syncs)
    if (!isScheduledSync) {
      const ui = SpreadsheetApp.getUi();
      const confirmation = ui.alert(
        'Push Changes to Pipedrive',
        `You are about to push ${modifiedRows.length} modified ${modifiedRows.length === 1 ? 'row' : 'rows'} to Pipedrive. Continue?`,
        ui.ButtonSet.YES_NO
      );

      if (confirmation !== ui.Button.YES) {
        return;
      }

      // Show a loading message for manual syncs
      SpreadsheetApp.getActiveSpreadsheet().toast(`Pushing ${modifiedRows.length} changes to Pipedrive...`, 'Syncing');
    } else {
      // Log for scheduled syncs
      Logger.log(`Pushing ${modifiedRows.length} changes to Pipedrive as part of scheduled sync`);
    }

    // Process each modified row
    const results = [];
    const successRowIndices = []; // Track successful row indices for batch update later

    for (const row of modifiedRows) {
      try {
        // Construct the update URL and method based on entity type
        let updateUrl;
        let method;

        // Check if this is a field-related entity type
        if (entityType.endsWith('Fields')) {
          // Field endpoints use v1 API with PUT method
          updateUrl = `https://${subdomain}.pipedrive.com/v1/${entityType}/${row.id}?api_token=${apiKey}`;
          method = 'PUT';
        } else if (entityType === ENTITY_TYPES.LEADS) {
          // Leads use v1 API with PATCH method
          updateUrl = `https://${subdomain}.pipedrive.com/v1/${entityType}/${row.id}?api_token=${apiKey}`;
          method = 'PATCH';

          // Special handling for leads payload
          // Pipedrive is picky about leads fields format

          // Fix field naming issues - creator_id should be owner_id
          if (row.data.creator_id !== undefined) {
            row.data.owner_id = row.data.creator_id;
            delete row.data.creator_id;
            Logger.log(`Replaced creator_id with owner_id for lead`);
          }

          // Check for field format issues
          if (row.data.value && typeof row.data.value !== 'object') {
            // Value needs to be an object like { amount: 123, currency: "USD" }
            try {
              // Try to see if it's a number value that needs to be formatted
              const amount = parseFloat(row.data.value);
              if (!isNaN(amount)) {
                // Convert to the expected format with default currency
                row.data.value = {
                  amount: amount,
                  currency: "USD" // Default currency
                };
                Logger.log(`Formatted lead value as object: ${JSON.stringify(row.data.value)}`);
              } else {
                // If we can't parse it, remove it to avoid API errors
                delete row.data.value;
                Logger.log(`Removed invalid lead value field`);
              }
            } catch (e) {
              // If there's any error, remove the field
              delete row.data.value;
              Logger.log(`Removed problematic lead value field due to error: ${e.message}`);
            }
          }

          // Make sure person_id and organization_id are numbers, not strings
          if (row.data.person_id && typeof row.data.person_id === 'string') {
            row.data.person_id = parseInt(row.data.person_id, 10);
            if (isNaN(row.data.person_id)) {
              delete row.data.person_id;
              Logger.log(`Removed invalid person_id that couldn't be converted to a number`);
            }
          }

          if (row.data.organization_id && typeof row.data.organization_id === 'string') {
            row.data.organization_id = parseInt(row.data.organization_id, 10);
            if (isNaN(row.data.organization_id)) {
              delete row.data.organization_id;
              Logger.log(`Removed invalid organization_id that couldn't be converted to a number`);
            }
          }

          // Make sure label_ids is an array
          if (row.data.label_ids && !Array.isArray(row.data.label_ids)) {
            // If it's a string, try to convert to array
            if (typeof row.data.label_ids === 'string') {
              // If it looks like JSON array, parse it
              if (row.data.label_ids.startsWith('[') && row.data.label_ids.endsWith(']')) {
                try {
                  row.data.label_ids = JSON.parse(row.data.label_ids);
                  Logger.log(`Parsed label_ids from JSON string to array`);
                } catch (e) {
                  delete row.data.label_ids;
                  Logger.log(`Removed invalid label_ids JSON: ${e.message}`);
                }
              } else {
                // Try to convert comma-separated list
                row.data.label_ids = row.data.label_ids.split(',').map(id => parseInt(id.trim(), 10)).filter(id => !isNaN(id));

                if (row.data.label_ids.length === 0) {
                  delete row.data.label_ids;
                  Logger.log(`Removed empty label_ids array`);
                } else {
                  Logger.log(`Converted label_ids to array: ${JSON.stringify(row.data.label_ids)}`);
                }
              }
            } else {
              // Not a string, not an array - remove it
              delete row.data.label_ids;
              Logger.log(`Removed invalid label_ids that is neither an array nor a string`);
            }
          }
        } else if (entityType === ENTITY_TYPES.PRODUCTS) {
          // Products use API v2 with PATCH method but need special handling
          // Debug the row ID to understand its type and format
          Logger.log(`Product row ID: ${row.id} (type: ${typeof row.id})`);

          // If the ID is not numeric, it's likely not a valid product ID
          let productId = row.id;

          // Check if the product ID is numeric
          if (isNaN(Number(productId)) || productId === null || productId === undefined) {
            Logger.log(`Warning: Non-numeric product ID found: "${productId}". Trying to find proper ID.`);

            // Try to extract a numeric ID from the string if possible
            const numericMatch = productId.toString().match(/(\d+)/);
            if (numericMatch && numericMatch[1]) {
              productId = Number(numericMatch[1]);
              Logger.log(`Extracted numeric portion from ID: ${productId}`);
            } else {
              Logger.log(`Could not extract numeric ID, using original: ${productId}`);
            }
          }

          updateUrl = `https://${subdomain}.pipedrive.com/api/v2/products/${productId}?api_token=${apiKey}`;
          method = 'PATCH';

          // Special cleanup for products payload
          // Check for and remove any owner_id.name or other non-standard fields
          if (row.data['owner_id.name']) {
            Logger.log(`Removing owner_id.name field from product update payload`);
            delete row.data['owner_id.name'];
          }

          // Also check for 'Owner Name' field
          if (row.data['Owner Name']) {
            Logger.log(`Removing Owner Name field from product update payload`);
            delete row.data['Owner Name'];
          }
        } else {
          // Other main entities (deals, activities, organizations, persons) use v2 API with PATCH method
          updateUrl = `https://${subdomain}.pipedrive.com/api/v2/${entityType}/${row.id}?api_token=${apiKey}`;
          method = 'PATCH';
        }

        // Remove id field from updateData if it exists
        if (row.data.id) {
          Logger.log(`Removing id field from update data for entity ${row.id}`);
          delete row.data.id;
        }

        Logger.log(`API request to ${updateUrl} with method ${method}`);
        Logger.log(`Request payload: ${JSON.stringify(row.data)}`);

        // Send the update request to Pipedrive
        try {
          const response = UrlFetchApp.fetch(updateUrl, {
            method: method,
            contentType: 'application/json',
            payload: JSON.stringify(row.data),
            muteHttpExceptions: true
          });

          const responseCode = response.getResponseCode();
          const responseBody = JSON.parse(response.getContentText());

          if (responseCode === 200 && responseBody.success) {
            // Update was successful
            results.push({
              id: row.id,
              rowIndex: row.rowIndex,
              success: true
            });

            // Add to successful rows for batch update later
            successRowIndices.push(row.rowIndex);

            // Update the row status in the sheet
            try {
              // Convert to 1-based column index for getRange
              const statusColumn = trackingColumnIndex + 1;
              Logger.log(`Updating status to "Synced" at row ${row.rowIndex}, column ${statusColumn}`);

              // Set the value directly
              activeSheet.getRange(row.rowIndex, statusColumn).setValue('Synced');
            } catch (updateError) {
              Logger.log(`Error updating status cell: ${updateError.message}`);
            }
          } else {
            // Update failed
            const errorMessage = responseBody.error || 'Unknown error';
            results.push({
              id: row.id,
              rowIndex: row.rowIndex,
              success: false,
              error: errorMessage
            });

            // Mark the row as error in the sheet
            try {
              // Convert to 1-based column index for getRange
              const statusColumn = trackingColumnIndex + 1;
              activeSheet.getRange(row.rowIndex, statusColumn).setValue('Error');
            } catch (updateError) {
              Logger.log(`Error updating error status cell: ${updateError.message}`);
            }

            // Log the error
            Logger.log(`Error updating ${entityType} ${row.id}: ${errorMessage}`);
          }
        } catch (error) {
          Logger.log(`API request failed with error: ${error.message}`);
          // Mark the row as error in the sheet
          try {
            const statusColumn = trackingColumnIndex + 1;
            activeSheet.getRange(row.rowIndex, statusColumn).setValue('Error');
          } catch (updateError) {
            Logger.log(`Error updating error status cell: ${updateError.message}`);
          }
        }
      } catch (error) {
        // Handle API errors
        results.push({
          id: row.id,
          rowIndex: row.rowIndex,
          success: false,
          error: error.message
        });

        // Mark the row as error in the sheet
        try {
          // Convert to 1-based column index for getRange
          const statusColumn = trackingColumnIndex + 1;
          activeSheet.getRange(row.rowIndex, statusColumn).setValue('Error');
        } catch (updateError) {
          Logger.log(`Error updating error status cell: ${updateError.message}`);
        }

        // Log the error
        Logger.log(`Exception updating ${entityType} ${row.id}: ${error.message}`);
      }
    }

    // Perform a batch update for all successful rows to ensure they're marked as "Synced"
    if (successRowIndices.length > 0) {
      try {
        Logger.log(`Performing batch update for ${successRowIndices.length} successful rows`);

        // Create a range for each row and column combination
        const statusColumn = trackingColumnIndex + 1;
        const successRanges = successRowIndices.map(rowIndex =>
          activeSheet.getRange(rowIndex, statusColumn)
        );

        // Create array of values (all "Synced")
        const syncedValues = successRowIndices.map(() => ["Synced"]);

        // Update all cells in a batch
        for (let i = 0; i < successRanges.length; i++) {
          successRanges[i].setValue("Synced");
        }

        // Force flush to ensure updates are committed
        SpreadsheetApp.flush();

        Logger.log(`Batch update complete for ${successRowIndices.length} rows`);
      } catch (batchError) {
        Logger.log(`Error in batch status update: ${batchError.message}`);
      }
    }

    // Update last sync time
    const now = new Date();
    scriptProperties.setProperty(twoWaySyncLastSyncKey, now.toISOString());

    // Show results to the user (only for manual syncs)
    const successCount = results.filter(r => r.success).length;
    const errorCount = results.filter(r => !r.success).length;

    // Log results for scheduled syncs
    if (isScheduledSync) {
      if (errorCount === 0) {
        Logger.log(`Scheduled sync: Successfully updated ${successCount} ${entityType} in Pipedrive.`);
      } else {
        Logger.log(`Scheduled sync: Completed with some errors. Successful: ${successCount}, Failed: ${errorCount}`);
      }
    } else {
      // Show UI alerts for manual syncs
      const ui = SpreadsheetApp.getUi();
      if (errorCount === 0) {
        // All updates successful
        ui.alert(
          'Update Complete',
          `Successfully updated ${successCount} ${entityType} in Pipedrive.`,
          ui.ButtonSet.OK
        );
      } else {
        // Some updates failed
        let message = `Updates completed with some errors.\n\n`;
        message += `Successful: ${successCount}\n`;
        message += `Failed: ${errorCount}\n\n`;

        if (errorCount > 0) {
          message += `Errors have been marked in the sheet. Please check the rows marked as 'Error'.`;
        }

        ui.alert('Update Results', message, ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    // Handle any unexpected errors
    if (!isScheduledSync) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Error',
        `An error occurred: ${error.message}`,
        ui.ButtonSet.OK
      );
    }
    Logger.log(`Error in pushChangesToPipedrive: ${error.message}`);
  }
}

/**
 * Gets field mappings for a specific entity type
 * Maps column headers to Pipedrive API field keys
 * @param {string} entityType The entity type (deals, persons, etc.)
 * @return {Object} An object mapping column headers to API field keys
 */
function getFieldMappingsForEntity(entityType) {
  // Basic field mappings for each entity type
  const commonMappings = {
    'ID': 'id',
    'Name': 'name',
    'Owner': 'owner_id',
    'Organization': 'org_id',
    'Person': 'person_id',
    'Added': 'add_time',
    'Updated': 'update_time'
  };

  // Entity-specific mappings
  const entityMappings = {
    [ENTITY_TYPES.DEALS]: {
      'Value': 'value',
      'Currency': 'currency',
      'Title': 'title',
      'Pipeline': 'pipeline_id',
      'Stage': 'stage_id',
      'Status': 'status',
      'Expected Close Date': 'expected_close_date'
    },
    [ENTITY_TYPES.PERSONS]: {
      'Email': 'email',
      'Phone': 'phone',
      'First Name': 'first_name',
      'Last Name': 'last_name',
      'Organization': 'org_id'
    },
    [ENTITY_TYPES.ORGANIZATIONS]: {
      'Address': 'address',
      'Website': 'web'
    },
    [ENTITY_TYPES.ACTIVITIES]: {
      'Type': 'type',
      'Due Date': 'due_date',
      'Due Time': 'due_time',
      'Duration': 'duration',
      'Deal': 'deal_id',
      'Person': 'person_id',
      'Organization': 'org_id',
      'Note': 'note'
    },
    [ENTITY_TYPES.PRODUCTS]: {
      'Code': 'code',
      'Description': 'description',
      'Unit': 'unit',
      'Tax': 'tax',
      'Category': 'category',
      'Active': 'active_flag',
      'Selectable': 'selectable',
      'Visible To': 'visible_to',
      'First Price': 'first_price',
      'Cost': 'cost',
      'Prices': 'prices',
      'Owner Name': 'owner_id.name'  // Map "Owner Name" to owner_id.name so we can detect this field
    }
  };

  // Combine common mappings with entity-specific mappings
  return { ...commonMappings, ...(entityMappings[entityType] || {}) };
}

/**
 * Refreshes the styling of the Sync Status column
 * This is useful if the styling gets lost or if the user wants to reset it
 */
function refreshSyncStatusStyling() {
  try {
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();

    cleanupPreviousSyncStatusColumn(sheet, sheetName);

    // Check if two-way sync is enabled for this sheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${sheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;

    const twoWaySyncEnabled = scriptProperties.getProperty(twoWaySyncEnabledKey) === 'true';

    if (!twoWaySyncEnabled) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Two-way sync is not enabled for this sheet. Please enable it first in "Two-Way Sync Settings".',
        'Cannot Refresh Styling',
        5
      );
      return;
    }

    // Get the tracking column
    let trackingColumn = scriptProperties.getProperty(twoWaySyncTrackingColumnKey) || '';
    let trackingColumnIndex;

    if (trackingColumn) {
      // Convert column letter to index (0-based)
      trackingColumnIndex = columnLetterToIndex(trackingColumn);
    } else {
      // Try to find the Sync Status column by header name
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      trackingColumnIndex = headers.findIndex(header =>
        header && header.toString().toLowerCase().includes('sync status')
      );

      if (trackingColumnIndex === -1) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Could not find the Sync Status column. Please set up two-way sync again.',
          'Column Not Found',
          5
        );
        return;
      }
    }

    // Convert to 1-based index for getRange
    const columnPos = trackingColumnIndex + 1;

    // Style the header cell
    const headerCell = sheet.getRange(1, columnPos);
    headerCell.setBackground('#E8F0FE') // Light blue background
      .setFontWeight('bold')
      .setNote('This column tracks changes for two-way sync with Pipedrive');

    // Add only BORDERS to the entire status column (not background)
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const fullStatusColumn = sheet.getRange(1, columnPos, lastRow, 1);
    fullStatusColumn.setBorder(null, true, null, true, false, false, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

    // Add data validation for all cells in the status column (except header)
    if (lastRow > 1) {
      // Get all values from the first column to identify timestamps/separators
      const firstColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

      // Collect row indices of actual data rows (not timestamps/separators)
      const dataRowIndices = [];
      for (let i = 0; i < firstColumnValues.length; i++) {
        const value = firstColumnValues[i][0];
        const rowIndex = i + 2; // +2 because we start at row 2

        // Skip empty rows and rows that look like timestamps
        if (!value || (typeof value === 'string' &&
          (value.includes('Timestamp') || value.includes('Last synced')))) {
          continue;
        }

        dataRowIndices.push(rowIndex);
      }

      // Apply background color only to data rows
      dataRowIndices.forEach(rowIndex => {
        sheet.getRange(rowIndex, columnPos).setBackground('#F8F9FA'); // Light gray background
      });

      // Apply data validation only to actual data rows
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Not modified', 'Modified', 'Synced', 'Error'], true)
        .build();

      // Apply validation to each data row individually
      dataRowIndices.forEach(rowIndex => {
        sheet.getRange(rowIndex, columnPos).setDataValidation(rule);
      });

      // Clear and recreate conditional formatting
      // Get all existing conditional formatting rules
      const rules = sheet.getConditionalFormatRules();
      // Clear any existing rules for the tracking column
      const newRules = rules.filter(rule => {
        const ranges = rule.getRanges();
        return !ranges.some(range => range.getColumn() === columnPos);
      });

      // Create conditional format for "Modified" status
      const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(modifiedRule);

      // Create conditional format for "Synced" status
      const syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA')  // Light green background
        .setFontColor('#137333')   // Green text
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(syncedRule);

      // Create conditional format for "Error" status
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Error')
        .setBackground('#FCE8E6')  // Light red background
        .setFontColor('#D93025')   // Red text
        .setBold(true)             // Bold text for errors
        .setRanges([sheet.getRange(2, columnPos, lastRow - 1, 1)])
        .build();
      newRules.push(errorRule);

      // Apply all rules
      sheet.setConditionalFormatRules(newRules);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Sync Status column styling has been refreshed successfully.',
      'Styling Updated',
      5
    );
  } catch (error) {
    Logger.log(`Error refreshing sync status styling: ${error.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error refreshing styling: ${error.message}`,
      'Error',
      5
    );
  }
}

/**
 * Determines if a user is part of a team using the most reliable method
 * @param {string} email The user's email
 * @return {boolean} Whether the user is in a team
 */
function isUserInTeam(email) {
  try {
    if (!email) return false;
    
    // First check email-to-team map for quick lookup
    const docProps = PropertiesService.getDocumentProperties();
    const emailToTeamMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
    
    if (emailToTeamMapStr) {
      const emailToTeamMap = JSON.parse(emailToTeamMapStr);
      // Fix 1: Use lowercase for case-insensitive comparison
      if (emailToTeamMap[email.toLowerCase()]) {
        Logger.log(`User ${email} found in email map`);
        return true;
      }
    }
    
    // Fall back to checking teams data directly
    const teamsData = getTeamsData();
    for (const teamId in teamsData) {
      // Fix 2: Replace includes() with case-insensitive loop
      const memberEmails = teamsData[teamId].memberEmails || [];
      for (let i = 0; i < memberEmails.length; i++) {
        if (memberEmails[i].toLowerCase() === email.toLowerCase()) {
          // Update the map as it seems to be out of date
          updateEmailToTeamMap();
          Logger.log(`User ${email} found in team ${teamId} directly`);
          return true;
        }
      }
    }
    
    return false;
  } catch (e) {
    Logger.log(`Error checking if user is in team: ${e.message}`);
    return false;
  }
}
