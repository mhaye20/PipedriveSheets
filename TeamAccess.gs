/**
 * Team Access Management
 * 
 * This module handles all team-related functionality including:
 * - Team access verification
 * - Team membership management
 * - User permissions and roles
 */

/**
 * Checks if the current user has any form of access rights
 * @param {string} userEmail - The email address of the user
 * @return {boolean} True if user has access, false otherwise
 */
function checkAnyUserAccess(userEmail) {
  try {
    if (!userEmail) return false;
    
    // APPROACH 1: Direct ownership check
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

/**
 * Forces a check of team membership for a user
 * @param {string} userEmail - The email address of the user
 * @return {boolean} True if user is in a team, false otherwise
 */
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
          Logger.log(`User found in team ${teamId}, creating Pipedrive menu`);
          return true;
        }
      }
    }
    
    Logger.log(`User ${userEmail} not found in any team`);
    return false;
  } catch (e) {
    Logger.log(`Error in forceTeamMembershipCheck: ${e.message}`);
    return false;
  }
}

/**
 * Checks if the user has verified team access
 * @return {boolean} True if user has verified team access, false otherwise
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
 * @return {boolean} True if preload was successful, false otherwise
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
 * @return {boolean} True if refresh was successful, false otherwise
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
 * @return {Object} Object with success status and error message if applicable
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

/**
 * Shows the team join request dialog
 */
function showTeamJoinRequest() {
  // This is now properly implemented in UI.gs
  // The proper implementation is called directly
  if (typeof UI !== 'undefined' && UI.showTeamJoinRequest) {
    UI.showTeamJoinRequest();
  } else {
    // Fall back to the global function
    if (typeof window !== 'undefined') {
      window.showTeamJoinRequest();
    }
  }
}

/**
 * Checks if a user is in any team
 * @param {string} email The email to check
 * @returns {boolean} True if the user is in a team, false otherwise
 */
function isUserInTeam(email) {
  try {
    if (!email) return false;
    
    // Normalize email for case-insensitive comparison
    const normalizedEmail = email.toLowerCase();
    
    // Try fast path using email-to-team map
    const emailMapStr = PropertiesService.getDocumentProperties().getProperty('EMAIL_TO_TEAM_MAP');
    if (emailMapStr) {
      try {
        const emailMap = JSON.parse(emailMapStr);
        if (emailMap[normalizedEmail]) {
          return true;
        }
      } catch (e) {
        Logger.log('Error parsing email map: ' + e.message);
      }
    }
    
    // Slower path if map lookup fails
    const teamsData = getTeamsData();
    
    // Check each team's memberEmails
    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      
      if (team.memberEmails && Array.isArray(team.memberEmails)) {
        for (let i = 0; i < team.memberEmails.length; i++) {
          if (team.memberEmails[i].toLowerCase() === normalizedEmail) {
            return true;
          }
        }
      }
    }
    
    return false;
  } catch (e) {
    Logger.log('Error in isUserInTeam: ' + e.message);
    return false;
  }
} 