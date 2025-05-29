/**
 * Team Data Management
 * 
 * This module handles the team data storage and retrieval:
 * - Team creation, joining, and leaving
 * - Team members management
 * - Team preferences and settings
 * - Team activity logging
 */

/**
 * Gets all teams data from document properties
 * @return {Object} Teams data object or empty object if none exists
 */
function getTeamsData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const teamsDataStr = docProps.getProperty('TEAMS_DATA');
    
    if (!teamsDataStr) {
      return {};
    }
    
    return JSON.parse(teamsDataStr);
  } catch (e) {
    Logger.log(`Error in getTeamsData: ${e.message}`);
    return {};
  }
}

/**
 * Gets the user's team data
 * @param {string} userEmail - Email of the user
 * @param {Object} teamsData - Optional teams data (to avoid redundant calls)
 * @return {Object} Team data or null if user is not in a team
 */
function getUserTeam(userEmail, teamsData = null) {
  try {
    if (!userEmail) return null;
    
    // Normalize email to lowercase
    const normalizedEmail = userEmail.toLowerCase();
    
    // Get teams data if not provided
    if (!teamsData) {
      teamsData = getTeamsData();
    }
    
    // Fast path check using the email-to-team map
    const docProps = PropertiesService.getDocumentProperties();
    const emailMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
    
    if (emailMapStr) {
      try {
        const emailMap = JSON.parse(emailMapStr);
        const mappedTeamId = emailMap[normalizedEmail];
        
        if (mappedTeamId && teamsData[mappedTeamId]) {
          Logger.log(`User ${userEmail} found in team map for team ${mappedTeamId}`);
          
          // Get the team data
          const team = teamsData[mappedTeamId];
          
          // Determine if user is admin (legacy format first)
          const isAdmin = team.adminEmails && 
                         team.adminEmails.includes(normalizedEmail);
          
          return {
            teamId: mappedTeamId,
            name: team.name || 'Unnamed Team',
            role: isAdmin ? 'Admin' : 'Member',
            adminEmails: team.adminEmails || [],
            memberEmails: team.memberEmails || []
          };
        }
      } catch (e) {
        Logger.log(`Error parsing email map: ${e.message}`);
        // Continue to slower path check
      }
    }
    
    // Slower path - check each team
    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      const memberEmails = team.memberEmails || [];
      
      // Check if user's email is in this team (case insensitive)
      for (let i = 0; i < memberEmails.length; i++) {
        if (memberEmails[i].toLowerCase() === normalizedEmail) {
          // Determine if user is admin
          const isAdmin = team.adminEmails && 
                         team.adminEmails.some(email => email.toLowerCase() === normalizedEmail);
          
          Logger.log(`User ${userEmail} found in team ${teamId}`);
          return {
            teamId: teamId,
            name: team.name || 'Unnamed Team',
            role: isAdmin ? 'Admin' : 'Member',
            adminEmails: team.adminEmails || [],
            memberEmails: team.memberEmails || []
          };
        }
      }
      
      // Check the members object format (for newer implementations)
      if (team.members && team.members[normalizedEmail]) {
        const role = team.members[normalizedEmail];
        
        return {
          teamId: teamId,
          name: team.name || 'Unnamed Team',
          role: role,
          adminEmails: team.adminEmails || [],
          memberEmails: team.memberEmails || []
        };
      }
    }
    
    // User not found in any team
    Logger.log(`User ${userEmail} not found in any team`);
    return null;
  } catch (e) {
    Logger.log(`Error in getUserTeam: ${e.message}`);
    return null;
  }
}

/**
 * Saves teams data to document properties
 * @param {Object} teamsData - Teams data object to save
 * @return {boolean} True if successful, false otherwise
 */
function saveTeamsData(teamsData) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    docProps.setProperty('TEAMS_DATA', JSON.stringify(teamsData));
    return true;
  } catch (e) {
    Logger.log(`Error in saveTeamsData: ${e.message}`);
    return false;
  }
}

/**
 * Checks if the current user is the script owner/installer
 * @return {boolean} True if the user is the script owner/installer, false otherwise
 */
function isScriptOwner() {
  try {
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    return authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED;
  } catch (e) {
    Logger.log('Error checking script owner: ' + e.message);
    return false;
  }
}

/**
 * Creates a new team with the specified name
 * @param {string} teamName The name of the new team
 * @returns {Object} Result object with success/error information and team data
 */
function createTeam(teamName) {
  try {
    if (!teamName || teamName.trim() === '') {
      return { success: false, message: 'Team name is required.' };
    }

    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Could not determine user email. Please make sure you are logged in.' };
    }
    
    // Only allow script owners to create teams
    if (!isScriptOwner()) {
      return { 
        success: false, 
        message: 'Only the user who installed the extension can create teams. Please contact your administrator to get a team ID to join.' 
      };
    }

    // Check if user is already in a team
    const existingTeam = getUserTeam(userEmail);
    if (existingTeam) {
      return { success: false, message: 'You are already a member of a team. Please leave your current team before creating a new one.' };
    }

    // Generate a team ID
    const teamId = Utilities.getUuid();

    // Get or initialize teams data
    const teamsData = getTeamsData() || {};

    // Create the new team entry
    teamsData[teamId] = {
      name: teamName,
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
      
      // Log team creation activity
      logTeamActivity(teamId, 'team_created', userEmail, {
        teamName: teamName
      });
      
      // Store verification status in user properties for future sessions
      try {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('VERIFIED_TEAM_MEMBER', 'true');
        Logger.log('Set VERIFIED_TEAM_MEMBER flag for team creator');
      } catch (e) {
        Logger.log('Error setting verified team member flag: ' + e.message);
      }
      
      // Try to refresh the menu
      try {
        fixMenuAfterJoin();
      } catch (menuError) {
        Logger.log('Error refreshing menu: ' + menuError.message);
      }

      // Return success with team information
      return {
        success: true,
        teamId: teamId,
        team: teamsData[teamId]
      };
    } else {
      return { success: false, message: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error creating team: ' + e.message);
    return { success: false, message: 'An error occurred: ' + e.message };
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
      return { success: false, message: 'Could not determine user email. Please make sure you are logged in.' };
    }

    // Get existing teams data from properties
    const teamsData = getTeamsData();

    // Check if user is already in a team in the main data store
    const existingTeam = getUserTeam(userEmail, teamsData);
    if (existingTeam) {
      return { success: false, message: 'You are already a member of a team. Please leave your current team before joining another one.' };
    }

    // Check if the team exists
    if (!teamsData[teamId]) {
      return { success: false, message: 'Team not found. Please check the team ID and try again.' };
    }

    // Define maximum team members
    const MAX_TEAM_MEMBERS = 5; // Team plan allows up to 5 users total (admin + 4 members)
    
    // Check if the team is at maximum capacity
    if (teamsData[teamId].memberEmails && teamsData[teamId].memberEmails.length >= MAX_TEAM_MEMBERS) {
      return { success: false, message: `This team has reached the maximum of ${MAX_TEAM_MEMBERS} members allowed in the team plan.` };
    }

    // Initialize member arrays if they don't exist
    if (!teamsData[teamId].memberEmails) {
      teamsData[teamId].memberEmails = [];
    }
    
    // Add the user to the team
    teamsData[teamId].memberEmails.push(userEmail);

    // Save the updated teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      
      // Log member joined activity
      logTeamActivity(teamId, 'member_joined', userEmail, {
        teamName: teamsData[teamId].name
      });
      
      // Store verification status in user properties for future sessions
      try {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('VERIFIED_TEAM_MEMBER', 'true');
        Logger.log('Set VERIFIED_TEAM_MEMBER flag for user');
      } catch (e) {
        Logger.log('Error setting verified team member flag: ' + e.message);
      }
      
      // Try to refresh the menu
      try {
        fixMenuAfterJoin();
      } catch (menuError) {
        Logger.log('Error refreshing menu: ' + menuError.message);
      }
      
      return { success: true };
    } else {
      return { success: false, message: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log('Error joining team: ' + e.message);
    return { success: false, message: 'An error occurred: ' + e.message };
  }
}

/**
 * Updates the email-to-team mapping for quick access lookups
 * @returns {boolean} True if successful, false otherwise
 */
function updateEmailToTeamMap() {
  try {
    const teamsData = getTeamsData();
    if (!teamsData) return false;
    
    // Create a mapping of email addresses to team IDs
    const emailMap = {};
    
    // Process each team
    for (const teamId in teamsData) {
      const team = teamsData[teamId];
      
      // Process member emails
      if (team.memberEmails && Array.isArray(team.memberEmails)) {
        team.memberEmails.forEach(function(email) {
          // Always use lowercase for consistent lookups
          emailMap[email.toLowerCase()] = teamId;
        });
      }
    }
    
    // Save the map to document properties
    PropertiesService.getDocumentProperties().setProperty('EMAIL_TO_TEAM_MAP', JSON.stringify(emailMap));
    Logger.log('Email-to-team map updated successfully');
    
    return true;
  } catch (e) {
    Logger.log('Error updating email map: ' + e.message);
    return false;
  }
}

/**
 * Leaves the current team
 * @return {Object} Result object with success status
 */
function leaveTeam() {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email address. Please ensure you are signed in.' };
    }
    
    // Get existing teams data
    const teamsData = getTeamsData();
    
    // Check if user is in a team
    const userTeam = getUserTeam(userEmail, teamsData);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    const teamId = userTeam.teamId;
    const team = teamsData[teamId];
    
    // Check if user is an admin
    const isAdmin = userTeam.role === 'Admin';
    
    // Using the new members structure
    if (team.members) {
      // Check if user is the last admin
      let adminCount = 0;
      for (const email in team.members) {
        if (team.members[email] === 'Admin') {
          adminCount++;
        }
      }
      
      // Get member count
      const memberCount = Object.keys(team.members).length;
      
      if (isAdmin && adminCount === 1 && memberCount > 1) {
        // User is the last admin and there are other members
        return { 
          success: false, 
          message: 'You are the last admin of this team. Please promote another member to admin before leaving.' 
        };
      } else if (memberCount === 1) {
        // User is the last member, delete the team
        delete teamsData[teamId];
      } else {
        // Remove user from team
        delete team.members[userEmail];
      }
    } 
    // Fallback for legacy structure
    else if (team.memberEmails) {
      // Check if user is the last admin
      const isLegacyAdmin = team.adminEmails && team.adminEmails.some(email => email.toLowerCase() === userEmail.toLowerCase());
      const otherAdmins = isLegacyAdmin && team.adminEmails.filter(email => email.toLowerCase() !== userEmail.toLowerCase());
      
      if (isLegacyAdmin && (!otherAdmins || otherAdmins.length === 0)) {
        // If user is the last admin, check if there are other members
        const otherMembers = team.memberEmails.filter(email => email.toLowerCase() !== userEmail.toLowerCase());
        
        if (otherMembers.length > 0) {
          return { 
            success: false, 
            message: 'You are the last admin of this team. Please promote another member to admin before leaving.' 
          };
        } else {
          // If user is the last member, delete the team
          delete teamsData[teamId];
        }
      } else {
        // Remove user from team
        team.memberEmails = team.memberEmails.filter(email => email.toLowerCase() !== userEmail.toLowerCase());
        
        // Also remove from admin list if applicable
        if (isLegacyAdmin) {
          team.adminEmails = team.adminEmails.filter(email => email.toLowerCase() !== userEmail.toLowerCase());
        }
      }
    }
    
    // Log the activity before saving (in case team gets deleted)
    const wasTeamDeleted = !teamsData[teamId]; // Check if team was deleted
    if (!wasTeamDeleted) {
      logTeamActivity(teamId, 'member_left', userEmail, {
        teamName: team.name
      });
    }
    
    // Save teams data
    saveTeamsData(teamsData);
    
    // Update the email-to-team map
    updateEmailToTeamMap();
    
    // Clear the verification flag since user is no longer in a team
    try {
      const userProperties = PropertiesService.getUserProperties();
      userProperties.deleteProperty('VERIFIED_TEAM_MEMBER');
      Logger.log('Cleared VERIFIED_TEAM_MEMBER flag after leaving team');
    } catch (e) {
      Logger.log('Error clearing verified team member flag: ' + e.message);
    }
    
    return { success: true, message: 'You have successfully left the team.' };
  } catch (e) {
    Logger.log(`Error in leaveTeam: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Adds a new member to the team
 * @param {string} email The email of the member to add
 * @returns {Object} Result object with success status
 */
function addTeamMember(email) {
  try {
    if (!email || !email.includes('@')) {
      return { success: false, message: 'Please enter a valid email address.' };
    }
    
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Check if user is admin
    if (userTeam.role !== 'Admin') {
      return { success: false, message: 'Only team admins can add members.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    
    // Check if member is already in the team
    if (teamsData[teamId].memberEmails.includes(email)) {
      return { success: false, message: 'This person is already a member of the team.' };
    }
    
    // Define maximum team members
    const MAX_TEAM_MEMBERS = 5; // Team plan allows up to 5 users total (admin + 4 members)
    
    // Check if team is at capacity
    if (teamsData[teamId].memberEmails.length >= MAX_TEAM_MEMBERS) {
      return { success: false, message: `Team has reached the maximum of ${MAX_TEAM_MEMBERS} members allowed in the team plan.` };
    }
    
    // Add the member
    teamsData[teamId].memberEmails.push(email);
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      
      // Log member added activity
      logTeamActivity(teamId, 'member_added', userEmail, {
        targetEmail: email
      });
      
      return { success: true, message: 'Member added successfully.' };
    } else {
      return { success: false, message: 'Failed to save team data.' };
    }
  } catch (e) {
    Logger.log(`Error in addTeamMember: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Deletes a team (admin only)
 * @returns {Object} Result object with success status
 */
function deleteTeam() {
  try {
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Check if user is admin
    if (userTeam.role !== 'Admin') {
      return { success: false, message: 'Only team admins can delete a team.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    
    // Delete the team
    delete teamsData[teamId];
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      return { success: true, message: 'Team deleted successfully.' };
    } else {
      return { success: false, message: 'Failed to delete team.' };
    }
  } catch (e) {
    Logger.log(`Error in deleteTeam: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Promotes a team member to admin
 * @param {string} email The email of the member to promote
 * @returns {Object} Result object with success status
 */
function promoteTeamMember(email) {
  try {
    if (!email) {
      return { success: false, message: 'Invalid email address.' };
    }
    
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Check if user is admin
    if (userTeam.role !== 'Admin') {
      return { success: false, message: 'Only team admins can promote members.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    const team = teamsData[teamId];
    
    // Check if member is in the team
    if (!team.memberEmails.includes(email)) {
      return { success: false, message: 'This person is not a member of the team.' };
    }
    
    // Check if member is already an admin
    if (team.adminEmails && team.adminEmails.includes(email)) {
      return { success: false, message: 'This person is already an admin.' };
    }
    
    // Initialize adminEmails array if needed
    if (!team.adminEmails) {
      team.adminEmails = [];
    }
    
    // Add to admin list
    team.adminEmails.push(email);
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Log member promoted activity
      logTeamActivity(teamId, 'member_promoted', userEmail, {
        targetEmail: email
      });
      
      return { success: true, message: 'Member promoted to admin successfully.' };
    } else {
      return { success: false, message: 'Failed to update team data.' };
    }
  } catch (e) {
    Logger.log(`Error in promoteTeamMember: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Demotes a team admin to regular member
 * @param {string} email The email of the admin to demote
 * @returns {Object} Result object with success status
 */
function demoteTeamMember(email) {
  try {
    if (!email) {
      return { success: false, message: 'Invalid email address.' };
    }
    
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Check if user is admin
    if (userTeam.role !== 'Admin') {
      return { success: false, message: 'Only team admins can demote members.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    const team = teamsData[teamId];
    
    // Check if member is an admin
    if (!team.adminEmails || !team.adminEmails.includes(email)) {
      return { success: false, message: 'This person is not an admin.' };
    }
    
    // Check if this is the last admin
    if (team.adminEmails.length === 1) {
      return { success: false, message: 'Cannot demote the last admin of the team. Promote another member first.' };
    }
    
    // Check if the target user is the script owner
    // We'll use the team creator field as a proxy for script owner since the script owner is typically the team creator
    if (team.createdBy && team.createdBy.toLowerCase() === email.toLowerCase()) {
      return { 
        success: false, 
        message: 'Cannot demote the script owner/installer who created this team. This is a protected admin account.' 
      };
    }
    
    // Get the script owner status of the target user
    try {
      // If this is an attempt to demote another admin who's the script owner
      let targetIsScriptOwner = false;
      if (email.toLowerCase() !== userEmail.toLowerCase()) { // Not trying to demote self
        // This call would need to run as the target user, which isn't possible
        // However, we'll use the team creator field as a strong proxy
        targetIsScriptOwner = (team.createdBy && team.createdBy.toLowerCase() === email.toLowerCase());
      }
      
      if (targetIsScriptOwner) {
        return { 
          success: false, 
          message: 'Cannot demote the script owner/installer. This is a protected admin account.' 
        };
      }
    } catch (e) {
      Logger.log(`Error checking script owner status for ${email}: ${e.message}`);
      // Continue with the demotion since we couldn't verify
    }
    
    // Remove from admin list
    team.adminEmails = team.adminEmails.filter(e => e !== email);
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Log member demoted activity
      logTeamActivity(teamId, 'member_demoted', userEmail, {
        targetEmail: email
      });
      
      return { success: true, message: 'Admin demoted to member successfully.' };
    } else {
      return { success: false, message: 'Failed to update team data.' };
    }
  } catch (e) {
    Logger.log(`Error in demoteTeamMember: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Removes a member from the team
 * @param {string} email The email of the member to remove
 * @returns {Object} Result object with success status
 */
function removeTeamMember(email) {
  try {
    if (!email) {
      return { success: false, message: 'Invalid email address.' };
    }
    
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Check if user is admin
    if (userTeam.role !== 'Admin') {
      return { success: false, message: 'Only team admins can remove members.' };
    }
    
    // Prevent removing yourself
    if (email === userEmail) {
      return { success: false, message: 'You cannot remove yourself. Use the Leave Team option instead.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    const team = teamsData[teamId];
    
    // Check if trying to remove the script owner/team creator
    if (team.createdBy && team.createdBy.toLowerCase() === email.toLowerCase()) {
      return { 
        success: false, 
        message: 'Cannot remove the script owner/installer who created this team. This is a protected account.' 
      };
    }
    
    // Check if member is in the team
    if (!team.memberEmails.includes(email)) {
      return { success: false, message: 'This person is not a member of the team.' };
    }
    
    // Remove from member list
    team.memberEmails = team.memberEmails.filter(e => e !== email);
    
    // Also remove from admin list if needed
    if (team.adminEmails && team.adminEmails.includes(email)) {
      team.adminEmails = team.adminEmails.filter(e => e !== email);
    }
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Update email map
      updateEmailToTeamMap();
      
      // Log member removed activity
      logTeamActivity(teamId, 'member_removed', userEmail, {
        targetEmail: email
      });
      
      return { success: true, message: 'Member removed successfully.' };
    } else {
      return { success: false, message: 'Failed to update team data.' };
    }
  } catch (e) {
    Logger.log(`Error in removeTeamMember: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Renames a team (script owner/installer only)
 * @param {string} newName The new name for the team
 * @returns {Object} Result object with success status
 */
function renameTeam(newName) {
  try {
    if (!newName || newName.trim() === '') {
      return { success: false, message: 'Team name is required.' };
    }
    
    // Get current user
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, message: 'Unable to determine your email. Please make sure you are logged in.' };
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return { success: false, message: 'You are not a member of any team.' };
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    const teamId = userTeam.teamId;
    const team = teamsData[teamId];
    
    // Check if user is the script owner/team creator
    if (!team.createdBy || team.createdBy.toLowerCase() !== userEmail.toLowerCase()) {
      return { 
        success: false, 
        message: 'Only the script owner/installer who created this team can rename it.' 
      };
    }
    
    // Update the team name
    team.name = newName.trim();
    
    // Save the teams data
    if (saveTeamsData(teamsData)) {
      // Log team renamed activity
      logTeamActivity(teamId, 'team_renamed', userEmail, {
        newName: newName.trim(),
        oldName: team.name // Note: this will be the new name since we already updated it
      });
      
      return { success: true, message: 'Team renamed successfully.' };
    } else {
      return { success: false, message: 'Failed to update team data.' };
    }
  } catch (e) {
    Logger.log(`Error in renameTeam: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Activity Logging System
 */

/**
 * Logs a team activity
 * @param {string} teamId - Team ID
 * @param {string} action - Action performed (join, leave, add_member, etc.)
 * @param {string} actorEmail - Email of user who performed the action
 * @param {Object} details - Additional details about the action
 */
function logTeamActivity(teamId, action, actorEmail, details = {}) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const activityKey = `TEAM_ACTIVITY_${teamId}`;
    
    // Get existing activities
    let activities = [];
    const existingActivities = docProps.getProperty(activityKey);
    if (existingActivities) {
      activities = JSON.parse(existingActivities);
    }
    
    // Create new activity entry
    const activity = {
      id: Utilities.getUuid(),
      timestamp: new Date().toISOString(),
      action: action,
      actor: actorEmail,
      details: details
    };
    
    // Add to the beginning of the array (most recent first)
    activities.unshift(activity);
    
    // Keep only the last 50 activities to prevent data bloat
    if (activities.length > 50) {
      activities = activities.slice(0, 50);
    }
    
    // Save back to properties
    docProps.setProperty(activityKey, JSON.stringify(activities));
    
    Logger.log(`Logged activity for team ${teamId}: ${action} by ${actorEmail}`);
  } catch (e) {
    Logger.log(`Error logging team activity: ${e.message}`);
  }
}

/**
 * Gets recent team activities
 * @param {string} teamId - Team ID
 * @param {number} limit - Maximum number of activities to return (default: 20)
 * @return {Array} Array of activity objects
 */
function getTeamActivities(teamId, limit = 20) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const activityKey = `TEAM_ACTIVITY_${teamId}`;
    
    const existingActivities = docProps.getProperty(activityKey);
    if (!existingActivities) {
      return [];
    }
    
    const activities = JSON.parse(existingActivities);
    return activities.slice(0, limit);
  } catch (e) {
    Logger.log(`Error getting team activities: ${e.message}`);
    return [];
  }
}

/**
 * Formats an activity for display
 * @param {Object} activity - Activity object
 * @return {Object} Formatted activity with description and formatted time
 */
function formatTeamActivity(activity) {
  const timeAgo = getTimeAgo(activity.timestamp);
  const actor = activity.actor;
  
  let description = '';
  
  switch (activity.action) {
    case 'team_created':
      description = `${actor} created the team`;
      break;
    case 'member_joined':
      description = `${actor} joined the team`;
      break;
    case 'member_added':
      description = `${actor} added ${activity.details.targetEmail} to the team`;
      break;
    case 'member_left':
      description = `${actor} left the team`;
      break;
    case 'member_removed':
      description = `${actor} removed ${activity.details.targetEmail} from the team`;
      break;
    case 'member_promoted':
      description = `${actor} promoted ${activity.details.targetEmail} to admin`;
      break;
    case 'member_demoted':
      description = `${actor} demoted ${activity.details.targetEmail} to member`;
      break;
    case 'team_renamed':
      description = `${actor} renamed the team to "${activity.details.newName}"`;
      break;
    case 'settings_changed':
      description = `${actor} modified ${activity.details.setting} settings`;
      break;
    case 'sync_performed':
      const entityType = activity.details.entityType || 'data';
      const sheetName = activity.details.sheetName || 'sheet';
      const itemCount = activity.details.itemCount || 0;
      description = `${actor} synced ${itemCount} ${entityType.toLowerCase()} to "${sheetName}"`;
      break;
    case 'columns_updated':
      const columnEntityType = activity.details.entityType || 'data';
      const columnSheetName = activity.details.sheetName || 'sheet';
      const columnCount = activity.details.columnCount || 0;
      description = `${actor} updated column preferences for ${columnEntityType.toLowerCase()} (${columnCount} columns) in "${columnSheetName}"`;
      break;
    case 'filter_changed':
      description = `${actor} changed filter to ${activity.details.filterName}`;
      break;
    case 'trigger_created':
      description = `${actor} created a ${activity.details.frequency} sync trigger`;
      break;
    case 'trigger_deleted':
      description = `${actor} deleted a sync trigger`;
      break;
    case 'twoway_sync_enabled':
      description = `${actor} enabled two-way sync`;
      break;
    case 'twoway_sync_disabled':
      description = `${actor} disabled two-way sync`;
      break;
    default:
      description = `${actor} performed ${activity.action}`;
  }
  
  return {
    ...activity,
    description: description,
    timeAgo: timeAgo
  };
}

/**
 * Helper function to get time ago string
 * @param {string} timestamp - ISO timestamp
 * @return {string} Human readable time ago
 */
function getTimeAgo(timestamp) {
  const now = new Date();
  const then = new Date(timestamp);
  const diffMs = now.getTime() - then.getTime();
  
  const diffSeconds = Math.floor(diffMs / 1000);
  const diffMinutes = Math.floor(diffSeconds / 60);
  const diffHours = Math.floor(diffMinutes / 60);
  const diffDays = Math.floor(diffHours / 24);
  
  if (diffSeconds < 60) {
    return 'just now';
  } else if (diffMinutes < 60) {
    return `${diffMinutes} minute${diffMinutes > 1 ? 's' : ''} ago`;
  } else if (diffHours < 24) {
    return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
  } else if (diffDays < 7) {
    return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
  } else {
    return then.toLocaleDateString();
  }
}

/**
 * Gets formatted team activities for the current user's team
 * @return {Array} Array of formatted activity objects for display
 */
function getFormattedTeamActivities() {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to determine user email');
    }
    
    // Get user's team
    const userTeam = getUserTeam(userEmail);
    if (!userTeam) {
      return []; // User is not in a team
    }
    
    // Get raw activities
    const rawActivities = getTeamActivities(userTeam.teamId, 20);
    
    // Format activities for display
    const formattedActivities = rawActivities.map(activity => formatTeamActivity(activity));
    
    return formattedActivities;
  } catch (e) {
    Logger.log(`Error in getFormattedTeamActivities: ${e.message}`);
    throw new Error(`Failed to load team activities: ${e.message}`);
  }
}

// Explicitly export these functions to ensure they're globally accessible
this.addTeamMember = addTeamMember;
this.deleteTeam = deleteTeam;
this.promoteTeamMember = promoteTeamMember;
this.demoteTeamMember = demoteTeamMember;
this.removeTeamMember = removeTeamMember;
this.joinTeam = joinTeam;
this.renameTeam = renameTeam;
this.logTeamActivity = logTeamActivity;
this.getTeamActivities = getTeamActivities;
this.formatTeamActivity = formatTeamActivity;
this.getFormattedTeamActivities = getFormattedTeamActivities; 