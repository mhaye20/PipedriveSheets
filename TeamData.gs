/**
 * Team Data Management
 * 
 * This module handles the team data storage and retrieval:
 * - Team creation, joining, and leaving
 * - Team members management
 * - Team preferences and settings
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
    const MAX_TEAM_MEMBERS = 50;
    
    // Check if the team is at maximum capacity
    if (teamsData[teamId].memberEmails && teamsData[teamId].memberEmails.length >= MAX_TEAM_MEMBERS) {
      return { success: false, message: `This team has reached the maximum of ${MAX_TEAM_MEMBERS} members.` };
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
    
    // Save teams data
    saveTeamsData(teamsData);
    
    // Update the email-to-team map
    updateEmailToTeamMap();
    
    return { success: true, message: 'You have successfully left the team.' };
  } catch (e) {
    Logger.log(`Error in leaveTeam: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// Explicitly export the joinTeam function to ensure it's globally accessible
// This is not technically needed as all functions in Google Apps Script are globally accessible by default
// But we're adding this to be explicit
this.joinTeam = joinTeam; 