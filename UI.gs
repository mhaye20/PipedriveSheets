/**
 * User Interface
 * 
 * This module handles all UI-related functions:
 * - Showing dialogs and sidebars
 * - Building UI components
 * - Managing user interactions
 */

/**
 * Shows the settings dialog
 */
function showSettings() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Get current settings
    const apiKey = docProps.getProperty('PIPEDRIVE_API_KEY') || '';
    const subdomain = docProps.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
    const entityType = docProps.getProperty('PIPEDRIVE_ENTITY_TYPE') || ENTITY_TYPES.DEALS;
    const filterId = docProps.getProperty('PIPEDRIVE_FILTER_ID') || '';
    const sheetName = docProps.getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    const enableTimestamp = docProps.getProperty('ENABLE_TIMESTAMP') === 'true';
    
    // Create the HTML from a template
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsDialog');
    
    // Pass settings to the template
    htmlTemplate.apiKey = apiKey;
    htmlTemplate.subdomain = subdomain;
    htmlTemplate.entityType = entityType;
    htmlTemplate.filterId = filterId;
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.enableTimestamp = enableTimestamp;
    
    // Get filters if API key and subdomain are configured
    if (apiKey && subdomain) {
      try {
        htmlTemplate.filters = getPipedriveFilters();
      } catch (e) {
        htmlTemplate.error = `Failed to load filters: ${e.message}`;
        htmlTemplate.filters = [];
      }
    } else {
      htmlTemplate.filters = [];
    }
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Pipedrive Settings')
      .setWidth(600)
      .setHeight(550);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Pipedrive Settings');
  } catch (e) {
    Logger.log(`Error in showSettings: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open settings dialog: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the column selector dialog
 */
function showColumnSelector() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Get entity type and sheet name
    const entityType = docProps.getProperty('PIPEDRIVE_ENTITY_TYPE') || ENTITY_TYPES.DEALS;
    const sheetName = docProps.getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Check if API key is configured
    const apiKey = docProps.getProperty('PIPEDRIVE_API_KEY');
    if (!apiKey) {
      SpreadsheetApp.getUi().alert(
        'API Key Required',
        'Please configure your Pipedrive API key in the settings first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      showSettings();
      return;
    }
    
    // Get fields based on entity type
    let fields = [];
    
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fields = getDealFields();
        break;
      case ENTITY_TYPES.PERSONS:
        fields = getPersonFields();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fields = getOrganizationFields();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        fields = getActivityFields();
        break;
      case ENTITY_TYPES.LEADS:
        fields = getLeadFields();
        break;
      case ENTITY_TYPES.PRODUCTS:
        fields = getProductFields();
        break;
    }
    
    // Process fields to create UI-friendly structure
    const availableColumns = extractFields(fields);
    
    // Get currently selected columns
    const selectedColumns = getTeamAwareColumnPreferences(entityType, sheetName) || [];
    
    // Show the column selector UI
    showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName);
  } catch (e) {
    Logger.log(`Error in showColumnSelector: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Extracts fields into a UI-friendly structure
 * @param {Array} fields - The fields to extract
 * @param {string} parentPath - Parent path for nested fields
 * @param {string} parentName - Parent name for nested fields
 * @return {Array} Processed fields
 */
function extractFields(fields, parentPath = '', parentName = '') {
  try {
    const result = [];
    
    // If this is an array of field definitions like from Pipedrive API
    if (Array.isArray(fields)) {
      fields.forEach(field => {
        // Skip unsupported field types
        if (field.field_type === 'picture' ||
            field.field_type === 'file' ||
            field.field_type === 'visibleto') {
          return;
        }
        
        // Ensure the field has a key
        if (!field.key) return;
        
        // Construct path
        const path = parentPath ? `${parentPath}.${field.key}` : field.key;
        
        // Construct readable name
        const name = parentName ? `${parentName} > ${field.name}` : field.name;
        
        // Add the field to the result
        result.push({
          path: path,
          name: name,
          type: field.field_type,
          options: field.options
        });
        
        // Handle nested fields (like fields within products, etc.)
        if (field.fields) {
          const nestedFields = extractFields(field.fields, path, name);
          result.push(...nestedFields);
        }
      });
    }
    // If this is an object with nested fields
    else if (typeof fields === 'object' && fields !== null) {
      Object.keys(fields).forEach(key => {
        // Skip certain keys that aren't proper fields
        if (['id', 'success', 'error', 'key'].includes(key)) {
          return;
        }
        
        // Get the field
        const field = fields[key];
        
        // Skip non-objects
        if (typeof field !== 'object' || field === null) {
          return;
        }
        
        // Construct path
        const path = parentPath ? `${parentPath}.${key}` : key;
        
        // If field has a name property, it's likely a proper field
        if (field.name) {
          // Construct readable name
          const name = parentName ? `${parentName} > ${field.name}` : field.name;
          
          // Add the field to the result
          result.push({
            path: path,
            name: name,
            type: field.field_type || typeof field,
            options: field.options
          });
        }
        
        // Recursively process nested fields
        const nestedFields = extractFields(field, path, field.name || parentName);
        result.push(...nestedFields);
      });
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error in extractFields: ${e.message}`);
    return [];
  }
}

/**
 * Shows the column selector UI
 * @param {Array} availableColumns - Available columns
 * @param {Array} selectedColumns - Selected columns
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
function showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName) {
  try {
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('ColumnSelector');
    
    // Pass data to the template
    htmlTemplate.availableColumns = availableColumns;
    htmlTemplate.selectedColumns = selectedColumns;
    htmlTemplate.entityType = entityType;
    htmlTemplate.sheetName = sheetName;
    
    // Add entity type label
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        htmlTemplate.entityTypeLabel = 'Deal';
        break;
      case ENTITY_TYPES.PERSONS:
        htmlTemplate.entityTypeLabel = 'Person';
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        htmlTemplate.entityTypeLabel = 'Organization';
        break;
      case ENTITY_TYPES.ACTIVITIES:
        htmlTemplate.entityTypeLabel = 'Activity';
        break;
      case ENTITY_TYPES.LEADS:
        htmlTemplate.entityTypeLabel = 'Lead';
        break;
      case ENTITY_TYPES.PRODUCTS:
        htmlTemplate.entityTypeLabel = 'Product';
        break;
      default:
        htmlTemplate.entityTypeLabel = 'Record';
    }
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Select Columns')
      .setWidth(800)
      .setHeight(600);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Columns');
  } catch (e) {
    Logger.log(`Error in showColumnSelectorUI: ${e.message}`);
    throw e;
  }
}

/**
 * Saves column preferences for a specific entity type and sheet
 * @param {Array} columns - Array of column keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {boolean} True if successful, false otherwise
 */
function saveColumnPreferences(columns, entityType, sheetName) {
  return saveTeamAwareColumnPreferences(columns, entityType, sheetName);
}

/**
 * Shows the sync status dialog
 * @param {string} sheetName - Sheet name
 */
function showSyncStatus(sheetName) {
  try {
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('SyncStatus');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    
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
 * Shows the team manager dialog
 * @param {boolean} joinOnly - Whether to show only the join team UI
 */
function showTeamManager(joinOnly = false) {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Unable to determine your email address. Please ensure you are signed in.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    
    // Get user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    
    // Determine if user is in a team
    const hasTeam = userTeam !== null;
    
    // If join-only mode and user isn't in a team, use the simpler join dialog
    if (joinOnly && !hasTeam) {
      return showTeamJoinRequest();
    }
    
    // Create a direct HTML string rather than using a complex template
    // This is more robust and less likely to have loading issues
    let html = `
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
        }
        h3 {
          margin-top: 0;
          margin-bottom: 15px;
          color: #4285F4;
        }
        .card {
          border: 1px solid #ddd;
          border-radius: 4px;
          padding: 15px;
          margin-bottom: 20px;
          background-color: #f9f9f9;
        }
        .section-title {
          font-weight: bold;
          margin-bottom: 10px;
        }
        input[type="text"] {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-bottom: 10px;
          box-sizing: border-box;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 8px 15px;
          border-radius: 4px;
          cursor: pointer;
          margin-right: 10px;
          margin-bottom: 10px;
        }
        button:hover {
          background-color: #3367D6;
        }
        button.secondary {
          background-color: #f1f1f1;
          color: #333;
        }
        button.secondary:hover {
          background-color: #e4e4e4;
        }
        button.danger {
          background-color: #EA4335;
        }
        button.danger:hover {
          background-color: #D62516;
        }
        .footer {
          margin-top: 20px;
          display: flex;
          justify-content: flex-end;
        }
        .tab-container {
          display: flex;
          margin-bottom: 20px;
          border-bottom: 1px solid #ddd;
        }
        .tab {
          padding: 10px 15px;
          cursor: pointer;
        }
        .tab.active {
          color: #4285F4;
          font-weight: bold;
          border-bottom: 2px solid #4285F4;
        }
        .tab-content {
          display: none;
        }
        .tab-content.active {
          display: block;
        }
        .member-list {
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-top: 10px;
          max-height: 200px;
          overflow-y: auto;
        }
        .member {
          padding: 8px 12px;
          border-bottom: 1px solid #eee;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .member:last-child {
          border-bottom: none;
        }
        .status-message {
          padding: 10px;
          margin: 10px 0;
          border-radius: 4px;
        }
        .status-success {
          background-color: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }
        .status-error {
          background-color: #f8d7da;
          color: #721c24;
          border: 1px solid #f5c6cb;
        }
      </style>
      
      <div class="container">
        <h3>${hasTeam ? 'Team Management' : 'Team Access'}</h3>
        
        <div id="status-container"></div>
        
        <div class="tab-container">
          ${hasTeam ? '' : '<div class="tab active" data-tab="join">Join Team</div>'}
          ${hasTeam ? '' : '<div class="tab" data-tab="create">Create Team</div>'}
          ${hasTeam ? '<div class="tab active" data-tab="manage">Manage Team</div>' : ''}
        </div>
    `;
    
    // Join Team Tab
    if (!hasTeam) {
      html += `
        <div id="join-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Join an Existing Team</div>
            <p>Enter the team ID provided by your team administrator:</p>
            <input type="text" id="team-id-input" placeholder="Team ID">
            <div class="footer">
              <button id="join-team-button">Join Team</button>
            </div>
          </div>
        </div>
        
        <div id="create-tab" class="tab-content">
          <div class="card">
            <div class="section-title">Create a New Team</div>
            <p>Create a new team to share Pipedrive configuration with your colleagues:</p>
            <input type="text" id="team-name-input" placeholder="Team Name">
            <div class="footer">
              <button id="create-team-button">Create Team</button>
            </div>
          </div>
        </div>
      `;
    }
    
    // Manage Team Tab
    if (hasTeam) {
      const teamName = userTeam.name || 'Unnamed Team';
      const teamId = userTeam.teamId || '';
      const userRole = userTeam.role || 'Member';
      const isAdmin = userRole === 'Admin';
      
      // Get team members
      const members = [];
      if (teamsData[teamId]) {
        if (teamsData[teamId].members) {
          // New format
          Object.keys(teamsData[teamId].members).forEach(email => {
            members.push({
              email: email,
              role: teamsData[teamId].members[email]
            });
          });
        } else if (teamsData[teamId].memberEmails) {
          // Legacy format
          teamsData[teamId].memberEmails.forEach(email => {
            const isAdmin = teamsData[teamId].adminEmails && 
                           teamsData[teamId].adminEmails.includes(email);
            members.push({
              email: email,
              role: isAdmin ? 'Admin' : 'Member'
            });
          });
        }
      }
      
      html += `
        <div id="manage-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Current Team</div>
            <div style="margin-bottom: 15px;">
              <div><strong>Team Name:</strong> ${teamName}</div>
              <div><strong>Team ID:</strong> <span id="team-id-display">${teamId}</span></div>
              <div><strong>Your Role:</strong> ${userRole}</div>
            </div>
            
            <button id="copy-team-id" class="secondary">Copy Team ID</button>
            
            <div class="section-title" style="margin-top: 20px;">Team Members</div>
            <div class="member-list">
      `;
      
      if (members.length > 0) {
        members.forEach(member => {
          html += `
            <div class="member">
              <div>${member.email}</div>
              <div>
                <span style="margin-right: 10px;">${member.role}</span>
                ${isAdmin && member.email !== userEmail ? 
                  `<button class="secondary remove-member" data-email="${member.email}">Remove</button>` : ''}
              </div>
            </div>
          `;
        });
      } else {
        html += '<div class="member">No members found</div>';
      }
      
      html += `
            </div>
      `;
      
      if (isAdmin) {
        html += `
            <div class="section-title" style="margin-top: 15px;">Add Team Member</div>
            <input type="text" id="new-member-email" placeholder="Email Address">
            <button id="add-member-button">Add Member</button>
        `;
      }
      
      html += `
            <div style="margin-top: 20px;">
              <button id="leave-team-button" class="danger">Leave Team</button>
              ${isAdmin ? '<button id="delete-team-button" class="danger">Delete Team</button>' : ''}
            </div>
          </div>
        </div>
      `;
    }
    
    // Footer
    html += `
        <div class="footer">
          <button id="close-button" class="secondary">Close</button>
        </div>
      </div>
      
      <script>
        // Document ready
        (function() {
          // Tab switching
          document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', function() {
              // Update active tab
              document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
              this.classList.add('active');
              
              // Show corresponding content
              const tabId = this.dataset.tab + '-tab';
              document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
              });
              document.getElementById(tabId).classList.add('active');
            });
          });
          
          // Status message helper
          window.showStatus = function(message, type) {
            const container = document.getElementById('status-container');
            const statusDiv = document.createElement('div');
            statusDiv.className = 'status-message status-' + type;
            statusDiv.textContent = message;
            container.innerHTML = '';
            container.appendChild(statusDiv);
            
            if (type === 'success') {
              setTimeout(() => statusDiv.remove(), 5000);
            }
          };
          
          // Close button
          document.getElementById('close-button').addEventListener('click', function() {
            google.script.host.close();
          });
    `;
    
    // Add join-specific handlers
    if (!hasTeam) {
      html += `
          // Join team handler
          document.getElementById('join-team-button').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-input').value.trim();
            if (!teamId) {
              showStatus('Please enter a team ID', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Joining team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Successfully joined the team!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('join-team-button').disabled = false;
                  showStatus(result.message || 'Error joining team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('join-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .joinTeam(teamId);
          });
          
          // Create team handler
          document.getElementById('create-team-button').addEventListener('click', function() {
            const teamName = document.getElementById('team-name-input').value.trim();
            if (!teamName) {
              showStatus('Please enter a team name', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Creating team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team created successfully!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('create-team-button').disabled = false;
                  showStatus(result.message || 'Error creating team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('create-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .createTeam(teamName);
          });
      `;
    }
    
    // Add team management handlers
    if (hasTeam) {
      html += `
          // Copy team ID handler
          document.getElementById('copy-team-id').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-display').textContent;
            
            // Create a temporary input element
            const tempInput = document.createElement('input');
            tempInput.value = teamId;
            document.body.appendChild(tempInput);
            tempInput.select();
            document.execCommand('copy');
            document.body.removeChild(tempInput);
            
            showStatus('Team ID copied to clipboard!', 'success');
          });
          
          // Leave team handler
          document.getElementById('leave-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to leave this team? You will lose access to shared configurations.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Leaving team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('You have left the team.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('leave-team-button').disabled = false;
                  showStatus(result.message || 'Error leaving team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('leave-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .leaveTeam();
          });
      `;
      
      // Add admin-specific handlers
      if (userTeam && userTeam.role === 'Admin') {
        html += `
          // Add member handler
          document.getElementById('add-member-button').addEventListener('click', function() {
            const email = document.getElementById('new-member-email').value.trim();
            if (!email) {
              showStatus('Please enter an email address', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Adding member...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Member added successfully.', 'success');
                  document.getElementById('new-member-email').value = '';
                  document.getElementById('add-member-button').disabled = false;
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('add-member-button').disabled = false;
                  showStatus(result.message || 'Error adding member', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('add-member-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .addTeamMember(email);
          });
          
          // Remove member handlers
          document.querySelectorAll('.remove-member').forEach(button => {
            button.addEventListener('click', function() {
              const email = this.dataset.email;
              if (!confirm('Are you sure you want to remove ' + email + ' from the team?')) {
                return;
              }
              
              this.disabled = true;
              showStatus('Removing member...', 'info');
              
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    showStatus('Member removed successfully.', 'success');
                    setTimeout(() => window.top.location.reload(), 2000);
                  } else {
                    document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                    showStatus(result.message || 'Error removing member', 'error');
                  }
                })
                .withFailureHandler(function(error) {
                  document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                  showStatus('Error: ' + error.message, 'error');
                })
                .removeTeamMember(email);
            });
          });
          
          // Delete team handler
          document.getElementById('delete-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to delete this team? This will remove access for all team members and cannot be undone.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Deleting team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team has been deleted.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('delete-team-button').disabled = false;
                  showStatus(result.message || 'Error deleting team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('delete-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .deleteTeam();
          });
        `;
      }
    }
    
    // Close script tag
    html += `
        })(); // End of self-executing function
      </script>
    `;
    
    // Create HTML service object
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setTitle(hasTeam ? 'Team Management' : 'Join Team')
      .setWidth(joinOnly ? 450 : 600)
      .setHeight(joinOnly ? 400 : 550);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Join Team');
  } catch (e) {
    Logger.log(`Error in showTeamManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open team manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reopens the team manager after an action
 */
function reopenTeamManager() {
  showTeamManager(false);
}

/**
 * Saves team-aware column preferences
 * @param {Array} columns - Column keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {boolean} True if successful, false otherwise
 */
function saveTeamAwareColumnPreferences(columns, entityType, sheetName) {
  try {
    if (!entityType || !sheetName) {
      return false;
    }
    
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    // Generate storage key
    const key = `COLUMN_PREFS_${entityType}_${sheetName}`;
    
    if (userTeam && userTeam.settings && userTeam.settings.shareColumns) {
      // Team member with shared column preferences - store in team data
      const teamsData = getTeamsData();
      const teamId = userTeam.teamId;
      
      if (teamsData[teamId]) {
        teamsData[teamId].columnPreferences = teamsData[teamId].columnPreferences || {};
        teamsData[teamId].columnPreferences[key] = columns;
        saveTeamsData(teamsData);
      }
    } else {
      // Individual preference - store in user properties
      const userProps = PropertiesService.getUserProperties();
      userProps.setProperty(key, JSON.stringify(columns));
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTeamAwareColumnPreferences: ${e.message}`);
    return false;
  }
}

/**
 * Gets team-aware column preferences
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of column keys
 */
function getTeamAwareColumnPreferences(entityType, sheetName) {
  try {
    if (!entityType || !sheetName) {
      return [];
    }
    
    const key = `COLUMN_PREFS_${entityType}_${sheetName}`;
    
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    if (userTeam && userTeam.settings && userTeam.settings.shareColumns) {
      // Check team preferences first
      if (userTeam.columnPreferences && userTeam.columnPreferences[key]) {
        return userTeam.columnPreferences[key];
      }
    }
    
    // Fall back to user preferences
    const userProps = PropertiesService.getUserProperties();
    const prefsJson = userProps.getProperty(key);
    
    if (prefsJson) {
      return JSON.parse(prefsJson);
    }
    
    return [];
  } catch (e) {
    Logger.log(`Error in getTeamAwareColumnPreferences: ${e.message}`);
    return [];
  }
}

/**
 * Gets team-aware Pipedrive filters
 * @return {Array} Array of filter objects
 */
function getTeamAwarePipedriveFilters() {
  try {
    // Get user filters
    const userFilters = getPipedriveFilters();
    
    // Check if user is in a team with shared filters
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    // If not in a team or team doesn't share filters, return user filters
    if (!userTeam || !userTeam.settings || !userTeam.settings.shareFilters) {
      return userFilters;
    }
    
    // Return team filters if available
    return userFilters;
  } catch (e) {
    Logger.log(`Error in getTeamAwarePipedriveFilters: ${e.message}`);
    return [];
  }
}

/**
 * Saves settings for the add-on
 * @param {string} apiKey - Pipedrive API key
 * @param {string} entityType - Entity type
 * @param {string} filterId - Filter ID
 * @param {string} subdomain - Pipedrive subdomain
 * @param {string} sheetName - Sheet name
 * @param {boolean} enableTimestamp - Whether to enable timestamp
 * @return {boolean} True if successful, false otherwise
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp = false) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Save the settings
    if (apiKey) docProps.setProperty('PIPEDRIVE_API_KEY', apiKey);
    if (entityType) docProps.setProperty('PIPEDRIVE_ENTITY_TYPE', entityType);
    if (subdomain) docProps.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
    if (sheetName) docProps.setProperty('EXPORT_SHEET_NAME', sheetName);
    
    // Only update filter ID if provided (may be empty intentionally)
    if (filterId !== undefined) {
      docProps.setProperty('PIPEDRIVE_FILTER_ID', filterId);
    }
    
    // Handle boolean property
    docProps.setProperty('ENABLE_TIMESTAMP', enableTimestamp ? 'true' : 'false');
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveSettings: ${e.message}`);
    throw e;
  }
}

/**
 * Shows the help/about dialog
 */
function showHelp() {
  try {
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('Help');
    
    // Get version
    const versionsStr = PropertiesService.getScriptProperties().getProperty('VERSION_HISTORY') || '[]';
    const versions = JSON.parse(versionsStr);
    
    // Get current version (latest in history)
    const currentVersion = versions.length > 0 ? versions[0].version : '1.0.0';
    
    // Pass data to the template
    htmlTemplate.currentVersion = currentVersion;
    htmlTemplate.versionHistory = versions;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Help & About')
      .setWidth(600)
      .setHeight(500);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Help & About');
  } catch (e) {
    Logger.log(`Error in showHelp: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open help dialog: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the trigger manager dialog
 */
function showTriggerManager() {
  try {
    // Get current triggers
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = PropertiesService.getDocumentProperties().getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Get all triggers for this sheet
    const triggers = getTriggersForSheet(sheetName);
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TriggerManager');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.triggers = triggers;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Schedule Sync')
      .setWidth(600)
      .setHeight(500);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Sync');
  } catch (e) {
    Logger.log(`Error in showTriggerManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open trigger manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets triggers for a specific sheet
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of trigger info objects
 */
function getTriggersForSheet(sheetName) {
  try {
    const result = [];
    const triggers = ScriptApp.getProjectTriggers();
    
    // Format each trigger into a more user-friendly object
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'syncSheetFromTrigger') {
        const triggerInfo = getTriggerInfo(trigger);
        
        // Only include if it's for this sheet or for all sheets
        if (!triggerInfo.sheetName || triggerInfo.sheetName === sheetName) {
          result.push(triggerInfo);
        }
      }
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error in getTriggersForSheet: ${e.message}`);
    return [];
  }
}

/**
 * Gets information about a trigger
 * @param {Trigger} trigger - Trigger object
 * @return {Object} Trigger info object
 */
function getTriggerInfo(trigger) {
  try {
    // Get trigger ID
    const triggerId = trigger.getUniqueId();
    
    // Get trigger properties from storage
    const userProps = PropertiesService.getUserProperties();
    const triggerPropsJson = userProps.getProperty(`TRIGGER_${triggerId}`);
    const triggerProps = triggerPropsJson ? JSON.parse(triggerPropsJson) : {};
    
    // Get trigger type
    const triggerType = trigger.getEventType();
    let schedule = '';
    
    if (triggerType === ScriptApp.EventType.CLOCK) {
      // Time-based trigger
      const hour = trigger.getAtHour();
      const minute = trigger.getAtMinute();
      const weekDay = trigger.getWeekDay();
      
      if (hour !== null && minute !== null) {
        // Daily or weekly trigger
        const timeStr = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
        
        if (weekDay !== null) {
          // Weekly trigger
          const weekDays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          schedule = `Weekly on ${weekDays[weekDay]} at ${timeStr}`;
        } else {
          // Daily trigger
          schedule = `Daily at ${timeStr}`;
        }
      } else {
        // Hourly trigger
        const frequency = trigger.getAtHour(); // Re-using this field for an hourly frequency
        schedule = `Every ${frequency} hour(s)`;
      }
    }
    
    // Combine info
    return {
      id: triggerId,
      schedule: schedule,
      sheetName: triggerProps.sheetName || 'All sheets',
      createdAt: triggerProps.createdAt || 'Unknown',
      entityType: triggerProps.entityType || 'Unknown',
      description: triggerProps.description || '',
      lastRun: triggerProps.lastRun || 'Never'
    };
  } catch (e) {
    Logger.log(`Error in getTriggerInfo: ${e.message}`);
    return {
      id: 'unknown',
      schedule: 'Unknown',
      sheetName: 'Unknown',
      createdAt: 'Unknown',
      lastRun: 'Never'
    };
  }
}

/**
 * Shows the two-way sync settings dialog
 */
function showTwoWaySyncSettings() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Get current settings
    const enableTwoWaySync = docProps.getProperty('ENABLE_TWO_WAY_SYNC') === 'true';
    const trackingColumn = docProps.getProperty('SYNC_TRACKING_COLUMN') || '';
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
    
    // Pass data to the template
    htmlTemplate.enableTwoWaySync = enableTwoWaySync;
    htmlTemplate.trackingColumn = trackingColumn;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Two-Way Sync Settings')
      .setWidth(500)
      .setHeight(400);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
  } catch (e) {
    Logger.log(`Error in showTwoWaySyncSettings: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open two-way sync settings: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Saves two-way sync settings
 * @param {boolean} enableTwoWaySync - Whether to enable two-way sync
 * @param {string} trackingColumn - Column letter for tracking changes
 * @return {boolean} True if successful, false otherwise
 */
function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Save settings
    docProps.setProperty('ENABLE_TWO_WAY_SYNC', enableTwoWaySync ? 'true' : 'false');
    
    if (trackingColumn) {
      docProps.setProperty('SYNC_TRACKING_COLUMN', trackingColumn);
    } else {
      docProps.deleteProperty('SYNC_TRACKING_COLUMN');
    }
    
    // If two-way sync is enabled, set up the onEdit trigger
    if (enableTwoWaySync) {
      setupOnEditTrigger();
    } else {
      removeOnEditTrigger();
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTwoWaySyncSettings: ${e.message}`);
    throw e;
  }
}

/**
 * Shows a simple dialog for joining a team
 * This is shown when a user without a team tries to use the add-on
 */
function showTeamJoinRequest() {
  try {
    const html = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 400px;
          margin: 0 auto;
        }
        h3 {
          margin-top: 0;
          color: #4285F4;
        }
        p {
          margin-bottom: 15px;
        }
        .form-group {
          margin-bottom: 15px;
        }
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: bold;
        }
        input[type="text"] {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 8px 15px;
          border-radius: 4px;
          cursor: pointer;
        }
        button:hover {
          background-color: #3367D6;
        }
        .error {
          color: #d32f2f;
          margin-top: 5px;
        }
        .options {
          margin-top: 20px;
          text-align: center;
        }
        .options a {
          color: #4285F4;
          text-decoration: none;
        }
        .options a:hover {
          text-decoration: underline;
        }
      </style>
      <div class="container">
        <h3>Join a Team</h3>
        <p>Enter a team ID to join an existing team:</p>
        
        <div class="form-group">
          <label for="team-id">Team ID</label>
          <input type="text" id="team-id" placeholder="Enter team ID">
          <div id="error-message" class="error"></div>
        </div>
        
        <button id="join-button">Join Team</button>
        
        <div class="options">
          <p>Don't have a team ID? <a href="#" id="create-team-link">Create a new team</a></p>
        </div>
      </div>
      
      <script>
        // Join team button handler
        document.getElementById('join-button').addEventListener('click', function() {
          // Get team ID
          const teamId = document.getElementById('team-id').value.trim();
          if (!teamId) {
            document.getElementById('error-message').textContent = 'Please enter a team ID';
            return;
          }
          
          // Call server function
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                // Refresh the page to show the full menu
                window.top.location.reload();
              } else {
                document.getElementById('error-message').textContent = result.message || 'Error joining team';
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('error-message').textContent = 'Error: ' + error.message;
            })
            .joinTeam(teamId);
        });
        
        // Create team link handler
        document.getElementById('create-team-link').addEventListener('click', function(e) {
          e.preventDefault();
          
          // Show the team management dialog in create mode
          google.script.run.showTeamManager(false);
          google.script.host.close();
        });
      </script>
    `)
    .setWidth(450)
    .setHeight(300)
    .setTitle('Join Team');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Join Team');
  } catch (e) {
    Logger.log('Error in showTeamJoinRequest: ' + e.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to show join request: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
} 