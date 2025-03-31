/**
 * Team Manager UI Helper
 * 
 * This module provides helper functions for the team manager UI:
 * - Styles for the team manager dialog
 * - Scripts for the team manager dialog
 */

var TeamManagerUI = TeamManagerUI || {};

/**
 * Gets the styles for the team manager dialog
 * @return {string} CSS styles
 */
TeamManagerUI.getStyles = function() {
  return `
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
      flex: 1;
    }
    
    .member-email {
      flex: 1;
    }
    
    .member-actions {
      display: flex;
      gap: 8px;
    }
    
    .member-action {
      font-size: 12px;
      padding: 2px 8px;
      min-height: 24px;
    }
    
    .add-member-form {
      margin-top: 16px;
      display: flex;
      gap: 8px;
    }
    
    .add-member-input {
      flex: 1;
    }
    
    .status-message {
      margin-top: 16px;
      padding: 12px;
      border-radius: 4px;
      font-size: 13px;
      display: none;
    }
    
    .status-success {
      background-color: #e6f4ea;
      color: var(--success-color);
      border: 1px solid #ceead6;
      display: block;
    }
    
    .status-error {
      background-color: #fce8e6;
      color: var(--error-color);
      border: 1px solid #f8ccc9;
      display: block;
    }
    
    .status-info {
      background-color: #e8f0fe;
      color: var(--primary-color);
      border: 1px solid #c4dafc;
      display: block;
    }
    
    .loading {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: rgba(255, 255, 255, 0.8);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    }
    
    .spinner {
      border: 4px solid rgba(66, 133, 244, 0.2);
      border-radius: 50%;
      border-top: 4px solid var(--primary-color);
      width: 40px;
      height: 40px;
      animation: spin 0.8s linear infinite;
    }
    
    .hidden {
      display: none !important;
    }
  `;
};

/**
 * Gets the scripts for the team manager dialog
 * @return {string} JavaScript code
 */
TeamManagerUI.getScripts = function() {
  return `
    // Tab switching
    document.querySelectorAll('.tab').forEach(tab => {
      tab.addEventListener('click', function() {
        // Update active tab
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        this.classList.add('active');
        
        // Show corresponding section
        const tabId = this.dataset.tab;
        document.querySelectorAll('.tab-content').forEach(section => {
          section.classList.remove('active');
        });
        document.getElementById(tabId + '-tab').classList.add('active');
      });
    });
    
    // Helper function to show status messages
    function showStatus(message, type) {
      const container = document.getElementById('status-container');
      
      // Clear any existing message
      container.innerHTML = '';
      
      // Create new message element
      const statusEl = document.createElement('div');
      statusEl.className = 'status-message status-' + type;
      statusEl.textContent = message;
      container.appendChild(statusEl);
      
      // Auto-dismiss success messages after 5 seconds
      if (type === 'success') {
        setTimeout(() => {
          statusEl.remove();
        }, 5000);
      }
    }
    
    // Show loading overlay
    function showLoading() {
      document.getElementById('loading-container').classList.remove('hidden');
    }
    
    // Hide loading overlay
    function hideLoading() {
      document.getElementById('loading-container').classList.add('hidden');
    }
    
    // Set button loading state
    function setButtonLoading(button, isLoading) {
      if (isLoading) {
        button.classList.add('button-loading');
        
        // Add spinner if not already present
        if (!button.querySelector('.loading-spinner')) {
          const spinner = document.createElement('span');
          spinner.className = 'loading-spinner';
          button.insertBefore(spinner, button.firstChild);
        }
        
        button.disabled = true;
      } else {
        button.classList.remove('button-loading');
        
        // Remove spinner if present
        const spinner = button.querySelector('.loading-spinner');
        if (spinner) {
          spinner.remove();
        }
        
        button.disabled = false;
      }
    }
    
    // Join team handler
    if (document.getElementById('join-team-button')) {
      document.getElementById('join-team-button').addEventListener('click', function() {
        const teamId = document.getElementById('team-id-input').value.trim();
        if (!teamId) {
          showStatus('Please enter a team ID', 'error');
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            setButtonLoading(document.getElementById('join-team-button'), false);
            
            if (result.success) {
              showStatus('Successfully joined team!', 'success');
              
              // Wait a moment, then reload
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error joining team', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(document.getElementById('join-team-button'), false);
            showStatus('Error: ' + error.message, 'error');
          })
          .joinTeam(teamId);
      });
    }
    
    // Create team handler
    if (document.getElementById('create-team-button')) {
      document.getElementById('create-team-button').addEventListener('click', function() {
        const teamName = document.getElementById('team-name-input').value.trim();
        if (!teamName) {
          showStatus('Please enter a team name', 'error');
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            setButtonLoading(document.getElementById('create-team-button'), false);
            
            if (result.success) {
              showStatus('Team created successfully!', 'success');
              
              // Wait a moment, then reload
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error creating team', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(document.getElementById('create-team-button'), false);
            showStatus('Error: ' + error.message, 'error');
          })
          .createTeam(teamName);
      });
    }
    
    // Copy team ID handler
    if (document.getElementById('copy-team-id')) {
      document.getElementById('copy-team-id').addEventListener('click', function() {
        const teamId = document.getElementById('team-id-display').textContent;
        navigator.clipboard.writeText(teamId).then(function() {
          showStatus('Team ID copied to clipboard!', 'success');
        }, function() {
          // Fallback for browsers that don't support clipboard API
          const tempInput = document.createElement('input');
          tempInput.value = teamId;
          document.body.appendChild(tempInput);
          tempInput.select();
          document.execCommand('copy');
          document.body.removeChild(tempInput);
          showStatus('Team ID copied to clipboard!', 'success');
        });
      });
    }
    
    // Leave team handler
    if (document.getElementById('leave-team-button')) {
      document.getElementById('leave-team-button').addEventListener('click', function() {
        if (!confirm('Are you sure you want to leave this team? You will lose access to shared configurations.')) {
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            setButtonLoading(document.getElementById('leave-team-button'), false);
            
            if (result.success) {
              showStatus('You have left the team.', 'success');
              
              // Wait a moment, then reload
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error leaving team', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(document.getElementById('leave-team-button'), false);
            showStatus('Error: ' + error.message, 'error');
          })
          .leaveTeam();
      });
    }
    
    // Delete team handler (admin only)
    if (document.getElementById('delete-team-button')) {
      document.getElementById('delete-team-button').addEventListener('click', function() {
        if (!confirm('Are you sure you want to delete this team? This will remove access for all team members and cannot be undone.')) {
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            setButtonLoading(document.getElementById('delete-team-button'), false);
            
            if (result.success) {
              showStatus('Team has been deleted.', 'success');
              
              // Wait a moment, then reload
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error deleting team', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(document.getElementById('delete-team-button'), false);
            showStatus('Error: ' + error.message, 'error');
          })
          .deleteTeam();
      });
    }
    
    // Add member handler (admin only)
    if (document.getElementById('add-member-button')) {
      document.getElementById('add-member-button').addEventListener('click', function() {
        const email = document.getElementById('new-member-email').value.trim();
        if (!email) {
          showStatus('Please enter an email address', 'error');
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            setButtonLoading(document.getElementById('add-member-button'), false);
            
            if (result.success) {
              showStatus('Member added successfully.', 'success');
              document.getElementById('new-member-email').value = '';
              
              // Reload the page to show new member
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error adding member', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(document.getElementById('add-member-button'), false);
            showStatus('Error: ' + error.message, 'error');
          })
          .addTeamMember(email);
      });
    }
    
    // Promote member handler
    document.querySelectorAll('.promote-member').forEach(button => {
      button.addEventListener('click', function() {
        const email = this.dataset.email;
        if (!confirm(\`Are you sure you want to promote \${email} to admin?\`)) {
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              showStatus('Member promoted to admin successfully.', 'success');
              
              // Reload the page to show updated role
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error promoting member', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            showStatus('Error: ' + error.message, 'error');
          })
          .promoteTeamMember(email);
      });
    });
    
    // Demote member handler
    document.querySelectorAll('.demote-member').forEach(button => {
      button.addEventListener('click', function() {
        const email = this.dataset.email;
        if (!confirm(\`Are you sure you want to remove admin privileges from \${email}?\`)) {
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              showStatus('Admin privileges removed successfully.', 'success');
              
              // Reload the page to show updated role
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error removing admin privileges', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            showStatus('Error: ' + error.message, 'error');
          })
          .demoteTeamMember(email);
      });
    });
    
    // Remove member handler
    document.querySelectorAll('.remove-member').forEach(button => {
      button.addEventListener('click', function() {
        const email = this.dataset.email;
        if (!confirm(\`Are you sure you want to remove \${email} from the team?\`)) {
          return;
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              showStatus('Member removed successfully.', 'success');
              
              // Reload the page to show updated members list
              setTimeout(() => {
                google.script.run.showTeamManager();
                google.script.host.close();
              }, 2000);
            } else {
              showStatus(result.message || 'Error removing member', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            showStatus('Error: ' + error.message, 'error');
          })
          .removeTeamMember(email);
      });
    });
    
    // Close button handler
    document.getElementById('close-button').addEventListener('click', function() {
      google.script.host.close();
    });
    
    // Hide loading indicator on initial load
    hideLoading();
  `;
}; 