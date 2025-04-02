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
      --primary-hover: #5294ff;
      --primary-dark: #3367d6;
      --primary-light: #e8f0fe;
      --success-color: #0f9d58;
      --success-light: #e6f4ea;
      --warning-color: #f4b400;
      --warning-light: #fef7e0;
      --error-color: #ea4335;
      --error-hover: #d93025;
      --error-light: #fce8e6;
      --text-dark: #202124;
      --text-medium: #5f6368;
      --text-light: #80868b;
      --bg-light: #f8f9fa;
      --bg-white: #ffffff;
      --border-color: #dadce0;
      --border-light: #f1f3f4;
      --section-bg: #f8f9fa;
      --shadow-small: 0 1px 2px rgba(60,64,67,0.1), 0 1px 3px rgba(60,64,67,0.12);
      --shadow-medium: 0 2px 6px rgba(60,64,67,0.15), 0 1px 2px rgba(60,64,67,0.2);
      --shadow-large: 0 4px 8px rgba(60,64,67,0.2), 0 2px 4px rgba(60,64,67,0.15);
      --transition-fast: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      --transition-medium: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      --border-radius-small: 4px;
      --border-radius-medium: 8px;
      --border-radius-large: 12px;
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
      padding: 0;
      font-size: 14px;
      background-color: var(--bg-light);
    }
    
    #main-container {
      max-width: 100%;
      margin: 0 auto;
      display: flex;
      flex-direction: column;
      min-height: 100vh;
    }
    
    .header {
      padding: 16px 20px 0;
    }
    
    .footer {
      padding: 16px 20px;
      border-top: 1px solid var(--border-color);
      margin-top: auto;
      text-align: right;
    }
    
    h3 {
      font-size: 20px;
      font-weight: 500;
      margin-bottom: 16px;
      color: var(--text-dark);
    }
    
    h4 {
      font-size: 16px;
      font-weight: 500;
      margin: 0;
      color: var(--text-dark);
    }
    
    h5 {
      font-size: 14px;
      font-weight: 500;
      margin: 0;
      color: var(--text-dark);
    }
    
    p {
      margin-bottom: 16px;
      color: var(--text-medium);
    }

    .material-icons {
      font-size: 20px;
      vertical-align: middle;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    @keyframes fadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }

    @keyframes scaleIn {
      from { transform: scale(0.95); opacity: 0; }
      to { transform: scale(1); opacity: 1; }
    }
    
    @keyframes fadeOut {
      from { opacity: 1; transform: translateY(0); }
      to { opacity: 0; transform: translateY(10px); }
    }
    
    @keyframes slideOutLeft {
      from { transform: translateX(0); opacity: 1; }
      to { transform: translateX(-30px); opacity: 0; }
    }
    
    .fade-out {
      animation: fadeOut 0.5s forwards;
    }
    
    .slide-out-left {
      animation: slideOutLeft 0.5s forwards;
    }
    
    .fade-in {
      animation: fadeIn 0.5s;
    }

    .loading-spinner {
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid rgba(255,255,255,0.5);
      border-radius: 50%;
      border-top-color: white;
      animation: spin 0.8s cubic-bezier(0.4, 0, 0.2, 1) infinite;
      margin-right: 8px;
      vertical-align: middle;
    }

    .button-text {
      display: inline-block;
      vertical-align: middle;
    }

    .button-loading {
      background-color: var(--primary-hover) !important;
      cursor: not-allowed !important;
      opacity: 0.85;
      transition: var(--transition-fast);
    }
    
    .section {
      margin-bottom: 24px;
    }
    
    .section-header {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom: 1px solid var(--border-light);
    }
    
    .section-header .material-icons {
      color: var(--primary-color);
    }
    
    .form-container {
      max-width: 100%;
    }
    
    .user-info {
      background-color: var(--primary-light);
      padding: 12px 16px;
      border-radius: var(--border-radius-medium);
      margin-bottom: 20px;
      display: flex;
      align-items: center;
      font-size: 13px;
      box-shadow: var(--shadow-small);
      animation: scaleIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .user-info .material-icons {
      margin-right: 12px;
      color: var(--primary-color);
      font-size: 24px;
    }
    
    .card {
      background-color: var(--bg-white);
      border-radius: var(--border-radius-medium);
      margin-bottom: 20px;
      overflow: hidden;
      box-shadow: var(--shadow-small);
      transition: var(--transition-medium);
      animation: scaleIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .card:hover {
      box-shadow: var(--shadow-medium);
    }
    
    .card-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 16px 20px;
      background-color: var(--primary-light);
      border-bottom: 1px solid var(--border-light);
    }
    
    .card-header .material-icons {
      color: var(--primary-color);
      margin-right: 8px;
      font-size: 24px;
    }
    
    .card-header h4 {
      color: var(--primary-dark);
    }
    
    .card-body {
      padding: 20px;
    }
    
    .form-row {
      display: flex;
      gap: 16px;
      margin-bottom: 16px;
    }
    
    .form-group {
      margin-bottom: 20px;
      flex: 1;
    }
    
    .form-group:last-child {
      margin-bottom: 0;
    }
    
    .label {
      font-weight: 500;
      margin-bottom: 6px;
      color: var(--text-medium);
      font-size: 13px;
    }
    
    .input-container {
      position: relative;
      display: flex;
      align-items: center;
    }
    
    .input-icon {
      position: absolute;
      left: 10px;
      color: var(--text-light);
      font-size: 18px;
    }
    
    input, select {
      width: 100%;
      padding: 10px 12px 10px 36px;
      border: 1px solid var(--border-color);
      border-radius: var(--border-radius-small);
      font-size: 14px;
      transition: var(--transition-fast);
      background-color: var(--bg-white);
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
      margin-top: 6px;
    }
    
    .button-container {
      display: flex;
      justify-content: flex-end;
      margin-top: 20px;
    }
    
    button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      border: none;
      border-radius: var(--border-radius-small);
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: var(--transition-fast);
      padding: 10px 20px;
      min-height: 36px;
    }
    
    button .material-icons {
      font-size: 18px;
    }
    
    .button-primary {
      background-color: var(--primary-color);
      color: white;
    }
    
    .button-primary:hover {
      background-color: var(--primary-hover);
      box-shadow: var(--shadow-medium);
    }
    
    .button-secondary {
      background-color: transparent;
      color: var(--primary-color);
      border: 1px solid var(--primary-color);
    }
    
    .button-secondary:hover {
      background-color: var(--primary-light);
    }
    
    .button-danger {
      background-color: var(--error-color);
      color: white;
    }
    
    .button-danger:hover {
      background-color: var(--error-hover);
      box-shadow: var(--shadow-medium);
    }
    
    .icon-button {
      background-color: transparent;
      color: var(--text-medium);
      padding: 4px;
      min-height: unset;
      border-radius: 50%;
      width: 32px;
      height: 32px;
    }
    
    .icon-button:hover {
      background-color: var(--bg-light);
      color: var(--primary-color);
    }
    
    .icon-button.promote-member:hover {
      background-color: var(--success-light);
      color: var(--success-color);
    }
    
    .icon-button.demote-member:hover, 
    .icon-button.remove-member:hover {
      background-color: var(--error-light);
      color: var(--error-color);
    }
    
    .tab-container {
      flex: 1;
      padding: 0 20px;
    }
    
    .tabs {
      display: flex;
      border-bottom: 1px solid var(--border-color);
      margin-bottom: 16px;
    }
    
    .tab {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 12px 16px;
      cursor: pointer;
      font-weight: 500;
      color: var(--text-medium);
      border-bottom: 2px solid transparent;
      transition: var(--transition-fast);
      border-top-left-radius: var(--border-radius-small);
      border-top-right-radius: var(--border-radius-small);
    }
    
    .tab:hover {
      color: var(--primary-color);
      background-color: var(--primary-light);
    }
    
    .tab.active {
      color: var(--primary-color);
      border-bottom-color: var(--primary-color);
      background-color: var(--primary-light);
    }
    
    .tab-content {
      padding: 16px 0;
      display: none;
      animation: fadeIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .tab-content.active {
      display: block;
    }
    
    .team-dashboard {
      animation: scaleIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .team-name-container {
      display: flex;
      align-items: center;
    }
    
    .team-id-section {
      background-color: var(--bg-light);
      padding: 12px 16px;
      border-radius: var(--border-radius-small);
      margin-bottom: 20px;
      border-left: 3px solid var(--primary-color);
    }
    
    .team-id-display-container {
      display: flex;
      align-items: center;
      gap: 8px;
      margin: 6px 0;
    }
    
    .badge {
      font-size: 11px;
      padding: 2px 10px;
      border-radius: 12px;
      background-color: var(--bg-light);
      color: var(--text-medium);
      display: inline-flex;
      align-items: center;
      gap: 4px;
    }
    
    .badge.admin {
      background-color: var(--primary-light);
      color: var(--primary-color);
    }
    
    .badge .material-icons {
      font-size: 14px;
    }
    
    .team-members {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    
    .team-member {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 12px;
      border-radius: var(--border-radius-small);
      background-color: var(--bg-light);
      transition: var(--transition-fast);
      transform-origin: center;
      overflow: hidden;
      min-height: 54px; /* Ensure consistent height for animation */
    }
    
    .team-member:hover {
      background-color: var(--primary-light);
      box-shadow: var(--shadow-small);
    }
    
    .team-member.removing {
      background-color: var(--error-light);
      border-left: 3px solid var(--error-color);
    }
    
    .member-info {
      display: flex;
      align-items: center;
      gap: 12px;
      flex: 1;
    }
    
    .member-info .material-icons {
      color: var(--text-medium);
    }
    
    .member-email {
      font-weight: 500;
      margin-bottom: 2px;
    }
    
    .member-actions {
      display: flex;
      gap: 4px;
    }
    
    .add-member-form {
      display: flex;
      gap: 12px;
      align-items: center;
    }
    
    .add-member-form .input-container {
      flex: 1;
    }
    
    .team-actions {
      display: flex;
      gap: 12px;
      justify-content: flex-end;
      margin-top: 24px;
      padding-top: 16px;
      border-top: 1px solid var(--border-light);
    }
    
    .status-message {
      margin: 0 20px 16px;
      padding: 12px 16px;
      border-radius: var(--border-radius-small);
      font-size: 13px;
      display: none;
      animation: scaleIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .status-success {
      background-color: var(--success-light);
      color: var(--success-color);
      border-left: 3px solid var(--success-color);
      display: block;
    }
    
    .status-error {
      background-color: var(--error-light);
      color: var(--error-color);
      border-left: 3px solid var(--error-color);
      display: block;
    }
    
    .status-info {
      background-color: var(--primary-light);
      color: var(--primary-color);
      border-left: 3px solid var(--primary-color);
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
      backdrop-filter: blur(3px);
      transition: var(--transition-medium);
      opacity: 1;
    }
    
    .loading.hidden {
      opacity: 0;
      pointer-events: none;
    }
    
    .spinner-container {
      display: flex;
      justify-content: center;
      align-items: center;
      width: 80px;
      height: 80px;
      background-color: var(--bg-white);
      border-radius: 50%;
      box-shadow: var(--shadow-large);
      animation: scaleIn 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .spinner {
      border: 4px solid rgba(66, 133, 244, 0.2);
      border-radius: 50%;
      border-top: 4px solid var(--primary-color);
      width: 40px;
      height: 40px;
      animation: spin 1s cubic-bezier(0.4, 0, 0.2, 1) infinite;
    }
    
    .hidden {
      display: none !important;
    }
    
    .no-members-message {
      text-align: center;
      padding: 20px;
      background-color: var(--bg-light);
      border-radius: var(--border-radius-small);
      color: var(--text-medium);
      font-style: italic;
      margin-top: 16px;
      border: 1px dashed var(--border-color);
      display: none;
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
    document.querySelectorAll('.tab').forEach(function(tab) {
      tab.addEventListener('click', function() {
        // Update active tab
        document.querySelectorAll('.tab').forEach(function(t) { t.classList.remove('active'); });
        this.classList.add('active');
        
        // Show corresponding section
        var tabId = this.dataset.tab;
        document.querySelectorAll('.tab-content').forEach(function(section) {
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
        setTimeout(function() {
          statusEl.classList.add('hidden');
          setTimeout(function() {
            statusEl.remove();
          }, 300);
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
          
          // If the button has an icon, hide it during loading
          const icon = button.querySelector('.material-icons');
          if (icon) icon.style.display = 'none';
          
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
        
        // If the button has an icon, show it again
        const icon = button.querySelector('.material-icons');
        if (icon) icon.style.display = '';
        
        button.disabled = false;
      }
    }
    
    // Join team handler
    if (document.getElementById('join-team-button')) {
      document.getElementById('join-team-button').addEventListener('click', function() {
        const teamId = document.getElementById('team-id-input').value.trim();
        if (!teamId) {
          showStatus('Please enter a team ID', 'error');
          document.getElementById('team-id-input').focus();
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
              setTimeout(function() {
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
          document.getElementById('team-name-input').focus();
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
              setTimeout(function() {
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
              
              // Wait a moment, then reload with the join tab active
              setTimeout(function() {
                google.script.run.showTeamManager(true); // Pass true to show join tab
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
              setTimeout(function() {
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
          document.getElementById('new-member-email').focus();
          return;
        }
        
        if (!email.includes('@')) {
          showStatus('Please enter a valid email address', 'error');
          document.getElementById('new-member-email').focus();
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
              
              // Update the UI directly without requiring a reload
              const membersContainer = document.querySelector('.team-members');
              if (membersContainer) {
                // If we have a "no members" message, remove it
                const noMembersMessage = document.querySelector('.no-members-message');
                if (noMembersMessage) {
                  noMembersMessage.classList.add('fade-out');
                  setTimeout(function() {
                    noMembersMessage.remove();
                  }, 300);
                }
                
                // Create a new member element
                const newMember = document.createElement('div');
                newMember.className = 'team-member';
                newMember.style.opacity = '0';
                newMember.style.transform = 'translateY(20px)';
                
                var userEmail = document.querySelector('.user-info strong').textContent;
                var isCurrentUserAdmin = document.querySelector('.badge.admin') !== null;
                
                // Use a similar structure to the existing HTML
                newMember.innerHTML = 
                  '<div class="member-info">' +
                    '<i class="material-icons">person</i>' +
                    '<div>' +
                      '<div class="member-email">' + email + '</div>' +
                      '<div class="badge">Member</div>' +
                    '</div>' +
                  '</div>' +
                  (isCurrentUserAdmin ? 
                  '<div class="member-actions">' +
                    '<button class="icon-button promote-member" data-email="' + email + '" title="Make Admin">' +
                      '<i class="material-icons">upgrade</i>' +
                    '</button>' +
                    '<button class="icon-button remove-member" data-email="' + email + '" title="Remove Member">' +
                      '<i class="material-icons">person_remove</i>' +
                    '</button>' +
                  '</div>'
                   : '');
                
                // Add the new element to the container
                membersContainer.appendChild(newMember);
                
                // Trigger animation after a small delay
                setTimeout(function() {
                  newMember.style.transition = 'all 0.5s cubic-bezier(0.2, 0.8, 0.2, 1)';
                  newMember.style.opacity = '1';
                  newMember.style.transform = 'translateY(0)';
                }, 50);
                
                // Add event listeners to the new buttons
                if (isCurrentUserAdmin) {
                  var promoteButton = newMember.querySelector('.promote-member');
                  var removeButton = newMember.querySelector('.remove-member');
                  
                  if (promoteButton) {
                    promoteButton.addEventListener('click', function() {
                      var email = this.dataset.email;
                      if (!confirm('Are you sure you want to promote ' + email + ' to admin?')) {
                        return;
                      }
                      
                      setButtonLoading(this, true);
                      showLoading();
                      
                      google.script.run
                        .withSuccessHandler(function(result) {
                          hideLoading();
                          setButtonLoading(promoteButton, false);
                          
                          if (result.success) {
                            showStatus('Member promoted to admin successfully.', 'success');
                            
                            // Reload the page to show updated role
                            setTimeout(function() {
                              google.script.run.showTeamManager();
                              google.script.host.close();
                            }, 2000);
                          } else {
                            showStatus(result.message || 'Error promoting member', 'error');
                          }
                        })
                        .withFailureHandler(function(error) {
                          hideLoading();
                          setButtonLoading(promoteButton, false);
                          showStatus('Error: ' + error.message, 'error');
                        })
                        .promoteTeamMember(email);
                    });
                  }
                  
                  if (removeButton) {
                    removeButton.addEventListener('click', function() {
                      var email = this.dataset.email;
                      if (!confirm('Are you sure you want to remove ' + email + ' from the team?')) {
                        return;
                      }
                      
                      // Get the member element
                      var memberElement = this.closest('.team-member');
                      
                      // Add a visual indication that this member is being removed
                      if (memberElement) {
                        memberElement.classList.add('removing');
                      }
                      
                      setButtonLoading(this, true);
                      showLoading();
                      
                      google.script.run
                        .withSuccessHandler(function(result) {
                          hideLoading();
                          
                          if (result.success) {
                            if (memberElement) {
                              // Add slide-out animation
                              memberElement.classList.add('slide-out-left');
                              
                              // After animation completes, remove the element
                              setTimeout(function() {
                                // Add fade-out effect
                                memberElement.style.height = memberElement.offsetHeight + 'px';
                                memberElement.style.minHeight = '0';
                                memberElement.style.height = '0';
                                memberElement.style.marginTop = '0';
                                memberElement.style.marginBottom = '0';
                                memberElement.style.padding = '0';
                                memberElement.style.opacity = '0';
                                
                                // Finally remove after height animation
                                setTimeout(function() {
                                  memberElement.remove();
                                }, 300);
                              }, 500);
                            }
                            
                            // Show success message
                            showStatus('Member removed successfully.', 'success');
                          } else {
                            // If there was an error, reset the button and element
                            setButtonLoading(removeButton, false);
                            if (memberElement) {
                              memberElement.classList.remove('removing');
                            }
                            showStatus(result.message || 'Error removing member', 'error');
                          }
                        })
                        .withFailureHandler(function(error) {
                          hideLoading();
                          setButtonLoading(removeButton, false);
                          
                          // Reset the element if there was an error
                          if (memberElement) {
                            memberElement.classList.remove('removing');
                          }
                          showStatus('Error: ' + error.message, 'error');
                        })
                        .removeTeamMember(email);
                    });
                  }
                }
              }
              
              // Clear the input field
              document.getElementById('new-member-email').value = '';
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
    document.querySelectorAll('.promote-member').forEach(function(button) {
      button.addEventListener('click', function() {
        var email = this.dataset.email;
        if (!confirm('Are you sure you want to promote ' + email + ' to admin?')) {
          return;
        }
        
        // Get the member element
        var memberElement = this.closest('.team-member');
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              showStatus('Member promoted to admin successfully.', 'success');
              
              // Update UI directly without reloading
              if (memberElement) {
                // Find the badge element
                var badgeElement = memberElement.querySelector('.badge');
                var roleIconElement = memberElement.querySelector('.member-info .material-icons');
                
                if (badgeElement) {
                  // Add animation effect
                  badgeElement.style.transition = 'all 0.3s';
                  badgeElement.style.transform = 'scale(1.1)';
                  badgeElement.style.opacity = '0.5';
                  
                  // Update after a small delay
                  setTimeout(function() {
                    badgeElement.classList.add('admin');
                    badgeElement.textContent = 'Admin';
                    badgeElement.style.opacity = '1';
                    
                    // Reset transform after update
                    setTimeout(function() {
                      badgeElement.style.transform = 'scale(1)';
                    }, 150);
                  }, 300);
                }
                
                // Update the icon as well
                if (roleIconElement) {
                  roleIconElement.textContent = 'admin_panel_settings';
                }
                
                // Replace the promote button with demote button
                var actionsContainer = memberElement.querySelector('.member-actions');
                if (actionsContainer) {
                  // Create new demote button
                  var demoteButton = document.createElement('button');
                  demoteButton.className = 'icon-button demote-member';
                  demoteButton.title = 'Remove Admin';
                  demoteButton.dataset.email = email;
                  demoteButton.innerHTML = '<i class="material-icons">remove_moderator</i>';
                  
                  // Replace promote button with demote
                  button.style.opacity = '0';
                  setTimeout(function() {
                    button.remove();
                    actionsContainer.insertBefore(demoteButton, actionsContainer.firstChild);
                    
                    // Add the demote event listener to the new button
                    demoteButton.addEventListener('click', function() {
                      if (!confirm('Are you sure you want to remove admin privileges from ' + email + '?')) {
                        return;
                      }
                      
                      var memberEl = this.closest('.team-member');
                      setButtonLoading(this, true);
                      showLoading();
                      
                      google.script.run
                        .withSuccessHandler(function(result) {
                          hideLoading();
                          setButtonLoading(demoteButton, false);
                          
                          if (result.success) {
                            showStatus('Admin privileges removed successfully.', 'success');
                            
                            // Update UI without reloading
                            if (memberEl) {
                              var badge = memberEl.querySelector('.badge');
                              var roleIcon = memberEl.querySelector('.member-info .material-icons');
                              
                              if (badge) {
                                // Animation for role change
                                badge.style.transition = 'all 0.3s';
                                badge.style.transform = 'scale(0.9)';
                                badge.style.opacity = '0.5';
                                
                                setTimeout(function() {
                                  badge.classList.remove('admin');
                                  badge.textContent = 'Member';
                                  badge.style.opacity = '1';
                                  
                                  setTimeout(function() {
                                    badge.style.transform = 'scale(1)';
                                  }, 150);
                                }, 300);
                              }
                              
                              // Update icon
                              if (roleIcon) {
                                roleIcon.textContent = 'person';
                              }
                              
                              // Replace demote button with promote
                              var actions = memberEl.querySelector('.member-actions');
                              if (actions) {
                                var promoteButton = document.createElement('button');
                                promoteButton.className = 'icon-button promote-member';
                                promoteButton.title = 'Make Admin';
                                promoteButton.dataset.email = email;
                                promoteButton.innerHTML = '<i class="material-icons">upgrade</i>';
                                
                                demoteButton.style.opacity = '0';
                                setTimeout(function() {
                                  demoteButton.remove();
                                  actions.insertBefore(promoteButton, actions.firstChild);
                                  
                                  // Add event listener to the new promote button
                                  promoteButton.addEventListener('click', function() {
                                    var email = this.dataset.email;
                                    if (!confirm('Are you sure you want to promote ' + email + ' to admin?')) {
                                      return;
                                    }
                                    
                                    setButtonLoading(this, true);
                                    showLoading();
                                    
                                    google.script.run
                                      .withSuccessHandler(function(res) {
                                        hideLoading();
                                        setButtonLoading(promoteButton, false);
                                        
                                        if (res.success) {
                                          showStatus('Member promoted to admin successfully.', 'success');
                                          
                                          // Reload the page to show updated role after delay
                                          setTimeout(function() {
                                            google.script.run.showTeamManager();
                                            google.script.host.close();
                                          }, 2000);
                                        } else {
                                          showStatus(res.message || 'Error promoting member', 'error');
                                        }
                                      })
                                      .withFailureHandler(function(err) {
                                        hideLoading();
                                        setButtonLoading(promoteButton, false);
                                        showStatus('Error: ' + err.message, 'error');
                                      })
                                      .promoteTeamMember(email);
                                  });
                                }, 300);
                              }
                            }
                          } else {
                            showStatus(result.message || 'Error removing admin privileges', 'error');
                          }
                        })
                        .withFailureHandler(function(error) {
                          hideLoading();
                          setButtonLoading(demoteButton, false);
                          showStatus('Error: ' + error.message, 'error');
                        })
                        .demoteTeamMember(email);
                    });
                  }, 300);
                }
              }
            } else {
              setButtonLoading(button, false);
              showStatus(result.message || 'Error promoting member', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(button, false);
            showStatus('Error: ' + error.message, 'error');
          })
          .promoteTeamMember(email);
      });
    });
    
    // Demote member handler
    document.querySelectorAll('.demote-member').forEach(function(button) {
      button.addEventListener('click', function() {
        var email = this.dataset.email;
        if (!confirm('Are you sure you want to remove admin privileges from ' + email + '?')) {
          return;
        }
        
        // Get the member element
        var memberElement = this.closest('.team-member');
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              showStatus('Admin privileges removed successfully.', 'success');
              
              // Update UI directly instead of reloading
              if (memberElement) {
                // Find the badge element
                var badgeElement = memberElement.querySelector('.badge');
                var roleIconElement = memberElement.querySelector('.member-info .material-icons');
                
                if (badgeElement) {
                  // Add animation effect
                  badgeElement.style.transition = 'all 0.3s';
                  badgeElement.style.transform = 'scale(0.9)';
                  badgeElement.style.opacity = '0.5';
                  
                  // Update after a small delay
                  setTimeout(function() {
                    badgeElement.classList.remove('admin');
                    badgeElement.textContent = 'Member';
                    badgeElement.style.opacity = '1';
                    
                    // Reset transform after update
                    setTimeout(function() {
                      badgeElement.style.transform = 'scale(1)';
                    }, 150);
                  }, 300);
                }
                
                // Update the icon
                if (roleIconElement) {
                  roleIconElement.textContent = 'person';
                }
                
                // Replace the demote button with promote button
                var actionsContainer = memberElement.querySelector('.member-actions');
                if (actionsContainer) {
                  // Create new promote button
                  var promoteButton = document.createElement('button');
                  promoteButton.className = 'icon-button promote-member';
                  promoteButton.title = 'Make Admin';
                  promoteButton.dataset.email = email;
                  promoteButton.innerHTML = '<i class="material-icons">upgrade</i>';
                  
                  // Replace demote button with promote
                  button.style.opacity = '0';
                  setTimeout(function() {
                    button.remove();
                    actionsContainer.insertBefore(promoteButton, actionsContainer.firstChild);
                    
                    // Add the promote event listener to the new button
                    promoteButton.addEventListener('click', function() {
                      if (!confirm('Are you sure you want to promote ' + email + ' to admin?')) {
                        return;
                      }
                      
                      var memberEl = this.closest('.team-member');
                      setButtonLoading(this, true);
                      showLoading();
                      
                      google.script.run
                        .withSuccessHandler(function(result) {
                          hideLoading();
                          setButtonLoading(promoteButton, false);
                          
                          if (result.success) {
                            showStatus('Member promoted to admin successfully.', 'success');
                            
                            // Update UI without reloading
                            if (memberEl) {
                              var badge = memberEl.querySelector('.badge');
                              var roleIcon = memberEl.querySelector('.member-info .material-icons');
                              
                              if (badge) {
                                // Animation for role change
                                badge.style.transition = 'all 0.3s';
                                badge.style.transform = 'scale(1.1)';
                                badge.style.opacity = '0.5';
                                
                                setTimeout(function() {
                                  badge.classList.add('admin');
                                  badge.textContent = 'Admin';
                                  badge.style.opacity = '1';
                                  
                                  setTimeout(function() {
                                    badge.style.transform = 'scale(1)';
                                  }, 150);
                                }, 300);
                              }
                              
                              // Update icon
                              if (roleIcon) {
                                roleIcon.textContent = 'admin_panel_settings';
                              }
                              
                              // Replace promote button with demote
                              var actions = memberEl.querySelector('.member-actions');
                              if (actions) {
                                var demoteButton = document.createElement('button');
                                demoteButton.className = 'icon-button demote-member';
                                demoteButton.title = 'Remove Admin';
                                demoteButton.dataset.email = email;
                                demoteButton.innerHTML = '<i class="material-icons">remove_moderator</i>';
                                
                                promoteButton.style.opacity = '0';
                                setTimeout(function() {
                                  promoteButton.remove();
                                  actions.insertBefore(demoteButton, actions.firstChild);
                                  
                                  // Add event listener to the new demote button
                                  demoteButton.addEventListener('click', function() {
                                    var email = this.dataset.email;
                                    if (!confirm('Are you sure you want to remove admin privileges from ' + email + '?')) {
                                      return;
                                    }
                                    
                                    setButtonLoading(this, true);
                                    showLoading();
                                    
                                    google.script.run
                                      .withSuccessHandler(function(res) {
                                        hideLoading();
                                        setButtonLoading(demoteButton, false);
                                        
                                        if (res.success) {
                                          showStatus('Admin privileges removed successfully.', 'success');
                                          
                                          // Reload the page to show updated role after delay
                                          setTimeout(function() {
                                            google.script.run.showTeamManager();
                                            google.script.host.close();
                                          }, 2000);
                                        } else {
                                          showStatus(res.message || 'Error removing admin privileges', 'error');
                                        }
                                      })
                                      .withFailureHandler(function(err) {
                                        hideLoading();
                                        setButtonLoading(demoteButton, false);
                                        showStatus('Error: ' + err.message, 'error');
                                      })
                                      .demoteTeamMember(email);
                                  });
                                }, 300);
                              }
                            }
                          } else {
                            showStatus(result.message || 'Error promoting member', 'error');
                          }
                        })
                        .withFailureHandler(function(error) {
                          hideLoading();
                          setButtonLoading(promoteButton, false);
                          showStatus('Error: ' + error.message, 'error');
                        })
                        .promoteTeamMember(email);
                    });
                  }, 300);
                }
              }
            } else {
              setButtonLoading(button, false);
              showStatus(result.message || 'Error removing admin privileges', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(button, false);
            showStatus('Error: ' + error.message, 'error');
          })
          .demoteTeamMember(email);
      });
    });
    
    // Remove member handler
    document.querySelectorAll('.remove-member').forEach(function(button) {
      button.addEventListener('click', function() {
        var email = this.dataset.email;
        if (!confirm('Are you sure you want to remove ' + email + ' from the team?')) {
          return;
        }
        
        // Get the member element
        var memberElement = this.closest('.team-member');
        
        // Add a visual indication that this member is being removed
        if (memberElement) {
          memberElement.classList.add('removing');
        }
        
        setButtonLoading(this, true);
        showLoading();
        
        google.script.run
          .withSuccessHandler(function(result) {
            hideLoading();
            
            if (result.success) {
              if (memberElement) {
                // Add slide-out animation
                memberElement.classList.add('slide-out-left');
                
                // Get the members container to check if this is the last member
                var membersContainer = memberElement.closest('.team-members');
                var remainingMembers = membersContainer.querySelectorAll('.team-member:not(.slide-out-left)').length - 1;
                
                // After animation completes, remove the element
                setTimeout(function() {
                  // Add fade-out effect
                  memberElement.style.height = memberElement.offsetHeight + 'px';
                  memberElement.style.minHeight = '0';
                  memberElement.style.height = '0';
                  memberElement.style.marginTop = '0';
                  memberElement.style.marginBottom = '0';
                  memberElement.style.padding = '0';
                  memberElement.style.opacity = '0';
                  
                  // Finally remove after height animation
                  setTimeout(function() {
                    memberElement.remove();
                    
                    // Check if we need to show the "no members" message
                    if (remainingMembers === 0) {
                      var noMembersMessage = document.createElement('div');
                      noMembersMessage.className = 'no-members-message fade-in';
                      noMembersMessage.textContent = 'No team members found.';
                      membersContainer.appendChild(noMembersMessage);
                      
                      // Display the message
                      setTimeout(function() {
                        noMembersMessage.style.display = 'block';
                      }, 100);
                    }
                  }, 300);
                }, 500);
              }
              
              // Show success message
              showStatus('Member removed successfully.', 'success');
              
              // We don't need to reload the page since we've updated the UI directly
            } else {
              // If there was an error, reset the button and element
              setButtonLoading(button, false);
              if (memberElement) {
                memberElement.classList.remove('removing');
              }
              showStatus(result.message || 'Error removing member', 'error');
            }
          })
          .withFailureHandler(function(error) {
            hideLoading();
            setButtonLoading(button, false);
            
            // Reset the element if there was an error
            if (memberElement) {
              memberElement.classList.remove('removing');
            }
            showStatus('Error: ' + error.message, 'error');
          })
          .removeTeamMember(email);
      });
    });
    
    // Close button handler
    document.getElementById('close-button').addEventListener('click', function() {
      google.script.host.close();
    });
    
    // Add keyboard event listener for Enter key on input fields
    document.querySelectorAll('input').forEach(function(input) {
      input.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          
          // Find the submit button in the current tab and click it
          var tabContent = this.closest('.tab-content');
          if (tabContent) {
            var submitButton = tabContent.querySelector('.button-primary');
            if (submitButton) {
              submitButton.click();
            }
          }
        }
      });
    });
    
    // Hide loading indicator on initial load
    hideLoading();
    
    // Initialize tooltip hover behavior for icon buttons
    document.querySelectorAll('[title]').forEach(function(el) {
      el.addEventListener('mouseover', function() {
        var title = this.getAttribute('title');
        if (!title) return;
        
        // Create tooltip element
        var tooltip = document.createElement('div');
        tooltip.className = 'tooltip-popup';
        tooltip.textContent = title;
        
        // Add styles
        tooltip.style.position = 'absolute';
        tooltip.style.backgroundColor = 'rgba(0, 0, 0, 0.7)';
        tooltip.style.color = 'white';
        tooltip.style.padding = '4px 8px';
        tooltip.style.borderRadius = '4px';
        tooltip.style.fontSize = '12px';
        tooltip.style.zIndex = '10000';
        tooltip.style.transition = 'opacity 0.2s';
        tooltip.style.opacity = '0';
        
        // Position the tooltip
        document.body.appendChild(tooltip);
        
        // Position based on the element's position
        var rect = this.getBoundingClientRect();
        tooltip.style.top = (rect.bottom + 5) + 'px';
        tooltip.style.left = (rect.left + rect.width/2 - tooltip.offsetWidth/2) + 'px';
        
        // Show the tooltip
        setTimeout(function() {
          tooltip.style.opacity = '1';
        }, 10);
        
        // Remove the original title to prevent default tooltip
        this.setAttribute('data-original-title', title);
        this.removeAttribute('title');
        
        // Remove tooltip when mouse leaves
        this.addEventListener('mouseout', function mouseOutHandler() {
          this.setAttribute('title', this.getAttribute('data-original-title'));
          this.removeAttribute('data-original-title');
          tooltip.style.opacity = '0';
          setTimeout(function() {
            if (tooltip.parentNode) {
              document.body.removeChild(tooltip);
            }
          }, 200);
          this.removeEventListener('mouseout', mouseOutHandler);
        });
      });
    });
  `;
};

/**
 * Shows the team management UI.
 * @param {boolean} joinOnly - Whether to show only the join team section.
 */
TeamManagerUI.showTeamManager = function(joinOnly = false) {
  try {
    // Get the active user's email
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('Unable to retrieve your email address. Please make sure you are logged in.');
    }

    // Get team data
    var teamAccess = new TeamAccess();
    var hasTeam = teamAccess.isUserInTeam(userEmail);
    var teamName = '';
    var teamId = '';
    var teamMembers = [];
    var userRole = '';

    if (hasTeam) {
      var teamData = teamAccess.getUserTeamData(userEmail);
      teamName = teamData.name;
      teamId = teamData.id;
      teamMembers = teamAccess.getTeamMembers(teamId);
      userRole = teamData.role;
    }

    // Create the HTML template
    var template = HtmlService.createTemplateFromFile('TeamManager');
    
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
    var htmlOutput = template.evaluate()
      .setWidth(500)
      .setHeight(hasTeam ? 600 : 400)
      .setTitle(hasTeam ? 'Team Management' : 'Team Access')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Team Access');
  } catch (error) {
    Logger.log('Error in showTeamManager: ' + error.message);
    showError('An error occurred while loading the team management interface: ' + error.message);
  }
};

// Export functions to be globally accessible
this.showTeamManager = TeamManagerUI.showTeamManager; 