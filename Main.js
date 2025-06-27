/**
 * Pipedrive to Google Sheets™ Integration
 * 
 * This script connects to Pipedrive API and fetches data based on filters
 * to populate a Google Sheet™ with the requested fields.
 * 
 * Version: 2.0.0
 * Author: Your Name
 * License: MIT
 */

// OAuth scopes needed for the add-on
const OAUTH_SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.container.ui",
  "https://www.googleapis.com/auth/script.scriptapp",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/script.send_mail"
];

// Service namespaces
if (typeof SyncService === 'undefined') {
  var SyncService = {};
}

if (typeof UI === 'undefined') {
  var UI = {};
}

/**
 * Cache for verified users
 */
let VERIFIED_USERS = {};

/**
 * Runs when the add-on is installed
 * @param {Object} e The event parameter for simple onInstall trigger
 */
function onInstall(e) {
  // Call onOpen to set up the menu
  onOpen(e);
  
  // Set a flag to show welcome message on first use
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('FIRST_INSTALL', 'true');
  
  // Track the installation
  trackInstallation(e);
  
  // Send welcome email to new user
  sendWelcomeEmail();
}

/**
 * Tracks installation to backend analytics
 * @param {Object} e The event parameter from onInstall trigger
 */
function trackInstallation(e) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const domain = userEmail ? userEmail.split('@')[1] : 'unknown';
    
    const installData = {
      email: userEmail || 'unknown',
      domain: domain,
      installTime: new Date().toISOString(),
      source: e?.source || 'marketplace', // Allow custom source or default to marketplace
      authMode: e?.authMode || 'unknown'
    };
    
    // Get backend URL from script properties or use default
    const scriptProperties = PropertiesService.getScriptProperties();
    const backendUrl = scriptProperties.getProperty('BACKEND_URL') || 'https://pipedrivesheets-backend.vercel.app';
    
    // Send tracking data to backend
    UrlFetchApp.fetch(`${backendUrl}/api/track-install`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(installData),
      muteHttpExceptions: true // Don't throw errors that could interrupt installation
    });
    
    console.log('Installation tracked successfully');
  } catch (error) {
    // Silently fail - don't interrupt the installation process
    console.log('Failed to track installation:', error);
  }
}

/**
 * Sends a welcome email to new users who install the add-on
 */
function sendWelcomeEmail() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return;
    
    const subject = "Welcome to PipedriveSheets! 🎉";
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #4285F4;">Thanks for being an early user! 🚀</h2>
        
        <p>Hey there!</p>
        
        <p>I just wanted to personally thank you for installing PipedriveSheets. As an early user, you're helping shape this tool into something amazing for the entire community.</p>
        
        <p>Here's what makes this special: <strong>I'm committed to responding to every email within 24 hours</strong>. Seriously! If you run into any issues, have feature requests, or just want to share feedback, don't hesitate to reach out.</p>
        
        <p>Some quick tips to get started:</p>
        <ul>
          <li>Look for the "Pipedrive" menu in your Google Sheets™</li>
          <li>Start with "Get Started with Pipedrive" to connect your account</li>
          <li>The Help section has guides for all the features</li>
        </ul>
        
        <p>I'm genuinely excited to see how you use PipedriveSheets and would love to hear about your experience. Every piece of feedback helps me make it better for everyone.</p>
        
        <p>Thanks again for being part of this journey!</p>
        
        <p>Best,<br>
        Mike Haye<br>
        Founder, PipedriveSheets<br>
        <a href="mailto:support@pipedrivesheets.com">support@pipedrivesheets.com</a></p>
        
        <hr style="margin: 30px 0; border: none; border-top: 1px solid #eee;">
        <p style="font-size: 12px; color: #666;">
          P.S. This email was sent because you installed the PipedriveSheets Google Sheets™ add-on. 
          If you have any questions, just reply to this email!
        </p>
      </div>
    `;
    
    // Send the email from your custom domain
    MailApp.sendEmail({
      to: userEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: 'PipedriveSheets Support',
      replyTo: 'support@pipedrivesheets.com'
    });
    
  } catch (error) {
    // Silently fail - don't interrupt the installation process
    console.log('Failed to send welcome email:', error);
  }
}

/**
 * Creates the menu when the spreadsheet opens
 * @param {Object} e The event parameter (optional)
 */
function onOpen(e) {
  try {
    // First check if user was previously verified as a team member or completed initialization
    const userProperties = PropertiesService.getUserProperties();
    const wasVerified = userProperties.getProperty('VERIFIED_TEAM_MEMBER') === 'true';
    const wasInitialized = userProperties.getProperty('PIPEDRIVE_INITIALIZED') === 'true';
    
    if (wasVerified || wasInitialized) {
      // User was previously verified or initialized, show full menu immediately
      createPipedriveMenu();
      checkForPaymentSuccess();
      return;
    }
    
    // Try to get user email without prompting
    const userEmail = Session.getActiveUser().getEmail();
    
    if (userEmail) {
      // User email is available, try automatic initialization
      
      // Preload verified users
      preloadVerifiedUsers();
      
      // Check access
      if (checkAnyUserAccess(userEmail) || 
          hasVerifiedTeamAccess() || 
          forceTeamMembershipCheck(userEmail)) {
        
        // User has team access, create full menu
        createPipedriveMenu();
        
        // Store verification status for future sessions
        userProperties.setProperty('VERIFIED_TEAM_MEMBER', 'true');
        userProperties.setProperty('PIPEDRIVE_INITIALIZED', 'true');
        
        // Check if user just completed a payment
        checkForPaymentSuccess();
        return;
      }
      
      // Check if user has previously completed OAuth (can use without team)
      try {
        const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
        if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED) {
          // User has authorized - give them access to individual features
          createPipedriveMenu();
          userProperties.setProperty('PIPEDRIVE_INITIALIZED', 'true');
          checkForPaymentSuccess();
          return;
        }
      } catch (e) {
        // Continue with initialization menu
      }
    }
    
    // If we can't auto-initialize, show the initialization menu
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Pipedrive')
      .addItem('🚀 Get Started with Pipedrive', 'initializePipedriveMenu')
      .addToUi();
      
    // Check if this is the first time opening
    const hasSeenWelcome = userProperties.getProperty('HAS_SEEN_WELCOME');
    
    if (!hasSeenWelcome && !userEmail) {
      // Show welcome message for first-time users
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '👋 Welcome! Click "Pipedrive" → "Get Started" to begin.',
        'Welcome to PipedriveSheets',
        5
      );
      userProperties.setProperty('HAS_SEEN_WELCOME', 'true');
    }
    
    // Check if this is the first install
    const isFirstInstall = userProperties.getProperty('FIRST_INSTALL');
    if (isFirstInstall === 'true') {
      userProperties.deleteProperty('FIRST_INSTALL');
      // You can add special first-install logic here if needed
    }
      
  } catch (error) {
    
    // Fallback to initialization menu
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Pipedrive')
      .addItem('Initialize Pipedrive Menu', 'initializePipedriveMenu')
      .addToUi();
  }
}

/**
 * Check if user just returned from successful payment
 */
function checkForPaymentSuccess() {
  try {
    // Get current URL to check for success parameter
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const url = spreadsheet.getUrl();
    
    // This is a simple check - in a real implementation you might want to 
    // use a more sophisticated method to detect the redirect
    // For now, we'll use a timed approach or let the user manually check
    
    // Alternative: Set a flag in Properties when payment starts, check when it completes
    const userProperties = PropertiesService.getUserProperties();
    const paymentInProgress = userProperties.getProperty('PAYMENT_IN_PROGRESS');
    
    if (paymentInProgress) {
      // Clear the flag
      userProperties.deleteProperty('PAYMENT_IN_PROGRESS');
      
      // Small delay to ensure webhook has processed
      Utilities.sleep(2000);
      
      // Check if subscription was upgraded
      const currentPlan = PaymentService.getCurrentPlan();
      if (currentPlan.plan !== 'free') {
        // Show success message
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Payment Successful! 🎉',
          `Welcome to ${currentPlan.details.name}! Your premium features are now active.\n\n` +
          `✅ ${currentPlan.details.features.join('\n✅ ')}\n\n` +
          'You can now access all premium features from the Pipedrive menu.',
          ui.ButtonSet.OK
        );
      }
    }
  } catch (error) {
    // Silently handle errors in this check
  }
}

/**
 * This function will be called when the user clicks "Initialize Pipedrive Menu"
 * At this point, we'll definitely have their email and can do proper verification
 */
function initializePipedriveMenu() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Track manual initialization (if not already tracked)
    const userProperties = PropertiesService.getUserProperties();
    const hasTrackedInstall = userProperties.getProperty('HAS_TRACKED_INSTALL');
    if (!hasTrackedInstall) {
      trackInstallation({ authMode: 'FULL', source: 'manual' });
      userProperties.setProperty('HAS_TRACKED_INSTALL', 'true');
    }
    
    // Preload verified users
    preloadVerifiedUsers();
    
    // Check if user has team access (team owner or verified team member)
    if (checkAnyUserAccess(userEmail) || 
        hasVerifiedTeamAccess() || 
        forceTeamMembershipCheck(userEmail)) {
      
      // User has team access - show full menu with team features
      createPipedriveMenu();
      
      // Store verification status for future sessions
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('VERIFIED_TEAM_MEMBER', 'true');
      userProperties.setProperty('PIPEDRIVE_INITIALIZED', 'true');
      
      // Show a toast notification instead of a modal
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '✅ Pipedrive menu ready! Check the menu bar above.',
        'Initialization Complete',
        3
      );
      
      return true;
    } else {
      // Check if user is an invited team member who needs verification
      if (isUserInTeam(userEmail)) {
        // User is in a team but needs verification - show team access dialog
        verifyTeamAccess();
        return false;
      } else {
        // User is not part of any team - give individual access after OAuth
        createPipedriveMenu();
        
        // Store that user has completed initialization for future sessions
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('PIPEDRIVE_INITIALIZED', 'true');
        
        // Special handling for test users - grant them team access after initialization
        if (TEST_USERS.includes(userEmail.toLowerCase())) {
          userProperties.setProperty('VERIFIED_TEAM_MEMBER', 'true');
          SpreadsheetApp.getActiveSpreadsheet().toast(
            '✅ Pipedrive menu ready with full team access! Check the menu bar above.',
            'Test User Initialization Complete', 
            3
          );
        } else {
          // Show a toast notification 
          SpreadsheetApp.getActiveSpreadsheet().toast(
            '✅ Pipedrive menu ready! Check the menu bar above.',
            'Initialization Complete', 
            3
          );
        }
        
        return true;
      }
    }
  } catch (e) {
    
    // Show error to user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'An error occurred while initializing Pipedrive menu: ' + e.message, ui.ButtonSet.OK);
    return false;
  }
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
      .addItem('⚙️ Pipedrive Settings', 'showSettingsTab')
      .addItem('🔁 Two-Way Sync Settings', 'showTwoWaySyncSettings')
      .addItem('👥 Team Management', 'showTeamManager')
      .addSeparator()
      .addItem('⏱️ Schedule Sync', 'showTriggerManager')
      .addSeparator();
  
  // For subscription menu items, we'll show both options since onOpen has limited authorization
  // The individual functions will handle showing appropriate content based on actual subscription status
  menu.addItem('💎 Upgrade Plan', 'showUpgradeDialog')
      .addItem('💳 Manage Subscription', 'showManageSubscription');
  
  menu.addItem('ℹ️ Help & About', 'showHelp');
      
  menu.addToUi();
}

/**
 * Shows the upgrade dialog
 */
function showUpgradeDialog() {
  // Check if user already has a subscription
  try {
    const currentPlan = PaymentService.getCurrentPlan();
    
    // Only redirect to subscription management if user has Team plan
    // Pro users should be able to upgrade to Team plan
    if (currentPlan.plan === 'team') {
      // User already has a Team subscription, redirect to manage subscription
      SpreadsheetApp.getUi().alert(
        'Active Team Subscription',
        `You already have an active ${currentPlan.details.name} subscription. Redirecting to subscription management...`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      showManageSubscription();
      return;
    }
  } catch (error) {
  }
  
  // Show upgrade dialog for free users and Pro users (who can upgrade to Team)
  PaymentService.showUpgradeDialog();
}

/**
 * Creates a checkout session - called from HTML dialog
 * @param {string} planType - The plan type (e.g., 'pro_monthly', 'pro_annual', etc.)
 * @return {string} The checkout URL or null if error
 */
function createCheckoutSession(planType) {
  return PaymentService.createCheckoutSession(planType);
}

/**
 * Sets a flag indicating payment is in progress
 */
function setPaymentInProgress() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('PAYMENT_IN_PROGRESS', 'true');
}

/**
 * Gets the current plan - called from HTML dialog
 * @return {Object} The current plan details
 */
function getCurrentPlan() {
  return PaymentService.getCurrentPlan();
}

/**
 * Shows the manage subscription dialog
 */
function showManageSubscription() {
  // Let the dialog handle all plan types including inherited team plans
  PaymentService.showManageSubscriptionDialog();
}

/**
 * Creates a customer portal session - called from HTML dialog
 * @return {string} The portal URL or null if error
 */
function createCustomerPortalSession() {
  return PaymentService.createCustomerPortalSession();
}

/**
 * Clears the subscription cache - called from HTML dialog
 */
function clearSubscriptionCache() {
  PaymentService.clearSubscriptionCache();
}

/**
 * Checks if a user has any type of access
 * @param {string} userEmail The email address to check
 * @returns {boolean} True if the user has access, false otherwise
 */
function checkAnyUserAccess(userEmail) {
  try {
    if (!userEmail) return false;
    
    // APPROACH 1: Direct ownership check
    try {
      const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
      if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED) {
        return true;
      }
    } catch (e) {
    }
    
    // APPROACH 2: Check email map in document properties-`
    try {
      const docProps = PropertiesService.getDocumentProperties();
      const emailMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
      
      if (emailMapStr) {
        const emailMap = JSON.parse(emailMapStr);
        
        // Simplified check - just use lowercase consistently
        if (emailMap[userEmail.toLowerCase()]) {
          return true;
        }
      }
    } catch (mapError) {
    }
    
    // APPROACH 3: Direct check of teams data
    try {
      if (isUserInTeam(userEmail)) {
        return true;
      }
    } catch (teamError) {
    }
    
    return false;
  } catch (e) {
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
    
    
    // Try getting from document properties directly
    const docProps = PropertiesService.getDocumentProperties();
    const emailToTeamMapStr = docProps.getProperty('EMAIL_TO_TEAM_MAP');
    
    if (emailToTeamMapStr) {
      const emailToTeamMap = JSON.parse(emailToTeamMapStr);
      
      if (emailToTeamMap[userEmail.toLowerCase()]) {
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
          return true;
        }
      }
    }
    
    return false;
  } catch (e) {
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
        return true;
      }
    } catch (e) {
    }
    
    // Directly check if user is in a team
    if (isUserInTeam(userEmail)) {
      return true;
    }
    
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Preloads verified user data to ensure faster access checks
 * @return {boolean} True if preload was successful, false otherwise
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
      
    }
    return true;
  } catch (e) {
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
        }
      } catch (e) {
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
        
      } catch (e) {
      }
      
      return { success: true };
    } else {
      return { success: false, error: 'You are not a member of any team. Please join a team first.' };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Shows the team join request dialog
 */
function showTeamJoinRequest() {
  // Delegate to the UI.gs implementation
  if (typeof UI !== 'undefined' && typeof UI.showTeamJoinRequest === 'function') {
    UI.showTeamJoinRequest();
  } else {
    // Fallback to team manager if UI.gs not properly loaded
    showTeamManager(true); // Show team manager in join-only mode
  }
}

/**
 * Refreshes the menu after joining a team
 */
function refreshMenuAfterJoin() {
  try {
    const ui = SpreadsheetApp.getUi();
    const userEmail = Session.getActiveUser().getEmail();
    const teamsData = getTeamsData();
    const userTeam = getUserTeam(userEmail, teamsData);

    // Show debug info
    if (userTeam) {
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
    return false;
  }
}

/**
 * Force a reauthorization of the script with all required permissions
 */
function forceReauthorize() {
  try {
    // First try to get user info to verify authentication
    const userEmail = Session.getActiveUser().getEmail();
    
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
    
    
    // Show confirmation
    ui.alert(
      'Data Cleared',
      'All team data has been successfully deleted. Please reload the page to see the changes.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to clear team data: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
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
    SpreadsheetApp.getUi().alert('Error', 'Failed to show join team dialog: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * Shows the support ticket request form
 */
function showSupportTicket() {
  try {
    const html = HtmlService.createTemplateFromFile('SupportTicket');
    const htmlOutput = html.evaluate()
      .setWidth(550)
      .setHeight(650);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Submit Support Request');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open support form: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the user's support tickets
 */
function showMyTickets() {
  try {
    viewSupportTickets();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to load support tickets: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Submits a support ticket
 * @param {Object} ticketData - The ticket form data
 * @return {Object} Result with success status and ticket ID
 */
function submitSupportTicket(ticketData) {
  return SupportTicketService.createTicket(ticketData);
}

/**
 * Gets user information for form pre-population
 * @return {Object} User information
 */
function getUserInfo() {
  return SupportTicketService.getUserInfo();
}

/**
 * Views a specific support ticket
 * @param {string} ticketId - The ticket ID to view
 */
function viewSupportTicket(ticketId) {
  try {
    const html = HtmlService.createTemplateFromFile('SupportTicketView');
    html.ticketId = ticketId;
    
    const htmlOutput = html.evaluate()
      .setWidth(700)
      .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Support Ticket #${ticketId}`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open ticket: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Views all support tickets for the current user
 */
function viewSupportTickets() {
  try {
    const html = HtmlService.createTemplateFromFile('SupportTicketsList');
    
    const htmlOutput = html.evaluate()
      .setWidth(800)
      .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'My Support Tickets');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to load tickets: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets tickets for the current user
 * @return {Array} Array of user's tickets
 */
function getUserTickets() {
  return SupportTicketService.getUserTickets();
}

/**
 * Gets a specific ticket
 * @param {string} ticketId - The ticket ID
 * @return {Object|null} The ticket object or null
 */
function getTicket(ticketId) {
  return SupportTicketService.getTicket(ticketId);
}

/**
 * Adds a message to a ticket
 * @param {string} ticketId - The ticket ID
 * @param {string} message - The message content
 * @return {Object} Result with success status
 */
function addTicketMessage(ticketId, message) {
  return SupportTicketService.addMessage(ticketId, message, false);
}

/**
 * Gets all tickets (admin function)
 * @return {Array} Array of all tickets
 */
function getAllTickets() {
  return SupportTicketService.getAllTickets();
}

/**
 * Updates ticket status (admin function)
 * @param {string} ticketId - The ticket ID
 * @param {string} status - New status
 * @return {Object} Result with success status
 */
function updateTicketStatus(ticketId, status) {
  return SupportTicketService.updateTicketStatus(ticketId, status);
}

/**
 * Shows the admin support ticket interface
 * (Only accessible by admin users)
 */
function showSupportTicketAdmin() {
  try {
    // Check if user is admin
    if (!SupportTicketService.isAdminUser()) {
      SpreadsheetApp.getUi().alert('Access Denied', 'You do not have admin privileges.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const html = HtmlService.createTemplateFromFile('SupportTicketAdmin');
    const htmlOutput = html.evaluate()
      .setWidth(900)
      .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Support Ticket Admin Panel');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open admin panel: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
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
    throw e;
  }
}

var SettingsDialogUI = SettingsDialogUI || {};

/**
 * Shows the settings dialog with the columns tab active
 * This function is called from the menu
 */
function showColumnsTab() {
  try {
    SettingsDialogUI.showSettings('columns');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column settings: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the settings dialog with the settings tab active
 * This function is called from the menu
 */
function showSettingsTab() {
  try {
    SettingsDialogUI.showSettings('settings');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open filter settings: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the team manager dialog
 * This function is called from the menu and from other dialogs
 * @param {boolean} joinOnly - Whether to show only the join team section
 */
function showTeamManager(joinOnly) {
  try {
    TeamManagerUI.showTeamManager(joinOnly);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to open team manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}