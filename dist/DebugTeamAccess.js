/**
 * Debug Team Access Issues
 * 
 * This temporary debugging script helps diagnose team management access issues
 */

/**
 * Debug function to check all team access conditions
 * Call this from the Apps Script editor or add a menu item
 */
function debugTeamAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    console.log('=== TEAM ACCESS DEBUG ===');
    console.log('User Email:', userEmail);
    
    // 1. Check subscription status
    console.log('\n1. SUBSCRIPTION STATUS:');
    const plan = PaymentService.getCurrentPlan();
    console.log('- Plan object:', JSON.stringify(plan, null, 2));
    console.log('- plan.plan:', plan.plan);
    console.log('- plan.status:', plan.status);
    console.log('- plan.isInherited:', plan.isInherited);
    
    // 2. Check team membership
    console.log('\n2. TEAM MEMBERSHIP:');
    const hasTeam = isUserInTeam(userEmail);
    console.log('- isUserInTeam():', hasTeam);
    
    if (hasTeam) {
      const userTeam = getUserTeam(userEmail);
      console.log('- getUserTeam():', JSON.stringify(userTeam, null, 2));
    }
    
    // 3. Check script owner status
    console.log('\n3. SCRIPT OWNER STATUS:');
    const isOwner = isScriptOwner();
    console.log('- isScriptOwner():', isOwner);
    
    // 4. Check individual components of isScriptOwner
    console.log('\n4. SCRIPT OWNER BREAKDOWN:');
    
    // Test user check
    const isTestUser = TEST_USERS.includes(userEmail.toLowerCase());
    console.log('- Is test user:', isTestUser);
    
    if (isTestUser) {
      const userProperties = PropertiesService.getUserProperties();
      const isInitialized = userProperties.getProperty('PIPEDRIVE_INITIALIZED') === 'true';
      console.log('- Test user initialized:', isInitialized);
    }
    
    // Original installer check
    try {
      const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
      const isOriginalInstaller = authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.ENABLED;
      console.log('- Is original installer:', isOriginalInstaller);
      console.log('- Auth status:', authInfo.getAuthorizationStatus());
    } catch (e) {
      console.log('- Auth check error:', e.toString());
    }
    
    // Team subscription check
    try {
      const hasTeamSub = plan.plan === 'team' && !plan.isInherited && plan.status === 'active';
      console.log('- Has active Team subscription:', hasTeamSub);
      console.log('  - plan === "team":', plan.plan === 'team');
      console.log('  - !isInherited:', !plan.isInherited);
      console.log('  - status === "active":', plan.status === 'active');
    } catch (e) {
      console.log('- Team subscription check error:', e.toString());
    }
    
    // 5. Check what UI should be shown
    console.log('\n5. UI LOGIC:');
    console.log('- Should show create team tabs:', !hasTeam && isOwner);
    console.log('- Should show join section:', !hasTeam);
    console.log('- Should show create section within join:', isOwner);
    console.log('- Should show full team management:', hasTeam);
    
    // 6. Check properties
    console.log('\n6. USER PROPERTIES:');
    const userProperties = PropertiesService.getUserProperties();
    const allProps = userProperties.getProperties();
    console.log('- All user properties:');
    for (const [key, value] of Object.entries(allProps)) {
      console.log(`  ${key}: ${value}`);
    }
    
    console.log('\n=== END DEBUG ===');
    
    // Return summary for easy viewing
    return {
      userEmail: userEmail,
      plan: plan,
      hasTeam: hasTeam,
      isScriptOwner: isOwner,
      shouldShowCreateOptions: !hasTeam && isOwner
    };
    
  } catch (error) {
    console.error('Debug function error:', error);
    return { error: error.toString() };
  }
}

/**
 * Alternative debug function that can be called from menu
 */
function debugTeamAccessWithAlert() {
  const result = debugTeamAccess();
  const ui = SpreadsheetApp.getUi();
  ui.alert('Debug Results', JSON.stringify(result, null, 2), ui.ButtonSet.OK);
}

/**
 * Test the exact conditions used in TeamManager
 */
function testTeamManagerConditions() {
  try {
    console.log('=== TEAM MANAGER CONDITIONS TEST ===');
    
    const userEmail = Session.getActiveUser().getEmail();
    const plan = PaymentService.getCurrentPlan();
    const hasTeam = isUserInTeam(userEmail);
    const isOwner = isScriptOwner();
    
    console.log('Variables passed to template:');
    console.log('- userEmail:', userEmail);
    console.log('- plan:', JSON.stringify(plan, null, 2));
    console.log('- hasTeam:', hasTeam);
    console.log('- isScriptOwner:', isOwner);
    
    // Test the exact template conditions
    const showTabNavigation = !hasTeam && isOwner;
    const showJoinTab = !hasTeam;
    const showCreateTab = isOwner;
    const showTeamManagement = hasTeam;
    
    console.log('\nTemplate condition results:');
    console.log('- Show tab navigation (!hasTeam && isScriptOwner):', showTabNavigation);
    console.log('- Show join tab (!hasTeam):', showJoinTab);
    console.log('- Show create tab (isScriptOwner):', showCreateTab);
    console.log('- Show team management (hasTeam):', showTeamManagement);
    
    return {
      showTabNavigation,
      showJoinTab, 
      showCreateTab,
      showTeamManagement,
      variables: { userEmail, plan, hasTeam, isScriptOwner: isOwner }
    };
    
  } catch (error) {
    console.error('Test conditions error:', error);
    return { error: error.toString() };
  }
}