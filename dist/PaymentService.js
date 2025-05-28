/**
 * PaymentService - Handles Stripe integration and license validation
 * This runs in Google Apps Script environment
 */

const PaymentService = {
  // Your backend API URL (update after deployment)
  API_URL: 'https://pipedrive-sheets.vercel.app/api',
  
  /**
   * Get current user's subscription status
   */
  getSubscriptionStatus() {
    try {
      const userEmail = Session.getActiveUser().getEmail();
      const userId = Session.getTemporaryActiveUserKey(); // Unique per user per script
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/subscription/status`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify({
          email: userEmail,
          googleUserId: userId,
          scriptId: ScriptApp.getScriptId()
        })
      });
      
      const data = JSON.parse(response.getContentText());
      
      // Cache the result for 1 minute (reduced from 5) to detect cancellations faster
      const cache = CacheService.getUserCache();
      cache.put('subscription_status', JSON.stringify(data), 60);
      
      return data;
    } catch (error) {
      console.error('Error fetching subscription status:', error);
      return {
        plan: 'free',
        status: 'active',
        error: error.toString()
      };
    }
  },
  
  /**
   * Create a Stripe checkout session for upgrading
   */
  createCheckoutSession(planType) {
    try {
      let userEmail = Session.getActiveUser().getEmail();
      const userId = Session.getTemporaryActiveUserKey() || 'anonymous-' + Utilities.getUuid();
      
      // If email is not available, try to get it from effective user
      if (!userEmail) {
        userEmail = Session.getEffectiveUser().getEmail();
      }
      
      // If still no email, prompt user
      if (!userEmail) {
        const ui = SpreadsheetApp.getUi();
        const response = ui.prompt(
          'Email Required',
          'Please enter your email address to continue with the upgrade:',
          ui.ButtonSet.OK_CANCEL
        );
        
        if (response.getSelectedButton() === ui.Button.OK) {
          userEmail = response.getResponseText();
          
          // Basic email validation
          if (!userEmail || !userEmail.includes('@')) {
            throw new Error('Please enter a valid email address');
          }
        } else {
          throw new Error('Email is required to process payment');
        }
      }
      
      Logger.log('Creating checkout session:');
      Logger.log('Email: ' + userEmail);
      Logger.log('User ID: ' + userId);
      Logger.log('Script ID: ' + ScriptApp.getScriptId());
      Logger.log('Plan Type: ' + planType);
      Logger.log('API URL: ' + this.API_URL);
      
      // Get the current spreadsheet URL for proper redirect
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetUrl = spreadsheet.getUrl();
      
      const payload = {
        email: userEmail,
        googleUserId: userId,
        scriptId: ScriptApp.getScriptId(),
        planType: planType, // 'pro_monthly', 'pro_annual', 'team_monthly', 'team_annual'
        successUrl: spreadsheetUrl + '?upgrade=success', // Redirect back to this spreadsheet
        cancelUrl: spreadsheetUrl + '?upgrade=cancelled'
      };
      
      Logger.log('Payload: ' + JSON.stringify(payload));
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/create-checkout-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true // This will return the error response instead of throwing
      });
      
      Logger.log('Response code: ' + response.getResponseCode());
      Logger.log('Response text: ' + response.getContentText());
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        Logger.log('Error response: ' + errorText);
        
        try {
          const errorData = JSON.parse(errorText);
          throw new Error(errorData.error || 'Failed to create payment session');
        } catch (e) {
          throw new Error('Failed to create payment session: ' + errorText);
        }
      }
      
      const data = JSON.parse(response.getContentText());
      
      if (!data.checkoutUrl) {
        Logger.log('No checkout URL in response: ' + JSON.stringify(data));
        throw new Error('No checkout URL returned from server');
      }
      
      return data.checkoutUrl;
    } catch (error) {
      Logger.log('Error in createCheckoutSession: ' + error.toString());
      Logger.log('Error stack: ' + error.stack);
      console.error('Error creating checkout session:', error);
      throw error;
    }
  },
  
  /**
   * Check if user has access to a specific feature
   */
  hasFeatureAccess(feature) {
    // First check if user is part of a team
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    if (userEmail && isUserInTeam(userEmail)) {
      // User is in a team - check if the team owner has a Team plan
      const userTeam = getUserTeam(userEmail);
      Logger.log('User team data: ' + JSON.stringify(userTeam));
      
      if (userTeam) {
        const teamsData = getTeamsData();
        const team = teamsData[userTeam.teamId];
        Logger.log('Team data: ' + JSON.stringify(team));
        
        if (team) {
          // Use createdBy if available, otherwise use first admin as team owner
          const teamOwnerEmail = team.createdBy || (team.adminEmails && team.adminEmails[0]);
          
          if (!teamOwnerEmail) {
            Logger.log('[hasFeatureAccess] No team owner found for team: ' + userTeam.teamId);
            return false;
          }
          
          Logger.log('Checking team owner subscription for: ' + teamOwnerEmail);
          // Check cache first for team owner's subscription
          const cache = CacheService.getUserCache();
          const cacheKey = 'team_owner_sub_' + teamOwnerEmail.toLowerCase();
          const cachedOwnerStatus = cache.get(cacheKey);
          
          let ownerHasTeamPlan = false;
          
          if (cachedOwnerStatus) {
            try {
              const cachedData = JSON.parse(cachedOwnerStatus);
              if (cachedData.plan === 'team') {
                if (cachedData.status === 'active' && !cachedData.cancelAt) {
                  ownerHasTeamPlan = true;
                } else if (cachedData.cancelAt) {
                  const cancelDate = new Date(cachedData.cancelAt);
                  const now = new Date();
                  ownerHasTeamPlan = now < cancelDate;
                }
              }
            } catch (e) {
              // Old cache format, just check if it's 'team'
              ownerHasTeamPlan = cachedOwnerStatus === 'team';
            }
          } else {
            // Check the subscription status of the team creator/owner
            try {
              // Normalize the team owner email for consistent checking
              const normalizedOwnerEmail = teamOwnerEmail.toLowerCase();
              const currentUserEmail = userEmail.toLowerCase();
              
              // Special case: if current user IS the team owner, use their own subscription status
              if (normalizedOwnerEmail === currentUserEmail) {
                Logger.log('Current user is the team owner, using their subscription status');
                const currentUserStatus = this.getSubscriptionStatus();
                
                if (currentUserStatus.plan === 'team') {
                  if (currentUserStatus.status === 'active' && !currentUserStatus.cancelAt) {
                    ownerHasTeamPlan = true;
                  } else if (currentUserStatus.cancelAt) {
                    const cancelDate = new Date(currentUserStatus.cancelAt);
                    const now = new Date();
                    ownerHasTeamPlan = now < cancelDate;
                  }
                }
                
                // Cache it
                cache.put(cacheKey, JSON.stringify(currentUserStatus), 300);
                ownerData = currentUserStatus;
              } else {
                // For other team members, we need to check the owner's subscription
                // This is a limitation - we can't get other users' Google User IDs
                // The backend should handle email-only lookups or store the mapping
                
                const response = UrlFetchApp.fetch(`${this.API_URL}/subscription/status`, {
                  method: 'POST',
                  headers: {
                    'Content-Type': 'application/json',
                  },
                  payload: JSON.stringify({
                    email: normalizedOwnerEmail,
                    googleUserId: '', // We can't get other users' IDs
                    scriptId: ScriptApp.getScriptId()
                  }),
                  muteHttpExceptions: true
                });
                
                if (response.getResponseCode() === 200) {
                  const ownerData = JSON.parse(response.getContentText());
                  Logger.log('Team owner subscription data: ' + JSON.stringify(ownerData));
                  
                  // Check if owner has team plan
                  if (ownerData.plan === 'team') {
                    if (ownerData.status === 'active' && !ownerData.cancelAt) {
                      // Active subscription without cancellation
                      ownerHasTeamPlan = true;
                    } else if (ownerData.cancelAt) {
                      // Subscription is canceled but check if still within valid period
                      const cancelDate = new Date(ownerData.cancelAt);
                      const now = new Date();
                      ownerHasTeamPlan = now < cancelDate;
                      Logger.log('Team owner subscription canceled, valid until: ' + ownerData.cancelAt + ', still valid: ' + ownerHasTeamPlan);
                    }
                  }
                  
                  // Cache the full owner data for use in getCurrentPlan
                  cache.put(cacheKey, JSON.stringify(ownerData), 300);
                } else {
                  Logger.log('Failed to get team owner subscription. Response code: ' + response.getResponseCode());
                  Logger.log('Response: ' + response.getContentText());
                }
              }
            } catch (error) {
              Logger.log('Error checking team owner subscription: ' + error.message);
            }
          }
          
          if (ownerHasTeamPlan) {
            // Team owner has active Team plan, grant access to all team features
            return true;
          }
        }
      }
    }
    
    // Fall back to checking individual subscription
    const cache = CacheService.getUserCache();
    const cachedStatus = cache.get('subscription_status');
    
    let subscription;
    if (cachedStatus) {
      subscription = JSON.parse(cachedStatus);
    } else {
      subscription = this.getSubscriptionStatus();
    }
    
    const featureMatrix = {
      'two_way_sync': ['pro', 'team'],
      'scheduled_sync': ['pro', 'team'],
      'unlimited_rows': ['pro', 'team'],
      'team_features': ['team'],
      'shared_filters': ['team'],
      'bulk_operations': ['pro', 'team']
    };
    
    const allowedPlans = featureMatrix[feature] || [];
    return allowedPlans.includes(subscription.plan);
  },
  
  /**
   * Get user's current plan details
   */
  getCurrentPlan() {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    // Check if user is a team member with inherited Team plan
    if (userEmail && isUserInTeam(userEmail)) {
      Logger.log('[getCurrentPlan] User is in team, checking for inherited access');
      const userTeam = getUserTeam(userEmail);
      if (userTeam) {
        Logger.log('[getCurrentPlan] User team: ' + JSON.stringify(userTeam));
        const teamsData = getTeamsData();
        const team = teamsData[userTeam.teamId];
        
        if (team) {
          // Use createdBy if available, otherwise use first admin as team owner
          const teamOwnerEmail = team.createdBy || (team.adminEmails && team.adminEmails[0]);
          
          if (!teamOwnerEmail) {
            Logger.log('[getCurrentPlan] No team owner found for team: ' + userTeam.teamId);
            // Fall back to individual subscription check
          } else {
            Logger.log('[getCurrentPlan] Team owner email: ' + teamOwnerEmail);
            
            // Check if team owner has active Team plan
            const cache = CacheService.getUserCache();
            const cacheKey = 'team_owner_sub_' + teamOwnerEmail.toLowerCase();
            const cachedOwnerStatus = cache.get(cacheKey);
          
          let ownerHasTeamPlan = false;
          let ownerData = null;
          
          if (cachedOwnerStatus) {
            try {
              ownerData = JSON.parse(cachedOwnerStatus);
              if (ownerData.plan === 'team') {
                if (ownerData.status === 'active' && !ownerData.cancelAt) {
                  ownerHasTeamPlan = true;
                } else if (ownerData.cancelAt) {
                  const cancelDate = new Date(ownerData.cancelAt);
                  const now = new Date();
                  ownerHasTeamPlan = now < cancelDate;
                }
              }
            } catch (e) {
              // Old cache format, just check if it's 'team'
              ownerHasTeamPlan = cachedOwnerStatus === 'team';
            }
          } else {
            try {
              const normalizedOwnerEmail = teamOwnerEmail.toLowerCase();
              const currentUserEmail = userEmail.toLowerCase();
              
              // Special case: if current user IS the team owner, use their own subscription status
              if (normalizedOwnerEmail === currentUserEmail) {
                Logger.log('[getCurrentPlan] Current user is the team owner, using their subscription status');
                const currentUserStatus = this.getSubscriptionStatus();
                ownerData = currentUserStatus;
                
                if (currentUserStatus.plan === 'team') {
                  if (currentUserStatus.status === 'active' && !currentUserStatus.cancelAt) {
                    ownerHasTeamPlan = true;
                  } else if (currentUserStatus.cancelAt) {
                    const cancelDate = new Date(currentUserStatus.cancelAt);
                    const now = new Date();
                    ownerHasTeamPlan = now < cancelDate;
                  }
                }
                
                // Cache it
                cache.put(cacheKey, JSON.stringify(currentUserStatus), 300);
              } else {
                // For other team members, check the owner's subscription
                const response = UrlFetchApp.fetch(`${this.API_URL}/subscription/status`, {
                  method: 'POST',
                  headers: {
                    'Content-Type': 'application/json',
                  },
                  payload: JSON.stringify({
                    email: normalizedOwnerEmail,
                    googleUserId: '', // We can't get other users' IDs
                    scriptId: ScriptApp.getScriptId()
                  }),
                  muteHttpExceptions: true
                });
                
                Logger.log('[getCurrentPlan] API Response code: ' + response.getResponseCode());
                
                if (response.getResponseCode() === 200) {
                  ownerData = JSON.parse(response.getContentText());
                  Logger.log('[getCurrentPlan] Owner subscription data: ' + JSON.stringify(ownerData));
                  
                  // Check if owner has team plan
                  if (ownerData.plan === 'team') {
                    if (ownerData.status === 'active' && !ownerData.cancelAt) {
                      // Active subscription without cancellation
                      ownerHasTeamPlan = true;
                    } else if (ownerData.cancelAt) {
                      // Subscription is canceled but check if still within valid period
                      const cancelDate = new Date(ownerData.cancelAt);
                      const now = new Date();
                      ownerHasTeamPlan = now < cancelDate;
                      Logger.log('[getCurrentPlan] Team owner subscription canceled, valid until: ' + ownerData.cancelAt + ', still valid: ' + ownerHasTeamPlan);
                    }
                  }
                  
                  // Cache the full owner data
                  cache.put(cacheKey, JSON.stringify(ownerData), 300);
                } else {
                  Logger.log('[getCurrentPlan] Failed to get owner subscription: ' + response.getContentText());
                }
              }
            } catch (error) {
              Logger.log('Error checking team owner subscription: ' + error.message);
            }
          }
          
          if (ownerHasTeamPlan) {
            // Return team plan details with inherited status
            const result = {
              plan: 'team',
              status: 'active',
              isInherited: true,
              teamName: team.name,
              teamOwner: teamOwnerEmail,
              details: {
                name: 'Team',
                limits: {
                  rows: 5000,
                  filters: -1,
                  users: 5
                },
                features: ['two_way_sync', 'scheduled_sync', 'bulk_operations', 'team_features', 'shared_filters', 'admin_dashboard', 'priority_support']
              }
            };
            
            // If the team owner's subscription is canceled, include that info
            if (ownerData && ownerData.cancelAt) {
              result.cancelAt = ownerData.cancelAt;
              result.teamOwnerCanceled = true;
            }
            
            return result;
          }
          }
        }
      }
    }
    
    // Check if user is in a team even if they don't have team features
    // This is for the case where team owner doesn't have Team plan
    let teamInfo = null;
    if (userEmail && isUserInTeam(userEmail)) {
      const userTeam = getUserTeam(userEmail);
      if (userTeam) {
        const teamsData = getTeamsData();
        const team = teamsData[userTeam.teamId];
        if (team) {
          teamInfo = {
            teamName: team.name,
            teamOwner: team.createdBy || (team.adminEmails && team.adminEmails[0])
          };
        }
      }
    }
    
    // Get individual subscription status
    const subscription = this.getSubscriptionStatus();
    
    const planDetails = {
      'free': {
        name: 'Free',
        limits: {
          rows: 50,
          filters: 1,
          users: 1
        },
        features: ['manual_sync', 'basic_support']
      },
      'pro': {
        name: 'Pro',
        limits: {
          rows: 5000,
          filters: -1, // unlimited
          users: 1
        },
        features: ['two_way_sync', 'scheduled_sync', 'bulk_operations', 'priority_support']
      },
      'team': {
        name: 'Team',
        limits: {
          rows: 5000,
          filters: -1,
          users: 5
        },
        features: ['two_way_sync', 'scheduled_sync', 'bulk_operations', 'team_features', 'shared_filters', 'admin_dashboard', 'priority_support']
      }
    };
    
    const result = {
      ...subscription,
      details: planDetails[subscription.plan] || planDetails.free,
      cancelAt: subscription.cancelAt,
      canceledAt: subscription.canceledAt
    };
    
    // Add team info if user is in a team but doesn't have inherited access
    if (teamInfo && !result.isInherited) {
      result.teamName = teamInfo.teamName;
      result.teamOwner = teamInfo.teamOwner;
    }
    
    return result;
  },
  
  /**
   * Enforce row limits based on plan
   */
  enforceRowLimit(currentRows) {
    const plan = this.getCurrentPlan();
    const limit = plan.details.limits.rows;
    
    if (limit > 0 && currentRows > limit) {
      throw new Error(`Your ${plan.details.name} plan allows up to ${limit} rows. Please upgrade to sync more data.`);
    }
    
    return true;
  },
  
  /**
   * Show upgrade dialog
   */
  showUpgradeDialog() {
    const template = HtmlService.createTemplateFromFile('PaymentDialog');
    template.currentPlan = this.getCurrentPlan();
    
    const html = template.evaluate()
      .setWidth(500)
      .setHeight(600);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Upgrade Your Plan');
  },
  
  /**
   * Show manage subscription dialog
   */
  showManageSubscriptionDialog() {
    // Clear cache to ensure we show the latest subscription status
    this.clearSubscriptionCache();
    
    const template = HtmlService.createTemplateFromFile('ManageSubscription');
    template.currentPlan = this.getCurrentPlan();
    template.userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    const html = template.evaluate()
      .setWidth(450)
      .setHeight(400);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Subscription');
  },
  
  /**
   * Clear the subscription cache
   */
  clearSubscriptionCache() {
    const cache = CacheService.getUserCache();
    cache.remove('subscription_status');
    Logger.log('Subscription cache cleared');
  },
  
  /**
   * Create a Stripe customer portal session for managing subscription
   */
  createCustomerPortalSession() {
    try {
      const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      const userId = Session.getTemporaryActiveUserKey() || 'anonymous-' + Utilities.getUuid();
      const scriptId = ScriptApp.getScriptId();
      
      Logger.log('Creating customer portal session:');
      Logger.log('Email: ' + userEmail);
      Logger.log('User ID: ' + userId);
      Logger.log('Script ID: ' + scriptId);
      Logger.log('API URL: ' + this.API_URL);
      
      const payload = {
        email: userEmail,
        googleUserId: userId,
        scriptId: scriptId,
        returnUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl()
      };
      
      Logger.log('Portal payload: ' + JSON.stringify(payload));
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/create-portal-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      Logger.log('Portal response code: ' + response.getResponseCode());
      Logger.log('Portal response text: ' + response.getContentText());
      
      if (response.getResponseCode() === 404) {
        Logger.log('404 Error - endpoint not found. Backend may need redeployment.');
        throw new Error('Portal endpoint not available. Please try again later.');
      }
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        try {
          const errorData = JSON.parse(errorText);
          const errorMessage = errorData.error || 'Failed to create portal session';
          Logger.log('Portal error: ' + errorMessage);
          
          if (errorMessage.includes('No active subscription')) {
            throw new Error('No active subscription found. Please ensure you have completed payment.');
          }
          
          throw new Error(errorMessage);
        } catch (parseError) {
          Logger.log('Error parsing response: ' + parseError);
          throw new Error('Failed to create portal session');
        }
      }
      
      const data = JSON.parse(response.getContentText());
      
      if (!data.portalUrl) {
        Logger.log('No portal URL in response: ' + JSON.stringify(data));
        throw new Error('No portal URL returned from server');
      }
      
      return data.portalUrl;
      
    } catch (error) {
      Logger.log('Error in createCustomerPortalSession: ' + error.toString());
      Logger.log('Error stack: ' + error.stack);
      throw error;
    }
  }
};