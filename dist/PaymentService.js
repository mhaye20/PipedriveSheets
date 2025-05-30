/**
 * PaymentService - Handles Stripe recurring subscription integration and license validation
 * This runs in Google™ Apps Script environment
 * 
 * Features:
 * - Creates recurring subscription checkout sessions for monthly/annual plans
 * - Validates subscription status and feature access
 * - Manages customer portal access for subscription management
 * - Handles team subscription inheritance
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
      
      // Cache the result for 1 minute to detect subscription changes faster
      const cache = CacheService.getUserCache();
      cache.put('subscription_status', JSON.stringify(data), 60);
      
      // Ensure we return consistent subscription data
      return {
        plan: data.plan || 'free',
        status: data.status || 'active',
        expiresAt: data.expiresAt,
        features: data.features || [],
        cancelAt: data.cancelAt,
        canceledAt: data.canceledAt,
        message: data.message
      };
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
   * Create a Stripe checkout session for recurring subscription
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
      
      // Validate plan type for recurring subscriptions
      const validPlans = ['pro_monthly', 'pro_annual', 'team_monthly', 'team_annual'];
      if (!validPlans.includes(planType)) {
        throw new Error('Invalid subscription plan type');
      }
      
      // Get the current spreadsheet URL for proper redirect
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetUrl = spreadsheet.getUrl();
      
      const payload = {
        email: userEmail,
        googleUserId: userId,
        scriptId: ScriptApp.getScriptId(),
        planType: planType, // This will be used to select the correct Stripe Price ID
        successUrl: spreadsheetUrl + '?upgrade=success',
        cancelUrl: spreadsheetUrl + '?upgrade=cancelled'
      };
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/create-checkout-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        
        try {
          const errorData = JSON.parse(errorText);
          throw new Error(errorData.error || 'Failed to create subscription checkout session');
        } catch (e) {
          throw new Error('Failed to create subscription checkout session: ' + errorText);
        }
      }
      
      const data = JSON.parse(response.getContentText());
      
      if (!data.checkoutUrl) {
        throw new Error('No checkout URL returned from server');
      }
      
      return data.checkoutUrl;
    } catch (error) {
      console.error('Error creating recurring subscription checkout session:', error);
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
      
      if (userTeam) {
        const teamsData = getTeamsData();
        const team = teamsData[userTeam.teamId];
        
        if (team) {
          // Use createdBy if available, otherwise use first admin as team owner
          const teamOwnerEmail = team.createdBy || (team.adminEmails && team.adminEmails[0]);
          
          if (!teamOwnerEmail) {
            return false;
          }
          
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
                    }
                  }
                  
                  // Cache the full owner data for use in getCurrentPlan
                  cache.put(cacheKey, JSON.stringify(ownerData), 300);
                } else {
                }
              }
            } catch (error) {
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
    
    // Check if plan allows feature AND subscription is active
    if (!allowedPlans.includes(subscription.plan)) {
      return false;
    }
    
    // Check subscription status - must be active and not expired
    if (subscription.status !== 'active') {
      return false;
    }
    
    // Check if subscription is canceled and past due date
    if (subscription.cancelAt) {
      const cancelDate = new Date(subscription.cancelAt);
      const now = new Date();
      if (now >= cancelDate) {
        return false;
      }
    }
    
    return true;
  },
  
  /**
   * Get user's current plan details
   */
  getCurrentPlan() {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    // First, get the user's individual subscription status
    const individualSubscription = this.getSubscriptionStatus();
    
    // If user has their own Team subscription, return it directly
    if (individualSubscription.plan === 'team') {
      return {
        ...individualSubscription,
        details: {
          name: 'Team',
          limits: {
            rows: -1, // unlimited
            filters: -1,
            users: 5,
            columns: -1 // unlimited
          },
          features: ['two_way_sync', 'scheduled_sync', 'bulk_operations', 'team_features', 'shared_filters', 'admin_dashboard', 'priority_support']
        }
      };
    }
    
    // Check if user is a team member with inherited Team plan
    if (userEmail && isUserInTeam(userEmail)) {
      const userTeam = getUserTeam(userEmail);
      if (userTeam) {
        const teamsData = getTeamsData();
        const team = teamsData[userTeam.teamId];
        
        if (team) {
          // Use createdBy if available, otherwise use first admin as team owner
          const teamOwnerEmail = team.createdBy || (team.adminEmails && team.adminEmails[0]);
          
          if (!teamOwnerEmail) {
            // Fall back to individual subscription check
          } else {
            
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
                
                
                if (response.getResponseCode() === 200) {
                  ownerData = JSON.parse(response.getContentText());
                  
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
                    }
                  }
                  
                  // Cache the full owner data
                  cache.put(cacheKey, JSON.stringify(ownerData), 300);
                } else {
                }
              }
            } catch (error) {
            }
          }
          
          if (ownerHasTeamPlan) {
            // Check if current user is the team owner
            const normalizedOwnerEmail = teamOwnerEmail.toLowerCase();
            const currentUserEmail = userEmail.toLowerCase();
            const isTeamOwner = normalizedOwnerEmail === currentUserEmail;
            
            // Return team plan details with inherited status
            const result = {
              plan: 'team',
              status: 'active',
              isInherited: !isTeamOwner,  // Only inherited if user is NOT the team owner
              teamName: team.name,
              teamOwner: teamOwnerEmail,
              details: {
                name: 'Team',
                limits: {
                  rows: -1, // unlimited
                  filters: -1,
                  users: 5,
                  columns: -1 // unlimited
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
    
    // Use the individual subscription we already got at the beginning
    const subscription = individualSubscription;
    
    const planDetails = {
      'free': {
        name: 'Free',
        limits: {
          rows: 50,
          filters: 1,
          users: 1,
          columns: 5
        },
        features: ['manual_sync', 'basic_support']
      },
      'pro': {
        name: 'Pro',
        limits: {
          rows: -1, // unlimited
          filters: -1, // unlimited
          users: 1,
          columns: -1 // unlimited
        },
        features: ['two_way_sync', 'scheduled_sync', 'bulk_operations', 'priority_support']
      },
      'team': {
        name: 'Team',
        limits: {
          rows: -1, // unlimited
          filters: -1,
          users: 5,
          columns: -1 // unlimited
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
    
    // -1 means unlimited rows
    if (limit > 0 && currentRows > limit) {
      throw new Error(`Your ${plan.details.name} plan allows up to ${limit} rows. Please upgrade to sync more data.`);
    }
    
    return true;
  },

  /**
   * Enforce column limits based on plan
   */
  enforceColumnLimit(selectedColumns) {
    const plan = this.getCurrentPlan();
    const limit = plan.details.limits.columns;
    
    // -1 means unlimited columns
    if (limit > 0 && selectedColumns > limit) {
      throw new Error(`Your ${plan.details.name} plan allows up to ${limit} columns. Please upgrade to select more columns.`);
    }
    
    return true;
  },
  
  /**
   * Check if current user is a team admin
   */
  isTeamAdmin() {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    if (!userEmail || !isUserInTeam(userEmail)) {
      return false;
    }
    
    const userTeam = getUserTeam(userEmail);
    return userTeam && userTeam.role === 'Admin';
  },

  /**
   * Check if current user can modify settings (either not in team or is team admin)
   */
  canModifySettings() {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    if (!userEmail) {
      return false;
    }
    
    // If not in a team, user can modify their own settings
    if (!isUserInTeam(userEmail)) {
      return true;
    }
    
    // If in a team, only admins can modify settings
    return this.isTeamAdmin();
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
  },
  
  /**
   * Create a Stripe customer portal session for managing recurring subscription
   */
  createCustomerPortalSession() {
    try {
      const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      const userId = Session.getTemporaryActiveUserKey() || 'anonymous-' + Utilities.getUuid();
      const scriptId = ScriptApp.getScriptId();
      
      if (!userEmail) {
        throw new Error('Email address is required to access subscription management');
      }
      
      const payload = {
        email: userEmail,
        googleUserId: userId,
        scriptId: scriptId,
        returnUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl()
      };
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/create-portal-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 404) {
        throw new Error('Subscription management portal is not available. Please contact support.');
      }
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        try {
          const errorData = JSON.parse(errorText);
          const errorMessage = errorData.error || 'Failed to create portal session';
          
          if (errorMessage.includes('No active subscription')) {
            throw new Error('No active subscription found. Please ensure you have an active paid subscription.');
          }
          
          throw new Error(errorMessage);
        } catch (parseError) {
          throw new Error('Failed to access subscription management portal');
        }
      }
      
      const data = JSON.parse(response.getContentText());
      
      if (!data.portalUrl) {
        throw new Error('No portal URL returned from server');
      }
      
      return data.portalUrl;
      
    } catch (error) {
      console.error('Error creating customer portal session:', error);
      throw error;
    }
  }
};