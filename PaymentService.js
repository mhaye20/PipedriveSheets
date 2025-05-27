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
      
      // Cache the result for 5 minutes to reduce API calls
      const cache = CacheService.getUserCache();
      cache.put('subscription_status', JSON.stringify(data), 300);
      
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
    // First check cache
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
    
    return {
      ...subscription,
      details: planDetails[subscription.plan] || planDetails.free
    };
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
    const template = HtmlService.createTemplateFromFile('ManageSubscription');
    template.currentPlan = this.getCurrentPlan();
    template.userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    const html = template.evaluate()
      .setWidth(450)
      .setHeight(400);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Subscription');
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