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
      const userEmail = Session.getActiveUser().getEmail();
      const userId = Session.getTemporaryActiveUserKey();
      
      const response = UrlFetchApp.fetch(`${this.API_URL}/create-checkout-session`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        payload: JSON.stringify({
          email: userEmail,
          googleUserId: userId,
          scriptId: ScriptApp.getScriptId(),
          planType: planType, // 'pro_monthly', 'pro_annual', 'team_monthly', 'team_annual'
          successUrl: 'https://docs.google.com/spreadsheets', // Redirect after payment
          cancelUrl: 'https://docs.google.com/spreadsheets'
        })
      });
      
      const data = JSON.parse(response.getContentText());
      return data.checkoutUrl;
    } catch (error) {
      console.error('Error creating checkout session:', error);
      throw new Error('Failed to create payment session');
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
  }
};