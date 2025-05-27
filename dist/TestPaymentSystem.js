/**
 * Test script for verifying payment system integration
 * Run these functions in Google Apps Script editor to test
 */

// Test 1: Check if backend is accessible
function testBackendConnection() {
  try {
    const response = UrlFetchApp.fetch('https://pipedrive-sheets.vercel.app/');
    const data = JSON.parse(response.getContentText());
    
    console.log('Backend Status:', data.status);
    console.log('Service:', data.service);
    console.log('Timestamp:', data.timestamp);
    
    if (data.status === 'ok') {
      console.log('✅ Backend is running successfully!');
    } else {
      console.log('❌ Backend returned unexpected status');
    }
  } catch (error) {
    console.error('❌ Failed to connect to backend:', error.message);
    console.log('Make sure you have deployed to Vercel with the latest changes');
  }
}

// Test 2: Check subscription status
function testSubscriptionStatus() {
  try {
    const status = PaymentService.getSubscriptionStatus();
    console.log('Subscription Status:', JSON.stringify(status, null, 2));
    
    if (status.plan) {
      console.log('✅ Successfully retrieved subscription status');
      console.log('Current plan:', status.plan);
      console.log('Status:', status.status);
    } else {
      console.log('❌ Failed to get subscription status');
    }
  } catch (error) {
    console.error('❌ Error getting subscription status:', error.message);
  }
}

// Test 3: Check current plan details
function testGetCurrentPlan() {
  try {
    const plan = PaymentService.getCurrentPlan();
    console.log('Current Plan:', JSON.stringify(plan, null, 2));
    
    console.log('\nPlan Summary:');
    console.log('- Name:', plan.details.name);
    console.log('- Row Limit:', plan.details.limits.rows);
    console.log('- Filter Limit:', plan.details.limits.filters === -1 ? 'Unlimited' : plan.details.limits.filters);
    console.log('- Users:', plan.details.limits.users);
    console.log('- Features:', plan.details.features.join(', '));
    
    console.log('✅ Successfully retrieved plan details');
  } catch (error) {
    console.error('❌ Error getting plan details:', error.message);
  }
}

// Test 4: Check feature access
function testFeatureAccess() {
  const features = [
    'two_way_sync',
    'scheduled_sync',
    'unlimited_rows',
    'team_features',
    'shared_filters',
    'bulk_operations'
  ];
  
  console.log('Feature Access Check:');
  console.log('-------------------');
  
  features.forEach(feature => {
    try {
      const hasAccess = PaymentService.hasFeatureAccess(feature);
      console.log(`${feature}: ${hasAccess ? '✅ Enabled' : '❌ Disabled'}`);
    } catch (error) {
      console.log(`${feature}: ❌ Error - ${error.message}`);
    }
  });
}

// Test 5: Test row limit enforcement
function testRowLimits() {
  const testCases = [
    { rows: 10, desc: 'Small dataset (10 rows)' },
    { rows: 50, desc: 'Free plan limit (50 rows)' },
    { rows: 100, desc: 'Above free limit (100 rows)' },
    { rows: 5000, desc: 'Pro plan limit (5000 rows)' },
    { rows: 10000, desc: 'Above pro limit (10000 rows)' }
  ];
  
  console.log('Row Limit Tests:');
  console.log('----------------');
  
  testCases.forEach(test => {
    try {
      PaymentService.enforceRowLimit(test.rows);
      console.log(`✅ ${test.desc} - Allowed`);
    } catch (error) {
      console.log(`❌ ${test.desc} - Blocked: ${error.message}`);
    }
  });
}

// Test 6: Create test checkout session
function testCreateCheckout() {
  try {
    console.log('Creating checkout session for Pro Monthly plan...');
    
    const checkoutUrl = PaymentService.createCheckoutSession('pro_monthly');
    
    if (checkoutUrl) {
      console.log('✅ Checkout session created successfully!');
      console.log('Checkout URL:', checkoutUrl);
      console.log('\nTo test payment:');
      console.log('1. Open the URL in your browser');
      console.log('2. Use test card: 4242 4242 4242 4242');
      console.log('3. Any future expiry date and any CVC');
    } else {
      console.log('❌ Failed to create checkout session');
    }
  } catch (error) {
    console.error('❌ Error creating checkout:', error.message);
    console.log('\nPossible issues:');
    console.log('- Backend not deployed or accessible');
    console.log('- Stripe environment variables not set');
    console.log('- Price IDs not configured correctly');
  }
}

// Run all tests
function runAllPaymentTests() {
  console.log('=== PAYMENT SYSTEM TEST SUITE ===\n');
  
  console.log('1. Testing Backend Connection...');
  testBackendConnection();
  
  console.log('\n2. Testing Subscription Status...');
  testSubscriptionStatus();
  
  console.log('\n3. Testing Current Plan...');
  testGetCurrentPlan();
  
  console.log('\n4. Testing Feature Access...');
  testFeatureAccess();
  
  console.log('\n5. Testing Row Limits...');
  testRowLimits();
  
  console.log('\n=== TEST SUITE COMPLETE ===');
  console.log('\nTo test actual payment flow, run testCreateCheckout() separately');
}

// Utility function to clear cache
function clearPaymentCache() {
  const cache = CacheService.getUserCache();
  cache.remove('subscription_status');
  console.log('✅ Payment cache cleared');
}