/**
 * Backend server for handling Stripe payments and license validation
 * Deploy this to Vercel, Railway, or Render
 */

const express = require('express');
const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);
const cors = require('cors');
const { createClient } = require('@supabase/supabase-js');

const app = express();

// Initialize Supabase for license storage (or use MongoDB/PostgreSQL)
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

// Middleware
app.use(cors());
app.use(express.json({ 
  verify: (req, res, buf) => {
    req.rawBody = buf.toString('utf-8');
  }
}));

// Health check endpoint
app.get('/', (req, res) => {
  res.json({ 
    status: 'ok', 
    service: 'PipedriveSheets Payment Backend',
    timestamp: new Date().toISOString()
  });
});

// Stripe webhook endpoint
app.post('/webhook', async (req, res) => {
  const sig = req.headers['stripe-signature'];
  let event;

  try {
    event = stripe.webhooks.constructEvent(
      req.rawBody,
      sig,
      process.env.STRIPE_WEBHOOK_SECRET
    );
  } catch (err) {
    console.error('Webhook signature verification failed:', err);
    return res.status(400).send(`Webhook Error: ${err.message}`);
  }

  // Handle the event
  switch (event.type) {
    case 'checkout.session.completed':
      const session = event.data.object;
      await handleSuccessfulPayment(session);
      break;
      
    case 'customer.subscription.updated':
    case 'customer.subscription.deleted':
      const subscription = event.data.object;
      await handleSubscriptionChange(subscription);
      break;
      
    default:
      console.log(`Unhandled event type ${event.type}`);
  }

  res.json({ received: true });
});

// Create Stripe checkout session
app.post('/api/create-checkout-session', async (req, res) => {
  try {
    const { email, googleUserId, scriptId, planType, successUrl, cancelUrl } = req.body;
    
    // Validate required fields
    if (!email || !email.includes('@')) {
      return res.status(400).json({ error: 'Valid email address is required' });
    }
    
    if (!planType || !['pro_monthly', 'pro_annual', 'team_monthly', 'team_annual'].includes(planType)) {
      return res.status(400).json({ error: 'Invalid plan type' });
    }
    
    // Define price IDs (create these in Stripe Dashboard)
    const priceIds = {
      'pro_monthly': process.env.STRIPE_PRO_MONTHLY_PRICE_ID,
      'pro_annual': process.env.STRIPE_PRO_ANNUAL_PRICE_ID,
      'team_monthly': process.env.STRIPE_TEAM_MONTHLY_PRICE_ID,
      'team_annual': process.env.STRIPE_TEAM_ANNUAL_PRICE_ID
    };
    
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [{
        price: priceIds[planType],
        quantity: 1,
      }],
      mode: planType.includes('annual') ? 'subscription' : 'subscription',
      success_url: successUrl,
      cancel_url: cancelUrl,
      customer_email: email,
      metadata: {
        googleUserId,
        scriptId,
        planType
      },
      subscription_data: {
        metadata: {
          googleUserId,
          scriptId,
          planType
        }
      }
    });
    
    res.json({ checkoutUrl: session.url });
  } catch (error) {
    console.error('Error creating checkout session:', error);
    res.status(500).json({ error: error.message });
  }
});

// Check subscription status
app.post('/api/subscription/status', async (req, res) => {
  try {
    const { email, googleUserId, scriptId } = req.body;
    
    // Look up user's subscription in database
    const { data: subscription, error } = await supabase
      .from('subscriptions')
      .select('*')
      .eq('google_user_id', googleUserId)
      .eq('script_id', scriptId)
      .single();
    
    if (error || !subscription) {
      return res.json({
        plan: 'free',
        status: 'active',
        features: []
      });
    }
    
    // Verify subscription is still active with Stripe
    if (subscription.stripe_subscription_id) {
      const stripeSubscription = await stripe.subscriptions.retrieve(
        subscription.stripe_subscription_id
      );
      
      if (stripeSubscription.status !== 'active') {
        // Update database
        await supabase
          .from('subscriptions')
          .update({ status: stripeSubscription.status })
          .eq('id', subscription.id);
          
        if (stripeSubscription.status === 'canceled' || stripeSubscription.status === 'past_due') {
          return res.json({
            plan: 'free',
            status: 'inactive',
            features: []
          });
        }
      }
    }
    
    res.json({
      plan: subscription.plan_type,
      status: subscription.status,
      expiresAt: subscription.expires_at,
      features: getPlanFeatures(subscription.plan_type)
    });
    
  } catch (error) {
    console.error('Error checking subscription status:', error);
    res.status(500).json({ error: error.message });
  }
});

// Helper function to handle successful payments
async function handleSuccessfulPayment(session) {
  const { googleUserId, scriptId, planType } = session.metadata;
  
  // Extract plan name from planType (e.g., 'pro_monthly' -> 'pro')
  const plan = planType.split('_')[0];
  
  // Store or update subscription in database
  const { error } = await supabase
    .from('subscriptions')
    .upsert({
      google_user_id: googleUserId,
      script_id: scriptId,
      stripe_customer_id: session.customer,
      stripe_subscription_id: session.subscription,
      plan_type: plan,
      status: 'active',
      email: session.customer_email,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    }, {
      onConflict: 'google_user_id,script_id'
    });
    
  if (error) {
    console.error('Error storing subscription:', error);
  }
}

// Helper function to handle subscription changes
async function handleSubscriptionChange(subscription) {
  const { googleUserId, scriptId, planType } = subscription.metadata;
  
  const status = subscription.status === 'active' ? 'active' : 'inactive';
  
  await supabase
    .from('subscriptions')
    .update({
      status,
      updated_at: new Date().toISOString()
    })
    .eq('stripe_subscription_id', subscription.id);
}

// Helper function to get plan features
function getPlanFeatures(plan) {
  const features = {
    'pro': ['two_way_sync', 'scheduled_sync', 'unlimited_rows', 'bulk_operations'],
    'team': ['two_way_sync', 'scheduled_sync', 'unlimited_rows', 'bulk_operations', 'team_features', 'shared_filters']
  };
  
  return features[plan] || [];
}

// Create Stripe customer portal session
app.post('/api/create-portal-session', async (req, res) => {
  try {
    const { email, googleUserId, scriptId, returnUrl } = req.body;
    
    // Validate required fields
    if (!email || !email.includes('@')) {
      return res.status(400).json({ error: 'Valid email address is required' });
    }
    
    // Look up customer in database
    const { data: subscription, error } = await supabase
      .from('subscriptions')
      .select('stripe_customer_id')
      .eq('google_user_id', googleUserId)
      .eq('script_id', scriptId)
      .single();
    
    if (error || !subscription || !subscription.stripe_customer_id) {
      return res.status(404).json({ error: 'No active subscription found' });
    }
    
    // Create portal session
    const portalSession = await stripe.billingPortal.sessions.create({
      customer: subscription.stripe_customer_id,
      return_url: returnUrl || 'https://docs.google.com/spreadsheets'
    });
    
    res.json({ portalUrl: portalSession.url });
  } catch (error) {
    console.error('Error creating portal session:', error);
    res.status(500).json({ error: error.message });
  }
});

// For Vercel deployment
if (process.env.VERCEL) {
  module.exports = app;
} else {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
}