# PipedriveSheets Payment Integration Guide

## Overview

This guide explains how to implement the Stripe payment system for your PipedriveSheets Google Workspace Add-on. The integration is designed to be modern, secure, and provide a seamless upgrade experience for users.

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Google Sheets  │────▶│  Backend Server  │────▶│     Stripe      │
│    Add-on       │◀────│  (Node.js/Express)│◀────│   Payments      │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                        │
        ▼                        ▼
┌─────────────────┐     ┌──────────────────┐
│ PaymentService  │     │    Supabase      │
│      .js        │     │   (Database)     │
└─────────────────┘     └──────────────────┘
```

## Implementation Steps

### 1. Set Up Stripe (15 minutes)

1. Create a Stripe account at [stripe.com](https://stripe.com)
2. Create products in Stripe Dashboard:
   - **Pro Monthly**: $12.99/month
   - **Pro Annual**: $129/year
   - **Team Monthly**: $39.99/month
   - **Team Annual**: $399/year
3. Note down the price IDs for each product

### 2. Deploy Backend (30 minutes)

1. **Set up Supabase database**:
   ```bash
   # Go to supabase.com and create a free project
   # Run the SQL from backend/README.md in the SQL editor
   ```

2. **Deploy to Vercel**:
   ```bash
   cd backend
   npm install
   vercel
   ```

3. **Configure environment variables** in Vercel dashboard:
   - All Stripe keys and price IDs
   - Supabase URL and key

4. **Set up Stripe webhook** in Stripe Dashboard:
   - Endpoint: `https://your-app.vercel.app/webhook`
   - Events: `checkout.session.completed`, `customer.subscription.*`

### 3. Update Google Apps Script (10 minutes)

1. **Update PaymentService.js**:
   ```javascript
   const PaymentService = {
     API_URL: 'https://your-backend.vercel.app/api', // Your deployed URL
     ...
   }
   ```

2. **Deploy to Google Apps Script**:
   ```bash
   npm run deploy
   ```

## How It Works

### User Flow

1. **Free User Tries Premium Feature**:
   - User clicks "Two-Way Sync" or "Schedule Sync"
   - System checks subscription status
   - Shows upgrade prompt with plan comparison

2. **Upgrade Process**:
   - User clicks "Upgrade" in menu or feature prompt
   - Beautiful upgrade dialog shows current plan and options
   - User selects monthly/annual billing
   - Redirected to Stripe Checkout (secure payment)
   - After payment, returns to Google Sheets

3. **License Validation**:
   - Each feature check calls `PaymentService.hasFeatureAccess()`
   - Results cached for 5 minutes to reduce API calls
   - Subscription status verified with backend

### Feature Gates

The following features are gated by subscription:

| Feature | Free | Pro | Team |
|---------|------|-----|------|
| Manual Sync | ✓ | ✓ | ✓ |
| Row Limit | 50 | 5,000 | 5,000 |
| Filters | 1 | Unlimited | Unlimited |
| Two-Way Sync | ✗ | ✓ | ✓ |
| Scheduled Sync | ✗ | ✓ | ✓ |
| Team Features | ✗ | ✗ | ✓ |
| Users | 1 | 1 | 5 |

## Testing

### Test Stripe Integration

1. Use Stripe test mode (test cards available in Stripe docs)
2. Test card: `4242 4242 4242 4242`
3. Any future expiry date and CVC

### Test Feature Access

```javascript
// In Apps Script console
function testPaymentSystem() {
  // Test current plan
  const plan = PaymentService.getCurrentPlan();
  console.log('Current plan:', plan);
  
  // Test feature access
  console.log('Two-way sync:', PaymentService.hasFeatureAccess('two_way_sync'));
  console.log('Team features:', PaymentService.hasFeatureAccess('team_features'));
  
  // Test row limits
  try {
    PaymentService.enforceRowLimit(100); // Should fail on free plan
  } catch (e) {
    console.log('Row limit error:', e.message);
  }
}
```

## Security Considerations

1. **Never expose Stripe secret keys** in client code
2. **Always validate webhooks** with signature verification
3. **Use HTTPS** for all API endpoints
4. **Implement rate limiting** on backend endpoints
5. **Store minimal user data** (only what's necessary)

## Maintenance

### Monthly Tasks

1. Review Stripe Dashboard for failed payments
2. Check backend logs for errors
3. Monitor subscription churn rate
4. Update pricing if needed

### Handling Edge Cases

1. **Payment Failures**: Stripe automatically retries failed payments
2. **Subscription Cancellations**: User reverts to free plan
3. **Team Downgrades**: Limit team to 1 user, others become free users
4. **API Outages**: Features gracefully degrade to free tier

## Support

For payment-related support:
1. Check Stripe Dashboard for payment status
2. Verify webhook delivery in Stripe
3. Check backend logs for errors
4. Test with a new Google account if needed

## Revenue Projections

Based on typical SaaS conversion rates:
- Free to Paid: 2-5% conversion
- Monthly to Annual: 20% prefer annual
- Churn Rate: 5-10% monthly

Example with 1,000 users:
- 950 Free users
- 40 Pro users (30 monthly × $12.99 + 10 annual × $10.75)
- 10 Team users (8 monthly × $39.99 + 2 annual × $33.25)
- **Monthly Revenue**: ~$900
- **Annual Revenue**: ~$11,000

## Next Steps

1. Set up Stripe account and products
2. Deploy backend to Vercel
3. Update PaymentService.js with your backend URL
4. Test the full payment flow
5. Launch to users!

Remember: Start with test mode, thoroughly test all flows, then switch to live mode when ready.