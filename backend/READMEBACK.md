# PipedriveSheets Payment Backend

This backend service handles Stripe payments and license validation for the PipedriveSheets Google Workspace Add-on.

## Quick Deployment Guide

### Option 1: Deploy to Vercel (Recommended - Free)

1. **Install Vercel CLI**:
   ```bash
   npm i -g vercel
   ```

2. **Create Supabase Database** (Free tier):
   - Go to [supabase.com](https://supabase.com)
   - Create new project
   - Go to SQL Editor and run:
   ```sql
   CREATE TABLE subscriptions (
     id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
     google_user_id TEXT NOT NULL,
     script_id TEXT NOT NULL,
     stripe_customer_id TEXT,
     stripe_subscription_id TEXT,
     plan_type TEXT NOT NULL,
     status TEXT NOT NULL,
     email TEXT NOT NULL,
     created_at TIMESTAMPTZ DEFAULT NOW(),
     updated_at TIMESTAMPTZ DEFAULT NOW(),
     expires_at TIMESTAMPTZ,
     UNIQUE(google_user_id, script_id)
   );
   
   CREATE INDEX idx_google_user ON subscriptions(google_user_id);
   CREATE INDEX idx_stripe_sub ON subscriptions(stripe_subscription_id);
   ```

3. **Set up Stripe**:
   - Create products and prices in Stripe Dashboard
   - Get your API keys and webhook secret

4. **Deploy to Vercel**:
   ```bash
   cd backend
   vercel
   ```
   
5. **Set Environment Variables in Vercel**:
   ```
   STRIPE_SECRET_KEY=sk_live_...
   STRIPE_WEBHOOK_SECRET=whsec_...
   STRIPE_PRO_MONTHLY_PRICE_ID=price_...
   STRIPE_PRO_ANNUAL_PRICE_ID=price_...
   STRIPE_TEAM_MONTHLY_PRICE_ID=price_...
   STRIPE_TEAM_ANNUAL_PRICE_ID=price_...
   SUPABASE_URL=https://your-project.supabase.co
   SUPABASE_ANON_KEY=your-anon-key
   ```

6. **Configure Stripe Webhook**:
   - In Stripe Dashboard, add webhook endpoint: `https://pipedrive-sheets.vercel.app/webhook`
   - Select events: `checkout.session.completed`, `customer.subscription.updated`, `customer.subscription.deleted`

### Option 2: Deploy to Railway

1. Push code to GitHub
2. Connect Railway to your GitHub repo
3. Add environment variables
4. Deploy

### Option 3: Deploy to Render

1. Create new Web Service on Render
2. Connect GitHub repo
3. Set environment variables
4. Deploy

## Testing

1. **Test checkout**:
   ```bash
   curl -X POST https://pipedrive-sheets.vercel.app/api/create-checkout-session \
     -H "Content-Type: application/json" \
     -d '{
       "email": "test@example.com",
       "googleUserId": "test123",
       "scriptId": "script123",
       "planType": "pro_monthly",
       "successUrl": "https://example.com/success",
       "cancelUrl": "https://example.com/cancel"
     }'
   ```

2. **Test subscription status**:
   ```bash
   curl -X POST https://pipedrive-sheets.vercel.app/api/subscription/status \
     -H "Content-Type: application/json" \
     -d '{
       "email": "test@example.com",
       "googleUserId": "test123",
       "scriptId": "script123"
     }'
   ```

## Updating PaymentService.js

After deployment, update the `API_URL` in `PaymentService.js`:
```javascript
const PaymentService = {
  API_URL: 'https://pipedrive-sheets.vercel.app/api', // Update this
  ...
}
```

## Security Notes

- Never expose Stripe secret keys in client code
- Validate webhook signatures
- Use HTTPS for all endpoints
- Implement rate limiting for production