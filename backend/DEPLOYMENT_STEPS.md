# Deployment Steps for Backend Updates

## 1. Redeploy to Vercel

Run these commands in the backend directory:

```bash
cd backend
vercel --prod
```

This will deploy the updated server with:
- Proper Vercel configuration (vercel.json)
- Health check endpoint
- Correct module exports for Vercel

## 2. Test the Deployment

After deployment, test these endpoints:

### Test Health Check
```bash
curl https://pipedrive-sheets.vercel.app/
```

Expected response:
```json
{
  "status": "ok",
  "service": "PipedriveSheets Payment Backend",
  "timestamp": "2025-01-27T17:45:00.000Z"
}
```

### Test Checkout Session Creation
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

Expected response (if Stripe is configured correctly):
```json
{
  "checkoutUrl": "https://checkout.stripe.com/pay/cs_test_..."
}
```

### Test Subscription Status
```bash
curl -X POST https://pipedrive-sheets.vercel.app/api/subscription/status \
  -H "Content-Type: application/json" \
  -d '{
    "email": "test@example.com",
    "googleUserId": "test123",
    "scriptId": "script123"
  }'
```

Expected response (for new user):
```json
{
  "plan": "free",
  "status": "active",
  "features": []
}
```

## 3. Common Issues and Solutions

### Issue: 404 Not Found
- Make sure you ran `vercel --prod` in the backend directory
- Check that vercel.json is present
- Verify the deployment URL matches what's in PaymentService.js

### Issue: Stripe Errors
- Verify all environment variables are set in Vercel dashboard
- Check that price IDs match your Stripe products
- Ensure webhook secret is correct

### Issue: Database Errors
- Verify Supabase URL and anon key are correct
- Check that the subscriptions table was created
- Test database connection independently

## 4. Update PaymentService.js

Once deployment is successful, ensure PaymentService.js has the correct URL:

```javascript
const PaymentService = {
  API_URL: 'https://pipedrive-sheets.vercel.app/api',
  // ... rest of the code
}
```

## 5. Test in Google Sheets

After everything is deployed:

1. Open your Google Sheet with the add-on
2. Go to Pipedrive menu → Upgrade Plan
3. Try upgrading to Pro plan
4. Complete test payment with Stripe test card: 4242 4242 4242 4242

The system should:
- Show the upgrade dialog
- Redirect to Stripe checkout
- Process payment via webhook
- Update user's subscription status
- Unlock premium features