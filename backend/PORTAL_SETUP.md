# Setting Up Stripe Customer Portal

## Before Deploying

You need to enable and configure the Stripe Customer Portal in your Stripe Dashboard:

### 1. Enable Customer Portal in Stripe

1. Go to [Stripe Dashboard](https://dashboard.stripe.com/)
2. Navigate to **Settings** → **Billing** → **Customer portal**
3. Click **Activate portal** if not already active

### 2. Configure Portal Settings

In the Customer Portal settings, configure:

#### Business Information
- **Business name**: PipedriveSheets
- **Privacy policy**: Link to your privacy policy
- **Terms of service**: Link to your terms

#### Features to Enable
- ✅ **Invoices**: Customers can view and download
- ✅ **Update payment methods**: Allow customers to update cards
- ✅ **Cancel subscriptions**: Allow customers to cancel
- ✅ **Switch plans**: Optional - if you want customers to change plans

#### Cancellation Settings
- **Cancellation reason**: Required
- **Cancellation flow**: Immediate or at end of billing period
- **Proration**: Choose how to handle partial month refunds

### 3. Save Configuration

Click **Save** after configuring all settings.

## Deploy Backend Update

After configuring Stripe:

```bash
cd backend
vercel --prod
```

## Test Portal Access

```bash
# First, create a test subscription using the checkout flow
# Then test portal access with actual subscription data

curl -X POST https://pipedrive-sheets.vercel.app/api/create-portal-session \
  -H "Content-Type: application/json" \
  -d '{
    "email": "your-test-email@example.com",
    "googleUserId": "your-actual-google-user-id",
    "scriptId": "your-actual-script-id",
    "returnUrl": "https://docs.google.com/spreadsheets"
  }'
```

## Troubleshooting

### "No active subscription found"
- Make sure you're using the same email/userId/scriptId that was used for the subscription
- Check Supabase to verify the subscription record exists
- Ensure stripe_customer_id is populated in the database

### Portal Not Working
- Verify portal is activated in Stripe Dashboard
- Check that all required fields are configured
- Ensure your Stripe API keys have the necessary permissions

### Customer Can't Cancel
- Check cancellation settings in portal configuration
- Verify the subscription status is "active"
- Make sure cancellation flow is properly configured