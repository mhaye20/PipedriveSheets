# Backend Fix Notes for Team Member Subscription Lookup

## Problem
Team members cannot see their team owner's subscription status because:
1. The backend looks up subscriptions by `google_user_id` 
2. Team members can't provide the team owner's Google User ID
3. When lookup fails, it returns "free" plan

## Solution Implemented
Modified `/api/subscription/status` endpoint to:
1. First try lookup by `google_user_id` if provided
2. Fall back to email lookup if not found or if `google_user_id` is empty
3. Use case-insensitive email comparison

## Database Requirements
The `subscriptions` table needs:
- `email` column (already being populated on creation)
- Index on `(email, script_id)` for efficient lookups

## Deployment Steps
1. Deploy the updated `server.js` to Vercel
2. Ensure database has `email` column indexed
3. For existing subscriptions without email, run a migration to populate from Stripe

## Testing
After deployment, test that:
1. Team owners can still see their own subscriptions
2. Team members can see team owner's subscription status
3. Email lookups are case-insensitive