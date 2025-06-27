# Installation Tracking Setup

## Overview
The PipedriveSheets backend now includes installation tracking to monitor when users install the add-on from Google Workspace Marketplace.

## Database Setup

1. **Create the installations table in Supabase:**
   - Run the SQL from `schema.sql` in your Supabase SQL editor
   - The table tracks: email, domain, install time, source, and auth mode

## Environment Variables

Add to your `.env` file:
```
ANALYTICS_API_KEY=your-secret-analytics-key
```

## How It Works

### 1. Marketplace Installations
When a user installs from Google Workspace Marketplace:
- `onInstall()` trigger fires automatically
- Sends installation data to `/api/track-install` endpoint
- Records source as "marketplace"

### 2. Manual Installations
When a user manually installs the add-on:
- Tracked on first initialization
- Records source as "manual"
- Only tracks once per user

### 3. View Analytics
Access analytics at `/api/analytics/installations` with parameters:
- `apiKey`: Your ANALYTICS_API_KEY
- `timeRange`: 7d, 30d, 90d, or all
- Example: `GET /api/analytics/installations?apiKey=xxx&timeRange=30d`

### 4. Analytics Dashboard
Use `analytics-viewer.html` to view installation data:
1. Open the HTML file in a browser
2. Enter your API key
3. Select time range
4. View installation statistics

## Data Collected
- User email (for support purposes)
- Domain (to identify organizations)
- Installation timestamp
- Source (marketplace vs manual)
- Auth mode (FULL, LIMITED, NONE)

## Privacy Considerations
- Installation tracking helps improve the product
- Data is used only for analytics and support
- Never share or sell user data
- Consider adding privacy policy disclosure