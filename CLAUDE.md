# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Sheets Add-on that integrates with Pipedrive CRM. It's built using Google Apps Script with a modern JavaScript build process that allows the use of npm packages and ES6+ syntax.

## Build Commands

```bash
npm run build    # Webpack build to dist/
npm run deploy   # Build + clasp push to Google Apps Script
```

## Architecture

### Dual-Build System
- **Native GAS Files**: `.js` and `.html` files in root are copied directly to `dist/`
- **NPM Integration**: `src/index.js` bundles npm packages (pipedrive SDK) via webpack into `dist/main.bundle.js`, exposed as global `AppLib`

### Core Services
- **PipedriveAPI.js**: Direct API communication layer handling authentication and requests
- **SyncService.js**: Manages bidirectional synchronization between Sheets and Pipedrive
- **TeamAccess.js/TeamData.js**: Team collaboration features using PropertiesService
- **OAuth.js**: OAuth2 authentication flow implementation

### UI Pattern
HTML dialogs (`*Dialog.html`, `*Manager.html`) are served via `HtmlService` with corresponding UI handler files (`*UI.js`).

### Key Implementation Details
- Custom Axios adapter in `src/index.js` wraps Google's UrlFetchApp for npm compatibility
- Date/time fields require special handling - converted to YYYY-MM-DD and HH:MM:SS formats
- Address fields use structured format with subfields (address, address_subpremise, etc.)
- Uses Google PropertiesService for persistent storage (user/script properties)

## Development Notes

### When Making Changes
1. Edit appropriate files (check if functionality exists in native GAS files first)
2. Run `npm run build` to compile
3. Test in Google Sheets environment (no local testing available)
4. Changes to npm dependencies require webpack rebuild

### Important Constraints
- Google Apps Script V8 runtime limitations
- No native npm support - requires polyfills (URL, URLSearchParams)
- UrlFetchApp instead of standard HTTP libraries
- No filesystem access - use PropertiesService/CacheService
- UI must use HtmlService dialogs/sidebars

### Common Patterns
- Namespace objects for code organization (e.g., `SyncService.syncData()`)
- Error handling with try/catch and SpreadsheetApp.getUi().alert()
- Column detection using `HeaderService.getHeaderColumns()`
- Team data stored with `teamId_` prefix in script properties