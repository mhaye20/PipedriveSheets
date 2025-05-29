# PipedriveSheets

A Google Sheets™ Add-on that connects to Pipedrive CRM and syncs data between your spreadsheets and Pipedrive.

## Features

- **Fetch Pipedrive Data**: Import deals, persons, organizations, activities, leads, and products based on saved filters.
- **Customizable Columns**: Select which fields to display in your sheet.
- **Two-Way Sync**: Edit data in Google Sheets™ and push changes back to Pipedrive.
- **Team Collaboration**: Share filters and column preferences with your team.
- **Scheduled Sync**: Set up automatic synchronization on a schedule.
- **Detailed Tracking**: Track when data was last synced and monitor sync status.

## Installation

### From Google Workspace Marketplace™

1. Visit the [Google Workspace Marketplace™](https://workspace.google.com/marketplace) (will be available after approval)
2. Search for "PipedriveSheets"
3. Click "Install" and follow the prompts

### Manual Installation

For developers or for testing:

1. Open a Google Sheet™
2. Go to Extensions > Apps Script
3. Copy all the files from this repository into your Apps Script project
4. Save and reload your sheet
5. You'll see a "Pipedrive" menu in your sheet

## Structure

The code is organized into modules based on functionality:

- **Main.gs**: Entry point with menu creation and initialization
- **PipedriveAPI.gs**: Handles API communication with Pipedrive
- **SyncService.gs**: Manages data synchronization operations
- **TeamAccess.gs**: Team membership verification and access control
- **TeamData.gs**: Team data storage and management
- **UI.gs**: User interface dialogs and interactions
- **Utilities.gs**: Utility functions used throughout the application
- **HTML Templates**: Various .html files for dialog UIs

## Configuration

After installation:

1. Click "Pipedrive" in the menu
2. Select "Initialize Pipedrive Menu"
3. Go to "Filter Settings" to configure your API key and Pipedrive subdomain
4. Choose the entity type (deals, persons, etc.) and select a filter
5. Click "Save Settings"

## Team Features

To create or join a team:

1. Go to "Team Management" in the Pipedrive menu
2. Create a new team or enter a team ID to join an existing team
3. Team members can share filter configurations and column preferences

## Development

To contribute or modify:

1. Clone this repository
2. Make your changes to the relevant modules
3. Test thoroughly in your own Google Sheet™
4. Submit a pull request or publish your own version

## Versioning

- **Version 1.0.0**: Initial release with team collaboration features and two-way sync capability

## License

Apache License 2.0

## Support

For support or feature requests, please contact [support@pipedrivesheets](mailto:support@pipedrivesheets)

## Project Structure

The codebase is organized into the following files:

### Server-side Scripts

- **Main.gs**: Entry point with onOpen trigger and menu creation
- **PipedriveAPI.gs**: API communication with Pipedrive
- **SyncService.gs**: Core synchronization logic
- **TeamAccess.gs**: Team verification functions
- **TeamData.gs**: Team data storage functions
- **UI.gs**: User interface handlers
- **Utilities.gs**: Helper functions

### HTML Templates

- **SettingsDialog.html**: Main configuration dialog
- **ColumnSelector.html**: UI for selecting which Pipedrive fields to display
- **TeamManager.html**: Team creation and management interface
- **TwoWaySyncSettings.html**: Settings for bi-directional sync
- **TriggerManager.html**: Schedule configuration for automatic syncs
- **SyncStatus.html**: Live status display during sync operations
- **Help.html**: Documentation and support information

### Configuration

- **appsscript.json**: Manifest file with OAuth scopes and add-on metadata 

## FILE TREE: 

PipedriveSheets
├── .babelrc
├── .clasp.json
├── .claude
│   └── settings.local.json
├── .git
├── .github
│   └── workflows
│       └── claude.yml
├── .vscode
├── CLAUDE.md
├── ColumnSelector.html
├── ColumnSelectorUI.js
├── ComprehensiveFieldTest.js
├── Constants.js
├── HeaderService.js
├── Help.html
├── Main.js
├── ManageSubscription.html
├── OAuth.js
├── PAYMENT_INTEGRATION_GUIDE.md
├── PaymentDialog.html
├── PaymentService.js
├── PipedriveAPI.js
├── PipedriveDirectAPI.js
├── PipedriveNpm.js
├── PricingSummary.md
├── QuickSetup.js
├── README.md
├── SettingsDialog.html
├── SettingsDialogUI.js
├── SyncService.js
├── SyncStatus.html
├── TeamAccess.js
├── TeamData.js
├── TeamManager.html
├── TeamManagerUI.js
├── TestOrganizationFix.js
├── TestOrganizationPush.js
├── TestPaymentSystem.js
├── TestPipedriveFields.js
├── TriggerManager.html
├── TriggerManagerUI.js
├── TriggerManagerUI_Scripts.html
├── TriggerManagerUI_Styles.html
├── TwoWaySyncSettings.html
├── TwoWaySyncSettingsUI.js
├── UI.js
├── Utilities.js
├── appsscript.json
├── backend
│   ├── .env.example
│   ├── .gitignore
│   ├── .vercel
│   │   ├── README.txt
│   │   └── project.json
│   ├── DEPLOYMENT_STEPS.md
│   ├── PORTAL_SETUP.md
│   ├── READMEBACK.md
│   ├── package.json
│   ├── server.js
│   └── vercel.json
├── dist
├── gaswebpack.md
├── node_modules
├── package-lock.json
├── package.json
├── pipedriveAPIdocumentation.md
├── pipedriveNPMapiV1.md
├── pipedriveNPMapiV2.md
├── pipedrive_client-nodejs.md
├── src
│   ├── api.js
│   └── index.js
└── webpack.gas.js