/**
 * Constants for Pipedrive Integration
 * 
 * This module contains all constants used throughout the application:
 * - API URLs and OAuth settings
 * - Default values
 * - Cache keys
 * - Status values
 * - UI constants
 * - Colors
 * - Default column preferences
 */

// Entity types
const ENTITY_TYPES = {
  DEALS: 'deals',
  PERSONS: 'persons',
  ORGANIZATIONS: 'organizations',
  ACTIVITIES: 'activities',
  LEADS: 'leads',
  PRODUCTS: 'products'
};

// OAuth credentials
const PIPEDRIVE_CLIENT_ID = 'f48c99e028029bab';
const PIPEDRIVE_CLIENT_SECRET = '2d245de02052108d8c22d8f7ea8004bc00e7aac7';

// API URLs
const PIPEDRIVE_API_URL_PREFIX = 'https://';
const PIPEDRIVE_API_URL_SUFFIX = '.pipedrive.com/api/v1';
const DEFAULT_PIPEDRIVE_SUBDOMAIN = 'api';

// Default values
const DEFAULT_FILTER_ID = '';
const DEFAULT_SHEET_NAME = 'Sheet1';

// Cache keys
const CACHE_KEYS = {
  DEAL_FIELDS: 'dealFields',
  PERSON_FIELDS: 'personFields',
  ORGANIZATION_FIELDS: 'organizationFields',
  ACTIVITY_FIELDS: 'activityFields',
  LEAD_FIELDS: 'leadFields',
  PRODUCT_FIELDS: 'productFields'
};

// Cache durations (in seconds)
const CACHE_DURATIONS = {
  SHORT: 300,  // 5 minutes
  MEDIUM: 1800,  // 30 minutes
  LONG: 3600,  // 1 hour
  VERY_LONG: 86400  // 24 hours
};

// Status values
const STATUS = {
  SYNCING: 'syncing',
  READY: 'ready',
  ERROR: 'error'
};

// Sync phases
const SYNC_PHASES = {
  CONNECTING: '1',
  RETRIEVING: '2',
  WRITING: '3'
};

// Sync status
const SYNC_STATUS = {
  ACTIVE: 'active',
  COMPLETED: 'completed',
  ERROR: 'error',
  PENDING: 'pending',
  WARNING: 'warning'
};

// UI constants
const UI_CONSTANTS = {
  DIALOG_WIDTH: 600,
  DIALOG_HEIGHT: 400,
  TOAST_DURATION: 5,
  ERROR_TOAST_DURATION: 10
};

// Colors
const COLORS = {
  PRIMARY: '#4285f4',
  SUCCESS: '#0f9d58',
  WARNING: '#f4b400',
  ERROR: '#db4437',
  MODIFIED: '#FCE8E6',
  SYNCED: '#E6F4EA',
  HEADER: '#E8F0FE',
  BACKGROUND: '#F8F9FA',
  BORDER: '#DADCE0'
};

// Font colors
const FONT_COLORS = {
  MODIFIED: '#D93025',
  SYNCED: '#137333'
};

// Border styles
const BORDERS = {
  SYNC_STATUS: {
    top: null,
    right: true,
    bottom: null,
    left: true,
    vertical: false,
    horizontal: false,
    color: '#DADCE0',
    style: SpreadsheetApp.BorderStyle.SOLID
  }
};

// Default column preferences
const DEFAULT_COLUMNS = {
  COMMON: ['id', 'name', 'owner_id', 'created_at', 'updated_at'],
  DEALS: ['id', 'title', 'status', 'value', 'currency', 'owner_id', 'created_at', 'updated_at'],
  PERSONS: ['id', 'name', 'email', 'phone', 'owner_id', 'created_at', 'updated_at'],
  ORGANIZATIONS: ['id', 'name', 'address', 'owner_id', 'created_at', 'updated_at'],
  ACTIVITIES: ['id', 'type', 'due_date', 'duration', 'deal_id', 'person_id', 'org_id', 'note', 'created_at', 'updated_at'],
  LEADS: ['id', 'title', 'owner_id', 'person_id', 'organization_id', 'created_at', 'updated_at'],
  PRODUCTS: ['id', 'name', 'code', 'description', 'unit', 'tax', 'active_flag', 'created_at', 'updated_at']
};