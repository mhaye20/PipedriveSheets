/**
 * User Interface
 * 
 * This module handles all UI-related functions:
 * - Showing dialogs and sidebars
 * - Building UI components
 * - Managing user interactions
 */

/**
 * Shows settings dialog where users can configure filter ID and entity type
 */
function showSettings() {
  // Get the active sheet name
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeSheetName = activeSheet.getName();
  
  // Get current settings from properties
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Check if we're authenticated
  const accessToken = scriptProperties.getProperty('PIPEDRIVE_ACCESS_TOKEN');
  if (!accessToken) {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Not Connected',
      'You need to connect to Pipedrive before configuring settings. Connect now?',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      showAuthorizationDialog();
    }
    return;
  }
  
  const savedSubdomain = scriptProperties.getProperty('PIPEDRIVE_SUBDOMAIN') || DEFAULT_PIPEDRIVE_SUBDOMAIN;
  
  // Get sheet-specific settings using the sheet name
  const sheetFilterIdKey = `FILTER_ID_${activeSheetName}`;
  const sheetEntityTypeKey = `ENTITY_TYPE_${activeSheetName}`;
  
  const savedFilterId = scriptProperties.getProperty(sheetFilterIdKey) || '';
  const savedEntityType = scriptProperties.getProperty(sheetEntityTypeKey) || ENTITY_TYPES.DEALS;
  
  // Create HTML content for the settings dialog
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 5px; margin-top: 5px; }
      .button-container { margin-top: 20px; text-align: right; }
      button { padding: 8px 12px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      .note { font-size: 12px; color: #666; margin-top: 5px; }
      .domain-container { display: flex; align-items: center; }
      .domain-input { flex: 1; margin-right: 5px; }
      .domain-suffix { padding: 5px; background: #f0f0f0; border: 1px solid #ccc; }
      .loading { display: none; margin-right: 10px; }
      .loader { 
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid rgba(255,255,255,.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 1s ease-in-out infinite;
        vertical-align: middle;
      }
      .sheet-info {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        font-size: 14px;
        border-left: 4px solid #4285f4;
      }
      .connected-status {
        background-color: #e6f4ea;
        border-left: 4px solid #34a853;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
      }
      .connected-status i {
        color: #34a853;
        margin-right: 10px;
        font-size: 24px;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
    <h3>Pipedrive Integration Settings</h3>
    
    <div class="connected-status">
      <span style="color: #34a853; font-size: 24px; margin-right: 10px;">✓</span>
      <div>
        <strong>Connected to Pipedrive</strong><br>
        <span style="font-size: 12px;">Company: ${savedSubdomain}.pipedrive.com</span>
      </div>
      <button style="margin-left: auto; background-color: #f1f3f4; color: #202124;" onclick="reconnect()">Reconnect</button>
    </div>
    
    <div class="sheet-info">
      Configuring settings for sheet: <strong>"${activeSheetName}"</strong>
    </div>
    
    <form id="settingsForm">
      <label for="entityType">Entity Type:</label>
      <select id="entityType">
        <option value="deals" ${savedEntityType === 'deals' ? 'selected' : ''}>Deals</option>
        <option value="persons" ${savedEntityType === 'persons' ? 'selected' : ''}>Persons</option>
        <option value="organizations" ${savedEntityType === 'organizations' ? 'selected' : ''}>Organizations</option>
        <option value="activities" ${savedEntityType === 'activities' ? 'selected' : ''}>Activities</option>
        <option value="leads" ${savedEntityType === 'leads' ? 'selected' : ''}>Leads</option>
      </select>
      <p class="note">Select which entity type you want to sync from Pipedrive</p>
      
      <label for="filterId">Filter ID:</label>
      <input type="text" id="filterId" value="${savedFilterId}" />
      <p class="note">The filter ID can be found in the URL when viewing your filter in Pipedrive</p>
      
      <input type="hidden" id="sheetName" value="${activeSheetName}" />
      <p class="note">The data will be exported to the current sheet: "${activeSheetName}"</p>
      
      <div class="button-container">
        <span class="loading" id="saveLoading"><span class="loader"></span> Saving...</span>
        <button type="button" id="saveBtn" onclick="saveSettings()">Save Settings</button>
      </div>
    </form>
    
    <script>
      function saveSettings() {
        const entityType = document.getElementById('entityType').value;
        const filterId = document.getElementById('filterId').value;
        const sheetName = document.getElementById('sheetName').value;
        
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'inline-block';
        document.getElementById('saveBtn').disabled = true;
        
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            alert('Settings saved successfully!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            alert('Error saving settings: ' + error.message);
          })
          .saveSettings('', entityType, filterId, '', sheetName);
      }
      
      function reconnect() {
        google.script.host.close();
        google.script.run.showAuthorizationDialog();
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(520)
  .setTitle(`Pipedrive Settings for "${activeSheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Pipedrive Settings for "${activeSheetName}"`);
}

/**
 * Shows the column selector dialog
 */
function showColumnSelector() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Get entity type and sheet name
    const entityType = docProps.getProperty('PIPEDRIVE_ENTITY_TYPE') || ENTITY_TYPES.DEALS;
    const sheetName = docProps.getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Check if API key is configured
    const apiKey = docProps.getProperty('PIPEDRIVE_API_KEY');
    if (!apiKey) {
      SpreadsheetApp.getUi().alert(
        'API Key Required',
        'Please configure your Pipedrive API key in the settings first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      showSettings();
      return;
    }
    
    // Get fields based on entity type
    let fields = [];
    
    switch (entityType) {
      case ENTITY_TYPES.DEALS:
        fields = getDealFields();
        break;
      case ENTITY_TYPES.PERSONS:
        fields = getPersonFields();
        break;
      case ENTITY_TYPES.ORGANIZATIONS:
        fields = getOrganizationFields();
        break;
      case ENTITY_TYPES.ACTIVITIES:
        fields = getActivityFields();
        break;
      case ENTITY_TYPES.LEADS:
        fields = getLeadFields();
        break;
      case ENTITY_TYPES.PRODUCTS:
        fields = getProductFields();
        break;
    }
    
    // Process fields to create UI-friendly structure
    const availableColumns = extractFields(fields);
    
    // Get currently selected columns
    const selectedColumns = getTeamAwareColumnPreferences(entityType, sheetName) || [];
    
    // Show the column selector UI
    showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName);
  } catch (e) {
    Logger.log(`Error in showColumnSelector: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open column selector: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Extracts fields into a UI-friendly structure
 * @param {Array} fields - The fields to extract
 * @param {string} parentPath - Parent path for nested fields
 * @param {string} parentName - Parent name for nested fields
 * @return {Array} Processed fields
 */
function extractFields(fields, parentPath = '', parentName = '') {
  try {
    const result = [];
    
    // If this is an array of field definitions like from Pipedrive API
    if (Array.isArray(fields)) {
      fields.forEach(field => {
        // Skip unsupported field types
        if (field.field_type === 'picture' ||
            field.field_type === 'file' ||
            field.field_type === 'visibleto') {
          return;
        }
        
        // Ensure the field has a key
        if (!field.key) return;
        
        // Construct path
        const path = parentPath ? `${parentPath}.${field.key}` : field.key;
        
        // Construct readable name
        const name = parentName ? `${parentName} > ${field.name}` : field.name;
        
        // Add the field to the result
        result.push({
          path: path,
          name: name,
          type: field.field_type,
          options: field.options
        });
        
        // Handle nested fields (like fields within products, etc.)
        if (field.fields) {
          const nestedFields = extractFields(field.fields, path, name);
          result.push(...nestedFields);
        }
      });
    }
    // If this is an object with nested fields
    else if (typeof fields === 'object' && fields !== null) {
      Object.keys(fields).forEach(key => {
        // Skip certain keys that aren't proper fields
        if (['id', 'success', 'error', 'key'].includes(key)) {
          return;
        }
        
        // Get the field
        const field = fields[key];
        
        // Skip non-objects
        if (typeof field !== 'object' || field === null) {
          return;
        }
        
        // Construct path
        const path = parentPath ? `${parentPath}.${key}` : key;
        
        // If field has a name property, it's likely a proper field
        if (field.name) {
          // Construct readable name
          const name = parentName ? `${parentName} > ${field.name}` : field.name;
          
          // Add the field to the result
          result.push({
            path: path,
            name: name,
            type: field.field_type || typeof field,
            options: field.options
          });
        }
        
        // Recursively process nested fields
        const nestedFields = extractFields(field, path, field.name || parentName);
        result.push(...nestedFields);
      });
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error in extractFields: ${e.message}`);
    return [];
  }
}

/**
 * Shows the column selector UI
 * @param {Array} availableColumns - Available columns
 * @param {Array} selectedColumns - Selected columns
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 */
function showColumnSelectorUI(availableColumns, selectedColumns, entityType, sheetName) {
  const htmlContent = `
    <style>
      body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
      .container { display: flex; height: 400px; }
      .column { width: 50%; padding: 10px; box-sizing: border-box; }
      .scrollable { height: 340px; overflow-y: auto; border: 1px solid #ccc; padding: 5px; }
      .header { font-weight: bold; margin-bottom: 10px; }
      .item { padding: 5px; margin: 2px 0; cursor: pointer; border-radius: 3px; }
      .item:hover { background-color: #f0f0f0; }
      .selected { background-color: #e8f0fe; }
      .footer { margin-top: 10px; display: flex; justify-content: space-between; }
      .search { margin-bottom: 10px; width: 100%; padding: 5px; }
      button { padding: 8px 12px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
      button.secondary { background-color: #f0f0f0; color: #333; }
      .action-btns { display: flex; gap: 5px; align-items: center; }
      .category { font-weight: bold; margin-top: 5px; padding: 5px; background-color: #f0f0f0; }
      .nested { margin-left: 15px; }
      .info { font-size: 12px; color: #666; margin-bottom: 5px; }
      .loading { display: none; margin-right: 10px; }
      .loader { 
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid rgba(255,255,255,.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 1s ease-in-out infinite;
        vertical-align: middle;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
      .drag-handle {
        display: inline-block;
        width: 10px;
        height: 16px;
        background-image: radial-gradient(circle, #999 1px, transparent 1px);
        background-size: 3px 3px;
        background-position: 0 center;
        background-repeat: repeat;
        margin-right: 8px;
        cursor: grab;
        opacity: 0.5;
      }
      .selected:hover .drag-handle {
        opacity: 1;
      }
      .dragging {
        opacity: 0.4;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
      }
      .over {
        border-top: 2px solid #4285f4;
      }
      .sheet-info {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 4px;
        margin-bottom: 15px;
        font-size: 14px;
        border-left: 4px solid #4285f4;
      }
    </style>
    
    <div class="header">
      <div class="sheet-info">
        Configuring columns for <strong>${entityType}</strong> in sheet <strong>"${sheetName}"</strong>
      </div>
      <p class="info">Select columns from the left panel and add them to the right panel. Drag items in the right panel to reorder them.</p>
    </div>
    
    <div class="container">
      <div class="column">
        <input type="text" id="searchBox" class="search" placeholder="Search for columns...">
        <div class="header">Available Columns</div>
        <div id="availableList" class="scrollable">
          <!-- Available columns will be populated here by JavaScript -->
        </div>
      </div>
      
      <div class="column">
        <div class="header">Selected Columns</div>
        <div id="selectedList" class="scrollable">
          <!-- Selected columns will be populated here by JavaScript -->
        </div>
      </div>
    </div>
    
    <div class="footer">
      <div class="action-btns">
        <button class="secondary" id="debug">View Debug Info</button>
      </div>
      <div class="action-btns">
        <span class="loading" id="saveLoading"><span class="loader"></span> Saving...</span>
        <button class="secondary" id="cancel">Cancel</button>
        <button id="save">Save & Close</button>
      </div>
    </div>

    <script>
      // Initialize data
      let availableColumns = ${JSON.stringify(availableColumns)};
      let selectedColumns = ${JSON.stringify(selectedColumns)};
      const entityType = "${entityType}";
      const sheetName = "${sheetName}";
      
      // DOM elements
      const availableList = document.getElementById('availableList');
      const selectedList = document.getElementById('selectedList');
      const searchBox = document.getElementById('searchBox');
      
      // Render the lists
      function renderAvailableList(searchTerm = '') {
        availableList.innerHTML = '';
        
        // Group columns by parent key or top-level
        const topLevel = [];
        const nested = {};
        
        availableColumns.forEach(col => {
          if (!selectedColumns.some(selected => selected.key === col.key)) {
            if (col.name.toLowerCase().includes(searchTerm.toLowerCase())) {
              if (!col.isNested) {
                topLevel.push(col);
              } else {
                const parentKey = col.parentKey || 'unknown';
                if (!nested[parentKey]) {
                  nested[parentKey] = [];
                }
                nested[parentKey].push(col);
              }
            }
          }
        });
        
        // Add top-level columns first
        if (topLevel.length > 0) {
          const topLevelHeader = document.createElement('div');
          topLevelHeader.className = 'category';
          topLevelHeader.textContent = 'Main Fields';
          availableList.appendChild(topLevelHeader);
          
          topLevel.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.textContent = col.name;
            item.dataset.key = col.key;
            item.onclick = () => addColumn(col);
            availableList.appendChild(item);
          });
        }
        
        // Then add nested columns by parent
        for (const parentKey in nested) {
          if (nested[parentKey].length > 0) {
            const parentName = availableColumns.find(col => col.key === parentKey)?.name || parentKey;
            
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'category';
            categoryHeader.textContent = parentName;
            availableList.appendChild(categoryHeader);
            
            nested[parentKey].forEach(col => {
              const item = document.createElement('div');
              item.className = 'item nested';
              item.textContent = col.name;
              item.dataset.key = col.key;
              item.onclick = () => addColumn(col);
              availableList.appendChild(item);
            });
          }
        }
      }
      
      function renderSelectedList() {
        selectedList.innerHTML = '';
        selectedColumns.forEach((col, index) => {
          const item = document.createElement('div');
          item.className = 'item selected';
          item.dataset.key = col.key;
          item.dataset.index = index;
          item.draggable = true;
          
          // Add drag handle
          const dragHandle = document.createElement('span');
          dragHandle.className = 'drag-handle';
          dragHandle.innerHTML = '&nbsp;&nbsp;&nbsp;';
          item.appendChild(dragHandle);
          
          // Add column name
          const nameSpan = document.createElement('span');
          nameSpan.textContent = col.name;
          item.appendChild(nameSpan);
          
          item.ondragstart = handleDragStart;
          item.ondragover = handleDragOver;
          item.ondrop = handleDrop;
          item.ondragend = handleDragEnd;
          
          const removeBtn = document.createElement('span');
          removeBtn.textContent = ' ✕';
          removeBtn.style.color = 'red';
          removeBtn.style.float = 'right';
          removeBtn.style.cursor = 'pointer';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeColumn(col);
          };
          
          item.appendChild(removeBtn);
          selectedList.appendChild(item);
        });
      }
      
      // Drag and drop functionality
      let draggedItem = null;
      
      function handleDragStart(e) {
        draggedItem = this;
        this.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', this.dataset.index);
        
        // Add a small delay to make the visual change noticeable
        setTimeout(() => {
          this.style.opacity = '0.4';
        }, 0);
      }
      
      function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this.classList.add('over');
      }
      
      function handleDrop(e) {
        e.preventDefault();
        this.classList.remove('over');
        
        const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
        const toIndex = parseInt(this.dataset.index);
        
        if (fromIndex !== toIndex) {
          const item = selectedColumns[fromIndex];
          selectedColumns.splice(fromIndex, 1);
          selectedColumns.splice(toIndex, 0, item);
          renderSelectedList();
        }
      }
      
      function handleDragEnd() {
        this.classList.remove('dragging');
        document.querySelectorAll('.item').forEach(item => {
          item.classList.remove('over');
        });
      }
      
      // Column management
      function addColumn(column) {
        selectedColumns.push(column);
        renderAvailableList(searchBox.value);
        renderSelectedList();
      }
      
      function removeColumn(column) {
        selectedColumns = selectedColumns.filter(col => col.key !== column.key);
        renderAvailableList(searchBox.value);
        renderSelectedList();
      }
      
      // Event listeners
      document.getElementById('save').onclick = () => {
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'inline-block';
        document.getElementById('save').disabled = true;
        document.getElementById('cancel').disabled = true;
        
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('saveLoading').style.display = 'none';
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('save').disabled = false;
            document.getElementById('cancel').disabled = false;
            alert('Error saving column preferences: ' + error.message);
          })
          .saveColumnPreferences(selectedColumns, entityType, sheetName);
      };
      
      document.getElementById('cancel').onclick = () => {
        google.script.host.close();
      };
      
      document.getElementById('debug').onclick = () => {
        google.script.run.logDebugInfo();
        alert('Debug information has been logged to the Apps Script execution log. You can view it from View > Logs in the Apps Script editor.');
      };
      
      searchBox.oninput = () => {
        renderAvailableList(searchBox.value);
      };
      
      // Initial render
      renderAvailableList();
      renderSelectedList();
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(800)
    .setHeight(550)
    .setTitle(`Select Columns for ${entityType} in "${sheetName}"`);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Select Columns for ${entityType} in "${sheetName}"`);
}

/**
 * Saves column preferences for a specific entity type and sheet
 * @param {Array} columns - Array of column keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {boolean} True if successful, false otherwise
 */
function saveColumnPreferences(columns, entityType, sheetName) {
  return saveTeamAwareColumnPreferences(columns, entityType, sheetName);
}

/**
 * Shows the sync status dialog
 * @param {string} sheetName - Sheet name
 */
function showSyncStatus(sheetName) {
  try {
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('SyncStatus');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Syncing Data')
      .setWidth(400)
      .setHeight(300);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Syncing Data');
  } catch (e) {
    Logger.log(`Error in showSyncStatus: ${e.message}`);
    // Just log the error but don't throw, as this is a UI enhancement
  }
}

/**
 * Shows the team manager dialog
 * @param {boolean} joinOnly - Whether to show only the join team UI
 */
function showTeamManager(joinOnly = false) {
  try {
    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Unable to determine your email address. Please ensure you are signed in.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get teams data
    const teamsData = getTeamsData();
    
    // Get user's team
    const userTeam = getUserTeam(userEmail, teamsData);
    
    // Determine if user is in a team
    const hasTeam = userTeam !== null;
    
    // If join-only mode and user isn't in a team, use the simpler join dialog
    if (joinOnly && !hasTeam) {
      return showTeamJoinRequest();
    }
    
    // Create a direct HTML string rather than using a complex template
    // This is more robust and less likely to have loading issues
    let html = `
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
        }
        h3 {
          margin-top: 0;
          margin-bottom: 15px;
          color: #4285F4;
        }
        .card {
          border: 1px solid #ddd;
          border-radius: 4px;
          padding: 15px;
          margin-bottom: 20px;
          background-color: #f9f9f9;
        }
        .section-title {
          font-weight: bold;
          margin-bottom: 10px;
        }
        input[type="text"] {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-bottom: 10px;
          box-sizing: border-box;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 8px 15px;
          border-radius: 4px;
          cursor: pointer;
          margin-right: 10px;
          margin-bottom: 10px;
        }
        button:hover {
          background-color: #3367D6;
        }
        button.secondary {
          background-color: #f1f1f1;
          color: #333;
        }
        button.secondary:hover {
          background-color: #e4e4e4;
        }
        button.danger {
          background-color: #EA4335;
        }
        button.danger:hover {
          background-color: #D62516;
        }
        .footer {
          margin-top: 20px;
          display: flex;
          justify-content: flex-end;
        }
        .tab-container {
          display: flex;
          margin-bottom: 20px;
          border-bottom: 1px solid #ddd;
        }
        .tab {
          padding: 10px 15px;
          cursor: pointer;
        }
        .tab.active {
          color: #4285F4;
          font-weight: bold;
          border-bottom: 2px solid #4285F4;
        }
        .tab-content {
          display: none;
        }
        .tab-content.active {
          display: block;
        }
        .member-list {
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-top: 10px;
          max-height: 200px;
          overflow-y: auto;
        }
        .member {
          padding: 8px 12px;
          border-bottom: 1px solid #eee;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .member:last-child {
          border-bottom: none;
        }
        .status-message {
          padding: 10px;
          margin: 10px 0;
          border-radius: 4px;
        }
        .status-success {
          background-color: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }
        .status-error {
          background-color: #f8d7da;
          color: #721c24;
          border: 1px solid #f5c6cb;
        }
      </style>
      
      <div class="container">
        <h3>${hasTeam ? 'Team Management' : 'Team Access'}</h3>
        
        <div id="status-container"></div>
        
        <div class="tab-container">
          ${hasTeam ? '' : '<div class="tab active" data-tab="join">Join Team</div>'}
          ${hasTeam ? '' : '<div class="tab" data-tab="create">Create Team</div>'}
          ${hasTeam ? '<div class="tab active" data-tab="manage">Manage Team</div>' : ''}
        </div>
    `;
    
    // Join Team Tab
    if (!hasTeam) {
      html += `
        <div id="join-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Join an Existing Team</div>
            <p>Enter the team ID provided by your team administrator:</p>
            <input type="text" id="team-id-input" placeholder="Team ID">
            <div class="footer">
              <button id="join-team-button">Join Team</button>
            </div>
          </div>
        </div>
        
        <div id="create-tab" class="tab-content">
          <div class="card">
            <div class="section-title">Create a New Team</div>
            <p>Create a new team to share Pipedrive configuration with your colleagues:</p>
            <input type="text" id="team-name-input" placeholder="Team Name">
            <div class="footer">
              <button id="create-team-button">Create Team</button>
            </div>
          </div>
        </div>
      `;
    }
    
    // Manage Team Tab
    if (hasTeam) {
      const teamName = userTeam.name || 'Unnamed Team';
      const teamId = userTeam.teamId || '';
      const userRole = userTeam.role || 'Member';
      const isAdmin = userRole === 'Admin';
      
      // Get team members
      const members = [];
      if (teamsData[teamId]) {
        if (teamsData[teamId].members) {
          // New format
          Object.keys(teamsData[teamId].members).forEach(email => {
            members.push({
              email: email,
              role: teamsData[teamId].members[email]
            });
          });
        } else if (teamsData[teamId].memberEmails) {
          // Legacy format
          teamsData[teamId].memberEmails.forEach(email => {
            const isAdmin = teamsData[teamId].adminEmails && 
                           teamsData[teamId].adminEmails.includes(email);
            members.push({
              email: email,
              role: isAdmin ? 'Admin' : 'Member'
            });
          });
        }
      }
      
      html += `
        <div id="manage-tab" class="tab-content active">
          <div class="card">
            <div class="section-title">Current Team</div>
            <div style="margin-bottom: 15px;">
              <div><strong>Team Name:</strong> ${teamName}</div>
              <div><strong>Team ID:</strong> <span id="team-id-display">${teamId}</span></div>
              <div><strong>Your Role:</strong> ${userRole}</div>
            </div>
            
            <button id="copy-team-id" class="secondary">Copy Team ID</button>
            
            <div class="section-title" style="margin-top: 20px;">Team Members</div>
            <div class="member-list">
      `;
      
      if (members.length > 0) {
        members.forEach(member => {
          html += `
            <div class="member">
              <div>${member.email}</div>
              <div>
                <span style="margin-right: 10px;">${member.role}</span>
                ${isAdmin && member.email !== userEmail ? 
                  `<button class="secondary remove-member" data-email="${member.email}">Remove</button>` : ''}
              </div>
            </div>
          `;
        });
      } else {
        html += '<div class="member">No members found</div>';
      }
      
      html += `
            </div>
      `;
      
      if (isAdmin) {
        html += `
            <div class="section-title" style="margin-top: 15px;">Add Team Member</div>
            <input type="text" id="new-member-email" placeholder="Email Address">
            <button id="add-member-button">Add Member</button>
        `;
      }
      
      html += `
            <div style="margin-top: 20px;">
              <button id="leave-team-button" class="danger">Leave Team</button>
              ${isAdmin ? '<button id="delete-team-button" class="danger">Delete Team</button>' : ''}
            </div>
          </div>
        </div>
      `;
    }
    
    // Footer
    html += `
        <div class="footer">
          <button id="close-button" class="secondary">Close</button>
        </div>
      </div>
      
      <script>
        // Document ready
        (function() {
          // Tab switching
          document.querySelectorAll('.tab').forEach(tab => {
            tab.addEventListener('click', function() {
              // Update active tab
              document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
              this.classList.add('active');
              
              // Show corresponding content
              const tabId = this.dataset.tab + '-tab';
              document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
              });
              document.getElementById(tabId).classList.add('active');
            });
          });
          
          // Status message helper
          window.showStatus = function(message, type) {
            const container = document.getElementById('status-container');
            const statusDiv = document.createElement('div');
            statusDiv.className = 'status-message status-' + type;
            statusDiv.textContent = message;
            container.innerHTML = '';
            container.appendChild(statusDiv);
            
            if (type === 'success') {
              setTimeout(() => statusDiv.remove(), 5000);
            }
          };
          
          // Close button
          document.getElementById('close-button').addEventListener('click', function() {
            google.script.host.close();
          });
    `;
    
    // Add join-specific handlers
    if (!hasTeam) {
      html += `
          // Join team handler
          document.getElementById('join-team-button').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-input').value.trim();
            if (!teamId) {
              showStatus('Please enter a team ID', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Joining team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Successfully joined the team!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('join-team-button').disabled = false;
                  showStatus(result.message || 'Error joining team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('join-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .joinTeam(teamId);
          });
          
          // Create team handler
          document.getElementById('create-team-button').addEventListener('click', function() {
            const teamName = document.getElementById('team-name-input').value.trim();
            if (!teamName) {
              showStatus('Please enter a team name', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Creating team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team created successfully!', 'success');
                  // Use fixMenuAfterJoin to properly refresh the menu
                  google.script.run.fixMenuAfterJoin();
                  setTimeout(() => google.script.host.close(), 1000);
                } else {
                  document.getElementById('create-team-button').disabled = false;
                  showStatus(result.message || 'Error creating team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('create-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .createTeam(teamName);
          });
      `;
    }
    
    // Add team management handlers
    if (hasTeam) {
      html += `
          // Copy team ID handler
          document.getElementById('copy-team-id').addEventListener('click', function() {
            const teamId = document.getElementById('team-id-display').textContent;
            
            // Create a temporary input element
            const tempInput = document.createElement('input');
            tempInput.value = teamId;
            document.body.appendChild(tempInput);
            tempInput.select();
            document.execCommand('copy');
            document.body.removeChild(tempInput);
            
            showStatus('Team ID copied to clipboard!', 'success');
          });
          
          // Leave team handler
          document.getElementById('leave-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to leave this team? You will lose access to shared configurations.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Leaving team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('You have left the team.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('leave-team-button').disabled = false;
                  showStatus(result.message || 'Error leaving team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('leave-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .leaveTeam();
          });
      `;
      
      // Add admin-specific handlers
      if (userTeam && userTeam.role === 'Admin') {
        html += `
          // Add member handler
          document.getElementById('add-member-button').addEventListener('click', function() {
            const email = document.getElementById('new-member-email').value.trim();
            if (!email) {
              showStatus('Please enter an email address', 'error');
              return;
            }
            
            this.disabled = true;
            showStatus('Adding member...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Member added successfully.', 'success');
                  document.getElementById('new-member-email').value = '';
                  document.getElementById('add-member-button').disabled = false;
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('add-member-button').disabled = false;
                  showStatus(result.message || 'Error adding member', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('add-member-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .addTeamMember(email);
          });
          
          // Remove member handlers
          document.querySelectorAll('.remove-member').forEach(button => {
            button.addEventListener('click', function() {
              const email = this.dataset.email;
              if (!confirm('Are you sure you want to remove ' + email + ' from the team?')) {
                return;
              }
              
              this.disabled = true;
              showStatus('Removing member...', 'info');
              
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    showStatus('Member removed successfully.', 'success');
                    setTimeout(() => window.top.location.reload(), 2000);
                  } else {
                    document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                    showStatus(result.message || 'Error removing member', 'error');
                  }
                })
                .withFailureHandler(function(error) {
                  document.querySelectorAll('.remove-member').forEach(btn => btn.disabled = false);
                  showStatus('Error: ' + error.message, 'error');
                })
                .removeTeamMember(email);
            });
          });
          
          // Delete team handler
          document.getElementById('delete-team-button').addEventListener('click', function() {
            if (!confirm('Are you sure you want to delete this team? This will remove access for all team members and cannot be undone.')) {
              return;
            }
            
            this.disabled = true;
            showStatus('Deleting team...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showStatus('Team has been deleted.', 'success');
                  setTimeout(() => window.top.location.reload(), 2000);
                } else {
                  document.getElementById('delete-team-button').disabled = false;
                  showStatus(result.message || 'Error deleting team', 'error');
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('delete-team-button').disabled = false;
                showStatus('Error: ' + error.message, 'error');
              })
              .deleteTeam();
          });
        `;
      }
    }
    
    // Close script tag
    html += `
        })(); // End of self-executing function
      </script>
    `;
    
    // Create HTML service object
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setTitle(hasTeam ? 'Team Management' : 'Join Team')
      .setWidth(joinOnly ? 450 : 600)
      .setHeight(joinOnly ? 400 : 550);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, hasTeam ? 'Team Management' : 'Join Team');
  } catch (e) {
    Logger.log(`Error in showTeamManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open team manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reopens the team manager after an action
 */
function reopenTeamManager() {
  showTeamManager(false);
}

/**
 * Saves team-aware column preferences
 * @param {Array} columns - Column keys to save
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {boolean} True if successful, false otherwise
 */
function saveTeamAwareColumnPreferences(columns, entityType, sheetName) {
  try {
    if (!entityType || !sheetName) {
      return false;
    }
    
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    // Generate storage key
    const key = `COLUMN_PREFS_${entityType}_${sheetName}`;
    
    if (userTeam && userTeam.settings && userTeam.settings.shareColumns) {
      // Team member with shared column preferences - store in team data
      const teamsData = getTeamsData();
      const teamId = userTeam.teamId;
      
      if (teamsData[teamId]) {
        teamsData[teamId].columnPreferences = teamsData[teamId].columnPreferences || {};
        teamsData[teamId].columnPreferences[key] = columns;
        saveTeamsData(teamsData);
      }
    } else {
      // Individual preference - store in user properties
      const userProps = PropertiesService.getUserProperties();
      userProps.setProperty(key, JSON.stringify(columns));
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTeamAwareColumnPreferences: ${e.message}`);
    return false;
  }
}

/**
 * Gets team-aware column preferences
 * @param {string} entityType - Entity type
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of column keys
 */
function getTeamAwareColumnPreferences(entityType, sheetName) {
  try {
    if (!entityType || !sheetName) {
      return [];
    }
    
    const key = `COLUMN_PREFS_${entityType}_${sheetName}`;
    
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    if (userTeam && userTeam.settings && userTeam.settings.shareColumns) {
      // Check team preferences first
      if (userTeam.columnPreferences && userTeam.columnPreferences[key]) {
        return userTeam.columnPreferences[key];
      }
    }
    
    // Fall back to user preferences
    const userProps = PropertiesService.getUserProperties();
    const prefsJson = userProps.getProperty(key);
    
    if (prefsJson) {
      return JSON.parse(prefsJson);
    }
    
    return [];
  } catch (e) {
    Logger.log(`Error in getTeamAwareColumnPreferences: ${e.message}`);
    return [];
  }
}

/**
 * Gets team-aware Pipedrive filters
 * @return {Array} Array of filter objects
 */
function getTeamAwarePipedriveFilters() {
  try {
    // Get user filters
    const userFilters = getPipedriveFilters();
    
    // Check if user is in a team with shared filters
    const userEmail = Session.getActiveUser().getEmail();
    const userTeam = getUserTeam(userEmail);
    
    // If not in a team or team doesn't share filters, return user filters
    if (!userTeam || !userTeam.settings || !userTeam.settings.shareFilters) {
      return userFilters;
    }
    
    // Return team filters if available
    return userFilters;
  } catch (e) {
    Logger.log(`Error in getTeamAwarePipedriveFilters: ${e.message}`);
    return [];
  }
}

/**
 * Saves settings for the add-on
 * @param {string} apiKey - Pipedrive API key
 * @param {string} entityType - Entity type
 * @param {string} filterId - Filter ID
 * @param {string} subdomain - Pipedrive subdomain
 * @param {string} sheetName - Sheet name
 * @param {boolean} enableTimestamp - Whether to enable timestamp
 * @return {boolean} True if successful, false otherwise
 */
function saveSettings(apiKey, entityType, filterId, subdomain, sheetName, enableTimestamp = false) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // Save the settings
    if (apiKey) docProps.setProperty('PIPEDRIVE_API_KEY', apiKey);
    if (subdomain) docProps.setProperty('PIPEDRIVE_SUBDOMAIN', subdomain);
    if (sheetName) docProps.setProperty('EXPORT_SHEET_NAME', sheetName);
    
    // Save sheet-specific entity type (this is the key fix)
    if (entityType && sheetName) {
      const sheetEntityTypeKey = `ENTITY_TYPE_${sheetName}`;
      docProps.setProperty(sheetEntityTypeKey, entityType);
      // Still save global entity type for backward compatibility
      docProps.setProperty('PIPEDRIVE_ENTITY_TYPE', entityType);
    }
    
    // Only update filter ID if provided (may be empty intentionally)
    if (filterId !== undefined && sheetName) {
      // Also save filter ID with sheet-specific key
      const sheetFilterIdKey = `FILTER_ID_${sheetName}`;
      docProps.setProperty(sheetFilterIdKey, filterId);
      // Global filter ID for backward compatibility
      docProps.setProperty('PIPEDRIVE_FILTER_ID', filterId);
    }
    
    // Handle boolean property
    docProps.setProperty('ENABLE_TIMESTAMP', enableTimestamp ? 'true' : 'false');
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveSettings: ${e.message}`);
    throw e;
  }
}

/**
 * Shows the help/about dialog
 */
function showHelp() {
  try {
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('Help');
    
    // Get version
    const versionsStr = PropertiesService.getScriptProperties().getProperty('VERSION_HISTORY') || '[]';
    const versions = JSON.parse(versionsStr);
    
    // Get current version (latest in history)
    const currentVersion = versions.length > 0 ? versions[0].version : '1.0.0';
    
    // Pass data to the template
    htmlTemplate.currentVersion = currentVersion;
    htmlTemplate.versionHistory = versions;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Help & About')
      .setWidth(600)
      .setHeight(500);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Help & About');
  } catch (e) {
    Logger.log(`Error in showHelp: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open help dialog: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows the trigger manager dialog
 */
function showTriggerManager() {
  try {
    // Get current triggers
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = PropertiesService.getDocumentProperties().getProperty('EXPORT_SHEET_NAME') || DEFAULT_SHEET_NAME;
    
    // Get all triggers for this sheet
    const triggers = getTriggersForSheet(sheetName);
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TriggerManager');
    
    // Pass data to the template
    htmlTemplate.sheetName = sheetName;
    htmlTemplate.triggers = triggers;
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Schedule Sync')
      .setWidth(600)
      .setHeight(500);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Sync');
  } catch (e) {
    Logger.log(`Error in showTriggerManager: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open trigger manager: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets triggers for a specific sheet
 * @param {string} sheetName - Sheet name
 * @return {Array} Array of trigger info objects
 */
function getTriggersForSheet(sheetName) {
  try {
    const result = [];
    const triggers = ScriptApp.getProjectTriggers();
    
    // Format each trigger into a more user-friendly object
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'syncSheetFromTrigger') {
        const triggerInfo = getTriggerInfo(trigger);
        
        // Only include if it's for this sheet or for all sheets
        if (!triggerInfo.sheetName || triggerInfo.sheetName === sheetName) {
          result.push(triggerInfo);
        }
      }
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error in getTriggersForSheet: ${e.message}`);
    return [];
  }
}

/**
 * Gets information about a trigger
 * @param {Trigger} trigger - Trigger object
 * @return {Object} Trigger info object
 */
function getTriggerInfo(trigger) {
  try {
    // Get trigger ID
    const triggerId = trigger.getUniqueId();
    
    // Get trigger properties from storage
    const userProps = PropertiesService.getUserProperties();
    const triggerPropsJson = userProps.getProperty(`TRIGGER_${triggerId}`);
    const triggerProps = triggerPropsJson ? JSON.parse(triggerPropsJson) : {};
    
    // Get trigger type
    const triggerType = trigger.getEventType();
    let schedule = '';
    
    if (triggerType === ScriptApp.EventType.CLOCK) {
      // Time-based trigger
      const hour = trigger.getAtHour();
      const minute = trigger.getAtMinute();
      const weekDay = trigger.getWeekDay();
      
      if (hour !== null && minute !== null) {
        // Daily or weekly trigger
        const timeStr = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
        
        if (weekDay !== null) {
          // Weekly trigger
          const weekDays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          schedule = `Weekly on ${weekDays[weekDay]} at ${timeStr}`;
        } else {
          // Daily trigger
          schedule = `Daily at ${timeStr}`;
        }
      } else {
        // Hourly trigger
        const frequency = trigger.getAtHour(); // Re-using this field for an hourly frequency
        schedule = `Every ${frequency} hour(s)`;
      }
    }
    
    // Combine info
    return {
      id: triggerId,
      schedule: schedule,
      sheetName: triggerProps.sheetName || 'All sheets',
      createdAt: triggerProps.createdAt || 'Unknown',
      entityType: triggerProps.entityType || 'Unknown',
      description: triggerProps.description || '',
      lastRun: triggerProps.lastRun || 'Never'
    };
  } catch (e) {
    Logger.log(`Error in getTriggerInfo: ${e.message}`);
    return {
      id: 'unknown',
      schedule: 'Unknown',
      sheetName: 'Unknown',
      createdAt: 'Unknown',
      lastRun: 'Never'
    };
  }
}

/**
 * Shows the two-way sync settings dialog
 */
function showTwoWaySyncSettings() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Get current settings
    const enableTwoWaySync = docProps.getProperty('ENABLE_TWO_WAY_SYNC') === 'true';
    const trackingColumn = docProps.getProperty('SYNC_TRACKING_COLUMN') || '';
    
    // Create a settings object with all properties the template might need
    const settings = {
      enableTwoWaySync: enableTwoWaySync,
      trackingColumn: trackingColumn,
      sheetName: activeSheetName
    };
    
    // Create the HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('TwoWaySyncSettings');
    
    // Pass data to the template
    htmlTemplate.enableTwoWaySync = enableTwoWaySync;
    htmlTemplate.trackingColumn = trackingColumn;
    htmlTemplate.settings = settings; // Pass the entire settings object
    
    // Create the HTML from the template
    const html = htmlTemplate.evaluate()
      .setTitle('Two-Way Sync Settings')
      .setWidth(500)
      .setHeight(400);
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Two-Way Sync Settings');
  } catch (e) {
    Logger.log(`Error in showTwoWaySyncSettings: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', 'Failed to open two-way sync settings: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Saves two-way sync settings
 * @param {boolean} enableTwoWaySync - Whether to enable two-way sync
 * @param {string} trackingColumn - Column letter for tracking changes
 * @return {boolean} True if successful, false otherwise
 */
function saveTwoWaySyncSettings(enableTwoWaySync, trackingColumn) {
  try {
    const docProps = PropertiesService.getDocumentProperties(); // Make sure we use DocumentProperties
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = activeSheet.getName();
    
    // Save settings with sheet-specific keys (matching original script)
    const twoWaySyncEnabledKey = `TWOWAY_SYNC_ENABLED_${activeSheetName}`;
    const twoWaySyncTrackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${activeSheetName}`;
    
    // Save the settings
    docProps.setProperty(twoWaySyncEnabledKey, enableTwoWaySync ? 'true' : 'false');
    
    // Also save global setting for backward compatibility
    docProps.setProperty('ENABLE_TWO_WAY_SYNC', enableTwoWaySync ? 'true' : 'false');
    
    // Debug log to verify what we're saving
    Logger.log(`Saving two-way sync setting: ${twoWaySyncEnabledKey} = ${enableTwoWaySync ? 'true' : 'false'}`);
    
    if (trackingColumn) {
      // Save both sheet-specific and global tracking column
      docProps.setProperty(twoWaySyncTrackingColumnKey, trackingColumn);
      docProps.setProperty('SYNC_TRACKING_COLUMN', trackingColumn);
    } else {
      // If no column specified, clean up the properties
      docProps.deleteProperty(twoWaySyncTrackingColumnKey);
      docProps.deleteProperty('SYNC_TRACKING_COLUMN');
    }
    
    // If two-way sync is enabled, set up the onEdit trigger and immediately add the Sync Status column
    if (enableTwoWaySync) {
      // Set up the onEdit trigger
      setupOnEditTrigger();
      
      // Add Sync Status column if it doesn't exist yet
      addSyncStatusColumn(activeSheet, trackingColumn);
    } else {
      removeOnEditTrigger();
      // Consider removing the Sync Status column?
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error in saveTwoWaySyncSettings: ${e.message}`);
    throw e;
  }
}

/**
 * Adds a Sync Status column to the sheet
 * @param {Sheet} sheet - The sheet to add the column to
 * @param {string} specificColumn - Optional specific column letter to use
 */
function addSyncStatusColumn(sheet, specificColumn = '') {
  try {
    // First, check if there's already a Sync Status column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let syncStatusColumnIndex = -1;
    
    // Search for existing Sync Status column
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'Sync Status') {
        syncStatusColumnIndex = i;
        break;
      }
    }
    
    // If we found an existing column, use it
    if (syncStatusColumnIndex >= 0) {
      const columnLetter = columnToLetter(syncStatusColumnIndex + 1); // +1 because it's 1-based
      Logger.log(`Found existing Sync Status column at ${columnLetter}`);
      return;
    }
    
    // If we specified a specific column, use that
    if (specificColumn) {
      const columnIndex = columnLetterToIndex(specificColumn);
      
      // Check if this column already has a header
      if (columnIndex <= sheet.getLastColumn()) {
        const existingHeader = sheet.getRange(1, columnIndex).getValue();
        if (existingHeader) {
          // If there's already a header, append to the end instead
          Logger.log(`Column ${specificColumn} already has header "${existingHeader}". Will append Sync Status to the end instead.`);
          specificColumn = '';
        }
      }
    }
    
    // Add the Sync Status column
    let targetColumnIndex;
    if (specificColumn) {
      // Use the specified column
      targetColumnIndex = columnLetterToIndex(specificColumn);
    } else {
      // Append to the end
      targetColumnIndex = sheet.getLastColumn() + 1;
    }
    
    // Set the header
    sheet.getRange(1, targetColumnIndex).setValue('Sync Status');
    
    // Format the header cell
    sheet.getRange(1, targetColumnIndex)
         .setFontWeight('bold')
         .setBackground('#E8F0FE');
    
    // Save the column location for later
    const sheetName = sheet.getName();
    const trackingColumnKey = `TWOWAY_SYNC_TRACKING_COLUMN_${sheetName}`;
    const columnLetter = columnToLetter(targetColumnIndex);
    PropertiesService.getDocumentProperties().setProperty(trackingColumnKey, columnLetter);
    
    Logger.log(`Added Sync Status column at column ${columnLetter}`);
    
    // Add conditional formatting for the Sync Status column
    const lastRow = Math.max(sheet.getLastRow(), 100); // Format at least 100 rows
    if (lastRow > 1) {
      const statusRange = sheet.getRange(2, targetColumnIndex, lastRow - 1, 1);
      const rules = sheet.getConditionalFormatRules();
      
      // Create rule for "Modified" cells
      const modifiedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Modified')
        .setBackground('#FCE8E6') // Light red
        .setBold(true)
        .setRanges([statusRange])
        .build();
      
      // Create rule for "Synced" cells
      const syncedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Synced')
        .setBackground('#E6F4EA') // Light green
        .setRanges([statusRange])
        .build();
      
      // Create rule for "Error" cells
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('Error')
        .setBackground('#FCE8E6') // Light red
        .setBold(true)
        .setRanges([statusRange])
        .build();
      
      // Add the new rules to the existing rules
      rules.push(modifiedRule);
      rules.push(syncedRule);
      rules.push(errorRule);
      sheet.setConditionalFormatRules(rules);
    }
    
    // Add data validation for the Sync Status column
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Modified', 'Synced'], true)
      .build();
    
    if (lastRow > 1) {
      sheet.getRange(2, targetColumnIndex, lastRow - 1, 1).setDataValidation(validationRule);
    }
    
    return true;
  } catch (e) {
    Logger.log(`Error adding Sync Status column: ${e.message}`);
    return false;
  }
}

/**
 * Removes the onEdit trigger for two-way sync
 */
function removeOnEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('onEdit trigger removed');
      break;
    }
  }
}

// Make sure this exists in your Utilities.gs file or add it here
/**
 * Converts a column index to letter format (e.g., 1 = A, 27 = AA)
 * @param {number} column - The column index (1-based)
 * @return {string} The column letter
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Converts a column letter to index format (e.g., A = 1, AA = 27)
 * @param {string} letter - The column letter
 * @return {number} The column index (1-based)
 */
function columnLetterToIndex(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * Shows a simple dialog for joining a team
 * This is shown when a user without a team tries to use the add-on
 */
function showTeamJoinRequest() {
  try {
    const html = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        .container {
          max-width: 400px;
          margin: 0 auto;
        }
        h3 {
          margin-top: 0;
          color: #4285F4;
        }
        p {
          margin-bottom: 15px;
        }
        .form-group {
          margin-bottom: 15px;
        }
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: bold;
        }
        input[type="text"] {
          width: 100%;
          padding: 8px;
          border: 1px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
        }
        button {
          background-color: #4285F4;
          color: white;
          border: none;
          padding: 8px 15px;
          border-radius: 4px;
          cursor: pointer;
        }
        button:hover {
          background-color: #3367D6;
        }
        .error {
          color: #d32f2f;
          margin-top: 5px;
        }
        .options {
          margin-top: 20px;
          text-align: center;
        }
        .options a {
          color: #4285F4;
          text-decoration: none;
        }
        .options a:hover {
          text-decoration: underline;
        }
      </style>
      <div class="container">
        <h3>Join a Team</h3>
        <p>Enter a team ID to join an existing team:</p>
        
        <div class="form-group">
          <label for="team-id">Team ID</label>
          <input type="text" id="team-id" placeholder="Enter team ID">
          <div id="error-message" class="error"></div>
        </div>
        
        <button id="join-button">Join Team</button>
        
        <div class="options">
          <p>Don't have a team ID? <a href="#" id="create-team-link">Create a new team</a></p>
        </div>
      </div>
      
      <script>
        // Join team button handler
        document.getElementById('join-button').addEventListener('click', function() {
          // Get team ID
          const teamId = document.getElementById('team-id').value.trim();
          if (!teamId) {
            document.getElementById('error-message').textContent = 'Please enter a team ID';
            return;
          }
          
          // Call server function
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                // Refresh the page to show the full menu
                window.top.location.reload();
              } else {
                document.getElementById('error-message').textContent = result.message || 'Error joining team';
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('error-message').textContent = 'Error: ' + error.message;
            })
            .joinTeam(teamId);
        });
        
        // Create team link handler
        document.getElementById('create-team-link').addEventListener('click', function(e) {
          e.preventDefault();
          
          // Show the team management dialog in create mode
          google.script.run.showTeamManager(false);
          google.script.host.close();
        });
      </script>
    `)
    .setWidth(450)
    .setHeight(300)
    .setTitle('Join Team');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Join Team');
  } catch (e) {
    Logger.log('Error in showTeamJoinRequest: ' + e.message);
    SpreadsheetApp.getUi().alert('Error', 'Failed to show join request: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}