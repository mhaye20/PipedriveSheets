/**
 * Column Selector UI Helper
 * 
 * This module provides helper functions for the column selector UI:
 * - Styles for the column selector dialog
 * - Scripts for the column selector dialog
 */

var ColumnSelectorUI = ColumnSelectorUI || {};

/**
 * Gets the styles for the column selector dialog
 * @return {string} CSS styles
 */
ColumnSelectorUI.getStyles = function() {
  return `
    <style>
      :root {
        --primary-color: #4285f4;
        --primary-dark: #3367d6;
        --success-color: #0f9d58;
        --warning-color: #f4b400;
        --error-color: #db4437;
        --text-dark: #202124;
        --text-light: #5f6368;
        --bg-light: #f8f9fa;
        --border-color: #dadce0;
        --shadow: 0 1px 3px rgba(60,64,67,0.15);
        --shadow-hover: 0 4px 8px rgba(60,64,67,0.2);
        --transition: all 0.2s ease;
      }
      
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      
      body {
        font-family: 'Roboto', Arial, sans-serif;
        color: var(--text-dark);
        line-height: 1.5;
        margin: 0;
        padding: 16px;
        font-size: 14px;
      }
      
      .header {
        margin-bottom: 16px;
      }
      
      h3 {
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 12px;
        color: var(--text-dark);
      }
      
      .sheet-info {
        background-color: var(--bg-light);
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 14px;
        border-left: 4px solid var(--primary-color);
        display: flex;
        align-items: center;
      }
      
      .sheet-info svg {
        margin-right: 12px;
        fill: var(--primary-color);
      }
      
      .info {
        font-size: 13px;
        color: var(--text-light);
        margin-bottom: 16px;
      }
      
      .container {
        display: flex;
        height: 400px;
        gap: 16px;
      }
      
      .column {
        width: 50%;
        display: flex;
        flex-direction: column;
      }
      
      .column-header {
        font-weight: 500;
        margin-bottom: 8px;
        padding: 0 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      
      .column-count {
        font-size: 12px;
        color: var(--text-light);
        font-weight: normal;
      }
      
      .search {
        padding: 10px 12px;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        margin-bottom: 12px;
        font-size: 14px;
        width: 100%;
        transition: var(--transition);
      }
      
      .search:focus {
        outline: none;
        border-color: var(--primary-color);
        box-shadow: 0 0 0 2px rgba(66,133,244,0.2);
      }
      
      .scrollable {
        flex-grow: 1;
        overflow-y: auto;
        border: 1px solid var(--border-color);
        border-radius: 4px;
        background-color: white;
      }
      
      .item {
        padding: 8px 12px;
        margin: 2px 4px;
        cursor: pointer;
        border-radius: 4px;
        transition: var(--transition);
        display: flex;
        align-items: center;
      }
      
      .item:hover {
        background-color: rgba(66,133,244,0.08);
      }
      
      .item.selected {
        background-color: #e8f0fe;
        position: relative;
      }
      
      .item.selected:hover {
        background-color: #d2e3fc;
      }
      
      .category {
        font-weight: 500;
        padding: 8px 12px;
        background-color: var(--bg-light);
        margin: 8px 4px 4px 4px;
        border-radius: 4px;
        color: var(--text-dark);
        font-size: 13px;
      }
      
      .nested {
        margin-left: 16px;
        position: relative;
      }
      
      .nested::before {
        content: '';
        position: absolute;
        left: -8px;
        top: 50%;
        width: 6px;
        height: 1px;
        background-color: var(--border-color);
      }
      
      .footer {
        margin-top: 16px;
        display: flex;
        justify-content: space-between;
      }
      
      button {
        padding: 10px 16px;
        border: none;
        border-radius: 4px;
        font-weight: 500;
        cursor: pointer;
        font-size: 14px;
        transition: var(--transition);
      }
      
      button:focus {
        outline: none;
      }
      
      button:disabled {
        opacity: 0.7;
        cursor: not-allowed;
      }
      
      .primary-btn {
        background-color: var(--primary-color);
        color: white;
      }
      
      .primary-btn:hover {
        background-color: var(--primary-dark);
        box-shadow: var(--shadow-hover);
      }
      
      .secondary-btn {
        background-color: transparent;
        color: var(--primary-color);
      }
      
      .secondary-btn:hover {
        background-color: rgba(66,133,244,0.08);
      }
      
      .action-btns {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      
      .drag-handle {
        display: inline-block;
        width: 12px;
        height: 20px;
        background-image: radial-gradient(circle, var(--text-light) 1px, transparent 1px);
        background-size: 3px 3px;
        background-position: center;
        background-repeat: repeat;
        margin-right: 8px;
        cursor: grab;
        opacity: 0.5;
      }
      
      .selected:hover .drag-handle {
        opacity: 0.8;
      }
      
      .column-text {
        flex-grow: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      
      .add-btn, .remove-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
      }
      
      .add-btn {
        color: var(--success-color);
        background-color: rgba(15,157,88,0.1);
      }
      
      .remove-btn {
        color: var(--error-color);
        background-color: rgba(219,68,55,0.1);
      }
      
      .rename-btn {
        opacity: 0;
        margin-left: 4px;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        transition: var(--transition);
        flex-shrink: 0;
        color: var(--primary-color);
        background-color: rgba(66,133,244,0.1);
      }
      
      .item:hover .add-btn,
      .item:hover .remove-btn,
      .item:hover .rename-btn {
        opacity: 1;
      }
      
      .loading {
        display: none;
        align-items: center;
        margin-right: 8px;
      }
      
      .loader {
        display: inline-block;
        width: 18px;
        height: 18px;
        border: 2px solid rgba(66,133,244,0.2);
        border-radius: 50%;
        border-top-color: var(--primary-color);
        animation: spin 1s ease-in-out infinite;
      }
      
      .dragging {
        opacity: 0.4;
        box-shadow: var(--shadow-hover);
      }
      
      .over {
        border-top: 2px solid var(--primary-color);
      }
      
      .indicator {
        display: none;
        padding: 12px 16px;
        border-radius: 4px;
        margin-bottom: 16px;
        font-weight: 500;
      }
      
      .indicator.success {
        background-color: rgba(15,157,88,0.1);
        color: var(--success-color);
        border-left: 4px solid var(--success-color);
      }
      
      .indicator.error {
        background-color: rgba(219,68,55,0.1);
        color: var(--error-color);
        border-left: 4px solid var(--error-color);
      }
      
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
  `;
};

/**
 * Gets the scripts for the column selector dialog
 * @return {string} JavaScript code
 */
ColumnSelectorUI.getScripts = function() {
  return `
    <script>
      // DOM elements
      const availableList = document.getElementById('availableList');
      const selectedList = document.getElementById('selectedList');
      const searchBox = document.getElementById('searchBox');
      const availableCountEl = document.getElementById('availableCount');
      const selectedCountEl = document.getElementById('selectedCount');
      
      // Create a deep copy of the original available columns to ensure we don't lose data
      const originalAvailableColumns = JSON.parse(JSON.stringify(availableColumns));
      
      // Render the lists
      function renderAvailableList(searchTerm = '') {
        availableList.innerHTML = '';
        
        // Group columns by parent key or top-level
        const topLevel = [];
        const nested = {};
        let availableCount = 0;
        
        // Always use the original available columns array as source of truth
        originalAvailableColumns.forEach(col => {
          // Check if this column is already selected
          if (!selectedColumns.some(selected => selected.key === col.key)) {
            availableCount++;
            
            if (!searchTerm || col.name.toLowerCase().includes(searchTerm.toLowerCase())) {
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
        
        // Update available count
        availableCountEl.textContent = '(' + availableCount + ')';
        
        // Add top-level columns first
        if (topLevel.length > 0) {
          const topLevelHeader = document.createElement('div');
          topLevelHeader.className = 'category';
          topLevelHeader.textContent = 'Main Fields';
          availableList.appendChild(topLevelHeader);
          
          topLevel.forEach(col => {
            const item = document.createElement('div');
            item.className = 'item';
            item.dataset.key = col.key;
            
            const columnText = document.createElement('span');
            columnText.className = 'column-text';
            columnText.textContent = col.name;
            item.appendChild(columnText);
            
            const addBtn = document.createElement('span');
            addBtn.className = 'add-btn';
            addBtn.innerHTML = '+';
            addBtn.title = 'Add column';
            addBtn.onclick = (e) => {
              e.stopPropagation();
              addColumn(col);
            };
            item.appendChild(addBtn);
            
            item.onclick = () => addColumn(col);
            availableList.appendChild(item);
          });
        }
        
        // Then add nested columns by parent
        for (const parentKey in nested) {
          if (nested[parentKey].length > 0) {
            // Find the parent name from the original available columns
            const parentName = originalAvailableColumns.find(col => col.key === parentKey)?.name || parentKey;
            
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'category';
            categoryHeader.textContent = parentName;
            availableList.appendChild(categoryHeader);
            
            nested[parentKey].forEach(col => {
              const item = document.createElement('div');
              item.className = 'item nested';
              item.dataset.key = col.key;
              
              const columnText = document.createElement('span');
              columnText.className = 'column-text';
              columnText.textContent = col.name;
              item.appendChild(columnText);
              
              const addBtn = document.createElement('span');
              addBtn.className = 'add-btn';
              addBtn.innerHTML = '+';
              addBtn.title = 'Add column';
              addBtn.onclick = (e) => {
                e.stopPropagation();
                addColumn(col);
              };
              item.appendChild(addBtn);
              
              item.onclick = () => addColumn(col);
              availableList.appendChild(item);
            });
          }
        }
        
        // Show "no results" message if nothing matches search
        if (availableList.children.length === 0 && searchTerm) {
          const noResults = document.createElement('div');
          noResults.style.padding = '16px';
          noResults.style.textAlign = 'center';
          noResults.style.color = 'var(--text-light)';
          noResults.textContent = 'No matching columns found';
          availableList.appendChild(noResults);
        }
      }
      
      function renderSelectedList() {
        selectedList.innerHTML = '';
        
        // Update selected count
        selectedCountEl.textContent = '(' + selectedColumns.length + ')';
        
        if (selectedColumns.length === 0) {
          const emptyState = document.createElement('div');
          emptyState.style.padding = '16px';
          emptyState.style.textAlign = 'center';
          emptyState.style.color = 'var(--text-light)';
          emptyState.innerHTML = 'No columns selected yet<br>Select columns from the left panel';
          selectedList.appendChild(emptyState);
          return;
        }
        
        selectedColumns.forEach((col, index) => {
          const item = document.createElement('div');
          item.className = 'item selected';
          item.dataset.key = col.key;
          item.dataset.index = index;
          item.draggable = true;
          
          // Add drag handle
          const dragHandle = document.createElement('span');
          dragHandle.className = 'drag-handle';
          item.appendChild(dragHandle);
          
          // Add column name
          const columnText = document.createElement('span');
          columnText.className = 'column-text';
          columnText.textContent = col.customName || col.name;
          if (col.customName) {
            columnText.title = "Original field: " + col.name;
            columnText.style.fontStyle = 'italic';
          }
          item.appendChild(columnText);
          
          // Add remove button
          const removeBtn = document.createElement('span');
          removeBtn.className = 'remove-btn';
          removeBtn.innerHTML = '❌';
          removeBtn.title = 'Remove column';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            removeColumn(col);
          };
          item.appendChild(removeBtn);
          
          // Add rename button
          const renameBtn = document.createElement('span');
          renameBtn.className = 'rename-btn';
          renameBtn.innerHTML = '✏️';
          renameBtn.title = 'Rename column';
          renameBtn.onclick = (e) => {
            e.stopPropagation();
            renameColumn(index);
          };
          item.appendChild(renameBtn);
          
          // Set up drag events
          item.ondragstart = handleDragStart;
          item.ondragover = handleDragOver;
          item.ondrop = handleDrop;
          item.ondragend = handleDragEnd;
          
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
        // Find the column in the original available columns array to ensure we have complete data
        const fullColumn = originalAvailableColumns.find(col => col.key === column.key);
        if (fullColumn) {
          // Add a deep copy to prevent reference issues
          selectedColumns.push(JSON.parse(JSON.stringify(fullColumn)));
          renderAvailableList(searchBox.value);
          renderSelectedList();
          showStatus('success', 'Column added: ' + fullColumn.name);
        }
      }
      
      function removeColumn(column) {
        selectedColumns = selectedColumns.filter(col => col.key !== column.key);
        renderAvailableList(searchBox.value);
        renderSelectedList();
        showStatus('success', 'Column removed: ' + column.name);
      }
      
      function renameColumn(columnIndex) {
        const column = selectedColumns[columnIndex];
        const currentName = column.customName || column.name;
        const newName = prompt("Enter custom header name for '" + column.name + "'", currentName);
        
        if (newName !== null) {
          // Update the column with the custom name
          selectedColumns[columnIndex].customName = newName;
          renderSelectedList();
          showStatus('success', 'Column renamed to: ' + newName);
        }
      }
      
      function showStatus(type, message) {
        const indicator = document.getElementById('statusIndicator');
        indicator.className = 'indicator ' + type;
        indicator.textContent = message;
        indicator.style.display = 'block';
        
        // Auto-hide after a delay
        setTimeout(function() {
          indicator.style.display = 'none';
        }, 2000);
      }
      
      // Event listeners
      document.getElementById('saveBtn').onclick = () => {
        if (selectedColumns.length === 0) {
          showStatus('error', 'Please select at least one column');
          return;
        }
        
        // Show loading animation
        document.getElementById('saveLoading').style.display = 'flex';
        document.getElementById('saveBtn').disabled = true;
        document.getElementById('cancelBtn').disabled = true;
        
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('saveLoading').style.display = 'none';
            showStatus('success', 'Column preferences saved successfully!');
            
            // Close after a short delay to show the success message
            setTimeout(() => {
              google.script.host.close();
            }, 1500);
          })
          .withFailureHandler((error) => {
            document.getElementById('saveLoading').style.display = 'none';
            document.getElementById('saveBtn').disabled = false;
            document.getElementById('cancelBtn').disabled = false;
            showStatus('error', 'Error: ' + error.message);
          })
          .saveColumnPreferences(selectedColumns, entityType, sheetName);
      };
      
      document.getElementById('cancelBtn').onclick = () => {
        google.script.host.close();
      };
      
      document.getElementById('helpBtn').onclick = () => {
        const helpContent = 'Tips for selecting columns:\\n' +
                           '\\n• Search for specific columns using the search box' +
                           '\\n• Main fields are top-level Pipedrive fields' +
                           '\\n• Nested fields provide more detailed information' +
                           '\\n• Drag and drop to reorder selected columns' +
                           '\\n• The column order here determines the order in your sheet';
                           
        alert(helpContent);
      };
      
      searchBox.oninput = () => {
        renderAvailableList(searchBox.value);
      };
      
      // Initial render
      renderAvailableList();
      renderSelectedList();
    </script>
  `;
}; 