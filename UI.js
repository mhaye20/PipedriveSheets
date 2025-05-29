/**
 * User Interface
 * 
 * This module handles core UI functionality:
 * - Template processing
 * - HTML file inclusion
 * - Error handling
 */

// Use a single global declaration pattern
if (typeof UI === 'undefined') {
  var UI = {};
}

/**
 * Shows an error message to the user in a dialog
 * @param {string} message - The error message to display
 */
UI.showError = function(message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', message, ui.ButtonSet.OK);
  } catch (error) {
  }
};

/**
 * Process a template string with PHP-like syntax using a data object
 * @param {string} template - Template string with PHP-like syntax
 * @param {Object} data - Data object containing values to substitute
 * @return {string} - Processed template
 */
UI.processTemplate = function(template, data) {
  let processedTemplate = template;
  
  // Handle conditionals
  processedTemplate = UI.processIfElse(processedTemplate, data);
  
  // Handle foreach loops
  processedTemplate = UI.processForEach(processedTemplate, data);
  
  // Process simple variable substitutions like <?= variable ?>
  processedTemplate = processedTemplate.replace(/\<\?=\s*([^?>]+?)\s*\?>/g, (match, variable) => {
    try {
      // Handle nested properties like member.email or complex expressions
      const value = UI.evalInContext(variable, data) || '';
      return value;
    } catch (e) {
      return '';
    }
  });
  
  return processedTemplate;
};

/**
 * Process if/else statements in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
UI.processIfElse = function(template, data) {
  let result = template;
  
  // Handle if statements: <?php if (condition): ?>content<?php endif; ?>
  const IF_PATTERN = /\<\?php\s+if\s+\((.+?)\)\s*:\s*\?>([\s\S]*?)(?:\<\?php\s+else\s*:\s*\?>([\s\S]*?))?\<\?php\s+endif;\s*\?>/g;
  
  result = result.replace(IF_PATTERN, (match, condition, ifContent, elseContent = '') => {
    try {
      // Convert PHP-like condition to JavaScript
      let jsCondition = condition
        .replace(/\!\=/g, '!==')
        .replace(/\=\=/g, '===')
        .replace(/\!([\w\.]+)/g, '!$1');
      
      // Evaluate the condition in the context of the data object
      const conditionResult = UI.evalInContext(jsCondition, data);
      
      return conditionResult ? ifContent : elseContent;
    } catch (e) {
      return '';
    }
  });
  
  return result;
};

/**
 * Process foreach loops in PHP-like template syntax
 * @param {string} template - The template content
 * @param {Object} data - Data object
 * @return {string} - Processed template
 */
UI.processForEach = function(template, data) {
  let result = template;
  
  // Handle foreach loops: <?php foreach ($items as $item): ?>content<?php endforeach; ?>
  const FOREACH_PATTERN = /\<\?php\s+foreach\s+\(\$(\w+)\s+as\s+\$(\w+)\)\s*:\s*\?>([\s\S]*?)\<\?php\s+endforeach;\s*\?>/g;
  
  result = result.replace(FOREACH_PATTERN, (match, collection, item, content) => {
    try {
      const items = data[collection];
      if (!Array.isArray(items) || items.length === 0) {
        return '';
      }
      
      return items.map(itemData => {
        // Create a context with the item for nested variable replacement
        const itemContext = Object.assign({}, data, { [item]: itemData });
        
        // Replace item variables within the loop content
        let itemContent = content;
        
        // Process variables like <?= item.property ?>
        itemContent = itemContent.replace(/\<\?=\s*(\$?)([\w\.]+)\s*\?>/g, (m, dollar, varName) => {
          try {
            // If it's a loop item variable like $member or $member.property
            if (varName.startsWith(item + '.')) {
              const propPath = varName.substring(item.length + 1);
              return UI.evalPropertyPath(itemData, propPath) || '';
            } 
            // If it's just $member (the whole item)
            else if (varName === item) {
              return itemData || '';
            } 
            // Other variables from the parent context
            else {
              return UI.evalInContext(varName, itemContext) || '';
            }
          } catch (e) {
            return '';
          }
        });
        
        return itemContent;
      }).join('');
    } catch (e) {
      return '';
    }
  });
  
  return result;
};

/**
 * Evaluate a JavaScript expression in the context of a data object
 * @param {string} expr - The expression to evaluate
 * @param {Object} context - The context object
 * @return {*} - The result of the evaluation
 */
UI.evalInContext = function(expr, context) {
  try {
    // Handle simple variable access first
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(expr)) {
      return context[expr];
    }
    
    // Handle nested properties with dot notation
    if (/^[a-zA-Z_$][a-zA-Z0-9_$]*(\.[a-zA-Z_$][a-zA-Z0-9_$]*)+$/.test(expr)) {
      return UI.evalPropertyPath(context, expr);
    }
    
    // Handle comparisons and more complex expressions
    // Create a safe function to evaluate in the context
    const keys = Object.keys(context);
    const values = keys.map(key => context[key]);
    const evaluator = new Function(...keys, `return ${expr};`);
    return evaluator(...values);
  } catch (e) {
    return null;
  }
};

/**
 * Evaluate a property path on an object (e.g. "user.profile.name")
 * @param {Object} obj - The object to evaluate on
 * @param {string} path - The property path
 * @return {*} - The value at the path or undefined
 */
UI.evalPropertyPath = function(obj, path) {
  try {
    return path.split('.').reduce((o, p) => o && o[p], obj);
  } catch (e) {
    return undefined;
  }
};

/**
 * Includes an HTML file with the content of a script file.
 * @param {string} filename - The name of the file to include without the extension
 * @return {string} The content to be included
 */
UI.include = function(filename) {
  return HtmlService.createHtmlOutput(UI.getFileContent_(filename)).getContent();
};

/**
 * Gets the content of a script file.
 * @param {string} filename - The name of the file without the extension
 * @return {string} The content of the file
 * @private
 */
UI.getFileContent_ = function(filename) {
  if (filename === 'TeamManagerUI') {
    return `<style>${TeamManagerUI.getStyles()}</style><script>${TeamManagerUI.getScripts()}</script>`;
  } else if (filename === 'TriggerManagerUI') {
    const styles = HtmlService.createHtmlOutputFromFile('TriggerManagerUI_Styles').getContent();
    const scripts = HtmlService.createHtmlOutputFromFile('TriggerManagerUI_Scripts').getContent();
    return `${styles}${scripts}`;
  } else if (filename === 'TwoWaySyncSettingsUI') {
    return `<style>${TwoWaySyncSettingsUI.getStyles()}</style><script>${TwoWaySyncSettingsUI.getScripts()}</script>`;
  }
  return '';
};

// Export functions to be globally accessible
this.showError = UI.showError;
this.processTemplate = UI.processTemplate;
this.include = UI.include;