/**
 * @file SheetActions_Main.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * AGENT SHEET ACTIONS - Main Entry Point
 * ============================================
 * 
 * Executes native Google Sheets operations INSTANTLY:
 * - Charts (bar, column, line, pie, area, scatter, combo)
 * - Formatting (currency, percent, number, date, styles, borders, banding)
 * - Conditional formatting (rules-based highlighting, gradient/color scale)
 * - Data validation (dropdown, number, checkbox, date, text, custom)
 * - Filtering (criteria-based row filtering)
 * - Data writing (writeData, createTable)
 * - Sheet operations (freeze, hide/show, insert/delete, sort, etc.)
 * 
 * KEY: These actions do NOT use AI jobs - they execute immediately
 * 
 * ARCHITECTURE:
 * This file dispatches to specialized modules:
 * - SheetActions_Chart.gs - Chart creation
 * - SheetActions_Format.gs - Cell formatting
 * - SheetActions_Validation.gs - Data validation
 * - SheetActions_ConditionalFormat.gs - Conditional formatting
 * - SheetActions_Filter.gs - Filtering
 * - SheetActions_Data.gs - Data writing and tables
 * - SheetActions_Operations.gs - Sheet-level operations
 * - SheetActions_Utils.gs - Shared utilities
 */

var AgentSheetActions = (function() {
  
  // Supported sheet action types
  // NOTE: Must match SDK Agent tools in ai-sheet-backend/src/lib/agents/tools/index.ts
  var SHEET_ACTIONS = [
    'chart',
    'format',
    'conditionalFormat',
    'dataValidation',
    'filter',
    'writeData',
    'createTable',  // Legacy name
    'table',        // SDK Agent name (maps to createTable)
    'updateTable',  // Modify existing table range/properties
    'deleteTable',  // Delete a table
    'appendToTable', // Append rows to existing table
    'sheetOps'
  ];
  
  /**
   * Check if action is a sheet action (not AI processing)
   * @param {string} action - Action type
   * @return {boolean}
   */
  function isSheetAction(action) {
    return SHEET_ACTIONS.indexOf(action) !== -1;
  }
  
  /**
   * Execute a sheet action
   * @param {Object} step - Step configuration from AI
   * @return {Object} { success, result, instant: true } or { success: false, error }
   */
  function execute(step) {
    var action = step.action;
    Logger.log('[SheetActions] Executing: ' + action);
    
    try {
      var result;
      switch (action) {
        case 'chart':
          result = SheetActions_Chart.createChart(step);
          break;
        case 'format':
          result = SheetActions_Format.applyFormat(step);
          break;
        case 'conditionalFormat':
          result = SheetActions_ConditionalFormat.apply(step);
          break;
        case 'dataValidation':
          result = SheetActions_Validation.create(step);
          break;
        case 'filter':
          result = SheetActions_Filter.apply(step);
          break;
        case 'writeData':
          result = SheetActions_Data.writeData(step);
          break;
        case 'createTable':
        case 'table':  // SDK Agent alias for createTable
          result = SheetActions_Data.createTable(step);
          break;
        case 'updateTable':
          result = SheetActions_Data.updateTable(step);
          break;
        case 'deleteTable':
          result = SheetActions_Data.deleteTable(step);
          break;
        case 'appendToTable':
          result = SheetActions_Data.appendToTable(step);
          break;
        case 'sheetOps':
          result = SheetActions_Operations.execute(step);
          break;
        default:
          return { success: false, error: 'Unknown sheet action: ' + action };
      }
      
      Logger.log('[SheetActions] Completed: ' + action);
      return {
        success: true,
        action: action,
        result: result,
        instant: true  // No polling needed - completed immediately
      };
    } catch (e) {
      Logger.log('[SheetActions] Error: ' + e.message);
      return { success: false, error: e.message };
    }
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    isSheetAction: isSheetAction,
    execute: execute,
    SHEET_ACTIONS: SHEET_ACTIONS
  };
  
})();

// ============================================
// GLOBAL FUNCTION FOR FRONTEND
// ============================================

/**
 * Execute a sheet action from the frontend
 * Called by google.script.run.executeSheetAction(step)
 * 
 * @param {Object} step - Step configuration from frontend
 * @return {Object} Result { success: boolean, action?: string, result?: Object, error?: string }
 */
function executeSheetAction(step) {
  Logger.log('[executeSheetAction] Called with: ' + JSON.stringify(step).substring(0, 300));
  
  try {
    // Validate step
    if (!step || !step.action) {
      return { success: false, error: 'No action specified' };
    }
    
    // Check if it's a valid sheet action
    if (!AgentSheetActions.isSheetAction(step.action)) {
      return { success: false, error: 'Invalid sheet action: ' + step.action };
    }
    
    // Execute the sheet action
    var result = AgentSheetActions.execute(step);
    
    Logger.log('[executeSheetAction] Result: ' + JSON.stringify(result));
    return result;
    
  } catch (e) {
    Logger.log('[executeSheetAction] Error: ' + e.message);
    return { success: false, error: e.message };
  }
}
