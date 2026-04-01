/**
 * @file SheetActions_Operations.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Sheet Operations (NEW)
 * ============================================
 * 
 * Sheet-level operations including:
 * - Freeze panes (rows/columns)
 * - Hide/show rows/columns
 * - Insert/delete rows/columns
 * - Clear content/format/validation
 * - Sort data
 * - Row height / Column width
 * - Auto-resize columns
 * - Sheet properties (name, tab color)
 * - Group rows/columns
 */

var SheetActions_Operations = (function() {
  
  // ============================================
  // MAIN EXECUTE FUNCTION
  // ============================================
  
  /**
   * Execute sheet operation(s)
   * @param {Object} step - { config: { operation: "...", ... } } OR { config: { operations: [...] } }
   * @return {Object} Result with operation details
   */
  function execute(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    // Handle MULTIPLE operations if provided as an array
    if (config.operations && Array.isArray(config.operations)) {
      Logger.log('[SheetActions_Operations] Executing ' + config.operations.length + ' operations');
      var results = [];
      for (var i = 0; i < config.operations.length; i++) {
        var opConfig = config.operations[i];
        var opResult = _executeSingleOperation(sheet, opConfig);
        results.push(opResult);
      }
      return { operations: results, count: results.length };
    }
    
    // Single operation
    return _executeSingleOperation(sheet, config);
  }
  
  /**
   * Execute a single operation
   * @param {Sheet} sheet - Active sheet
   * @param {Object} config - Operation config
   * @return {Object} Result
   */
  function _executeSingleOperation(sheet, config) {
    var operation = config.operation;
    
    Logger.log('[SheetActions_Operations] Operation: ' + operation);
    
    switch (operation) {
      // ===== FREEZE PANES =====
      case 'freezeRows':
        sheet.setFrozenRows(config.rows || 1);
        return { frozen: 'rows', count: config.rows || 1 };
        
      case 'freezeColumns':
        sheet.setFrozenColumns(config.columns || 1);
        return { frozen: 'columns', count: config.columns || 1 };
        
      case 'freeze':
        // Freeze both rows and columns
        if (config.rows) sheet.setFrozenRows(config.rows);
        if (config.columns) sheet.setFrozenColumns(config.columns);
        return { frozen: { rows: config.rows || 0, columns: config.columns || 0 } };
        
      case 'unfreeze':
      case 'unfreezeRows':  // SDK Agent alias
      case 'unfreezeColumns':  // SDK Agent alias
        // Unfreeze based on operation name
        if (operation === 'unfreezeRows') {
          sheet.setFrozenRows(0);
          return { unfrozen: 'rows' };
        } else if (operation === 'unfreezeColumns') {
          sheet.setFrozenColumns(0);
          return { unfrozen: 'columns' };
        } else {
          // Full unfreeze
          sheet.setFrozenRows(0);
          sheet.setFrozenColumns(0);
          return { unfrozen: true };
        }
      
      // ===== HIDE/SHOW ROWS =====
      case 'hideRows':
        sheet.hideRows(config.startRow, config.numRows || 1);
        return { hidden: 'rows', start: config.startRow, count: config.numRows || 1 };
        
      case 'showRows':
        sheet.showRows(config.startRow, config.numRows || 1);
        return { shown: 'rows', start: config.startRow, count: config.numRows || 1 };
      
      // ===== HIDE/SHOW COLUMNS =====
      case 'hideColumns':
        var colIndex = _getColIndex(config.startColumn);
        sheet.hideColumns(colIndex, config.numColumns || 1);
        return { hidden: 'columns', start: config.startColumn, count: config.numColumns || 1 };
        
      case 'showColumns':
        var colIndex = _getColIndex(config.startColumn);
        sheet.showColumns(colIndex, config.numColumns || 1);
        return { shown: 'columns', start: config.startColumn, count: config.numColumns || 1 };
      
      // ===== INSERT ROWS/COLUMNS =====
      case 'insertRows':
        return _insertRows(sheet, config);
        
      case 'insertColumns':
        return _insertColumns(sheet, config);
      
      // ===== DELETE ROWS/COLUMNS =====
      case 'deleteRows':
        sheet.deleteRows(config.startRow, config.count || 1);
        return { deleted: 'rows', start: config.startRow, count: config.count || 1 };
        
      case 'deleteColumns':
        var colIndex = _getColIndex(config.startColumn);
        sheet.deleteColumns(colIndex, config.count || 1);
        return { deleted: 'columns', start: config.startColumn, count: config.count || 1 };
      
      // ===== CLEAR OPERATIONS =====
      case 'clear':
        var range = sheet.getRange(config.range);
        range.clear();
        return { cleared: 'all', range: config.range };
        
      case 'clearContent':
      case 'clearContents':
        var range = sheet.getRange(config.range);
        range.clearContent();
        return { cleared: 'content', range: config.range };
        
      case 'clearFormat':
      case 'clearFormatting':
        var range = sheet.getRange(config.range);
        range.clearFormat();
        return { cleared: 'format', range: config.range };
        
      case 'clearValidation':
      case 'clearDataValidations':
        var range = sheet.getRange(config.range);
        range.clearDataValidations();
        return { cleared: 'validation', range: config.range };
        
      case 'clearNotes':
        var range = sheet.getRange(config.range);
        range.clearNote();
        return { cleared: 'notes', range: config.range };
      
      // ===== SORT =====
      case 'sort':
        return _sortRange(sheet, config);
      
      // ===== ROW/COLUMN DIMENSIONS =====
      case 'resizeRows':  // SDK Agent alias
      case 'setRowHeight':
        var rowNum = config.row || config.rows || 1;
        var height = config.height || config.size || 21;
        sheet.setRowHeight(rowNum, height);
        return { rowHeight: height, row: rowNum };
        
      case 'setRowHeights':
        // Set height for multiple rows
        for (var r = config.startRow; r <= (config.endRow || config.startRow); r++) {
          sheet.setRowHeight(r, config.height);
        }
        return { rowHeight: config.height, startRow: config.startRow, endRow: config.endRow };
        
      case 'resizeColumns':  // SDK Agent alias
      case 'setColumnWidth':
        var colIndex = _getColIndex(config.column || config.columns || 'A');
        var width = config.width || config.size || 100;
        sheet.setColumnWidth(colIndex, width);
        return { columnWidth: width, column: config.column || config.columns };
        
      case 'setColumnWidths':
        // Set width for multiple columns
        var startCol = _getColIndex(config.startColumn);
        var endCol = _getColIndex(config.endColumn || config.startColumn);
        for (var c = startCol; c <= endCol; c++) {
          sheet.setColumnWidth(c, config.width);
        }
        return { columnWidth: config.width, startColumn: config.startColumn, endColumn: config.endColumn };
        
      case 'autoResizeColumn':
        var colIndex = _getColIndex(config.column);
        sheet.autoResizeColumn(colIndex);
        return { autoResized: 'column', column: config.column };
        
      case 'autoResizeColumns':
        return _autoResizeColumns(sheet, config);
        
      case 'autoResizeRows':
        sheet.autoResizeRows(config.startRow, config.numRows || 1);
        return { autoResized: 'rows', startRow: config.startRow, numRows: config.numRows || 1 };
      
      // ===== SHEET PROPERTIES =====
      case 'renameSheet':
        sheet.setName(config.name);
        return { renamed: true, name: config.name };
        
      case 'setTabColor':
        sheet.setTabColor(config.color);
        return { tabColor: config.color };
        
      case 'clearTabColor':
        sheet.setTabColor(null);
        return { tabColor: null };
      
      // ===== GROUPING =====
      case 'groupRows':
        return _groupRows(sheet, config);
        
      case 'groupColumns':
        return _groupColumns(sheet, config);
        
      case 'ungroupRows':
        var range = sheet.getRange(config.startRow + ':' + (config.endRow || config.startRow));
        range.shiftRowGroupDepth(-1);
        return { ungrouped: 'rows', start: config.startRow, end: config.endRow };
        
      case 'ungroupColumns':
        var startCol = _getColIndex(config.startColumn);
        var endCol = _getColIndex(config.endColumn || config.startColumn);
        var range = sheet.getRange(1, startCol, 1, endCol - startCol + 1);
        range.shiftColumnGroupDepth(-1);
        return { ungrouped: 'columns', start: config.startColumn, end: config.endColumn };
      
      // ===== PROTECTION (basic) =====
      case 'protect':  // SDK Agent alias
      case 'protectRange':
        // If range is provided, protect that range; otherwise protect header row (A1:lastCol1)
        var rangeToProtect = config.range || 'A1:' + String.fromCharCode(64 + sheet.getLastColumn()) + '1';
        var range = sheet.getRange(rangeToProtect);
        var protection = range.protect();
        if (config.description) protection.setDescription(config.description);
        if (config.warningOnly !== false) protection.setWarningOnly(true); // Default to warning only
        return { protected: true, range: rangeToProtect };
        
      case 'protectSheet':
        var protection = sheet.protect();
        if (config.description) protection.setDescription(config.description);
        return { protected: true, sheet: sheet.getName() };
        
      case 'unprotect':
        // Remove all protections from the sheet
        var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        protections.forEach(function(p) {
          if (!p.getDescription().includes('system')) { // Don't remove system protections
            p.remove();
          }
        });
        return { unprotected: true, count: protections.length };
      
      default:
        throw new Error('Unknown sheet operation: ' + operation);
    }
  }
  
  // ============================================
  // HELPER FUNCTIONS
  // ============================================
  
  /**
   * Get column index from string or number
   */
  function _getColIndex(column) {
    return typeof column === 'string' ? SheetActions_Utils.letterToColumn(column) : column;
  }
  
  /**
   * Insert rows
   */
  function _insertRows(sheet, config) {
    var count = config.count || 1;
    
    if (config.after) {
      sheet.insertRowsAfter(config.after, count);
    } else if (config.before) {
      sheet.insertRowsBefore(config.before, count);
    } else {
      // Default: insert at end
      sheet.insertRowsAfter(sheet.getLastRow(), count);
    }
    
    return { inserted: 'rows', count: count, after: config.after, before: config.before };
  }
  
  /**
   * Insert columns
   */
  function _insertColumns(sheet, config) {
    var count = config.count || 1;
    
    if (config.after) {
      var afterCol = _getColIndex(config.after);
      sheet.insertColumnsAfter(afterCol, count);
    } else if (config.before) {
      var beforeCol = _getColIndex(config.before);
      sheet.insertColumnsBefore(beforeCol, count);
    }
    
    return { inserted: 'columns', count: count, after: config.after, before: config.before };
  }
  
  /**
   * Sort range
   */
  function _sortRange(sheet, config) {
    var rangeStr = config.range || SheetActions_Utils.detectDataRange(sheet);
    var range = sheet.getRange(rangeStr);
    var sortSpec = config.sortBy || config.columns || [{ column: 1, ascending: true }];
    
    // Normalize sortSpec
    if (!Array.isArray(sortSpec)) sortSpec = [sortSpec];
    sortSpec = sortSpec.map(function(s) {
      return {
        column: typeof s.column === 'string' ? SheetActions_Utils.letterToColumn(s.column) : s.column,
        ascending: s.ascending !== false
      };
    });
    
    range.sort(sortSpec);
    
    return { sorted: true, range: rangeStr, sortBy: sortSpec };
  }
  
  /**
   * Auto-resize multiple columns
   */
  function _autoResizeColumns(sheet, config) {
    var startCol = _getColIndex(config.startColumn);
    var endCol = config.endColumn ? _getColIndex(config.endColumn) : startCol;
    
    for (var c = startCol; c <= endCol; c++) {
      sheet.autoResizeColumn(c);
    }
    
    return { autoResized: 'columns', start: config.startColumn, end: config.endColumn || config.startColumn };
  }
  
  /**
   * Group rows
   */
  function _groupRows(sheet, config) {
    var startRow = config.startRow;
    var endRow = config.endRow || startRow;
    var range = sheet.getRange(startRow + ':' + endRow);
    
    range.shiftRowGroupDepth(1);
    
    // Collapse if requested
    if (config.collapse) {
      // Note: collapse() is on the Group object, not directly available here
      // This is a limitation - would need to get the group first
      Logger.log('[SheetActions_Operations] Group created but collapse requires getting the group object');
    }
    
    return { grouped: 'rows', start: startRow, end: endRow };
  }
  
  /**
   * Group columns
   */
  function _groupColumns(sheet, config) {
    var startCol = _getColIndex(config.startColumn);
    var endCol = config.endColumn ? _getColIndex(config.endColumn) : startCol;
    var range = sheet.getRange(1, startCol, 1, endCol - startCol + 1);
    
    range.shiftColumnGroupDepth(1);
    
    return { grouped: 'columns', start: config.startColumn, end: config.endColumn || config.startColumn };
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    execute: execute
  };
  
})();
