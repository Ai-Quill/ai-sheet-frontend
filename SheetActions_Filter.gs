/**
 * @file SheetActions_Filter.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Filtering
 * ============================================
 * 
 * Applies filter criteria to data including:
 * - Text conditions (equals, contains, starts with, ends with)
 * - Number conditions (greater than, less than, between, etc.)
 * - Date conditions (before, after, equals) - NEW
 * - Array-based filters (equals any, not equals any) - NEW
 * - Custom formula filters - NEW
 * - Hidden/visible values - NEW
 */

var SheetActions_Filter = (function() {
  
  // ============================================
  // MAIN FILTER FUNCTION
  // ============================================
  
  /**
   * Apply filter to data
   * @param {Object} step - { dataRange, config: { criteria: [] } }
   * @return {Object} Result with filter details
   */
  function apply(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    var rangeStr = config.range || config.dataRange || step.dataRange || step.range || step.inputRange || SheetActions_Utils.detectDataRange(sheet);
    
    // CRITICAL: Filter must include header row (row 1)
    // If range starts at row 2 (e.g., A2:G31), adjust to start at row 1 (A1:G31)
    rangeStr = _ensureHeaderRowIncluded(rangeStr);
    
    var dataRange = sheet.getRange(rangeStr);
    
    Logger.log('[SheetActions_Filter] range=' + rangeStr);
    
    // Remove existing filter
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    
    // Create new filter
    var filter = dataRange.createFilter();
    
    // Apply criteria
    var criteria = config.criteria || step.criteria || [];
    criteria.forEach(function(criterion) {
      var colIndex = SheetActions_Utils.getColumnIndex(criterion.column);
      var filterCriteria = _buildFilterCriteria(criterion);
      
      if (filterCriteria && colIndex > 0) {
        filter.setColumnFilterCriteria(colIndex, filterCriteria);
        Logger.log('[SheetActions_Filter] Filter applied to column ' + colIndex);
      }
    });
    
    return { filterApplied: true, criteriaCount: criteria.length, range: rangeStr };
  }
  
  // ============================================
  // HELPER FUNCTIONS
  // ============================================
  
  /**
   * Ensure the filter range includes the header row (row 1)
   * Filters MUST start at the header row to work correctly
   * If range is "A2:G31", converts to "A1:G31"
   * @param {string} rangeStr - Range in A1 notation
   * @return {string} Adjusted range that includes row 1
   */
  function _ensureHeaderRowIncluded(rangeStr) {
    if (!rangeStr) return rangeStr;
    
    // Parse the range (e.g., "A2:G31" -> start="A2", end="G31")
    var parts = rangeStr.split(':');
    if (parts.length !== 2) return rangeStr;  // Not a proper range
    
    var startCell = parts[0];
    var endCell = parts[1];
    
    // Extract row number from start cell (e.g., "A2" -> row 2)
    var startRowMatch = startCell.match(/(\d+)$/);
    if (!startRowMatch) return rangeStr;
    
    var startRow = parseInt(startRowMatch[1], 10);
    
    // If start row is > 1, adjust to row 1
    if (startRow > 1) {
      var startCol = startCell.replace(/\d+$/, '');  // Extract column letters
      var adjustedRange = startCol + '1:' + endCell;
      Logger.log('[SheetActions_Filter] Adjusted range from ' + rangeStr + ' to ' + adjustedRange + ' to include header row');
      return adjustedRange;
    }
    
    return rangeStr;
  }
  
  // ============================================
  // FILTER CRITERIA BUILDER
  // ============================================
  
  /**
   * Build filter criteria from criterion object
   * @param {Object} criterion - { column, condition, value }
   * @return {FilterCriteria|null}
   */
  function _buildFilterCriteria(criterion) {
    // Support both 'condition' (GAS style) and 'operator' (SDK Agent style)
    var condition = (criterion.condition || criterion.operator || '').toLowerCase().replace(/[_\s]/g, '');
    var value = criterion.value;
    
    // Also support minValue/maxValue from SDK Agent
    if (criterion.minValue !== undefined) criterion.min = criterion.minValue;
    if (criterion.maxValue !== undefined) criterion.max = criterion.maxValue;
    
    Logger.log('[SheetActions_Filter] Criterion: column=' + criterion.column + ', condition=' + condition + ', value=' + value);
    
    switch (condition) {
      // ===== TEXT CONDITIONS =====
      case 'equals':
      case 'equal':
      case 'eq':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextEqualTo(String(value))
          .build();
        
      case 'notequals':
      case 'notequal':
      case 'neq':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextDoesNotContain(String(value))
          .build();
        
      case 'contains':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextContains(value)
          .build();
        
      case 'notcontains':
      case 'doesnotcontain':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextDoesNotContain(value)
          .build();
        
      case 'startswith':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextStartsWith(value)
          .build();
        
      case 'endswith':
        return SpreadsheetApp.newFilterCriteria()
          .whenTextEndsWith(value)
          .build();
      
      // ===== NUMBER CONDITIONS =====
      case 'greaterthan':
      case 'gt':
        return SpreadsheetApp.newFilterCriteria()
          .whenNumberGreaterThan(value)
          .build();
        
      case 'greaterthanorequal':
      case 'greaterthanorequalto':
      case 'gte':
        return SpreadsheetApp.newFilterCriteria()
          .whenNumberGreaterThanOrEqualTo(value)
          .build();
        
      case 'lessthan':
      case 'lt':
        return SpreadsheetApp.newFilterCriteria()
          .whenNumberLessThan(value)
          .build();
        
      case 'lessthanorequal':
      case 'lessthanorequalto':
      case 'lte':
        return SpreadsheetApp.newFilterCriteria()
          .whenNumberLessThanOrEqualTo(value)
          .build();
        
      case 'between':
        if (criterion.min !== undefined && criterion.max !== undefined) {
          return SpreadsheetApp.newFilterCriteria()
            .whenNumberBetween(criterion.min, criterion.max)
            .build();
        }
        return null;
      
      // ===== CELL STATE =====
      case 'isempty':
      case 'empty':
      case 'blank':
        return SpreadsheetApp.newFilterCriteria()
          .whenCellEmpty()
          .build();
        
      case 'isnotempty':
      case 'notempty':
      case 'notblank':
        return SpreadsheetApp.newFilterCriteria()
          .whenCellNotEmpty()
          .build();
      
      // ===== DATE CONDITIONS (NEW) =====
      case 'dateafter':
        return SpreadsheetApp.newFilterCriteria()
          .whenDateAfter(new Date(value))
          .build();
        
      case 'datebefore':
        return SpreadsheetApp.newFilterCriteria()
          .whenDateBefore(new Date(value))
          .build();
        
      case 'dateequal':
      case 'dateequalto':
        return SpreadsheetApp.newFilterCriteria()
          .whenDateEqualTo(new Date(value))
          .build();
        
      case 'datenotequal':
      case 'datenotequalto':
        return SpreadsheetApp.newFilterCriteria()
          .whenDateNotEqualTo(new Date(value))
          .build();
      
      // ===== ARRAY-BASED FILTERS (NEW) =====
      case 'equalsany':
      case 'inlist':
      case 'invaluelist':
        var values = Array.isArray(value) ? value : [value];
        return SpreadsheetApp.newFilterCriteria()
          .whenTextEqualToAny(values)
          .build();
        
      case 'notequalsany':
      case 'notinlist':
        var values = Array.isArray(value) ? value : [value];
        return SpreadsheetApp.newFilterCriteria()
          .whenTextNotEqualToAny(values)
          .build();
      
      // ===== HIDDEN/VISIBLE VALUES (NEW) =====
      case 'hidevalues':
      case 'hide':
        var hiddenValues = Array.isArray(value) ? value : [value];
        return SpreadsheetApp.newFilterCriteria()
          .setHiddenValues(hiddenValues)
          .build();
        
      case 'showonlyvalues':
      case 'showonly':
      case 'visible':
        var visibleValues = Array.isArray(value) ? value : [value];
        return SpreadsheetApp.newFilterCriteria()
          .setVisibleValues(visibleValues)
          .build();
      
      // ===== CUSTOM FORMULA (NEW) =====
      case 'formula':
      case 'customformula':
        return SpreadsheetApp.newFilterCriteria()
          .whenFormulaSatisfied(value)
          .build();
        
      default:
        Logger.log('[SheetActions_Filter] Unknown filter condition: ' + condition);
        return null;
    }
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    apply: apply
  };
  
})();
