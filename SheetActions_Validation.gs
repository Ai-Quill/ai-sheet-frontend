/**
 * @file SheetActions_Validation.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Data Validation
 * ============================================
 * 
 * Creates data validation rules on cells/ranges including:
 * - Dropdown lists (from values or range)
 * - Number validation (range, equals, greater than, etc.)
 * - Date validation (range, equals, before, after, etc.)
 * - Text validation (contains, equals, email, URL)
 * - Checkbox
 * - Custom formula
 * - Help text support
 */

var SheetActions_Validation = (function() {
  
  // ============================================
  // MAIN VALIDATION FUNCTION
  // ============================================
  
  /**
   * Create data validation on a range (supports comma-separated ranges)
   * @param {Object} step - { validationType, range, config }
   * @return {Object} Result with validation details
   */
  function create(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    var rangeStr = config.range || step.range || step.inputRange;
    var validationType = (config.validationType || step.validationType || '').toLowerCase();
    var options = config.options || config || {};
    var helpText = options.helpText || config.helpText;
    
    // Handle multiple comma-separated ranges
    var rangeStrings = rangeStr.split(',').map(function(r) { return r.trim(); }).filter(Boolean);
    Logger.log('[SheetActions_Validation] type=' + validationType + ', ranges=' + rangeStrings.length);
    
    rangeStrings.forEach(function(singleRangeStr) {
      var range = sheet.getRange(singleRangeStr);
      var validationBuilder = _buildValidation(validationType, options, config, step);
      
      if (validationBuilder) {
        // Add help text if provided (NEW)
        if (helpText) {
          validationBuilder.setHelpText(helpText);
        }
        
        var validation = validationBuilder.build();
        range.setDataValidation(validation);
        Logger.log('[SheetActions_Validation] Validation applied to ' + singleRangeStr);
      }
    });
    
    return { validationType: validationType, range: rangeStr, rangeCount: rangeStrings.length };
  }
  
  // ============================================
  // VALIDATION BUILDERS
  // ============================================
  
  /**
   * Build validation based on type
   * @param {string} validationType - Type of validation
   * @param {Object} options - Validation options
   * @param {Object} config - Full config object
   * @param {Object} step - Full step object
   * @return {DataValidationBuilder|null} Validation builder or null
   */
  function _buildValidation(validationType, options, config, step) {
    var builder = SpreadsheetApp.newDataValidation();
    
    switch (validationType) {
      // ===== DROPDOWN / LIST =====
      case 'dropdown':
      case 'list':
        var values = options.values || config.values || step.values || [];
        if (values.length > 0) {
          builder.requireValueInList(values, true)
                 .setAllowInvalid(options.allowInvalid !== false);
          return builder;
        }
        return null;
      
      // Dropdown from range (NEW)
      case 'rangedropdown':
      case 'listfromrange':
      case 'dropdownfromrange':
        var sourceRange = options.sourceRange || config.sourceRange;
        if (sourceRange) {
          var sourceRangeObj = SpreadsheetApp.getActiveSpreadsheet().getRange(sourceRange);
          builder.requireValueInRange(sourceRangeObj, true)
                 .setAllowInvalid(options.allowInvalid !== false);
          return builder;
        }
        return null;
      
      // ===== NUMBER VALIDATIONS =====
      case 'number':
      case 'numeric':
      case 'numberrange':
        return _buildNumberBetween(builder, options, config);
      
      // Number equals (NEW)
      case 'numberequal':
      case 'numberequalto':
        var value = options.value !== undefined ? options.value : config.value;
        if (value !== undefined) {
          builder.requireNumberEqualTo(value);
          return builder;
        }
        return null;
      
      // Number not equal (NEW)
      case 'numbernotequal':
      case 'numbernotequalto':
        var value = options.value !== undefined ? options.value : config.value;
        if (value !== undefined) {
          builder.requireNumberNotEqualTo(value);
          return builder;
        }
        return null;
      
      // Number greater than (NEW)
      case 'numbergreaterthan':
      case 'numbergt':
        var value = options.value !== undefined ? options.value : config.value;
        if (value !== undefined) {
          builder.requireNumberGreaterThan(value);
          return builder;
        }
        return null;
      
      // Number greater than or equal (existing but explicit)
      case 'numbergreaterthanorequal':
      case 'numbergte':
        var value = options.value !== undefined ? options.value : (options.min !== undefined ? options.min : config.min);
        if (value !== undefined) {
          builder.requireNumberGreaterThanOrEqualTo(value);
          return builder;
        }
        return null;
      
      // Number less than (NEW)
      case 'numberlessthan':
      case 'numberlt':
        var value = options.value !== undefined ? options.value : config.value;
        if (value !== undefined) {
          builder.requireNumberLessThan(value);
          return builder;
        }
        return null;
      
      // Number less than or equal (existing but explicit)
      case 'numberlessthanorequal':
      case 'numberlte':
        var value = options.value !== undefined ? options.value : (options.max !== undefined ? options.max : config.max);
        if (value !== undefined) {
          builder.requireNumberLessThanOrEqualTo(value);
          return builder;
        }
        return null;
      
      // Number not between (NEW)
      case 'numbernotbetween':
        var min = options.min !== undefined ? options.min : config.min;
        var max = options.max !== undefined ? options.max : config.max;
        if (min !== undefined && max !== undefined) {
          builder.requireNumberNotBetween(min, max);
          return builder;
        }
        return null;
      
      // ===== CHECKBOX =====
      case 'checkbox':
      case 'boolean':
        var customTrue = options.checkedValue || options.trueValue;
        var customFalse = options.uncheckedValue || options.falseValue;
        
        if (customTrue && customFalse) {
          builder.requireCheckbox(customTrue, customFalse);
        } else {
          builder.requireCheckbox();
        }
        return builder;
      
      // ===== DATE VALIDATIONS =====
      case 'date':
        return _buildDateBetween(builder, options, config);
      
      // Date equals (NEW)
      case 'dateequal':
      case 'dateequalto':
        var dateValue = options.date || config.date || options.value;
        if (dateValue) {
          builder.requireDateEqualTo(new Date(dateValue));
          return builder;
        }
        return null;
      
      // Date on or after (NEW)
      case 'dateonorafter':
        var dateValue = options.date || config.date || options.value || options.min;
        if (dateValue) {
          builder.requireDateOnOrAfter(new Date(dateValue));
          return builder;
        }
        return null;
      
      // Date on or before (NEW)
      case 'dateonorbefore':
        var dateValue = options.date || config.date || options.value || options.max;
        if (dateValue) {
          builder.requireDateOnOrBefore(new Date(dateValue));
          return builder;
        }
        return null;
      
      // Date not between (NEW)
      case 'datenotbetween':
        var start = options.start || config.start || options.min;
        var end = options.end || config.end || options.max;
        if (start && end) {
          builder.requireDateNotBetween(new Date(start), new Date(end));
          return builder;
        }
        return null;
      
      // ===== TEXT VALIDATIONS (NEW) =====
      case 'textcontains':
        var text = options.text || config.text || options.value;
        if (text) {
          builder.requireTextContains(text);
          return builder;
        }
        return null;
      
      case 'textnotcontains':
      case 'textdoesnotcontain':
        var text = options.text || config.text || options.value;
        if (text) {
          builder.requireTextDoesNotContain(text);
          return builder;
        }
        return null;
      
      case 'textequal':
      case 'textequalto':
        var text = options.text || config.text || options.value;
        if (text) {
          builder.requireTextEqualTo(text);
          return builder;
        }
        return null;
      
      // ===== TEXT LENGTH (SDK Agent) =====
      case 'textlength':
        // Requires text length to be between min and max
        var min = options.min !== undefined ? options.min : config.min;
        var max = options.max !== undefined ? options.max : config.max;
        if (min !== undefined && max !== undefined) {
          // Use custom formula for text length validation
          builder.requireFormulaSatisfied('=AND(LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))>=' + min + ',LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))<=' + max + ')');
          return builder;
        } else if (min !== undefined) {
          builder.requireFormulaSatisfied('=LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))>=' + min);
          return builder;
        } else if (max !== undefined) {
          builder.requireFormulaSatisfied('=LEN(INDIRECT(ADDRESS(ROW(),COLUMN())))<=' + max);
          return builder;
        }
        return null;
      
      // ===== EMAIL =====
      case 'email':
        builder.requireTextIsEmail();
        return builder;
      
      // ===== URL =====
      case 'url':
      case 'link':
        builder.requireTextIsUrl();
        return builder;
      
      // ===== CUSTOM FORMULA =====
      case 'custom':
      case 'formula':
      case 'customformula':  // SDK Agent alias
        var formula = options.formula || config.formula;
        if (formula) {
          builder.requireFormulaSatisfied(formula);
          return builder;
        }
        return null;
      
      default:
        Logger.log('[SheetActions_Validation] Unknown validation type: ' + validationType);
        return null;
    }
  }
  
  /**
   * Build number between validation
   */
  function _buildNumberBetween(builder, options, config) {
    var min = options.min !== undefined ? options.min : 
              config.min !== undefined ? config.min :
              options.minValue !== undefined ? options.minValue :
              config.minValue;
    var max = options.max !== undefined ? options.max : 
              config.max !== undefined ? config.max :
              options.maxValue !== undefined ? options.maxValue :
              config.maxValue;
    
    if (min !== undefined && max !== undefined) {
      builder.requireNumberBetween(min, max);
    } else if (min !== undefined) {
      builder.requireNumberGreaterThanOrEqualTo(min);
    } else if (max !== undefined) {
      builder.requireNumberLessThanOrEqualTo(max);
    } else {
      // Default: any number
      builder.requireNumberGreaterThanOrEqualTo(-999999999);
    }
    return builder;
  }
  
  /**
   * Build date between validation
   */
  function _buildDateBetween(builder, options, config) {
    var afterDate = options.after || options.min;
    var beforeDate = options.before || options.max;
    
    if (afterDate && beforeDate) {
      builder.requireDateBetween(new Date(afterDate), new Date(beforeDate));
    } else if (afterDate) {
      builder.requireDateAfter(new Date(afterDate));
    } else if (beforeDate) {
      builder.requireDateBefore(new Date(beforeDate));
    } else {
      builder.requireDate();
    }
    return builder;
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    create: create
  };
  
})();
