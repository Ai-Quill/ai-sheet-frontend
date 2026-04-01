/**
 * @file SheetActions_ConditionalFormat.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Conditional Formatting
 * ============================================
 * 
 * Applies conditional formatting rules including:
 * - Number conditions (greater than, less than, equals, between, etc.)
 * - Text conditions (contains, starts with, ends with, equals)
 * - Date conditions (before, after, equals, relative dates)
 * - Cell state (empty, not empty)
 * - Custom formulas
 * - Gradient / Color scale (NEW)
 */

var SheetActions_ConditionalFormat = (function() {
  
  // ============================================
  // MAIN CONDITIONAL FORMAT FUNCTION
  // ============================================
  
  /**
   * Apply conditional formatting rules (supports comma-separated ranges)
   * @param {Object} step - { range, config: { rules: [] } }
   * @return {Object} Result with rules count
   */
  function apply(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    var rangeStr = config.range || step.range || step.inputRange;
    var rules = config.rules || step.rules || [];
    var builtRules = [];
    
    Logger.log('[SheetActions_ConditionalFormat] === APPLY CONDITIONAL FORMAT ===');
    Logger.log('[SheetActions_ConditionalFormat] Range: ' + rangeStr);
    Logger.log('[SheetActions_ConditionalFormat] Rules count: ' + rules.length);
    Logger.log('[SheetActions_ConditionalFormat] Rules: ' + JSON.stringify(rules).substring(0, 500));
    
    if (!rules || rules.length === 0) {
      Logger.log('[SheetActions_ConditionalFormat] ⚠️ NO RULES provided! Nothing to apply.');
      return { rulesAdded: 0, range: rangeStr, warning: 'No rules provided' };
    }
    
    // Handle multiple comma-separated ranges
    var rangeStrings = rangeStr.split(',').map(function(r) { return r.trim(); }).filter(Boolean);
    Logger.log('[SheetActions_ConditionalFormat] Processing ' + rangeStrings.length + ' range(s)');
    
    rangeStrings.forEach(function(singleRangeStr) {
      var range = sheet.getRange(singleRangeStr);
      
      rules.forEach(function(rule, ruleIdx) {
        Logger.log('[SheetActions_ConditionalFormat] Rule ' + ruleIdx + ': ' + JSON.stringify(rule).substring(0, 200));
        
        // Check if this is a gradient/color scale rule (NEW)
        if (rule.type === 'gradient' || rule.type === 'colorScale') {
          var gradientRule = _buildGradientRule(range, rule);
          if (gradientRule) {
            builtRules.push(gradientRule);
            Logger.log('[SheetActions_ConditionalFormat] Rule ' + ruleIdx + ': ✅ Gradient rule built');
          }
          return;
        }
        
        // Standard conditional format rule
        var builder = SpreadsheetApp.newConditionalFormatRule().setRanges([range]);
        
        // Apply condition
        var conditionApplied = _applyCondition(builder, rule);
        
        if (conditionApplied) {
          // Apply format
          _applyFormat(builder, rule);
          builtRules.push(builder.build());
          Logger.log('[SheetActions_ConditionalFormat] Rule ' + ruleIdx + ': ✅ Applied (bg=' + (rule.backgroundColor || rule.background || 'none') + ')');
        } else {
          Logger.log('[SheetActions_ConditionalFormat] Rule ' + ruleIdx + ': ❌ Condition not applied');
        }
      });
    });
    
    // Add to existing rules (don't replace)
    var existingRules = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existingRules.concat(builtRules));
    
    return { rulesAdded: builtRules.length, range: rangeStr };
  }
  
  // ============================================
  // CONDITION BUILDERS
  // ============================================
  
  /**
   * Apply condition to the builder
   * @param {ConditionalFormatRuleBuilder} builder
   * @param {Object} rule - Rule with condition and value
   * @return {boolean} Whether condition was applied
   */
  function _applyCondition(builder, rule) {
    // Support both 'condition' (GAS style) and 'type' (SDK Agent style)
    var condition = (rule.condition || rule.type || '').toLowerCase().replace(/[_\s]/g, '');
    var value = rule.value;
    
    Logger.log('[SheetActions_ConditionalFormat] Processing rule: ' + JSON.stringify(rule));
    Logger.log('[SheetActions_ConditionalFormat] Normalized condition: "' + condition + '", value: ' + value);
    
    switch (condition) {
      // ===== NUMBER CONDITIONS =====
      case 'greaterthan':
      case 'gt':
        builder.whenNumberGreaterThan(value);
        return true;
        
      case 'greaterthanorequal':
      case 'greaterthanorequalto':
      case 'gte':
        builder.whenNumberGreaterThanOrEqualTo(value);
        return true;
        
      case 'lessthan':
      case 'lt':
        builder.whenNumberLessThan(value);
        return true;
        
      case 'lessthanorequal':
      case 'lessthanorequalto':
      case 'lte':
        builder.whenNumberLessThanOrEqualTo(value);
        return true;
        
      case 'equals':
      case 'equal':
      case 'eq':
      case 'equalto':  // SDK Agent alias
        if (typeof value === 'number') {
          builder.whenNumberEqualTo(value);
        } else {
          builder.whenTextEqualTo(String(value));
        }
        return true;
        
      // Number not equal (NEW)
      case 'notequals':
      case 'notequal':
      case 'neq':
      case 'numbernotequal':
        if (typeof value === 'number') {
          builder.whenNumberNotEqualTo(value);
        } else {
          builder.whenTextDoesNotContain(String(value));
        }
        return true;
        
      case 'between':
        if (rule.min !== undefined && rule.max !== undefined) {
          builder.whenNumberBetween(rule.min, rule.max);
          return true;
        }
        return false;
        
      // Number not between (NEW)
      case 'notbetween':
      case 'numbernotbetween':
        if (rule.min !== undefined && rule.max !== undefined) {
          builder.whenNumberNotBetween(rule.min, rule.max);
          return true;
        }
        return false;
        
      case 'negative':
        builder.whenNumberLessThan(0);
        return true;
        
      case 'positive':
        builder.whenNumberGreaterThan(0);
        return true;
        
      // ===== TEXT CONDITIONS =====
      case 'contains':
      case 'textcontains':  // SDK Agent alias
        builder.whenTextContains(value);
        return true;
        
      case 'notcontains':
      case 'doesnotcontain':
        builder.whenTextDoesNotContain(value);
        return true;
        
      case 'startswith':
        builder.whenTextStartsWith(value);
        return true;
        
      case 'endswith':
        builder.whenTextEndsWith(value);
        return true;
        
      // ===== DATE CONDITIONS (NEW) =====
      case 'dateafter':
        builder.whenDateAfter(new Date(value));
        return true;
        
      case 'datebefore':
        builder.whenDateBefore(new Date(value));
        return true;
        
      case 'dateequal':
      case 'dateequalto':
        builder.whenDateEqualTo(new Date(value));
        return true;
        
      // Relative dates (NEW)
      case 'dateistoday':
      case 'today':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.TODAY);
        return true;
        
      case 'dateistomorrow':
      case 'tomorrow':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.TOMORROW);
        return true;
        
      case 'dateisyesterday':
      case 'yesterday':
        builder.whenDateEqualTo(SpreadsheetApp.RelativeDate.YESTERDAY);
        return true;
        
      case 'dateinpastweek':
      case 'pastweek':
        builder.whenDateAfter(SpreadsheetApp.RelativeDate.PAST_WEEK);
        return true;
        
      case 'dateinpastmonth':
      case 'pastmonth':
        builder.whenDateAfter(SpreadsheetApp.RelativeDate.PAST_MONTH);
        return true;
        
      case 'dateinpastyear':
      case 'pastyear':
        builder.whenDateAfter(SpreadsheetApp.RelativeDate.PAST_YEAR);
        return true;
        
      // ===== CELL STATE =====
      case 'isempty':
      case 'empty':
      case 'blank':
        builder.whenCellEmpty();
        return true;
        
      case 'isnotempty':
      case 'notempty':
      case 'notblank':
        builder.whenCellNotEmpty();
        return true;
        
      // ===== SPECIAL CONDITIONS =====
      case 'max':
      case 'ismax':
      case 'maximum':
      case 'highest':
        Logger.log('[SheetActions_ConditionalFormat] MAX condition - using custom formula');
        builder.whenFormulaSatisfied('=INDIRECT(ADDRESS(ROW(),COLUMN()))=MAX(INDIRECT(ADDRESS(3,COLUMN())&":"&ADDRESS(14,COLUMN())))');
        return true;
        
      case 'min':
      case 'ismin':
      case 'minimum':
      case 'lowest':
        Logger.log('[SheetActions_ConditionalFormat] MIN condition - using custom formula');
        builder.whenFormulaSatisfied('=INDIRECT(ADDRESS(ROW(),COLUMN()))=MIN(INDIRECT(ADDRESS(3,COLUMN())&":"&ADDRESS(14,COLUMN())))');
        return true;
        
      case 'formula':
      case 'customformula':
        // SDK Agent uses 'formula' field, GAS style uses 'value'
        var formulaStr = rule.formula || value;
        if (formulaStr && typeof formulaStr === 'string') {
          Logger.log('[SheetActions_ConditionalFormat] Custom formula: ' + formulaStr);
          builder.whenFormulaSatisfied(formulaStr);
          return true;
        }
        return false;
        
      default:
        // Check if value is "MAX" or "MIN" as a special case
        if (String(value).toUpperCase() === 'MAX') {
          builder.whenFormulaSatisfied('=INDIRECT(ADDRESS(ROW(),COLUMN()))=MAX(INDIRECT(ADDRESS(3,COLUMN())&":"&ADDRESS(14,COLUMN())))');
          return true;
        } else if (String(value).toUpperCase() === 'MIN') {
          builder.whenFormulaSatisfied('=INDIRECT(ADDRESS(ROW(),COLUMN()))=MIN(INDIRECT(ADDRESS(3,COLUMN())&":"&ADDRESS(14,COLUMN())))');
          return true;
        }
        Logger.log('[SheetActions_ConditionalFormat] Unknown condition: ' + rule.condition);
        return false;
    }
  }
  
  /**
   * Apply format to the builder
   * @param {ConditionalFormatRuleBuilder} builder
   * @param {Object} rule - Rule with format options
   */
  function _applyFormat(builder, rule) {
    var format = rule.format || rule;
    
    if (format.backgroundColor || format.background) {
      builder.setBackground(format.backgroundColor || format.background);
    }
    if (format.textColor || format.fontColor || format.color) {
      builder.setFontColor(format.textColor || format.fontColor || format.color);
    }
    if (format.bold) builder.setBold(true);
    if (format.italic) builder.setItalic(true);
    if (format.strikethrough) builder.setStrikethrough(true);
    if (format.underline) builder.setUnderline(true);
  }
  
  // ============================================
  // GRADIENT / COLOR SCALE (NEW)
  // ============================================
  
  /**
   * Build a gradient/color scale rule
   * @param {Range} range - Range to apply gradient to
   * @param {Object} rule - Rule with gradient options
   * @return {ConditionalFormatRule|null}
   */
  function _buildGradientRule(range, rule) {
    var builder = SpreadsheetApp.newConditionalFormatRule().setRanges([range]);
    
    try {
      // Two-color scale
      if (rule.minColor && rule.maxColor && !rule.midColor) {
        builder.setGradientMinpoint(rule.minColor)
               .setGradientMaxpoint(rule.maxColor);
        Logger.log('[SheetActions_ConditionalFormat] Two-color gradient: ' + rule.minColor + ' to ' + rule.maxColor);
        return builder.build();
      }
      
      // Three-color scale
      if (rule.minColor && rule.midColor && rule.maxColor) {
        var minType = _getInterpolationType(rule.minType || 'MIN');
        var midType = _getInterpolationType(rule.midType || 'PERCENTILE');
        var maxType = _getInterpolationType(rule.maxType || 'MAX');
        
        builder.setGradientMinpointWithValue(rule.minColor, minType, rule.minValue || '')
               .setGradientMidpointWithValue(rule.midColor, midType, rule.midValue || '50')
               .setGradientMaxpointWithValue(rule.maxColor, maxType, rule.maxValue || '');
        Logger.log('[SheetActions_ConditionalFormat] Three-color gradient applied');
        return builder.build();
      }
      
      Logger.log('[SheetActions_ConditionalFormat] Gradient rule missing required colors');
      return null;
    } catch (e) {
      Logger.log('[SheetActions_ConditionalFormat] Error building gradient: ' + e.message);
      return null;
    }
  }
  
  /**
   * Get interpolation type enum from string
   */
  function _getInterpolationType(type) {
    var types = {
      'MIN': SpreadsheetApp.InterpolationType.MIN,
      'MAX': SpreadsheetApp.InterpolationType.MAX,
      'NUMBER': SpreadsheetApp.InterpolationType.NUMBER,
      'PERCENT': SpreadsheetApp.InterpolationType.PERCENT,
      'PERCENTILE': SpreadsheetApp.InterpolationType.PERCENTILE
    };
    return types[type.toUpperCase()] || SpreadsheetApp.InterpolationType.MIN;
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    apply: apply
  };
  
})();
