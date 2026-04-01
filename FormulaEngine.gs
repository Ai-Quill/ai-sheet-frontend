/**
 * @file FormulaEngine.gs
 * @description Intelligent Formula Composition Engine for Google Sheets
 * 
 * This is the main entry point for formula generation. It orchestrates
 * all the FormulaCatalog modules and provides intelligent formula composition.
 * 
 * ARCHITECTURE:
 * - FormulaCatalog_Lookup.gs - Lookup & Reference functions
 * - FormulaCatalog_Text.gs - Text manipulation functions
 * - FormulaCatalog_Math.gs - Math & statistics functions
 * - FormulaCatalog_Logic.gs - Logical & information functions
 * - FormulaCatalog_DateTime.gs - Date & time functions
 * - FormulaCatalog_Array.gs - Array & LAMBDA functions
 * - FormulaCatalog_Google.gs - Google-specific functions
 * - FormulaCatalog_Financial.gs - Financial functions
 * - FormulaCatalog_Engineering.gs - Engineering functions
 */

var FormulaEngine = (function() {
  'use strict';
  
  // ========== CATALOG REGISTRY ==========
  
  /**
   * Get function info from any catalog
   */
  function getFunctionInfo(functionName) {
    var upper = functionName.toUpperCase();
    
    // Search all catalogs
    var catalogs = [
      { name: 'Lookup', getter: typeof getLookupFunction !== 'undefined' ? getLookupFunction : null },
      { name: 'Text', getter: typeof getTextFunction !== 'undefined' ? getTextFunction : null },
      { name: 'Math', getter: typeof getMathFunction !== 'undefined' ? getMathFunction : null },
      { name: 'Logic', getter: typeof getLogicFunction !== 'undefined' ? getLogicFunction : null },
      { name: 'DateTime', getter: typeof getDateTimeFunction !== 'undefined' ? getDateTimeFunction : null },
      { name: 'Array', getter: typeof getArrayFunction !== 'undefined' ? getArrayFunction : null },
      { name: 'Google', getter: typeof getGoogleFunction !== 'undefined' ? getGoogleFunction : null },
      { name: 'Financial', getter: typeof getFinancialFunction !== 'undefined' ? getFinancialFunction : null },
      { name: 'Engineering', getter: typeof getEngineeringFunction !== 'undefined' ? getEngineeringFunction : null }
    ];
    
    for (var i = 0; i < catalogs.length; i++) {
      if (catalogs[i].getter) {
        var info = catalogs[i].getter(upper);
        if (info) {
          info._catalog = catalogs[i].name;
          return info;
        }
      }
    }
    
    return null;
  }
  
  /**
   * Get all available function names
   */
  function getAllFunctionNames() {
    var allFunctions = [];
    
    if (typeof getLookupFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getLookupFunctionNames());
    if (typeof getTextFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getTextFunctionNames());
    if (typeof getMathFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getMathFunctionNames());
    if (typeof getLogicFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getLogicFunctionNames());
    if (typeof getDateTimeFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getDateTimeFunctionNames());
    if (typeof getArrayFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getArrayFunctionNames());
    if (typeof getGoogleFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getGoogleFunctionNames());
    if (typeof getFinancialFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getFinancialFunctionNames());
    if (typeof getEngineeringFunctionNames !== 'undefined') allFunctions = allFunctions.concat(getEngineeringFunctionNames());
    
    return allFunctions;
  }
  
  // ========== PATTERN TO FORMULA MAPPING ==========
  
  /**
   * Pattern types and their formula strategies
   */
  var PATTERN_STRATEGIES = {
    
    // Numeric threshold classification (e.g., score -> grade)
    threshold: {
      analyze: function(pattern) {
        return pattern.rules && pattern.rules.some(function(r) {
          return /[<>=]/.test(r.condition);
        });
      },
      compose: function(pattern, cellRef) {
        return composeThresholdFormula(pattern, cellRef);
      }
    },
    
    // Keyword-based classification
    keyword: {
      analyze: function(pattern) {
        return pattern.rules && pattern.rules.some(function(r) {
          return r.keywords && r.keywords.length > 0;
        });
      },
      compose: function(pattern, cellRef) {
        return composeKeywordFormula(pattern, cellRef);
      }
    },
    
    // Exact value matching
    exact_match: {
      analyze: function(pattern) {
        return pattern.rules && pattern.rules.every(function(r) {
          return r.condition && !r.keywords && !/[<>=]/.test(r.condition);
        });
      },
      compose: function(pattern, cellRef) {
        return composeExactMatchFormula(pattern, cellRef);
      }
    },
    
    // Text contains matching
    contains: {
      analyze: function(pattern) {
        return pattern.patternType === 'contains' || 
               (pattern.rules && pattern.rules.some(function(r) {
                 return r.matchType === 'contains';
               }));
      },
      compose: function(pattern, cellRef) {
        return composeContainsFormula(pattern, cellRef);
      }
    },
    
    // Regex-based pattern matching
    regex: {
      analyze: function(pattern) {
        return pattern.patternType === 'regex' ||
               (pattern.rules && pattern.rules.some(function(r) {
                 return r.regex;
               }));
      },
      compose: function(pattern, cellRef) {
        return composeRegexFormula(pattern, cellRef);
      }
    },
    
    // Data extraction
    extraction: {
      analyze: function(pattern) {
        return pattern.patternType === 'extraction';
      },
      compose: function(pattern, cellRef) {
        return composeExtractionFormula(pattern, cellRef);
      }
    },
    
    // Lookup-based (VLOOKUP, XLOOKUP)
    lookup: {
      analyze: function(pattern) {
        return pattern.patternType === 'lookup' || pattern.lookupRange;
      },
      compose: function(pattern, cellRef) {
        return composeLookupFormula(pattern, cellRef);
      }
    },
    
    // Date-based classification
    date: {
      analyze: function(pattern) {
        return pattern.patternType === 'date' ||
               (pattern.rules && pattern.rules.some(function(r) {
                 return r.dateFunction;
               }));
      },
      compose: function(pattern, cellRef) {
        return composeDateFormula(pattern, cellRef);
      }
    },
    
    // Numeric calculation/transformation
    calculation: {
      analyze: function(pattern) {
        return pattern.patternType === 'calculation' || pattern.formula;
      },
      compose: function(pattern, cellRef) {
        return composeCalculationFormula(pattern, cellRef);
      }
    }
  };
  
  // ========== FORMULA COMPOSERS ==========
  
  /**
   * Compose IFS formula for numeric thresholds
   */
  function composeThresholdFormula(pattern, cellRef) {
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var conditions = [];
    
    // Sort rules by threshold value (descending for >= conditions)
    var sortedRules = pattern.rules.slice().sort(function(a, b) {
      var valA = parseThresholdValue(a.condition);
      var valB = parseThresholdValue(b.condition);
      if (isNaN(valA) || isNaN(valB)) return 0;
      return valB - valA;
    });
    
    sortedRules.forEach(function(rule) {
      if (rule.condition === 'default' || rule.condition === 'else') {
        conditions.push('TRUE,"' + escapeFormulaString(rule.result) + '"');
      } else {
        conditions.push(cellRef + rule.condition + ',"' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    // Ensure there's a default case
    if (!conditions.some(function(c) { return c.startsWith('TRUE'); })) {
      conditions.push('TRUE,""');
    }
    
    return '=IFS(' + conditions.join(',') + ')';
  }
  
  /**
   * Compose REGEXMATCH formula for keyword classification
   */
  function composeKeywordFormula(pattern, cellRef) {
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var conditions = [];
    
    pattern.rules.forEach(function(rule) {
      if (rule.keywords && rule.keywords.length > 0) {
        // Create regex pattern from keywords (case-insensitive)
        var regexPattern = '(?i)(' + rule.keywords.map(escapeRegex).join('|') + ')';
        conditions.push('REGEXMATCH(' + cellRef + ',"' + regexPattern + '"),"' + escapeFormulaString(rule.result) + '"');
      } else if (rule.condition === 'default' || rule.condition === 'else') {
        conditions.push('TRUE,"' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    // Add default case if not present
    if (!conditions.some(function(c) { return c.startsWith('TRUE'); })) {
      conditions.push('TRUE,""');
    }
    
    return '=IFS(' + conditions.join(',') + ')';
  }
  
  /**
   * Compose SWITCH formula for exact matching
   */
  function composeExactMatchFormula(pattern, cellRef) {
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var cases = [];
    var defaultValue = '""';
    
    pattern.rules.forEach(function(rule) {
      if (rule.condition === 'default' || rule.condition === 'else') {
        defaultValue = '"' + escapeFormulaString(rule.result) + '"';
      } else {
        cases.push('"' + escapeFormulaString(rule.condition) + '","' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    if (cases.length > 0) {
      return '=SWITCH(' + cellRef + ',' + cases.join(',') + ',' + defaultValue + ')';
    }
    
    return null;
  }
  
  /**
   * Compose SEARCH-based formula for text contains
   */
  function composeContainsFormula(pattern, cellRef) {
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var conditions = [];
    
    pattern.rules.forEach(function(rule) {
      if (rule.searchText) {
        conditions.push('ISNUMBER(SEARCH("' + escapeFormulaString(rule.searchText) + '",' + cellRef + ')),"' + escapeFormulaString(rule.result) + '"');
      } else if (rule.condition === 'default' || rule.condition === 'else') {
        conditions.push('TRUE,"' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    if (!conditions.some(function(c) { return c.startsWith('TRUE'); })) {
      conditions.push('TRUE,""');
    }
    
    return '=IFS(' + conditions.join(',') + ')';
  }
  
  /**
   * Compose REGEXEXTRACT or REGEXMATCH formula
   */
  function composeRegexFormula(pattern, cellRef) {
    if (pattern.extractPattern) {
      // Extraction mode
      return '=REGEXEXTRACT(' + cellRef + ',"' + escapeFormulaString(pattern.extractPattern) + '")';
    }
    
    // Classification mode
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var conditions = [];
    
    pattern.rules.forEach(function(rule) {
      if (rule.regex) {
        conditions.push('REGEXMATCH(' + cellRef + ',"' + escapeFormulaString(rule.regex) + '"),"' + escapeFormulaString(rule.result) + '"');
      } else if (rule.condition === 'default' || rule.condition === 'else') {
        conditions.push('TRUE,"' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    if (!conditions.some(function(c) { return c.startsWith('TRUE'); })) {
      conditions.push('TRUE,""');
    }
    
    return '=IFS(' + conditions.join(',') + ')';
  }
  
  /**
   * Compose data extraction formula
   */
  function composeExtractionFormula(pattern, cellRef) {
    switch (pattern.extractType) {
      case 'regex':
        return '=REGEXEXTRACT(' + cellRef + ',"' + escapeFormulaString(pattern.extractPattern) + '")';
      
      case 'left':
        return '=LEFT(' + cellRef + ',' + pattern.extractLength + ')';
      
      case 'right':
        return '=RIGHT(' + cellRef + ',' + pattern.extractLength + ')';
      
      case 'mid':
        return '=MID(' + cellRef + ',' + pattern.extractStart + ',' + pattern.extractLength + ')';
      
      case 'split':
        return '=INDEX(SPLIT(' + cellRef + ',"' + escapeFormulaString(pattern.delimiter) + '"),' + pattern.extractIndex + ')';
      
      case 'before':
        return '=LEFT(' + cellRef + ',FIND("' + escapeFormulaString(pattern.delimiter) + '",' + cellRef + ')-1)';
      
      case 'after':
        return '=MID(' + cellRef + ',FIND("' + escapeFormulaString(pattern.delimiter) + '",' + cellRef + ')+' + pattern.delimiter.length + ',LEN(' + cellRef + '))';
      
      default:
        if (pattern.extractPattern) {
          return '=REGEXEXTRACT(' + cellRef + ',"' + escapeFormulaString(pattern.extractPattern) + '")';
        }
    }
    
    return null;
  }
  
  /**
   * Compose VLOOKUP/XLOOKUP formula
   */
  function composeLookupFormula(pattern, cellRef) {
    if (!pattern.lookupRange) return null;
    
    // Prefer XLOOKUP if available (more modern)
    if (pattern.useXlookup !== false) {
      var formula = '=XLOOKUP(' + cellRef + ',' + pattern.lookupRange;
      formula += ',' + pattern.returnRange;
      formula += ',"' + (pattern.notFoundValue || 'Not Found') + '"';
      if (pattern.matchMode) formula += ',' + pattern.matchMode;
      formula += ')';
      return formula;
    }
    
    // Fall back to VLOOKUP
    var vlookup = '=VLOOKUP(' + cellRef + ',' + pattern.lookupRange;
    vlookup += ',' + pattern.returnColumn;
    vlookup += ',FALSE)';
    
    if (pattern.notFoundValue) {
      return '=IFNA(' + vlookup.substring(1) + ',"' + escapeFormulaString(pattern.notFoundValue) + '")';
    }
    
    return vlookup;
  }
  
  /**
   * Compose date-based formula
   */
  function composeDateFormula(pattern, cellRef) {
    if (!pattern.rules || pattern.rules.length === 0) return null;
    
    var conditions = [];
    
    pattern.rules.forEach(function(rule) {
      if (rule.dateFunction) {
        switch (rule.dateFunction) {
          case 'weekday':
            conditions.push(rule.dateCondition.replace('{cell}', 'WEEKDAY(' + cellRef + ')') + ',"' + escapeFormulaString(rule.result) + '"');
            break;
          case 'month':
            conditions.push(rule.dateCondition.replace('{cell}', 'MONTH(' + cellRef + ')') + ',"' + escapeFormulaString(rule.result) + '"');
            break;
          case 'year':
            conditions.push(rule.dateCondition.replace('{cell}', 'YEAR(' + cellRef + ')') + ',"' + escapeFormulaString(rule.result) + '"');
            break;
          case 'day':
            conditions.push(rule.dateCondition.replace('{cell}', 'DAY(' + cellRef + ')') + ',"' + escapeFormulaString(rule.result) + '"');
            break;
          case 'quarter':
            conditions.push(rule.dateCondition.replace('{cell}', 'ROUNDUP(MONTH(' + cellRef + ')/3,0)') + ',"' + escapeFormulaString(rule.result) + '"');
            break;
        }
      } else if (rule.condition === 'default') {
        conditions.push('TRUE,"' + escapeFormulaString(rule.result) + '"');
      }
    });
    
    if (conditions.length === 0) return null;
    
    if (!conditions.some(function(c) { return c.startsWith('TRUE'); })) {
      conditions.push('TRUE,""');
    }
    
    return '=IFS(' + conditions.join(',') + ')';
  }
  
  /**
   * Compose calculation formula
   */
  function composeCalculationFormula(pattern, cellRef) {
    if (pattern.formula) {
      // Replace placeholder with cell reference
      return '=' + pattern.formula.replace(/\{cell\}/g, cellRef).replace(/\{value\}/g, cellRef);
    }
    
    // Simple operations
    if (pattern.operation) {
      switch (pattern.operation) {
        case 'multiply':
          return '=' + cellRef + '*' + pattern.operand;
        case 'divide':
          return '=' + cellRef + '/' + pattern.operand;
        case 'add':
          return '=' + cellRef + '+' + pattern.operand;
        case 'subtract':
          return '=' + cellRef + '-' + pattern.operand;
        case 'percentage':
          return '=' + cellRef + '*' + (pattern.operand / 100);
        case 'round':
          return '=ROUND(' + cellRef + ',' + (pattern.decimals || 0) + ')';
        case 'abs':
          return '=ABS(' + cellRef + ')';
      }
    }
    
    return null;
  }
  
  // ========== HELPER FUNCTIONS ==========
  
  function parseThresholdValue(condition) {
    var match = condition.match(/[<>=]*\s*([\d.]+)/);
    return match ? parseFloat(match[1]) : NaN;
  }
  
  function escapeFormulaString(str) {
    if (!str) return '';
    return String(str).replace(/"/g, '""');
  }
  
  function escapeRegex(str) {
    return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  
  // ========== MAIN API ==========
  
  /**
   * Generate formula from pattern
   * @param {Object} pattern - Detected pattern with rules
   * @param {string} inputColumn - Input column letter (e.g., "A")
   * @param {string} outputColumn - Output column letter (e.g., "B")
   * @param {number} startRow - Starting row number
   * @returns {string|null} Formula template or null if cannot generate
   */
  function generateFormula(pattern, inputColumn, outputColumn, startRow) {
    if (!pattern || !inputColumn) {
      Logger.log('[FormulaEngine] Missing pattern or input column');
      return null;
    }
    
    // Cell reference placeholder - will be replaced with actual row
    var cellRef = inputColumn + '{ROW}';
    
    // Determine pattern type
    var patternType = pattern.patternType || detectPatternType(pattern);
    
    Logger.log('[FormulaEngine] Generating formula for pattern type: ' + patternType);
    
    // Get appropriate composer
    var strategy = PATTERN_STRATEGIES[patternType];
    if (!strategy) {
      Logger.log('[FormulaEngine] No strategy for pattern type: ' + patternType);
      return null;
    }
    
    // Compose formula
    var formula = strategy.compose(pattern, cellRef);
    
    if (formula) {
      Logger.log('[FormulaEngine] Generated formula: ' + formula);
    }
    
    return formula;
  }
  
  /**
   * Detect pattern type from pattern object
   */
  function detectPatternType(pattern) {
    for (var type in PATTERN_STRATEGIES) {
      if (PATTERN_STRATEGIES[type].analyze && PATTERN_STRATEGIES[type].analyze(pattern)) {
        return type;
      }
    }
    return 'exact_match'; // Default
  }
  
  /**
   * Apply formula to a range of cells
   * @param {string} formulaTemplate - Formula with {ROW} placeholder
   * @param {string} outputColumn - Output column letter
   * @param {number} startRow - Starting row
   * @param {number} endRow - Ending row
   * @returns {Object} Result with success status and details
   */
  function applyFormulaToRange(formulaTemplate, outputColumn, startRow, endRow) {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      var results = [];
      
      for (var row = startRow; row <= endRow; row++) {
        var formula = formulaTemplate.replace(/\{ROW\}/g, row);
        var cell = sheet.getRange(outputColumn + row);
        cell.setFormula(formula);
        results.push({ row: row, formula: formula });
      }
      
      return {
        success: true,
        rowsUpdated: endRow - startRow + 1,
        formulaTemplate: formulaTemplate,
        results: results
      };
      
    } catch (error) {
      Logger.log('[FormulaEngine] Error applying formula: ' + error.message);
      return {
        success: false,
        error: error.message
      };
    }
  }
  
  /**
   * Suggest the best formula approach for a task
   * @param {string} taskDescription - Natural language description
   * @param {Array} sampleData - Sample input/output data
   * @returns {Object} Suggestion with formula and explanation
   */
  function suggestFormula(taskDescription, sampleData) {
    var suggestion = {
      formula: null,
      explanation: '',
      confidence: 0,
      alternatives: []
    };
    
    var lowerTask = (taskDescription || '').toLowerCase();
    
    // Analyze task description for function hints
    var functionHints = [];
    
    if (lowerTask.includes('lookup') || lowerTask.includes('find') || lowerTask.includes('match')) {
      functionHints.push({ func: 'XLOOKUP', reason: 'Looking up values' });
      functionHints.push({ func: 'VLOOKUP', reason: 'Traditional lookup' });
    }
    
    if (lowerTask.includes('categorize') || lowerTask.includes('classify') || lowerTask.includes('grade')) {
      functionHints.push({ func: 'IFS', reason: 'Multiple conditions classification' });
      functionHints.push({ func: 'SWITCH', reason: 'Exact value mapping' });
    }
    
    if (lowerTask.includes('extract') || lowerTask.includes('parse')) {
      functionHints.push({ func: 'REGEXEXTRACT', reason: 'Pattern-based extraction' });
      functionHints.push({ func: 'SPLIT', reason: 'Delimiter-based extraction' });
    }
    
    if (lowerTask.includes('contains') || lowerTask.includes('search') || lowerTask.includes('filter')) {
      functionHints.push({ func: 'FILTER', reason: 'Filter rows by condition' });
      functionHints.push({ func: 'QUERY', reason: 'SQL-like data filtering' });
    }
    
    if (lowerTask.includes('sum') || lowerTask.includes('total') || lowerTask.includes('add up')) {
      functionHints.push({ func: 'SUMIF', reason: 'Conditional summing' });
      functionHints.push({ func: 'SUMIFS', reason: 'Multi-condition summing' });
    }
    
    if (lowerTask.includes('count') || lowerTask.includes('how many')) {
      functionHints.push({ func: 'COUNTIF', reason: 'Conditional counting' });
      functionHints.push({ func: 'COUNTIFS', reason: 'Multi-condition counting' });
    }
    
    if (lowerTask.includes('unique') || lowerTask.includes('distinct') || lowerTask.includes('deduplicate')) {
      functionHints.push({ func: 'UNIQUE', reason: 'Remove duplicates' });
    }
    
    if (lowerTask.includes('sort') || lowerTask.includes('order') || lowerTask.includes('rank')) {
      functionHints.push({ func: 'SORT', reason: 'Sort data' });
      functionHints.push({ func: 'SORTN', reason: 'Top N sorted' });
    }
    
    if (lowerTask.includes('date') || lowerTask.includes('day') || lowerTask.includes('month')) {
      functionHints.push({ func: 'DATE', reason: 'Date creation' });
      functionHints.push({ func: 'TEXT', reason: 'Date formatting' });
    }
    
    if (lowerTask.includes('text') || lowerTask.includes('string') || lowerTask.includes('format')) {
      functionHints.push({ func: 'TEXT', reason: 'Number/date formatting' });
      functionHints.push({ func: 'CONCATENATE', reason: 'Join text' });
    }
    
    // Build suggestion from hints
    if (functionHints.length > 0) {
      var primary = functionHints[0];
      var funcInfo = getFunctionInfo(primary.func);
      
      suggestion.formula = funcInfo ? funcInfo.syntax : primary.func + '(...)';
      suggestion.explanation = primary.reason;
      suggestion.confidence = Math.min(0.9, 0.5 + (functionHints.length * 0.1));
      suggestion.alternatives = functionHints.slice(1).map(function(h) {
        return { function: h.func, reason: h.reason };
      });
    }
    
    return suggestion;
  }
  
  /**
   * Get statistics about available functions
   */
  function getCatalogStats() {
    var stats = {
      totalFunctions: 0,
      catalogs: {}
    };
    
    var catalogNames = ['Lookup', 'Text', 'Math', 'Logic', 'DateTime', 'Array', 'Google', 'Financial', 'Engineering'];
    var getters = {
      'Lookup': typeof getLookupFunctionNames !== 'undefined' ? getLookupFunctionNames : null,
      'Text': typeof getTextFunctionNames !== 'undefined' ? getTextFunctionNames : null,
      'Math': typeof getMathFunctionNames !== 'undefined' ? getMathFunctionNames : null,
      'Logic': typeof getLogicFunctionNames !== 'undefined' ? getLogicFunctionNames : null,
      'DateTime': typeof getDateTimeFunctionNames !== 'undefined' ? getDateTimeFunctionNames : null,
      'Array': typeof getArrayFunctionNames !== 'undefined' ? getArrayFunctionNames : null,
      'Google': typeof getGoogleFunctionNames !== 'undefined' ? getGoogleFunctionNames : null,
      'Financial': typeof getFinancialFunctionNames !== 'undefined' ? getFinancialFunctionNames : null,
      'Engineering': typeof getEngineeringFunctionNames !== 'undefined' ? getEngineeringFunctionNames : null
    };
    
    catalogNames.forEach(function(name) {
      var getter = getters[name];
      if (getter) {
        var count = getter().length;
        stats.catalogs[name] = count;
        stats.totalFunctions += count;
      }
    });
    
    return stats;
  }
  
  // ========== PUBLIC API ==========
  
  return {
    generateFormula: generateFormula,
    applyFormulaToRange: applyFormulaToRange,
    suggestFormula: suggestFormula,
    getFunctionInfo: getFunctionInfo,
    getAllFunctionNames: getAllFunctionNames,
    getCatalogStats: getCatalogStats,
    detectPatternType: detectPatternType,
    
    // Expose composers for direct use
    composers: {
      threshold: composeThresholdFormula,
      keyword: composeKeywordFormula,
      exactMatch: composeExactMatchFormula,
      contains: composeContainsFormula,
      regex: composeRegexFormula,
      extraction: composeExtractionFormula,
      lookup: composeLookupFormula,
      date: composeDateFormula,
      calculation: composeCalculationFormula
    }
  };
})();

// Log stats on load
(function() {
  try {
    var stats = FormulaEngine.getCatalogStats();
    Logger.log('🔧 FormulaEngine loaded - ' + stats.totalFunctions + ' functions available');
    Logger.log('   Catalogs: ' + JSON.stringify(stats.catalogs));
  } catch (e) {
    Logger.log('🔧 FormulaEngine loaded (catalogs loading...)');
  }
})();
