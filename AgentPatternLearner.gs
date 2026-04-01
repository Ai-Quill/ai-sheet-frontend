/**
 * ============================================
 * AGENT PATTERN LEARNER
 * ============================================
 * 
 * Analyzes AI results to detect patterns that could be expressed as formulas.
 * Goal: After AI processes data, suggest formulas so users don't need AI next time.
 * 
 * CHANGELOG:
 * - 1.0.0 (2026-01-16): Initial implementation
 *   - analyzeResultsForPattern() - Main pattern detection
 *   - detectThresholdPattern() - Numeric threshold detection
 *   - detectKeywordPattern() - Keyword-based classification
 *   - generateFormulaFromPattern() - Formula generation
 * 
 * ============================================
 */

/**
 * Main function: Analyze AI results for extractable patterns
 * @param {Object} params - Parameters object
 * @param {Array<Object>} params.inputs - Input data rows (multi-column objects)
 * @param {Array<string>} params.outputs - AI classification results
 * @param {Array<string>} params.headers - Column headers
 * @param {string} params.inputRange - Original input range
 * @param {string} params.outputColumn - Output column letter
 * @return {Object} Pattern analysis result
/**
 * Retrieve stored patterns from backend that match the current data shape
 * Called before running AI to check if a known pattern exists
 * 
 * @param {Array<string>} headers - Column headers
 * @param {string} patternType - e.g., 'threshold', 'keyword', 'classify'
 * @return {Array<Object>} Matching patterns with formula templates
 */
function retrieveStoredPatterns(headers, patternType) {
  try {
    var userId = getUserId();
    if (!userId) return [];
    
    var params = {
      userId: userId,
      resource: 'patterns',
      patternType: patternType || '',
      minAccuracy: '0.85'
    };
    
    var result = ApiClient.get('AGENT_LEARNING', params);
    
    if (result && result.patterns && result.patterns.length > 0) {
      Logger.log('[PatternLearner] Found ' + result.patterns.length + ' stored patterns');
      return result.patterns;
    }
    
    return [];
  } catch (e) {
    Logger.log('[PatternLearner] Failed to retrieve patterns: ' + e.message);
    return [];
  }
}

/**
 * Main function: Analyze AI results for extractable patterns
 * @param {Object} params - Parameters object
 * @param {Array<Object>} params.inputs - Input data rows (multi-column objects)
 * @param {Array<string>} params.outputs - AI classification results
 * @param {Array<string>} params.headers - Column headers
 * @param {string} params.inputRange - Original input range
 * @param {string} params.outputColumn - Output column letter
 * @return {Object} Pattern analysis result
 */
function analyzeResultsForPattern(params) {
  var inputs = params.inputs || [];
  var outputs = params.outputs || [];
  var headers = params.headers || [];
  var inputRange = params.inputRange || '';
  var outputColumn = params.outputColumn || '';
  
  if (inputs.length === 0 || outputs.length !== inputs.length) {
    return { hasPattern: false, reason: 'Invalid input/output data' };
  }
  
  // Get unique output categories
  var categories = getUniqueCategories(outputs);
  
  if (categories.length < 2 || categories.length > 10) {
    return { 
      hasPattern: false, 
      reason: categories.length < 2 
        ? 'All results are the same category' 
        : 'Too many categories for pattern detection'
    };
  }
  
  // Try different pattern detection strategies
  var thresholdPattern = detectThresholdPattern(inputs, outputs, categories, headers);
  if (thresholdPattern.hasPattern) {
    return thresholdPattern;
  }
  
  var keywordPattern = detectKeywordPattern(inputs, outputs, categories, headers);
  if (keywordPattern.hasPattern) {
    return keywordPattern;
  }
  
  var rangePattern = detectRangePattern(inputs, outputs, categories, headers);
  if (rangePattern.hasPattern) {
    return rangePattern;
  }
  
  return { 
    hasPattern: false, 
    reason: 'No clear pattern detected - classification depends on complex factors',
    categories: categories,
    breakdown: getCategoryBreakdown(outputs, categories)
  };
}

/**
 * Get unique categories from outputs
 */
function getUniqueCategories(outputs) {
  var seen = {};
  var categories = [];
  
  outputs.forEach(function(out) {
    var normalized = String(out || '').trim().toLowerCase();
    if (normalized && !seen[normalized]) {
      seen[normalized] = true;
      categories.push(String(out || '').trim());
    }
  });
  
  return categories;
}

/**
 * Get breakdown of categories
 */
function getCategoryBreakdown(outputs, categories) {
  var breakdown = {};
  categories.forEach(function(cat) {
    breakdown[cat] = 0;
  });
  
  outputs.forEach(function(out) {
    var normalized = String(out || '').trim();
    if (breakdown.hasOwnProperty(normalized)) {
      breakdown[normalized]++;
    }
  });
  
  return breakdown;
}

/**
 * Detect threshold-based patterns (e.g., size > 500 = Hot)
 */
function detectThresholdPattern(inputs, outputs, categories, headers) {
  // Find numeric columns
  var numericColumns = findNumericColumns(inputs, headers);
  
  if (numericColumns.length === 0) {
    return { hasPattern: false, reason: 'No numeric columns for threshold detection' };
  }
  
  // For each numeric column, check if thresholds predict categories
  for (var i = 0; i < numericColumns.length; i++) {
    var colInfo = numericColumns[i];
    var pattern = checkThresholdForColumn(inputs, outputs, categories, colInfo);
    
    if (pattern.hasPattern && pattern.accuracy >= 0.9) {
      return pattern;
    }
  }
  
  return { hasPattern: false };
}

/**
 * Find columns that contain numeric data
 */
function findNumericColumns(inputs, headers) {
  var columns = [];
  
  if (inputs.length === 0) return columns;
  
  var firstRow = inputs[0];
  var keys = Object.keys(firstRow);
  
  keys.forEach(function(key, index) {
    var numericCount = 0;
    var total = 0;
    
    inputs.forEach(function(row) {
      var val = row[key];
      if (val !== null && val !== undefined && val !== '') {
        total++;
        var num = parseFloat(String(val).replace(/[,$]/g, ''));
        if (!isNaN(num)) {
          numericCount++;
        }
      }
    });
    
    // Consider it numeric if >80% are numbers
    if (total > 0 && numericCount / total >= 0.8) {
      columns.push({
        key: key,
        header: headers[index] || key,
        index: index
      });
    }
  });
  
  return columns;
}

/**
 * Check if a single numeric column can predict categories via thresholds
 */
function checkThresholdForColumn(inputs, outputs, categories, colInfo) {
  // Group values by category
  var categoryValues = {};
  categories.forEach(function(cat) {
    categoryValues[cat.toLowerCase()] = [];
  });
  
  inputs.forEach(function(row, i) {
    var val = row[colInfo.key];
    var num = parseFloat(String(val || '').replace(/[,$]/g, ''));
    var category = String(outputs[i] || '').trim().toLowerCase();
    
    if (!isNaN(num) && categoryValues[category]) {
      categoryValues[category].push(num);
    }
  });
  
  // Calculate min/max for each category
  var catStats = {};
  var sortedCategories = [];
  
  categories.forEach(function(cat) {
    var values = categoryValues[cat.toLowerCase()];
    if (values.length > 0) {
      catStats[cat] = {
        min: Math.min.apply(null, values),
        max: Math.max.apply(null, values),
        avg: values.reduce(function(a, b) { return a + b; }, 0) / values.length,
        count: values.length
      };
      sortedCategories.push({
        name: cat,
        avg: catStats[cat].avg
      });
    }
  });
  
  // Sort categories by average value
  sortedCategories.sort(function(a, b) { return b.avg - a.avg; });
  
  // Check if ranges don't overlap
  var thresholds = [];
  var hasOverlap = false;
  
  for (var i = 0; i < sortedCategories.length - 1; i++) {
    var current = sortedCategories[i];
    var next = sortedCategories[i + 1];
    
    var currentStats = catStats[current.name];
    var nextStats = catStats[next.name];
    
    // Check overlap
    if (currentStats.min <= nextStats.max) {
      hasOverlap = true;
      break;
    }
    
    // Threshold is midpoint between max of lower and min of higher
    var threshold = Math.round((nextStats.max + currentStats.min) / 2);
    thresholds.push({
      category: current.name,
      threshold: threshold
    });
  }
  
  if (hasOverlap) {
    return { hasPattern: false, reason: 'Overlapping ranges for ' + colInfo.header };
  }
  
  // Generate formula
  var formula = generateThresholdFormula(sortedCategories, thresholds, colInfo);
  
  // Calculate accuracy
  var correct = 0;
  inputs.forEach(function(row, i) {
    var val = row[colInfo.key];
    var num = parseFloat(String(val || '').replace(/[,$]/g, ''));
    var actualCategory = String(outputs[i] || '').trim();
    var predictedCategory = predictWithThreshold(num, sortedCategories, thresholds);
    
    if (predictedCategory.toLowerCase() === actualCategory.toLowerCase()) {
      correct++;
    }
  });
  
  var accuracy = inputs.length > 0 ? correct / inputs.length : 0;
  
  return {
    hasPattern: true,
    patternType: 'threshold',
    column: colInfo.header,
    columnKey: colInfo.key,
    thresholds: thresholds,
    categories: sortedCategories.map(function(c) { return c.name; }),
    formula: formula,
    accuracy: accuracy,
    explanation: buildThresholdExplanation(colInfo.header, sortedCategories, thresholds, accuracy),
    stats: catStats
  };
}

/**
 * Predict category using threshold
 */
function predictWithThreshold(num, sortedCategories, thresholds) {
  if (isNaN(num)) return sortedCategories[sortedCategories.length - 1].name;
  
  for (var i = 0; i < thresholds.length; i++) {
    if (num >= thresholds[i].threshold) {
      return thresholds[i].category;
    }
  }
  
  return sortedCategories[sortedCategories.length - 1].name;
}

/**
 * Generate threshold formula (IFS-style)
 */
function generateThresholdFormula(sortedCategories, thresholds, colInfo) {
  // Use {cell} as placeholder for the cell reference
  var conditions = [];
  
  for (var i = 0; i < thresholds.length; i++) {
    conditions.push('{cell}>=' + thresholds[i].threshold + ',"' + thresholds[i].category + '"');
  }
  
  // Last category (lowest values)
  var lastCat = sortedCategories[sortedCategories.length - 1].name;
  conditions.push('TRUE,"' + lastCat + '"');
  
  return '=IFS(' + conditions.join(',') + ')';
}

/**
 * Build human-readable explanation
 */
function buildThresholdExplanation(columnName, sortedCategories, thresholds, accuracy) {
  var parts = [];
  
  for (var i = 0; i < thresholds.length; i++) {
    parts.push('"' + thresholds[i].category + '" when ' + columnName + ' ≥ ' + thresholds[i].threshold);
  }
  
  var lastCat = sortedCategories[sortedCategories.length - 1].name;
  parts.push('"' + lastCat + '" otherwise');
  
  var accuracyStr = Math.round(accuracy * 100) + '%';
  
  return 'Classification based on ' + columnName + ': ' + parts.join(', ') + '. (Accuracy: ' + accuracyStr + ')';
}

/**
 * Detect keyword-based patterns
 */
function detectKeywordPattern(inputs, outputs, categories, headers) {
  // Find text columns
  var textColumns = findTextColumns(inputs, headers);
  
  if (textColumns.length === 0) {
    return { hasPattern: false };
  }
  
  // For each category, find common keywords
  var categoryKeywords = {};
  categories.forEach(function(cat) {
    categoryKeywords[cat.toLowerCase()] = {};
  });
  
  // Collect words per category
  inputs.forEach(function(row, i) {
    var category = String(outputs[i] || '').trim().toLowerCase();
    if (!categoryKeywords[category]) return;
    
    textColumns.forEach(function(col) {
      var text = String(row[col.key] || '').toLowerCase();
      var words = text.match(/\b[a-z]{3,}\b/g) || [];
      
      words.forEach(function(word) {
        if (!categoryKeywords[category][word]) {
          categoryKeywords[category][word] = 0;
        }
        categoryKeywords[category][word]++;
      });
    });
  });
  
  // Find distinctive keywords per category
  var distinctiveKeywords = {};
  categories.forEach(function(cat) {
    var catLower = cat.toLowerCase();
    var keywords = categoryKeywords[catLower];
    var distinctive = [];
    
    Object.keys(keywords).forEach(function(word) {
      // Check if word is more common in this category
      var inOther = 0;
      categories.forEach(function(otherCat) {
        if (otherCat.toLowerCase() !== catLower) {
          inOther += categoryKeywords[otherCat.toLowerCase()][word] || 0;
        }
      });
      
      var inThis = keywords[word];
      if (inThis > 1 && inThis > inOther * 2) {
        distinctive.push({ word: word, count: inThis, ratio: inThis / (inOther + 1) });
      }
    });
    
    distinctive.sort(function(a, b) { return b.ratio - a.ratio; });
    distinctiveKeywords[cat] = distinctive.slice(0, 3);
  });
  
  // Check if keywords can predict categories
  var allEmpty = true;
  categories.forEach(function(cat) {
    if (distinctiveKeywords[cat].length > 0) allEmpty = false;
  });
  
  if (allEmpty) {
    return { hasPattern: false, reason: 'No distinctive keywords found' };
  }
  
  // Build REGEXMATCH formula
  var formula = buildKeywordFormula(categories, distinctiveKeywords);
  
  // Calculate accuracy
  var correct = 0;
  inputs.forEach(function(row, i) {
    var actualCategory = String(outputs[i] || '').trim().toLowerCase();
    var predictedCategory = predictWithKeywords(row, textColumns, distinctiveKeywords, categories);
    
    if (predictedCategory.toLowerCase() === actualCategory) {
      correct++;
    }
  });
  
  var accuracy = inputs.length > 0 ? correct / inputs.length : 0;
  
  // Only suggest if accuracy is decent
  if (accuracy < 0.7) {
    return { hasPattern: false, reason: 'Keyword pattern too weak (accuracy: ' + Math.round(accuracy * 100) + '%)' };
  }
  
  return {
    hasPattern: true,
    patternType: 'keyword',
    keywords: distinctiveKeywords,
    formula: formula,
    accuracy: accuracy,
    explanation: 'Classification based on keywords in text. Accuracy: ' + Math.round(accuracy * 100) + '%'
  };
}

/**
 * Find columns with text data
 */
function findTextColumns(inputs, headers) {
  var columns = [];
  
  if (inputs.length === 0) return columns;
  
  var firstRow = inputs[0];
  var keys = Object.keys(firstRow);
  
  keys.forEach(function(key, index) {
    var textCount = 0;
    var total = 0;
    var avgLength = 0;
    
    inputs.forEach(function(row) {
      var val = row[key];
      if (val !== null && val !== undefined && val !== '') {
        total++;
        var str = String(val);
        if (str.length > 10 && isNaN(parseFloat(str))) {
          textCount++;
          avgLength += str.length;
        }
      }
    });
    
    // Consider it text if >50% are text strings
    if (total > 0 && textCount / total >= 0.5) {
      columns.push({
        key: key,
        header: headers[index] || key,
        avgLength: textCount > 0 ? avgLength / textCount : 0
      });
    }
  });
  
  return columns;
}

/**
 * Predict category using keywords
 */
function predictWithKeywords(row, textColumns, distinctiveKeywords, categories) {
  var scores = {};
  categories.forEach(function(cat) { scores[cat] = 0; });
  
  var combinedText = '';
  textColumns.forEach(function(col) {
    combinedText += ' ' + String(row[col.key] || '').toLowerCase();
  });
  
  categories.forEach(function(cat) {
    var keywords = distinctiveKeywords[cat] || [];
    keywords.forEach(function(kw) {
      if (combinedText.indexOf(kw.word) !== -1) {
        scores[cat] += kw.ratio;
      }
    });
  });
  
  var best = categories[0];
  var bestScore = scores[best];
  
  categories.forEach(function(cat) {
    if (scores[cat] > bestScore) {
      best = cat;
      bestScore = scores[cat];
    }
  });
  
  return best;
}

/**
 * Build keyword-based formula
 */
function buildKeywordFormula(categories, distinctiveKeywords) {
  var conditions = [];
  
  categories.forEach(function(cat) {
    var keywords = distinctiveKeywords[cat] || [];
    if (keywords.length > 0) {
      var keywordPatterns = keywords.map(function(kw) { return kw.word; }).join('|');
      conditions.push('REGEXMATCH(LOWER({cell}),"' + keywordPatterns + '"),"' + cat + '"');
    }
  });
  
  if (conditions.length === 0) return null;
  
  conditions.push('TRUE,"Unknown"');
  
  return '=IFS(' + conditions.join(',') + ')';
}

/**
 * Detect range-based patterns (e.g., 0-50 = Low, 51-100 = High)
 */
function detectRangePattern(inputs, outputs, categories, headers) {
  // Similar to threshold but checks for explicit ranges
  // (Simplified - could be expanded)
  return { hasPattern: false };
}

/**
 * Apply learned formula to sheet
 * @param {string} formula - Formula template with {cell} placeholder
 * @param {string} inputColumn - Input column letter
 * @param {string} outputColumn - Output column letter
 * @param {number} startRow - Start row number
 * @param {number} endRow - End row number
 * @return {Object} Result with row count
 */
function applyLearnedFormula(formula, inputColumn, outputColumn, startRow, endRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var rowCount = endRow - startRow + 1;
  var formulas = [];
  
  for (var row = startRow; row <= endRow; row++) {
    var cellFormula = formula.replace(/\{cell\}/g, inputColumn + row);
    formulas.push([cellFormula]);
  }
  
  var outputRange = sheet.getRange(outputColumn + startRow + ':' + outputColumn + endRow);
  outputRange.setFormulas(formulas);
  
  return {
    success: true,
    rowCount: rowCount,
    formula: formula
  };
}

/**
 * Save a detected pattern to the backend for cross-session learning
 * Called from sidebar JS (savePatternToHistory)
 * 
 * @param {Object} pattern - Pattern object from analyzeResultsForPattern
 * @param {Object} jobData - Job context data
 * @return {Object} { success: boolean }
 */
function savePatternToBackend(pattern, jobData) {
  try {
    var userId = getUserId();
    if (!userId) {
      Logger.log('[PatternLearner] No userId available, skipping save');
      return { success: false, reason: 'no_user' };
    }
    
    var payload = {
      type: 'pattern',
      userId: userId,
      patternType: pattern.patternType || 'unknown',
      formulaTemplate: pattern.formula || null,
      accuracy: pattern.accuracy || 0,
      dataShape: {
        columns: jobData && jobData.headers ? jobData.headers : [],
        rowCount: jobData && jobData.rowCount ? jobData.rowCount : 0
      },
      columnTypes: pattern.columnTypes || null
    };
    
    var result = ApiClient.post('AGENT_LEARNING', payload);
    Logger.log('[PatternLearner] Pattern saved to backend: ' + (result.success ? 'OK' : 'FAILED'));
    return { success: true };
  } catch (e) {
    Logger.log('[PatternLearner] Failed to save pattern: ' + e.message);
    return { success: false, reason: e.message };
  }
}

/**
 * Submit analysis feedback to backend
 * Called from sidebar JS (feedback buttons)
 * 
 * @param {string} queryText - Original analysis question
 * @param {string} feedback - 'up' or 'down'
 * @param {string} correctionCategory - Optional category
 * @param {Object} dataShape - Optional data shape info
 * @return {Object} { success: boolean }
 */
function submitAnalysisFeedback(queryText, feedback, correctionCategory, dataShape) {
  try {
    var userId = getUserId();
    if (!userId) {
      return { success: false, reason: 'no_user' };
    }
    
    var payload = {
      type: 'feedback',
      userId: userId,
      queryText: queryText,
      feedback: feedback,
      correctionCategory: correctionCategory || null,
      dataShape: dataShape || null
    };
    
    var result = ApiClient.post('AGENT_LEARNING', payload);
    return { success: true };
  } catch (e) {
    Logger.log('[PatternLearner] Failed to submit feedback: ' + e.message);
    return { success: false, reason: e.message };
  }
}

/**
 * Retrieve user analysis preferences from backend.
 * Used to adjust analysis behavior based on accumulated feedback.
 * Only returns preferences with confidence >= 0.6, signal_count >= 3.
 * 
 * @return {Object} Map of preference_key -> preference_value
 */
function getUserAnalysisPreferences() {
  try {
    var userId = getUserId();
    if (!userId) return {};
    
    var result = ApiClient.get('AGENT_LEARNING', {
      userId: userId,
      resource: 'preferences'
    });
    
    if (result && result.preferences && result.preferences.length > 0) {
      var prefs = {};
      result.preferences.forEach(function(p) {
        prefs[p.preference_key] = p.preference_value;
      });
      Logger.log('[PatternLearner] Retrieved ' + Object.keys(prefs).length + ' user preferences');
      return prefs;
    }
    
    return {};
  } catch (e) {
    Logger.log('[PatternLearner] Failed to retrieve preferences: ' + e.message);
    return {};
  }
}
