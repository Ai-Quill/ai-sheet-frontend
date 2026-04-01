/**
 * @file AgentAnalyzer.gs
 * @version 2.1.0
 * @updated 2026-01-18
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.1.0 (2026-01-18): Enhanced proactive suggestions with pattern detection
 *   - Added generateProactiveSuggestions() with confidence scoring
 *   - Data pattern detection (numeric variance, text length, URL/email patterns)
 *   - Context-aware suggestion reasons
 *   - Exposed for Feedback UI integration
 * - 2.0.0 (2026-01-16): AI-powered analysis replaces dumb heuristics
 *   - Uses AI to understand data context and suggest relevant actions
 *   - No more hardcoded pattern matching
 *   - Suggestions are context-aware (understands "Company Name" vs "Description")
 * - 1.0.0 (2026-01-15): Initial proactive data analyzer (dumb heuristics)
 * 
 * ============================================
 * AGENT ANALYZER - AI-Powered Data Analysis
 * ============================================
 * 
 * Analyzes spreadsheet data using AI to provide intelligent suggestions:
 * - Understands column semantics (not just patterns)
 * - Generates contextually relevant suggestions
 * - Pattern detection with confidence scoring
 * - Learns what actions make sense for the data type
 */

// ============================================
// ANALYSIS CONSTANTS
// ============================================

var ANALYSIS_SAMPLE_SIZE = 10;  // Sample rows for AI analysis
var MAX_SUGGESTIONS = 3;

// ============================================
// PROACTIVE SUGGESTIONS WITH PATTERN DETECTION
// ============================================

/**
 * Generate proactive suggestions based on data pattern analysis
 * This is the enhanced version with confidence scoring
 * 
 * @param {Object} context - Unified context from AgentContext
 * @return {Array} Array of suggestions with confidence scores and reasons
 */
function generateProactiveSuggestions(context) {
  if (!context || !context.selectionInfo) {
    return [];
  }
  
  var suggestions = [];
  var selInfo = context.selectionInfo;
  var startRow = selInfo.dataStartRow || 2;
  
  // Get empty column for output
  var emptyColRaw = selInfo.emptyColumns?.[0];
  var outputCol = typeof emptyColRaw === 'object' ? emptyColRaw.column : emptyColRaw || 'E';
  
  // Analyze each data column
  var dataColumns = selInfo.columnsWithData || [];
  
  for (var i = 0; i < dataColumns.length && suggestions.length < MAX_SUGGESTIONS + 2; i++) {
    var col = dataColumns[i];
    var header = selInfo.headers?.find(function(h) { return h.column === col; });
    var headerName = header?.name || 'Column ' + col;
    
    // Get sample values for this column
    var sampleValues = getSampleValuesForColumn(col, startRow, 10);
    if (!sampleValues || sampleValues.length < 3) continue;
    
    // Analyze patterns in this column
    var patterns = analyzeColumnPatterns(sampleValues, headerName);
    
    // Generate suggestions based on detected patterns
    patterns.forEach(function(pattern) {
      if (suggestions.length < MAX_SUGGESTIONS + 2) {
        var suggestion = createSuggestionFromPattern(pattern, col, outputCol, headerName);
        if (suggestion) {
          suggestions.push(suggestion);
        }
      }
    });
  }
  
  // Sort by confidence (highest first) and take top MAX_SUGGESTIONS
  suggestions.sort(function(a, b) { return b.confidence - a.confidence; });
  return suggestions.slice(0, MAX_SUGGESTIONS);
}

/**
 * Analyze patterns in a column's data
 * Returns array of detected patterns with confidence scores
 */
function analyzeColumnPatterns(values, headerName) {
  var patterns = [];
  var nonEmpty = values.filter(function(v) { return v !== null && v !== undefined && String(v).trim() !== ''; });
  
  if (nonEmpty.length === 0) return patterns;
  
  var headerLower = (headerName || '').toLowerCase();
  
  // 1. Check for numeric data with variance (good for classification)
  var numericAnalysis = analyzeNumericPattern(nonEmpty);
  if (numericAnalysis.isNumeric && numericAnalysis.hasVariance) {
    patterns.push({
      type: 'numeric_variance',
      confidence: numericAnalysis.confidence,
      reason: 'Numeric column with ' + numericAnalysis.rangeDescription,
      suggestedAction: 'classify',
      data: numericAnalysis
    });
  }
  
  // 2. Check for text length (good for summarization)
  var textAnalysis = analyzeTextLength(nonEmpty);
  if (textAnalysis.hasLongText) {
    patterns.push({
      type: 'long_text',
      confidence: textAnalysis.confidence,
      reason: 'Text averaging ' + textAnalysis.avgLength + ' characters',
      suggestedAction: 'summarize',
      data: textAnalysis
    });
  }
  
  // 3. Check for email patterns
  var emailAnalysis = analyzeEmailPattern(nonEmpty);
  if (emailAnalysis.hasEmails) {
    patterns.push({
      type: 'email',
      confidence: emailAnalysis.confidence,
      reason: emailAnalysis.count + ' email addresses detected',
      suggestedAction: 'extract',
      extractType: 'domain',
      data: emailAnalysis
    });
  }
  
  // 4. Check for URL patterns
  var urlAnalysis = analyzeURLPattern(nonEmpty);
  if (urlAnalysis.hasURLs) {
    patterns.push({
      type: 'url',
      confidence: urlAnalysis.confidence,
      reason: urlAnalysis.count + ' URLs detected',
      suggestedAction: 'extract',
      extractType: 'domain',
      data: urlAnalysis
    });
  }
  
  // 5. Check header name for semantic hints
  var semanticHints = analyzeHeaderSemantics(headerLower);
  if (semanticHints.suggestedAction && !patterns.some(function(p) { return p.suggestedAction === semanticHints.suggestedAction; })) {
    patterns.push({
      type: 'semantic',
      confidence: semanticHints.confidence,
      reason: 'Column name suggests ' + semanticHints.reason,
      suggestedAction: semanticHints.suggestedAction,
      data: semanticHints
    });
  }
  
  return patterns;
}

/**
 * Analyze numeric patterns in values
 */
function analyzeNumericPattern(values) {
  var numbers = [];
  
  values.forEach(function(v) {
    var num = parseFloat(String(v).replace(/[,$]/g, ''));
    if (!isNaN(num)) numbers.push(num);
  });
  
  var ratio = numbers.length / values.length;
  
  if (ratio < 0.6) {
    return { isNumeric: false };
  }
  
  var min = Math.min.apply(null, numbers);
  var max = Math.max.apply(null, numbers);
  var range = max - min;
  var avg = numbers.reduce(function(a, b) { return a + b; }, 0) / numbers.length;
  
  // Calculate variance
  var variance = numbers.reduce(function(sum, n) { return sum + Math.pow(n - avg, 2); }, 0) / numbers.length;
  var stdDev = Math.sqrt(variance);
  var coeffVariation = avg !== 0 ? stdDev / Math.abs(avg) : 0;
  
  // Has meaningful variance if CV > 0.3 (30% variation)
  var hasVariance = coeffVariation > 0.3 && range > 0;
  
  return {
    isNumeric: true,
    hasVariance: hasVariance,
    confidence: Math.min(0.95, ratio * 0.8 + (hasVariance ? 0.15 : 0)),
    min: min,
    max: max,
    avg: Math.round(avg),
    rangeDescription: 'range ' + formatNumber(min) + ' to ' + formatNumber(max)
  };
}

/**
 * Analyze text length patterns
 */
function analyzeTextLength(values) {
  var lengths = values.map(function(v) { return String(v).length; });
  var avgLength = Math.round(lengths.reduce(function(a, b) { return a + b; }, 0) / lengths.length);
  var longCount = lengths.filter(function(l) { return l > 100; }).length;
  var ratio = longCount / values.length;
  
  return {
    hasLongText: ratio > 0.4 && avgLength > 80,
    confidence: Math.min(0.9, ratio * 0.7 + (avgLength > 150 ? 0.2 : 0.1)),
    avgLength: avgLength,
    longRatio: ratio
  };
}

/**
 * Analyze email patterns
 */
function analyzeEmailPattern(values) {
  var emailRegex = /[\w.-]+@[\w.-]+\.\w+/;
  var count = values.filter(function(v) { return emailRegex.test(String(v)); }).length;
  var ratio = count / values.length;
  
  return {
    hasEmails: ratio > 0.3,
    confidence: Math.min(0.95, ratio),
    count: count,
    ratio: ratio
  };
}

/**
 * Analyze URL patterns
 */
function analyzeURLPattern(values) {
  var urlRegex = /https?:\/\/[^\s]+/;
  var count = values.filter(function(v) { return urlRegex.test(String(v)); }).length;
  var ratio = count / values.length;
  
  return {
    hasURLs: ratio > 0.3,
    confidence: Math.min(0.95, ratio),
    count: count,
    ratio: ratio
  };
}

/**
 * Analyze header name for semantic hints
 */
function analyzeHeaderSemantics(headerLower) {
  var classifications = ['status', 'category', 'type', 'priority', 'lead', 'tier', 'segment', 'stage', 'score', 'rating'];
  var summaries = ['description', 'summary', 'notes', 'details', 'content', 'body', 'text', 'bio', 'about'];
  var extractions = ['email', 'url', 'website', 'phone', 'address', 'name'];
  
  for (var i = 0; i < classifications.length; i++) {
    if (headerLower.includes(classifications[i])) {
      return {
        suggestedAction: 'classify',
        confidence: 0.7,
        reason: headerLower + ' typically needs categorization'
      };
    }
  }
  
  for (var j = 0; j < summaries.length; j++) {
    if (headerLower.includes(summaries[j])) {
      return {
        suggestedAction: 'summarize',
        confidence: 0.75,
        reason: headerLower + ' typically contains text to summarize'
      };
    }
  }
  
  for (var k = 0; k < extractions.length; k++) {
    if (headerLower.includes(extractions[k])) {
      return {
        suggestedAction: 'extract',
        confidence: 0.7,
        reason: headerLower + ' may contain extractable data'
      };
    }
  }
  
  return { suggestedAction: null };
}

/**
 * Create a suggestion object from detected pattern
 */
function createSuggestionFromPattern(pattern, inputCol, outputCol, headerName) {
  var icons = {
    classify: '🏷️',
    summarize: '📝',
    extract: '🔍'
  };
  
  var titles = {
    classify: 'Classify ' + headerName,
    summarize: 'Summarize ' + headerName,
    extract: 'Extract from ' + headerName
  };
  
  var command;
  var description;
  
  switch (pattern.suggestedAction) {
    case 'classify':
      if (pattern.type === 'numeric_variance') {
        command = 'Classify as High/Medium/Low based on column ' + inputCol + ' to column ' + outputCol;
        description = pattern.reason + ' - ideal for threshold-based classification';
      } else {
        command = 'Classify as Hot/Warm/Cold based on column ' + inputCol + ' to column ' + outputCol;
        description = pattern.reason;
      }
      break;
      
    case 'summarize':
      command = 'Summarize column ' + inputCol + ' to column ' + outputCol;
      description = pattern.reason + ' - condense for quick review';
      break;
      
    case 'extract':
      if (pattern.extractType === 'domain' && pattern.type === 'email') {
        command = 'Extract domains from column ' + inputCol + ' to column ' + outputCol;
        description = pattern.reason + ' - get company domains';
      } else if (pattern.extractType === 'domain' && pattern.type === 'url') {
        command = 'Extract domains from column ' + inputCol + ' to column ' + outputCol;
        description = pattern.reason + ' - extract website domains';
      } else {
        command = 'Extract key info from column ' + inputCol + ' to column ' + outputCol;
        description = pattern.reason;
      }
      break;
      
    default:
      return null;
  }
  
  return {
    type: 'proactive',
    icon: icons[pattern.suggestedAction] || '✨',
    title: titles[pattern.suggestedAction] || 'Process ' + headerName,
    description: description,
    command: command,
    confidence: pattern.confidence,
    reason: pattern.reason,
    patternType: pattern.type,
    priority: Math.round((1 - pattern.confidence) * 10)
  };
}

/**
 * Get sample values for a specific column
 */
function getSampleValuesForColumn(col, startRow, count) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = Math.min(startRow + count - 1, sheet.getLastRow());
    
    if (lastRow < startRow) return [];
    
    var range = sheet.getRange(col + startRow + ':' + col + lastRow);
    var values = range.getValues().flat();
    
    return values;
  } catch (e) {
    Logger.log('[Analyzer] Error getting column samples: ' + e.message);
    return [];
  }
}

/**
 * Format large numbers for display
 */
function formatNumber(num) {
  if (num >= 1000000) return (num / 1000000).toFixed(1) + 'M';
  if (num >= 1000) return (num / 1000).toFixed(1) + 'K';
  return num.toString();
}

// ============================================
// MAIN ANALYSIS FUNCTION
// ============================================

/**
 * Analyze spreadsheet data and generate suggestions
 * FORMULA-FIRST: Check for native formula solutions before using AI
 * 
 * Now includes pattern detection for proactive suggestions with confidence scores
 * 
 * @param {Object} context - Selection context from getAgentContext()
 * @return {Object} Analysis results with suggestions
 */
function analyzeDataForSuggestions(context) {
  if (!context || !context.selectionInfo) {
    Logger.log('[Analyzer] No context provided');
    return { suggestions: [], analysis: {} };
  }
  
  try {
    var startTime = Date.now();
    
    // Get sample data for analysis
    var sampleData = getSampleDataForAI(context);
    if (!sampleData || sampleData.rows.length === 0) {
      return { suggestions: [], analysis: { empty: true } };
    }
    
    // First try: Pattern-based suggestions (fast, no AI call needed)
    var patternSuggestions = generateProactiveSuggestions(context);
    Logger.log('[Analyzer] Pattern detection found ' + patternSuggestions.length + ' suggestions');
    
    // If pattern detection found high-confidence suggestions, use them
    var highConfidence = patternSuggestions.filter(function(s) { return s.confidence >= 0.7; });
    
    var suggestions;
    var method;
    
    if (highConfidence.length >= 2) {
      // Use pattern-based suggestions (faster, no API call)
      suggestions = patternSuggestions;
      method = 'pattern';
      Logger.log('[Analyzer] Using pattern-based suggestions (high confidence)');
    } else {
      // Fallback to AI for more nuanced analysis
      suggestions = getAISuggestions(sampleData, context);
      method = 'ai';
      Logger.log('[Analyzer] Using AI-based suggestions');
      
      // Merge with any pattern suggestions that AI might have missed
      if (patternSuggestions.length > 0) {
        var aiTypes = suggestions.map(function(s) { return s.command?.split(' ')[0]?.toLowerCase(); });
        patternSuggestions.forEach(function(ps) {
          var psType = ps.command?.split(' ')[0]?.toLowerCase();
          if (!aiTypes.includes(psType) && suggestions.length < MAX_SUGGESTIONS) {
            suggestions.push(ps);
          }
        });
      }
    }
    
    var analysis = {
      method: method,
      sampleSize: sampleData.rows.length,
      headers: sampleData.headers,
      suggestionCount: suggestions.length,
      patternCount: patternSuggestions.length,
      duration: Date.now() - startTime
    };
    
    Logger.log('[Analyzer] Analysis completed in ' + analysis.duration + 'ms: ' + 
               suggestions.length + ' suggestions (' + method + ')');
    
    return {
      suggestions: suggestions,
      analysis: analysis
    };
    
  } catch (e) {
    Logger.log('[Analyzer] Error: ' + e.message);
    // Fallback to basic suggestions if analysis fails
    return { 
      suggestions: getFallbackSuggestions(context), 
      analysis: { error: e.message, fallback: true } 
    };
  }
}

/**
 * Get pre-built formula suggestions based on data analysis
 * These are instant solutions that don't require AI
 */
function getPrebuiltFormulaSuggestions(context) {
  var suggestions = [];
  var selInfo = context.selectionInfo;
  
  if (!selInfo) return suggestions;
  
  var inputCol = selInfo.sourceColumn || selInfo.columnsWithData?.[0] || 'A';
  var emptyColRaw = selInfo.emptyColumns?.[0];
  var outputCol = typeof emptyColRaw === 'object' ? emptyColRaw.column : emptyColRaw || 'E';
  var startRow = selInfo.dataStartRow || 2;
  var cellRef = inputCol + startRow;
  
  // Analyze sample data to detect type
  var sampleData = getSampleDataForAI(context);
  if (!sampleData || sampleData.rows.length === 0) return suggestions;
  
  // Get first column values for analysis
  var firstColValues = sampleData.rows.map(function(row) {
    return row[0];
  }).filter(function(v) { return v !== null && v !== undefined && v !== ''; });
  
  if (firstColValues.length === 0) return suggestions;
  
  // Check if data looks like emails
  var emailPattern = /[\w.-]+@[\w.-]+\.\w+/;
  var emailCount = firstColValues.filter(function(v) { return emailPattern.test(String(v)); }).length;
  if (emailCount > firstColValues.length * 0.3) {
    // Data has emails - suggest domain extraction
    suggestions.push({
      type: 'formula',
      icon: '⚡',
      title: 'Extract Email Domain',
      description: 'Get domain from email addresses',
      action: 'apply-formula',
      formula: '=IFERROR(REGEXEXTRACT(' + cellRef + ',"@([\\w.-]+)"),"")',
      inputColumn: inputCol,
      outputColumn: outputCol,
      startRow: startRow,
      excelCompatible: false,
      priority: 0
    });
  }
  
  // Check if data looks like URLs
  var urlPattern = /https?:\/\//;
  var urlCount = firstColValues.filter(function(v) { return urlPattern.test(String(v)); }).length;
  if (urlCount > firstColValues.length * 0.3) {
    suggestions.push({
      type: 'formula',
      icon: '⚡',
      title: 'Extract Domain from URL',
      description: 'Get domain name from URLs',
      action: 'apply-formula',
      formula: '=IFERROR(REGEXEXTRACT(' + cellRef + ',"https?://([^/]+)"),"")',
      inputColumn: inputCol,
      outputColumn: outputCol,
      startRow: startRow,
      excelCompatible: false,
      priority: 0
    });
  }
  
  // Check if data is mostly numeric
  var numericCount = firstColValues.filter(function(v) {
    var num = parseFloat(String(v).replace(/[,$]/g, ''));
    return !isNaN(num);
  }).length;
  
  if (numericCount > firstColValues.length * 0.7) {
    // Numeric data - suggest threshold classification
    suggestions.push({
      type: 'formula',
      icon: '⚡',
      title: 'Classify by Thresholds',
      description: 'Categorize numbers into ranges',
      action: 'apply-formula',
      formula: '=IFS(' + cellRef + '>=1000,"Large",' + cellRef + '>=100,"Medium",TRUE,"Small")',
      inputColumn: inputCol,
      outputColumn: outputCol,
      startRow: startRow,
      excelCompatible: true,
      priority: 0
    });
  }
  
  // Check if data is text that could be cleaned
  var textValues = firstColValues.filter(function(v) {
    var str = String(v);
    return str.length > 3 && isNaN(parseFloat(str));
  });
  
  if (textValues.length > firstColValues.length * 0.5) {
    // Check for messy text (extra spaces, inconsistent case)
    var messyCount = textValues.filter(function(v) {
      var str = String(v);
      return str !== str.trim() || str !== str.toLowerCase() && str !== str.toUpperCase();
    }).length;
    
    if (messyCount > textValues.length * 0.3) {
      suggestions.push({
        type: 'formula',
        icon: '⚡',
        title: 'Clean & Format Text',
        description: 'Trim spaces, proper case',
        action: 'apply-formula',
        formula: '=PROPER(TRIM(' + cellRef + '))',
        inputColumn: inputCol,
        outputColumn: outputCol,
        startRow: startRow,
        excelCompatible: true,
        priority: 0
      });
    }
  }
  
  return suggestions;
}

// ============================================
// DATA SAMPLING FOR AI
// ============================================

/**
 * Get structured sample data for AI analysis
 * 
 * @param {Object} context - Selection context
 * @return {Object} Structured data with headers and sample rows
 */
function getSampleDataForAI(context) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var selInfo = context.selectionInfo;
    
    // Get headers
    var headers = [];
    if (selInfo && selInfo.headers) {
      headers = selInfo.headers.map(function(h) { 
        return { column: h.column, name: h.name || h.column };
      });
    }
    
    // Get data range
    var dataRange = selInfo?.dataRange;
    var dataRows = [];
    
    if (dataRange) {
      var rangeObj = sheet.getRange(dataRange);
      dataRows = rangeObj.getValues().slice(0, ANALYSIS_SAMPLE_SIZE);
    } else {
      // Fallback: get data from detected region
      var startRow = selInfo?.dataStartRow || 2;
      var endRow = Math.min(startRow + ANALYSIS_SAMPLE_SIZE - 1, selInfo?.dataEndRow || startRow + 9);
      var lastCol = sheet.getLastColumn();
      
      if (lastCol > 0 && endRow >= startRow) {
        var range = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol);
        dataRows = range.getValues();
      }
    }
    
    // Find empty columns for output suggestions
    // emptyColumns is array of objects {column: 'E', header: '...'} - extract just letters
    var emptyColumnsRaw = selInfo?.emptyColumns || [];
    var emptyColumns = emptyColumnsRaw.map(function(c) {
      return typeof c === 'object' ? c.column : c;
    });
    
    return {
      headers: headers,
      rows: dataRows,
      emptyColumns: emptyColumns,
      rowCount: selInfo?.dataRowCount || dataRows.length
    };
    
  } catch (e) {
    Logger.log('[Analyzer] getSampleDataForAI error: ' + e.message);
    return { headers: [], rows: [], emptyColumns: [] };
  }
}

// ============================================
// AI-POWERED SUGGESTIONS
// ============================================

/**
 * Use AI to generate contextual suggestions
 * Uses the user's selected model (not hardcoded)
 * 
 * @param {Object} sampleData - Structured sample data
 * @param {Object} context - Selection context
 * @return {Array} AI-generated suggestions
 */
function getAISuggestions(sampleData, context) {
  // Build context for AI
  var dataDescription = buildDataDescription(sampleData);
  
  // Get user's selected model and API key
  var provider = getAgentModel() || 'GEMINI';
  var apiKey = getUserApiKey(provider);
  
  if (!apiKey) {
    Logger.log('[Analyzer] No API key for ' + provider + ', using fallback');
    return getFallbackSuggestions(context);
  }
  
  // Build prompt
  var prompt = buildAnalysisPrompt(dataDescription, sampleData);
  
  try {
    // Use the user's selected model
    var response = callAIForAnalysis(prompt, provider, apiKey);
    
    if (response && response.suggestions) {
      // Validate and format suggestions
      return formatAISuggestions(response.suggestions, sampleData, context);
    }
    
  } catch (e) {
    Logger.log('[Analyzer] AI call failed: ' + e.message);
  }
  
  return getFallbackSuggestions(context);
}

/**
 * Build a description of the data for AI
 * Clearly distinguishes DATA columns from EMPTY columns
 */
function buildDataDescription(sampleData) {
  var description = 'SPREADSHEET DATA:\n\n';
  
  // Separate data columns from empty columns
  var dataColumns = [];
  var emptyColumns = sampleData.emptyColumns || [];
  
  // Headers - mark which have data
  if (sampleData.headers && sampleData.headers.length > 0) {
    description += 'COLUMNS WITH DATA (use as INPUT):\n';
    sampleData.headers.forEach(function(h) {
      if (emptyColumns.indexOf(h.column) === -1) {
        dataColumns.push(h.column);
        description += '- Column ' + h.column + ': "' + h.name + '"\n';
      }
    });
    description += '\n';
  } else if (sampleData.rows && sampleData.rows.length > 0) {
    // No headers - infer columns from sample data
    description += 'COLUMNS WITH DATA (use as INPUT):\n';
    var rowLength = sampleData.rows[0].length;
    for (var i = 0; i < rowLength; i++) {
      var col = String.fromCharCode(65 + i); // A, B, C, ...
      if (emptyColumns.indexOf(col) === -1) {
        // Check if this column has actual data
        var hasData = sampleData.rows.some(function(row) {
          return row[i] && String(row[i]).trim().length > 3;
        });
        if (hasData) {
          dataColumns.push(col);
          description += '- Column ' + col + ' (has data)\n';
        }
      }
    }
    description += '\n';
  }
  
  // Empty columns for output
  if (emptyColumns.length > 0) {
    description += 'EMPTY COLUMNS (use for OUTPUT): ' + emptyColumns.join(', ') + '\n\n';
  }
  
  // Sample rows (formatted as table with column labels)
  if (sampleData.rows && sampleData.rows.length > 0 && dataColumns.length > 0) {
    description += 'SAMPLE DATA:\n';
    description += 'Columns: ' + dataColumns.join(' | ') + '\n';
    
    sampleData.rows.forEach(function(row, i) {
      var rowStr = [];
      dataColumns.forEach(function(col) {
        var colIndex = col.charCodeAt(0) - 65; // A=0, B=1, etc.
        var val = String(row[colIndex] || '').substring(0, 40);
        if (val.length === 40) val += '...';
        rowStr.push(val || '(empty)');
      });
      description += (i + 1) + '. ' + rowStr.join(' | ') + '\n';
    });
  }
  
  return description;
}

/**
 * Build the analysis prompt for AI
 * IMPORTANT: Ensures diverse suggestions including CLASSIFY when data fits
 */
function buildAnalysisPrompt(dataDescription, sampleData) {
  var outputCol = sampleData.emptyColumns[0] || 'E';
  var emptyColumns = sampleData.emptyColumns || [];
  
  // Get data columns from headers, or infer from sample data
  var dataColumns = sampleData.headers
    .filter(function(h) { return emptyColumns.indexOf(h.column) === -1; })
    .map(function(h) { return h.column; });
  
  // Fallback: if no headers, infer columns from sample data width
  if (dataColumns.length === 0 && sampleData.rows && sampleData.rows.length > 0) {
    var rowLength = sampleData.rows[0].length;
    for (var i = 0; i < rowLength; i++) {
      var col = String.fromCharCode(65 + i); // A, B, C, ...
      if (emptyColumns.indexOf(col) === -1) {
        // Check if this column has actual data
        var hasData = sampleData.rows.some(function(row) {
          return row[i] && String(row[i]).trim().length > 3;
        });
        if (hasData) dataColumns.push(col);
      }
    }
    Logger.log('[Analyzer] Inferred data columns from sample data: ' + dataColumns.join(', '));
  }
  
  return `You are a data analysis assistant. Analyze this spreadsheet and suggest 2-3 DIVERSE tasks.

${dataDescription}

CRITICAL RULES:
1. INPUT must be from columns WITH DATA: ${dataColumns.join(', ')}
2. OUTPUT must go to an EMPTY column: ${sampleData.emptyColumns?.join(', ') || outputCol}
3. NEVER use empty columns as input - they have no data!
4. Command format: "[Action] [input description] to column [output]"

SUGGEST DIVERSE TASK TYPES - ALWAYS INCLUDE CLASSIFY IF DATA FITS:
1. CLASSIFY - For leads, companies, products, text that can be categorized
2. SUMMARIZE - For long text descriptions that need condensing
3. EXTRACT - For pulling specific info like names, dates, keywords

COMMAND FORMAT EXAMPLES:
- "Classify as Hot/Warm/Cold based on column B to column ${outputCol}"
- "Classify as High/Medium/Low priority based on column A to column ${outputCol}"
- "Summarize column B to column ${outputCol}"
- "Extract key topics from column B to column ${outputCol}"

REQUIRED COMMAND STRUCTURE:
1. Action verb FIRST (Classify, Summarize, Extract)
2. For Classify: categories AFTER "as" (e.g., "as Hot/Warm/Cold")
3. "based on column X" or "column X" for input
4. "to column ${outputCol}" at the END (ALWAYS!)

COMMAND PATTERNS:
✓ "Classify as Hot/Warm/Cold based on column B to column ${outputCol}"
✓ "Classify as Startup/SMB/Enterprise based on column A to column ${outputCol}"
✓ "Summarize column B to column ${outputCol}"
✓ "Extract company names from column B to column ${outputCol}"
✗ "Classify in column D" (wrong preposition - use "to column")
✗ "Summarize text" (missing output column!)

DIVERSITY REQUIREMENT:
- Suggest AT LEAST ONE "Classify" task if data contains business entities, leads, or categorizable items
- Suggest AT LEAST ONE "Summarize" task if data has text descriptions
- Each suggestion must use DIFFERENT action verbs when possible

OUTPUT FORMAT (JSON only, no markdown):
{
  "suggestions": [
    {
      "title": "Classify Leads",
      "description": "Categorize leads by potential value",
      "command": "Classify as Hot/Warm/Cold based on column B to column ${outputCol}",
      "icon": "🏷️",
      "priority": 1
    },
    {
      "title": "Summarize Descriptions",
      "description": "Condense long descriptions",
      "command": "Summarize column B to column ${outputCol}",
      "icon": "📝",
      "priority": 2
    }
  ]
}`;
}

/**
 * Call AI API for analysis using user's selected model
 * Supports: GEMINI, CHATGPT, CLAUDE, GROQ
 * 
 * @param {string} prompt - Analysis prompt
 * @param {string} provider - AI provider (GEMINI, CHATGPT, etc.)
 * @param {string} apiKey - User's API key
 * @return {Object|null} Parsed response or null
 */
function callAIForAnalysis(prompt, provider, apiKey) {
  var url, payload, options;
  
  switch (provider) {
    case 'GEMINI':
      url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
      payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.3,
          maxOutputTokens: 500,
          responseMimeType: 'application/json'
        }
      };
      break;
      
    case 'CHATGPT':
      url = 'https://api.openai.com/v1/chat/completions';
      payload = {
        model: 'gpt-5-mini',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.3,
        max_completion_tokens: 500,
        response_format: { type: 'json_object' }
      };
      break;
      
    case 'CLAUDE':
      url = 'https://api.anthropic.com/v1/messages';
      payload = {
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 500,
        messages: [{ role: 'user', content: prompt }]
      };
      break;
      
    case 'GROQ':
      url = 'https://api.groq.com/openai/v1/chat/completions';
      payload = {
        model: 'meta-llama/llama-4-scout-17b-16e-instruct',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.3,
        max_completion_tokens: 500,
        response_format: { type: 'json_object' }
      };
      break;
      
    default:
      Logger.log('[Analyzer] Unknown provider: ' + provider);
      return null;
  }
  
  // Build request options
  options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  // Add auth headers for non-Gemini providers
  if (provider === 'CHATGPT' || provider === 'GROQ') {
    options.headers = { 'Authorization': 'Bearer ' + apiKey };
  } else if (provider === 'CLAUDE') {
    options.headers = {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    };
  }
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log('[Analyzer] ' + provider + ' API error: ' + responseCode);
      return null;
    }
    
    var result = JSON.parse(response.getContentText());
    var text = extractTextFromResponse(result, provider);
    
    if (text) {
      return parseJSONResponse(text);
    }
    
  } catch (e) {
    Logger.log('[Analyzer] API call error: ' + e.message);
  }
  
  return null;
}

/**
 * Extract text content from different API response formats
 */
function extractTextFromResponse(result, provider) {
  try {
    switch (provider) {
      case 'GEMINI':
        return result.candidates?.[0]?.content?.parts?.[0]?.text;
        
      case 'CHATGPT':
      case 'GROQ':
        return result.choices?.[0]?.message?.content;
        
      case 'CLAUDE':
        return result.content?.[0]?.text;
        
      default:
        return null;
    }
  } catch (e) {
    return null;
  }
}

/**
 * Parse JSON response, handling markdown code blocks
 */
function parseJSONResponse(text) {
  if (!text) return null;
  
  try {
    text = text.trim();
    
    // Remove markdown code blocks if present
    if (text.startsWith('```json')) {
      text = text.replace(/```json\n?/, '').replace(/```$/, '');
    } else if (text.startsWith('```')) {
      text = text.replace(/```\n?/, '').replace(/```$/, '');
    }
    
    return JSON.parse(text);
  } catch (e) {
    Logger.log('[Analyzer] Failed to parse JSON: ' + e.message);
    Logger.log('[Analyzer] Raw text: ' + text.substring(0, 200));
    return null;
  }
}

/**
 * Format and validate AI suggestions
 * Filters out suggestions that reference empty columns as input
 */
function formatAISuggestions(suggestions, sampleData, context) {
  if (!Array.isArray(suggestions)) return [];
  
  var emptyColumns = sampleData.emptyColumns || [];
  var dataColumns = sampleData.headers
    .filter(function(h) { return emptyColumns.indexOf(h.column) === -1; })
    .map(function(h) { return h.column; });
  
  return suggestions
    .filter(function(s) {
      if (!s || !s.title || !s.command) return false;
      
      // Extract input column from command (e.g., "based on column D", "column B to")
      var inputMatch = s.command.match(/(?:based on|from)\s+column\s+([A-Z])/i) ||
                       s.command.match(/column\s+([A-Z])(?:\s+to|$)/i);
      
      if (inputMatch) {
        var inputCol = inputMatch[1].toUpperCase();
        // Reject if input column is empty
        if (emptyColumns.indexOf(inputCol) !== -1) {
          Logger.log('[Analyzer] Filtered out suggestion - input column ' + inputCol + ' is empty: ' + s.command);
          return false;
        }
        // Reject if input column is not in data columns
        if (dataColumns.length > 0 && dataColumns.indexOf(inputCol) === -1) {
          Logger.log('[Analyzer] Filtered out suggestion - input column ' + inputCol + ' has no data: ' + s.command);
          return false;
        }
      }
      
      return true;
    })
    .map(function(s) {
      return {
        type: 'ai',
        icon: s.icon || '✨',
        title: String(s.title).substring(0, 40),
        description: String(s.description || '').substring(0, 100),
        action: 'ai-suggestion',
        command: s.command,
        priority: s.priority || 3
      };
    })
    .slice(0, MAX_SUGGESTIONS);
}

// ============================================
// FALLBACK SUGGESTIONS (if AI fails)
// ============================================

/**
 * Generate basic fallback suggestions without AI
 * ALWAYS includes diverse options: Classify, Summarize, Extract
 */
function getFallbackSuggestions(context) {
  var suggestions = [];
  var selInfo = context?.selectionInfo;
  
  if (!selInfo) return suggestions;
  
  // Get first data column and empty column
  var dataCol = selInfo.columnsWithData?.[0] || 'A';
  var emptyColRaw = selInfo.emptyColumns?.[0];
  var emptyCol = (typeof emptyColRaw === 'object' ? emptyColRaw?.column : emptyColRaw) || 'E';
  
  if (selInfo.dataRowCount > 0) {
    // Always offer Classify first - most useful for business data
    suggestions.push({
      type: 'fallback',
      icon: '🏷️',
      title: 'Classify Data',
      description: 'Categorize items (Hot/Warm/Cold)',
      action: 'classify',
      command: 'Classify as Hot/Warm/Cold based on column ' + dataCol + ' to column ' + emptyCol,
      priority: 1
    });
    
    // Summarize for text content
    suggestions.push({
      type: 'fallback',
      icon: '📝',
      title: 'Summarize',
      description: 'Summarize the content in column ' + dataCol,
      action: 'summarize',
      command: 'Summarize column ' + dataCol + ' to column ' + emptyCol,
      priority: 2
    });
    
    // Extract for structured data
    suggestions.push({
      type: 'fallback',
      icon: '🔍',
      title: 'Extract Key Info',
      description: 'Extract key information from column ' + dataCol,
      action: 'extract',
      command: 'Extract key info from column ' + dataCol + ' to column ' + emptyCol,
      priority: 3
    });
  }
  
  return suggestions.slice(0, 3);
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Convert column index to letter (0 -> A, 1 -> B, etc.)
 */
function colToLetter(col) {
  var letter = '';
  while (col >= 0) {
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col = Math.floor(col / 26) - 1;
  }
  return letter;
}

// ============================================
// FORMULA-FIRST SUGGESTIONS (Using FormulaEngine)
// ============================================

/**
 * Check if a task can be done with native formulas
 * Uses FormulaEngine for intelligent formula generation
 * 
 * @param {Object} context - Selection context
 * @param {string} command - User command or suggestion
 * @return {Object|null} Formula suggestion or null
 */
function checkForFormulaSolution(context, command) {
  if (!context || !command) return null;
  
  var lower = command.toLowerCase();
  var selInfo = context.selectionInfo;
  
  // Build context for FormulaEngine
  var inputCol = selInfo?.sourceColumn || 'A';
  var emptyCol = selInfo?.emptyColumns?.[0];
  var outputCol = typeof emptyCol === 'object' ? emptyCol.column : emptyCol || 'E';
  var startRow = selInfo?.dataStartRow || 2;
  var cellRef = inputCol + startRow;
  
  var engineContext = {
    currentCell: cellRef,
    inputColumn: inputCol,
    outputColumn: outputCol,
    startRow: startRow,
    columns: {}
  };
  
  // Add column info
  if (selInfo?.headers) {
    selInfo.headers.forEach(function(h) {
      engineContext.columns[h.column] = {
        header: h.name,
        type: detectColumnType(context, h.column)
      };
    });
  }
  
  // Try to match formula-friendly tasks
  var formulaResult = matchFormulaTask(lower, cellRef, engineContext);
  
  if (formulaResult && formulaResult.formula) {
    return {
      type: 'formula',
      formulaType: formulaResult.taskType,
      formula: formulaResult.formula,
      description: formulaResult.description,
      excelCompatible: formulaResult.excelCompatible !== false,
      inputColumn: inputCol,
      outputColumn: outputCol,
      startRow: startRow
    };
  }
  
  return null;
}

/**
 * Match common formula tasks and generate formulas
 * This is GENERIC - not hardcoded cases
 */
function matchFormulaTask(command, cellRef, context) {
  var lower = command.toLowerCase();
  
  // ========== LOOKUP PATTERNS ==========
  if (lower.match(/vlookup|look\s*up|find.*in|search.*in|get.*from/)) {
    // Try to extract lookup details
    var lookupMatch = lower.match(/(?:from|in)\s+(?:sheet\s*)?["']?(\w+)["']?/i);
    var sheetName = lookupMatch ? lookupMatch[1] : 'Sheet2';
    
    // Use FormulaEngine's VLOOKUP pattern
    if (typeof FormulaPatterns !== 'undefined' && FormulaPatterns.vlookupSafe) {
      var result = FormulaPatterns.vlookupSafe(cellRef, sheetName + '!A:B', 2, context);
      if (result.success) {
        return {
          formula: result.formula,
          description: 'Lookup value from ' + sheetName,
          taskType: 'lookup',
          excelCompatible: true
        };
      }
    }
  }
  
  // ========== EXTRACTION PATTERNS ==========
  var extractionPatterns = {
    email: { regex: '[\\w.+-]+@[\\w.-]+\\.[a-zA-Z]{2,}', desc: 'Extract email address' },
    phone: { regex: '\\+?[\\d\\s\\-\\(\\)]{10,}', desc: 'Extract phone number' },
    url: { regex: 'https?://[^\\s]+', desc: 'Extract URL' },
    domain: { regex: '@([\\w.-]+)', desc: 'Extract email domain' },
    hashtag: { regex: '#\\w+', desc: 'Extract hashtag' },
    number: { regex: '[\\d,]+\\.?\\d*', desc: 'Extract number' },
    firstName: { formula: '=IFERROR(TRIM(LEFT({cell},FIND(" ",{cell}&" ")-1)),{cell})', desc: 'Extract first name', excel: true },
    lastName: { formula: '=IFERROR(TRIM(MID({cell},FIND("*",SUBSTITUTE({cell}," ","*",LEN({cell})-LEN(SUBSTITUTE({cell}," ",""))))+1,100)),"")', desc: 'Extract last name', excel: true }
  };
  
  for (var key in extractionPatterns) {
    var pattern = extractionPatterns[key];
    var keyRegex = new RegExp('extract|get|pull|find').test(lower) && lower.includes(key.replace(/([A-Z])/g, ' $1').toLowerCase().trim());
    
    if (keyRegex || (key === 'email' && lower.match(/email/)) || 
        (key === 'phone' && lower.match(/phone|call|mobile/)) ||
        (key === 'url' && lower.match(/url|link|website/)) ||
        (key === 'domain' && lower.match(/domain/)) ||
        (key === 'firstName' && lower.match(/first\s*name/)) ||
        (key === 'lastName' && lower.match(/last\s*name|surname/))) {
      
      if (pattern.formula) {
        return {
          formula: pattern.formula.replace(/\{cell\}/g, cellRef),
          description: pattern.desc,
          taskType: 'extract_' + key,
          excelCompatible: pattern.excel !== false
        };
      } else {
        return {
          formula: '=IFERROR(REGEXEXTRACT(' + cellRef + ',"' + pattern.regex + '"),"")',
          description: pattern.desc,
          taskType: 'extract_' + key,
          excelCompatible: false
        };
      }
    }
  }
  
  // ========== CLEANING PATTERNS ==========
  var cleaningPatterns = {
    trim: { formula: '=TRIM({cell})', desc: 'Remove extra spaces', excel: true },
    clean: { formula: '=CLEAN(TRIM({cell}))', desc: 'Clean text', excel: true },
    proper: { formula: '=PROPER(TRIM({cell}))', desc: 'Title Case', excel: true },
    upper: { formula: '=UPPER(TRIM({cell}))', desc: 'UPPERCASE', excel: true },
    lower: { formula: '=LOWER(TRIM({cell}))', desc: 'lowercase', excel: true },
    removeNumbers: { formula: '=REGEXREPLACE({cell},"[0-9]","")', desc: 'Remove numbers', excel: false },
    onlyNumbers: { formula: '=REGEXREPLACE({cell},"[^0-9.]","")', desc: 'Keep only numbers', excel: false },
    removeSpecial: { formula: '=REGEXREPLACE({cell},"[^a-zA-Z0-9\\s]","")', desc: 'Remove special chars', excel: false }
  };
  
  for (var cleanKey in cleaningPatterns) {
    var cleanPattern = cleaningPatterns[cleanKey];
    
    if ((cleanKey === 'trim' && lower.match(/trim|remove\s*space|whitespace/)) ||
        (cleanKey === 'clean' && lower.match(/clean|fix/)) ||
        (cleanKey === 'proper' && lower.match(/proper|title\s*case|capitalize/)) ||
        (cleanKey === 'upper' && lower.match(/upper|caps/)) ||
        (cleanKey === 'lower' && lower.match(/lower/)) ||
        (cleanKey === 'removeNumbers' && lower.match(/remove.*number/)) ||
        (cleanKey === 'onlyNumbers' && lower.match(/only.*number|keep.*number/)) ||
        (cleanKey === 'removeSpecial' && lower.match(/remove.*special|remove.*symbol/))) {
      
      return {
        formula: cleanPattern.formula.replace(/\{cell\}/g, cellRef),
        description: cleanPattern.desc,
        taskType: 'clean_' + cleanKey,
        excelCompatible: cleanPattern.excel
      };
    }
  }
  
  // ========== CLASSIFICATION PATTERNS ==========
  // Check for explicit thresholds in command
  var thresholdMatch = parseThresholdsFromCommand(lower);
  if (thresholdMatch) {
    return {
      formula: buildIFSFormula(thresholdMatch, cellRef),
      description: 'Classify by thresholds',
      taskType: 'classify_threshold',
      excelCompatible: true
    };
  }
  
  // Check for keyword-based classification
  var keywordMatch = parseKeywordsFromCommand(lower);
  if (keywordMatch) {
    return {
      formula: buildKeywordIFSFormula(keywordMatch, cellRef),
      description: 'Classify by keywords',
      taskType: 'classify_keyword',
      excelCompatible: false  // Uses REGEXMATCH
    };
  }
  
  // ========== DATE PATTERNS ==========
  if (lower.match(/age|how\s*old|days?\s*(since|ago|until)/)) {
    return {
      formula: '=DATEDIF(' + cellRef + ',TODAY(),"D")',
      description: 'Calculate days since date',
      taskType: 'date_diff',
      excelCompatible: true
    };
  }
  
  if (lower.match(/year|month|day/) && lower.match(/extract|get/)) {
    var datePart = lower.match(/year/) ? 'YEAR' : lower.match(/month/) ? 'MONTH' : 'DAY';
    return {
      formula: '=' + datePart + '(' + cellRef + ')',
      description: 'Extract ' + datePart.toLowerCase() + ' from date',
      taskType: 'date_extract',
      excelCompatible: true
    };
  }
  
  // ========== MATH PATTERNS ==========
  if (lower.match(/sum|total|add.*up/)) {
    return {
      formula: '=SUM(' + cellRef.replace(/\d+/, '') + ':' + cellRef.replace(/\d+/, '') + ')',
      description: 'Sum values in column',
      taskType: 'math_sum',
      excelCompatible: true
    };
  }
  
  if (lower.match(/average|mean/)) {
    return {
      formula: '=AVERAGE(' + cellRef.replace(/\d+/, '') + ':' + cellRef.replace(/\d+/, '') + ')',
      description: 'Average of values',
      taskType: 'math_average',
      excelCompatible: true
    };
  }
  
  if (lower.match(/count|how\s*many/)) {
    return {
      formula: '=COUNTA(' + cellRef.replace(/\d+/, '') + ':' + cellRef.replace(/\d+/, '') + ')',
      description: 'Count non-empty cells',
      taskType: 'math_count',
      excelCompatible: true
    };
  }
  
  if (lower.match(/round/)) {
    var decimals = (lower.match(/(\d+)\s*decimal/) || [0, 2])[1];
    return {
      formula: '=ROUND(' + cellRef + ',' + decimals + ')',
      description: 'Round to ' + decimals + ' decimals',
      taskType: 'math_round',
      excelCompatible: true
    };
  }
  
  // ========== TEXT COMBINATION ==========
  if (lower.match(/combin|concat|join|merge/)) {
    var delimiter = lower.match(/with\s+["'](.+?)["']/) ? RegExp.$1 : ', ';
    return {
      formula: '=TEXTJOIN("' + delimiter + '",TRUE,' + cellRef + ')',
      description: 'Join text with delimiter',
      taskType: 'text_join',
      excelCompatible: true
    };
  }
  
  if (lower.match(/split|separate/)) {
    var splitDelim = lower.match(/by\s+["'](.+?)["']/) ? RegExp.$1 : ',';
    return {
      formula: '=SPLIT(' + cellRef + ',"' + splitDelim + '")',
      description: 'Split text by delimiter',
      taskType: 'text_split',
      excelCompatible: false
    };
  }
  
  // ========== LENGTH/SIZE ==========
  if (lower.match(/length|how\s*long|char.*count/)) {
    return {
      formula: '=LEN(' + cellRef + ')',
      description: 'Count characters',
      taskType: 'text_length',
      excelCompatible: true
    };
  }
  
  if (lower.match(/word.*count|how\s*many\s*word/)) {
    return {
      formula: '=IF(LEN(TRIM(' + cellRef + '))=0,0,LEN(TRIM(' + cellRef + '))-LEN(SUBSTITUTE(TRIM(' + cellRef + ')," ",""))+1)',
      description: 'Count words',
      taskType: 'text_wordcount',
      excelCompatible: true
    };
  }
  
  return null;
}

/**
 * Parse threshold values from command text
 * E.g., "<100 small, 100-500 medium, >500 large"
 */
function parseThresholdsFromCommand(command) {
  var patterns = [
    // <100 = small, 100-500 = medium
    /([<>]=?)\s*(\d+)\s*[=:→]?\s*["']?(\w+)["']?/gi,
    // classify as small if <100
    /["']?(\w+)["']?\s+(?:if|when)\s+([<>]=?)\s*(\d+)/gi
  ];
  
  var thresholds = [];
  var match;
  
  // Try first pattern
  var regex1 = /([<>]=?)\s*(\d+)\s*[=:→]?\s*["']?([a-zA-Z]+)["']?/gi;
  while ((match = regex1.exec(command)) !== null) {
    thresholds.push({
      operator: match[1],
      value: parseInt(match[2]),
      category: match[3]
    });
  }
  
  if (thresholds.length >= 2) {
    return thresholds;
  }
  
  return null;
}

/**
 * Parse keyword mappings from command
 * E.g., "enterprise/fortune = hot, startup = warm"
 */
function parseKeywordsFromCommand(command) {
  var patterns = /["']?([a-zA-Z\|\/]+)["']?\s*[=:→]\s*["']?(\w+)["']?/gi;
  var mappings = [];
  var match;
  
  while ((match = patterns.exec(command)) !== null) {
    var keywords = match[1].split(/[|\/]/);
    if (keywords.length > 0 && keywords[0].length > 2) {
      mappings.push({
        keywords: keywords,
        category: match[2]
      });
    }
  }
  
  return mappings.length >= 2 ? mappings : null;
}

/**
 * Build IFS formula from thresholds
 */
function buildIFSFormula(thresholds, cellRef) {
  // Sort by value descending for >= operators
  thresholds.sort(function(a, b) { return b.value - a.value; });
  
  var parts = thresholds.map(function(t) {
    return cellRef + t.operator + t.value + ',"' + t.category + '"';
  });
  
  parts.push('TRUE,"Other"');
  
  return '=IFS(' + parts.join(',') + ')';
}

/**
 * Build IFS formula with REGEXMATCH for keywords
 */
function buildKeywordIFSFormula(mappings, cellRef) {
  var parts = mappings.map(function(m) {
    var pattern = m.keywords.join('|');
    return 'REGEXMATCH(LOWER(' + cellRef + '),"' + pattern + '"),"' + m.category + '"';
  });
  
  parts.push('TRUE,"Other"');
  
  return '=IFS(' + parts.join(',') + ')';
}

/**
 * Detect column data type from sample values
 */
function detectColumnType(context, column) {
  // This would analyze sample data from the column
  // For now, return 'unknown' - could be enhanced
  return 'unknown';
}

/**
 * Generate formula-first suggestion for the suggestions panel
 */
function getFormulaSuggestion(formulaSolution, context) {
  var emptyColRaw = context.selectionInfo?.emptyColumns?.[0];
  var outputCol = typeof emptyColRaw === 'object' ? emptyColRaw.column : emptyColRaw;
  
  return {
    type: 'formula',
    icon: '⚡',
    title: 'Instant Formula',
    description: formulaSolution.description,
    action: 'apply-formula',
    formula: formulaSolution.formula,
    inputColumn: formulaSolution.inputColumn,
    outputColumn: outputCol || formulaSolution.outputColumn,
    startRow: formulaSolution.startRow,
    excelCompatible: formulaSolution.excelCompatible,
    priority: 0  // Highest priority
  };
}

// ============================================
// MODULE LOADED
// ============================================

Logger.log('📦 AgentAnalyzer v2.1 (Formula-First) loaded');
