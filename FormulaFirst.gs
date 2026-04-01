/**
 * @file FormulaFirst.gs
 * @version 1.1.0
 * @updated 2026-01-18
 * @description AI-Driven Formula-First Evaluation System
 * 
 * CHANGELOG:
 * - 1.1.0 (2026-01-18): Enhanced explanations for Decision Transparency UI
 *   - Rich, user-friendly explanations always provided
 *   - Context-aware reasoning for both formula and AI decisions
 *   - Better fallback explanations
 * - 1.0.0: Initial AI-driven formula evaluation
 * 
 * VISION: Use AI intelligence to DECIDE, not just to PROCESS.
 * 
 * One smart AI call upfront to determine:
 * - Can this be done with a native Google Sheets formula?
 * - If YES → Generate the formula (FREE, instant, auto-updating)
 * - If NO → Explain why and proceed with AI processing
 * 
 * This is GENERIC - no hardcoded rules for specific cases.
 * AI sees ALL data and makes the optimal decision.
 */

// ============================================
// MAIN ENTRY POINT
// ============================================

/**
 * Evaluate if a task can be solved with a formula
 * This is the GENERIC entry point - AI decides everything
 * 
 * @param {Object} context - Full task context
 * @param {string} context.command - User's original command
 * @param {string} context.taskType - Detected task type
 * @param {string} context.inputColumn - Suggested input column
 * @param {string} context.outputColumn - Target output column
 * @param {Array} context.sampleData - Sample input values
 * @param {string} context.categories - For classify tasks
 * @param {Array} context.allColumns - All available columns
 * @param {Object} context.columnHeaders - Column header names
 * @param {number} context.dataStartRow - First data row
 * @param {number} context.dataEndRow - Last data row
 * @return {Object} { canUseFormula, formula, explanation, inputColumn }
 */
function evaluateFormulaFirst(context) {
  Logger.log('[FormulaFirst] Evaluating with full context');
  
  // Quick rejection for tasks that ALWAYS need AI (semantic understanding)
  // With rich, user-friendly explanations
  // NOTE: 'translate' is NOT included here - let AI evaluation decide if GOOGLETRANSLATE() is sufficient
  var aiOnlyTasks = {
    'summarize': {
      explanation: 'Summarization requires understanding context, key points, and what matters most. AI can read and comprehend your text to create meaningful summaries.',
      tip: 'For a quick preview, you could use =LEFT(cell, 100) & "..." to show first 100 characters.'
    },
    'summarise': {
      explanation: 'Summarization requires understanding context, key points, and what matters most. AI can read and comprehend your text to create meaningful summaries.',
      tip: 'For a quick preview, you could use =LEFT(cell, 100) & "..." to show first 100 characters.'
    },
    'sentiment': {
      explanation: 'Sentiment analysis requires understanding emotional tone, sarcasm, and context that formulas cannot detect.',
      tip: 'AI can classify text as positive, negative, or neutral with high accuracy.'
    },
    'explain': {
      explanation: 'Explanations require understanding complex concepts and generating clear, contextual descriptions.',
      tip: 'AI excels at breaking down technical terms into simple language.'
    },
    'rewrite': {
      explanation: 'Rewriting text requires understanding meaning and style to produce natural-sounding alternatives.',
      tip: 'AI can adjust tone, simplify language, or make text more professional.'
    },
    'improve': {
      explanation: 'Improving text requires understanding quality, clarity, and context to make meaningful enhancements.',
      tip: 'AI can fix grammar, improve flow, and enhance readability.'
    },
    'generate content': {
      explanation: 'Content generation creates new, unique text that cannot be produced by formulas.',
      tip: 'For templated content, consider using =CONCATENATE() with variables.'
    }
  };
  
  var commandLower = (context.command || '').toLowerCase();
  
  for (var keyword in aiOnlyTasks) {
    if (commandLower.includes(keyword)) {
      var task = aiOnlyTasks[keyword];
      return {
        canUseFormula: false,
        explanation: task.explanation,
        tip: task.tip,
        reason: 'semantic_understanding'
      };
    }
  }
  
  // For everything else: Let AI decide with FULL context
  return evaluateWithAI(context);
}

// ============================================
// AI-DRIVEN EVALUATION (THE CORE)
// ============================================

/**
 * Let AI evaluate the best approach with FULL context
 * This is the GENERIC solution - AI sees everything and decides
 */
function evaluateWithAI(context) {
  Logger.log('[FormulaFirst] AI evaluation with full context');
  
  // Get user's API key and model
  var provider = getAgentModel() || 'GEMINI';
  var apiKey = getUserApiKey(provider);
  
  if (!apiKey) {
    Logger.log('[FormulaFirst] No API key available');
    return {
      canUseFormula: false,
      explanation: 'API key needed for smart formula detection. Using AI processing instead.',
      reason: 'no_api_key'
    };
  }
  
  // Build comprehensive context for AI
  var prompt = buildGenericFormulaPrompt(context);
  
  try {
    var response = callFormulaAI(prompt, provider, apiKey);
    
    if (response && response.canUseFormula && response.formula) {
      Logger.log('[FormulaFirst] AI generated formula: ' + response.formula);
      
      // Ensure explanation is rich and user-friendly
      var explanation = response.explanation || buildFormulaExplanation(response, context);
      
      return {
        canUseFormula: true,
        formula: response.formula,
        explanation: explanation,
        formulaType: response.formulaType || 'ai-generated',
        suggestedInputColumn: response.inputColumn,
        benefits: ['Auto-updates when data changes']
      };
    }
    
    // Build rich explanation for why AI is needed
    var aiExplanation = response?.explanation || buildAIRequiredExplanation(context);
    
    return {
      canUseFormula: false,
      explanation: aiExplanation,
      reason: response?.reason || 'semantic_understanding'
    };
    
  } catch (e) {
    Logger.log('[FormulaFirst] AI evaluation error: ' + e.message);
    return {
      canUseFormula: false,
      explanation: 'Proceeding with AI processing for reliable results.',
      reason: 'evaluation_error',
      error: e.message
    };
  }
}

/**
 * Build a user-friendly explanation for formula decisions
 */
function buildFormulaExplanation(response, context) {
  var formulaType = response.formulaType || 'custom';
  var taskType = context.taskType?.toLowerCase() || '';
  
  var explanations = {
    'threshold': 'Detected numeric thresholds in your data. Using IFS formula to classify based on values.',
    'regex': 'Found a pattern that can be extracted with REGEXEXTRACT formula.',
    'lookup': 'This lookup can be done with VLOOKUP/INDEX+MATCH formulas.',
    'text': 'Text transformation detected - using native text formulas.',
    'date': 'Date calculation detected - using date formulas.'
  };
  
  if (explanations[formulaType]) {
    return explanations[formulaType];
  }
  
  // Fallback based on task type
  if (taskType === 'classify' || taskType === 'categorize') {
    return 'Classification can be done with formula based on your data patterns.';
  }
  
  if (taskType === 'extract') {
    return 'Extraction pattern detected - using formula.';
  }
  
  return 'Formula identified for this task.';
}

/**
 * Build explanation for why AI is required
 */
function buildAIRequiredExplanation(context) {
  var taskType = context.taskType?.toLowerCase() || '';
  var command = context.command?.toLowerCase() || '';
  
  // Task-specific explanations
  if (taskType === 'classify' || taskType === 'categorize') {
    if (command.includes('company') || command.includes('business') || command.includes('industry')) {
      return 'Classifying companies requires understanding business context, industry knowledge, and semantic meaning that formulas cannot assess.';
    }
    return 'This classification requires understanding the meaning of your data - AI will analyze context to categorize accurately.';
  }
  
  if (taskType === 'extract') {
    return 'The information you want to extract requires understanding context and meaning, not just pattern matching.';
  }
  
  if (taskType === 'generate') {
    return 'Content generation creates unique text based on your data - this requires AI creativity.';
  }
  
  // Generic explanation
  return 'This task requires semantic understanding that formulas cannot provide. AI will analyze your data for accurate results.';
}

/**
 * Build the GENERIC prompt for AI formula evaluation
 * This prompt gives AI ALL the context to make the best decision
 * 
 * Uses unified context when available for richer information
 */
function buildGenericFormulaPrompt(context) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = context.dataStartRow || 2;
  var endRow = context.dataEndRow || Math.min(startRow + 9, sheet.getLastRow());
  
  // Use pre-built data summary if available (from unified context)
  if (context.dataSummary) {
    Logger.log('[FormulaFirst] Using pre-built data summary from unified context');
  }
  
  // Gather sample data from ALL relevant columns
  var columnSamples = {};
  var allCols = context.allColumns || context.inputColumns || [context.inputColumn];
  
  // If we have columnDetails from unified context, use them
  if (context.columnDetails && context.columnDetails.length > 0) {
    context.columnDetails.forEach(function(col) {
      columnSamples[col.column] = {
        header: col.header,
        samples: col.sampleValues || [],
        type: col.dataType,
        isEmpty: col.isEmpty,
        numericRatio: col.numericRatio
      };
    });
  } else {
    // Fallback: gather data manually
    for (var i = 0; i < Math.min(allCols.length, 6); i++) {
      var col = allCols[i];
      var header = context.columnHeaders?.[col] || 'Column ' + col;
      var samples = getSampleDataForFormula(col + startRow + ':' + col + endRow, 5);
      
      // Detect data type
      var isNumeric = samples.length > 0 && samples.filter(function(v) {
        return !isNaN(parseFloat(v));
      }).length > samples.length * 0.7;
      
      columnSamples[col] = {
        header: header,
        samples: samples.slice(0, 5),
        type: isNumeric ? 'numeric' : 'text',
        isEmpty: samples.length === 0
      };
    }
  }
  
  var prompt = `You are a Google Sheets formula expert. Your job is to decide the BEST way to fulfill a user's request.

## USER REQUEST
"${context.command}"

## AVAILABLE DATA
Output Column: ${context.outputColumn} (where results should go)
Data Range: Row ${startRow} to ${endRow}

### Columns Available:
${Object.keys(columnSamples).map(function(col) {
  var info = columnSamples[col];
  return `- Column ${col} ("${info.header}"): ${info.type}${info.isEmpty ? ' [EMPTY]' : ''}
    Samples: ${info.samples.slice(0, 3).map(function(s) { return '"' + String(s).substring(0, 30) + '"'; }).join(', ')}`;
}).join('\n')}

${context.categories ? '### Target Categories: ' + context.categories : ''}

## YOUR TASK
Decide: Can this be accomplished with a native Google Sheets formula?

FORMULA IS BEST WHEN:
- Task involves numeric thresholds (e.g., classify by size >1000 = "Hot")
- Task involves pattern extraction (emails, URLs, numbers)
- Task involves lookups or comparisons
- Task involves text manipulation (uppercase, trim, concatenate)
- Task involves date/time calculations
- Task involves SIMPLE TRANSLATION (use GOOGLETRANSLATE function)

AI IS REQUIRED WHEN:
- Task requires understanding meaning/context (summarize, sentiment, nuanced translation)
- Task requires external knowledge (e.g., "Is this company in tech sector?")
- Task requires nuanced judgment that varies by context
- Translation needs to preserve TONE, INTENT, or MARKETING MESSAGE (not just literal translation)

## RESPONSE FORMAT (JSON only, no markdown)
If formula is possible:
{
  "canUseFormula": true,
  "formula": "=IFS(D{ROW}>1000,\\"Hot\\",D{ROW}>100,\\"Warm\\",TRUE,\\"Cold\\")",
  "explanation": "Detected numeric values in Size column (D). Using thresholds: values over 1000 = Hot, over 100 = Warm, otherwise Cold.",
  "formulaType": "threshold",
  "inputColumn": "D"
}

For translation (simple, literal):
{
  "canUseFormula": true,
  "formula": "=GOOGLETRANSLATE(B{ROW},\\"en\\",\\"es\\")",
  "explanation": "Simple translation can use GOOGLETRANSLATE formula - it's free, instant, and auto-updates. Source language detected as English, target is Spanish.",
  "formulaType": "translate",
  "inputColumn": "B"
}

If AI is required:
{
  "canUseFormula": false,
  "explanation": "This requires understanding the business context of each company to classify leads - formulas cannot assess semantic meaning like industry, company size implications, or market position."
}

CRITICAL RULES:
1. Use {ROW} as placeholder for row number (will be replaced with actual row)
2. Choose the MOST APPROPRIATE column for the task (not necessarily the one user mentioned)
3. For numeric thresholds, analyze the data distribution to pick smart breakpoints
4. Always escape quotes in formulas with backslash
5. Be AGGRESSIVE about using formulas - they auto-update when data changes
6. ALWAYS provide a clear, user-friendly explanation that explains WHY (not just what)
7. For AI-required tasks, explain what semantic understanding is needed`;

  return prompt;
}

// ============================================
// AI COMMUNICATION
// ============================================

/**
 * Call AI for formula evaluation
 * Uses a lightweight, fast model call
 */
function callFormulaAI(prompt, provider, apiKey) {
  // Build payload using SecureRequest (centralized API key handling)
  // Note: apiKey param is legacy, we now get it from SecureRequest
  var payload = SecureRequest.buildPayloadWithUser(provider, {
    model: provider,
    input: prompt,
    specificModel: getFormulaModel(provider),
    taskType: 'CODE' // Optimize for code/formula generation
  });
  
  try {
    var result = ApiClient.post('QUERY', payload);
    var response = result.result || '';
    
    // Parse JSON response
    var jsonMatch = response.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    
    Logger.log('[FormulaFirst] Could not parse AI response: ' + response.substring(0, 200));
    return null;
    
  } catch (e) {
    Logger.log('[FormulaFirst] AI call error: ' + e.message);
    throw e;
  }
}

/**
 * Get the best model for formula generation
 * Uses centralized Config to ensure model names match the database
 */
function getFormulaModel(provider) {
  // Use centralized config to ensure model names are in sync with backend
  return Config.getDefaultModel(provider);
}

// ============================================
// FORMULA APPLICATION
// ============================================

/**
 * Apply a generated formula to the sheet
 * 
 * @param {Object} formulaResult - Result from evaluateFormulaFirst
 * @param {string} outputColumn - Target column letter
 * @param {number} startRow - First row to apply
 * @param {number} endRow - Last row to apply
 * @return {Object} { success, message, rowsAffected }
 */
function applyGeneratedFormula(formulaResult, outputColumn, startRow, endRow) {
  Logger.log('[FormulaFirst] Applying formula to ' + outputColumn + startRow + ':' + outputColumn + endRow);
  
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var formula = formulaResult.formula;
    
    // Apply formula to each row
    for (var row = startRow; row <= endRow; row++) {
      var rowFormula = formula.replace(/\{ROW\}/g, row);
      var cell = sheet.getRange(outputColumn + row);
      cell.setFormula(rowFormula);
    }
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: 'Formula applied to ' + (endRow - startRow + 1) + ' rows',
      rowsAffected: endRow - startRow + 1,
      formula: formula
    };
    
  } catch (e) {
    Logger.log('[FormulaFirst] Error applying formula: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Get sample data from a range for evaluation
 */
function getSampleDataForFormula(inputRange, sampleSize) {
  sampleSize = sampleSize || 10;
  
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(inputRange);
    var values = range.getValues();
    
    // Flatten and filter non-empty values
    var samples = [];
    for (var i = 0; i < values.length && samples.length < sampleSize; i++) {
      var val = values[i][0];
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        samples.push(val);
      }
    }
    
    return samples;
    
  } catch (e) {
    Logger.log('[FormulaFirst] Error getting sample data: ' + e.message);
    return [];
  }
}

/**
 * Get user's selected AI model for agent tasks
 */
function getAgentModel() {
  try {
    var settings = getUserSettings();
    return settings?.agentModel || 'GEMINI';
  } catch (e) {
    return 'GEMINI';
  }
}

/**
 * Get user's API key for a provider
 */
function getUserApiKey(provider) {
  try {
    var settings = getUserSettings();
    return settings?.[provider]?.apiKey || null;
  } catch (e) {
    return null;
  }
}
