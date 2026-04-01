/**
 * @file AgentTaskChain_Plan.gs
 * @version 2.0.0
 * @updated 2026-02-10
 * @author AISheeter Team
 * 
 * ============================================
 * AGENT TASK CHAIN - Plan Building
 * ============================================
 * 
 * Builds execution plans for individual chain steps:
 * - Determine input/output columns and ranges
 * - Handle multi-aspect output detection
 * - Format instructions and smart headers
 * 
 * Part of the AgentTaskChain module suite:
 * - AgentTaskChain_Parse.gs    (parsing & suggestions)
 * - AgentTaskChain_Execute.gs  (execution flow)
 * - AgentTaskChain_Plan.gs     (this file)
 * - AgentTaskChain_State.gs    (state management)
 * - AgentTaskChain_Analyze.gs  (analyze/chat steps)
 */

// ============================================
// PLAN BUILDING FOR CHAIN STEPS
// ============================================

/**
 * Build a plan for a chain step - REUSES the same plan format as single-step
 * This ensures chain steps use the same intelligent parsing/execution logic
 * 
 * @param {Object} step - The step definition
 * @param {Object} chainState - The chain state
 * @param {number} stepIndex - Current step index
 * @return {Object} Plan in the same format as parseCommandWithAI returns
 */
function buildPlanForChainStep(step, chainState, stepIndex) {
  Logger.log('[buildPlan] Building plan for step ' + (stepIndex + 1));
  Logger.log('[buildPlan] chainState.inputRange: ' + chainState.inputRange);
  Logger.log('[buildPlan] chainState.inputColumn: ' + chainState.inputColumn);
  Logger.log('[buildPlan] chainState.inputColumns: ' + JSON.stringify(chainState.inputColumns));
  Logger.log('[buildPlan] chainState.inputColumns type: ' + typeof chainState.inputColumns);
  Logger.log('[buildPlan] chainState.inputColumns isArray: ' + Array.isArray(chainState.inputColumns));
  Logger.log('[buildPlan] chainState.hasMultipleInputColumns: ' + chainState.hasMultipleInputColumns);
  
  var inputRange;
  var inputColumn;
  var inputColumns;
  var hasMultipleInputColumns = false;
  
  // Log step info
  Logger.log('[buildPlan] step.inputColumns: ' + JSON.stringify(step.inputColumns));
  Logger.log('[buildPlan] step.inputColumns type: ' + typeof step.inputColumns);
  Logger.log('[buildPlan] step.inputColumns isArray: ' + Array.isArray(step.inputColumns));
  Logger.log('[buildPlan] step.outputColumn: ' + step.outputColumn);
  
  // PRIORITY: Use step-level column configuration if available (from backend)
  // This allows each step to have its own input/output column assignments
  var stepInputCols = step.inputColumns;
  
  // DEFENSIVE: Handle cases where inputColumns might be object (from JSON serialization)
  if (stepInputCols && !Array.isArray(stepInputCols)) {
    if (typeof stepInputCols === 'object') {
      stepInputCols = Object.values(stepInputCols);
    } else if (typeof stepInputCols === 'string') {
      stepInputCols = stepInputCols.split(',');
    }
    Logger.log('[buildPlan] Converted step.inputColumns to array: ' + JSON.stringify(stepInputCols));
  }
  
  if (stepInputCols && Array.isArray(stepInputCols) && stepInputCols.length > 0) {
    Logger.log('[buildPlan] Using STEP-level inputColumns: ' + stepInputCols.join(','));
    inputColumns = stepInputCols;
    inputColumn = inputColumns[0];
    hasMultipleInputColumns = inputColumns.length > 1;
    
    // Build inputRange from the columns - use chainState row numbers
    var rowMatch = chainState.inputRange ? chainState.inputRange.match(/\d+/g) : null;
    var startRow = rowMatch ? rowMatch[0] : '2';
    var endRow = rowMatch && rowMatch.length > 1 ? rowMatch[rowMatch.length - 1] : String(parseInt(startRow) + (chainState.rowCount || 10) - 1);
    inputRange = inputColumns[0] + startRow + ':' + inputColumns[inputColumns.length - 1] + endRow;
    Logger.log('[buildPlan] Built inputRange from step columns: ' + inputRange);
  }
  // Fallback: Determine input source from chain state
  else if (step.usesResultOf && stepIndex > 0) {
    // Use output from previous step as input
    var prevStep = chainState.steps[stepIndex - 1];
    if (prevStep.outputRange) {
      inputRange = prevStep.outputRange;
      inputColumn = prevStep.outputRange.charAt(0);
      Logger.log('[buildPlan] Using previous step output: ' + inputRange);
    } else {
      // Fallback to original input range
      inputRange = chainState.inputRange;
      inputColumn = chainState.inputColumn;
    }
  } else {
    // Use original input from chain (step 1)
    inputRange = chainState.inputRange;
    inputColumn = chainState.inputColumn;
    inputColumns = chainState.inputColumns;
    hasMultipleInputColumns = chainState.hasMultipleInputColumns || false;
    
    // DEFENSIVE: Ensure inputColumns is an array
    if (inputColumns && !Array.isArray(inputColumns)) {
      Logger.log('[buildPlan] ⚠️ inputColumns was not an array, converting...');
      // Could be an object with numeric keys from JSON serialization
      if (typeof inputColumns === 'object') {
        inputColumns = Object.values(inputColumns);
      } else if (typeof inputColumns === 'string') {
        inputColumns = inputColumns.split(',');
      }
      Logger.log('[buildPlan] Converted inputColumns: ' + JSON.stringify(inputColumns));
    }
  }
  
  if (!inputRange) {
    Logger.log('[buildPlan] ❌ No input range available!');
    throw new Error('No input range available for step ' + (stepIndex + 1));
  }
  
  // Determine output column(s)
  // PRIORITY: Use step's outputColumn if specified by backend
  var outputColumn = step.outputColumn;
  Logger.log('[buildPlan] Step outputColumn from backend: ' + (outputColumn || 'not specified'));
  
  if (!outputColumn) {
    // Find next empty column after last used
    var lastCol = inputColumn || inputRange.charAt(0);
    if (stepIndex > 0 && chainState.steps[stepIndex - 1].outputRange) {
      lastCol = chainState.steps[stepIndex - 1].outputRange.charAt(0);
    }
    outputColumn = String.fromCharCode(lastCol.charCodeAt(0) + 1);
    Logger.log('[buildPlan] Auto-detected outputColumn: ' + outputColumn);
  }
  
  // MULTI-ASPECT DETECTION: Check if outputFormat indicates multiple aspects
  // Format like "Performance | UX | Pricing | Features" means 4 output columns needed
  var outputColumns = [outputColumn];
  
  // PRIORITY 0: NEVER expand if user explicitly requested single column
  if (step.explicitSingleColumn) {
    Logger.log('[buildPlan] User explicitly requested single column - NOT expanding');
    outputColumns = step.outputColumns && step.outputColumns.length > 0 ? step.outputColumns : [outputColumn];
  }
  // PRIORITY 1: Use step.outputColumns if provided by frontend (already expanded)
  else if (step.outputColumns && Array.isArray(step.outputColumns) && step.outputColumns.length > 0) {
    Logger.log('[buildPlan] Using outputColumns from frontend: ' + step.outputColumns.join(', '));
    outputColumns = step.outputColumns;
  } 
  // PRIORITY 2: Detect from outputFormat (backend expansion) - only if no explicit column request
  else {
    var outputFormat = step.outputFormat || '';
    if (outputFormat.indexOf('|') !== -1) {
      var aspects = outputFormat.split('|').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
      if (aspects.length > 1) {
        Logger.log('[buildPlan] Detected multi-aspect output: ' + aspects.length + ' aspects');
        // Expand to multiple columns
        outputColumns = [];
        var startColCode = outputColumn.charCodeAt(0);
        for (var i = 0; i < aspects.length; i++) {
          outputColumns.push(String.fromCharCode(startColCode + i));
        }
        Logger.log('[buildPlan] Expanded outputColumns: ' + outputColumns.join(', '));
      }
    }
  }
  
  // Build prompt - use step's prompt or description
  var prompt = step.prompt || step.description;
  if (!prompt.includes('{{input}}')) {
    prompt = prompt + '\n\nInput: {{input}}\n\nResult:';
  }
  
  // Apply output format instructions if specified
  var outputFormat = step.outputFormat;
  var customHint = step.customFormatHint;
  
  // Build format instruction
  var formatInstruction = null;
  
  // CRITICAL: Skip format instructions if outputFormat is multi-aspect metadata (contains '|')
  // The backend prompt already includes proper instructions for multi-aspect output
  var isMultiAspectFormat = outputFormat && outputFormat.indexOf('|') !== -1;
  
  if (isMultiAspectFormat) {
    Logger.log('[buildPlan] Multi-aspect outputFormat detected, skipping additional format instructions (prompt already contains them)');
  } else if (customHint && customHint.trim()) {
    var trimmedHint = customHint.trim();
    
    // Detect if custom hint looks like classification options (comma-separated short values)
    var looksLikeOptions = trimmedHint.includes(',') && 
      trimmedHint.split(',').every(function(opt) { 
        return opt.trim().length > 0 && opt.trim().length < 25; 
      });
    
    if (looksLikeOptions) {
      // User wants to categorize into specific options
      var options = trimmedHint.split(',').map(function(o) { return o.trim(); }).join(', ');
      formatInstruction = 'OUTPUT FORMAT: Your response MUST be exactly one of these values: ' + options + '\nDo NOT add any other text, explanation, or formatting - just output one of the exact values above.';
      Logger.log('[buildPlan] Applied classification options: ' + options);
    } else if (outputFormat && outputFormat !== 'auto' && outputFormat !== 'simple') {
      // Combine format type with custom requirements
      var baseInstruction = getOutputFormatInstruction(outputFormat);
      formatInstruction = baseInstruction + '\nAdditional requirements: ' + trimmedHint;
      Logger.log('[buildPlan] Applied format with custom requirements: ' + outputFormat + ' + ' + trimmedHint);
    } else {
      // Just custom requirements
      formatInstruction = 'OUTPUT FORMAT: ' + trimmedHint;
      Logger.log('[buildPlan] Applied custom format: ' + trimmedHint);
    }
  } else if (outputFormat && outputFormat !== 'auto' && outputFormat !== 'simple') {
    // No custom hint, but format type selected
    formatInstruction = getOutputFormatInstruction(outputFormat);
    Logger.log('[buildPlan] Applied output format: ' + outputFormat);
  }
  
  if (formatInstruction) {
    prompt = prompt + '\n\n' + formatInstruction;
  }
  
  // Build plan in standard format (same as parseCommandWithAI returns)
  var plan = {
    taskType: step.action || 'process',
    inputRange: inputRange,
    inputColumn: inputColumn,
    inputColumns: inputColumns,
    hasMultipleInputColumns: hasMultipleInputColumns,
    inputColumnHeaders: chainState.inputColumnHeaders || {},
    outputColumns: outputColumns,  // Use the expanded array (single or multi-aspect)
    prompt: prompt,
    summary: step.description,
    model: chainState.model || getAgentModel(),  // Use chain's model (respects user selection)
    rowCount: chainState.rowCount || 10,
    outputFormat: outputFormat,
    // Chain-specific metadata
    chainId: chainState.chainId,
    stepIndex: stepIndex
  };
  
  return plan;
}

/**
 * Get prompt instruction for output format
 * @param {string} format - Output format type
 * @return {string} Instruction to append to prompt
 */
function getOutputFormatInstruction(format) {
  var instructions = {
    'json': 'OUTPUT FORMAT: Return ONLY a valid JSON object. No markdown, no explanation, just the JSON.',
    'list': 'OUTPUT FORMAT: Return a comma-separated list or bullet points. Each item on the same line separated by commas, or use bullet points (- item).',
    'score_reason': 'OUTPUT FORMAT: Return in format "SCORE: [number 1-10] | REASON: [brief explanation]". Keep reason under 50 words.',
    'yes_no': 'OUTPUT FORMAT: Return in format "ANSWER: [Yes/No] | CONFIDENCE: [percentage]%". Example: "ANSWER: Yes | CONFIDENCE: 85%"'
  };
  return instructions[format] || null;
}

/**
 * Generate a smart column header from step description
 * Creates concise, meaningful headers like "Signals & Blockers", "Deal Health", "Next Action"
 * 
 * @param {string} description - Full step description
 * @param {string} action - Action type (extract, analyze, generate, etc.)
 * @return {string} Smart header text
 */
function generateSmartHeader(description, action) {
  if (!description) {
    return action ? capitalizeFirst(action) : 'Output';
  }
  
  // Step 1: Remove common action verbs to get the "what"
  var cleaned = description
    .replace(/^(extract|analyze|assess|generate|create|calculate|identify|produce|get|find|determine|evaluate)\s+/i, '')
    .replace(/\s+(from|for|based on|using|with|in|to)\s+.*/i, '')  // Remove "from notes", "for each", etc.
    .replace(/\s+and\s+/gi, ' & ')  // Replace "and" with "&"
    .replace(/,\s+/g, ', ')  // Normalize commas
    .trim();
  
  // Step 2: Extract key nouns (first 2-3 significant items)
  var items = cleaned.split(/[,&]/).map(function(s) { return s.trim(); }).filter(Boolean);
  
  if (items.length >= 2) {
    // Take first two items and join with "&"
    var header = items.slice(0, 2).map(function(item) {
      // Capitalize and clean each item
      return capitalizeFirst(item.replace(/^(a|an|the|one|each)\s+/i, '').trim());
    }).join(' & ');
    
    // Limit total length
    if (header.length > 25) {
      header = header.substring(0, 22) + '...';
    }
    return header;
  }
  
  if (items.length === 1 && items[0].length > 0) {
    var single = capitalizeFirst(items[0].replace(/^(a|an|the|one|each)\s+/i, ''));
    if (single.length > 25) {
      single = single.substring(0, 22) + '...';
    }
    return single;
  }
  
  // Fallback: First meaningful words from original description
  var words = description.split(/\s+/).filter(function(w) {
    return w.length > 2 && !/^(the|and|for|from|with|into|each|row|column|a|an|one)$/i.test(w);
  });
  
  if (words.length >= 2) {
    var fallback = capitalizeFirst(words.slice(0, 2).join(' '));
    if (fallback.length > 25) {
      fallback = fallback.substring(0, 22) + '...';
    }
    return fallback;
  }
  
  // Last resort: use action type
  return action ? capitalizeFirst(action) : 'Output';
}

/**
 * Capitalize first letter of a string
 */
function capitalizeFirst(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ============================================
// MODULE LOADED
// ============================================

Logger.log('📦 AgentTaskChain_Plan module loaded');
