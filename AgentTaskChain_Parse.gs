/**
 * @file AgentTaskChain_Parse.gs
 * @version 2.0.0
 * @updated 2026-02-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.0.0 (2026-02-10): Refactored from AgentTaskChain.gs (split into 5 modules)
 * - 1.0.0 (2026-01-15): Initial multi-step task chain support
 * 
 * ============================================
 * AGENT TASK CHAIN - Parsing & Suggestions
 * ============================================
 * 
 * Handles command parsing and suggestion fetching:
 * - Parse commands into multi-step workflows via backend AI
 * - Local fallback parsing for when API is unavailable
 * - Action type detection from command text
 * - Post-chain completion suggestions
 * 
 * Part of the AgentTaskChain module suite:
 * - AgentTaskChain_Parse.gs    (this file)
 * - AgentTaskChain_Execute.gs  (execution flow)
 * - AgentTaskChain_Plan.gs     (plan building)
 * - AgentTaskChain_State.gs    (state management)
 * - AgentTaskChain_Analyze.gs  (analyze/chat steps)
 * 
 * Depends on:
 * - Agent.gs (core agent functions)
 * - Config.gs (API endpoints)
 * - SecureRequest.gs (payload building)
 * - ApiClient.gs (HTTP requests)
 */

// ============================================
// TASK CHAIN PARSING
// ============================================

/**
 * Check if a command is multi-step and parse it
 * 
 * @param {string} command - User's natural language command
 * @param {Object} context - Current spreadsheet context
 * @return {Object} { isMultiStep, steps, summary }
 */
function parseTaskChain(command, context, model) {
  Logger.log('[TaskChain] ========== parseTaskChain() CALLED ==========');
  Logger.log('[TaskChain] Command: ' + (command || 'EMPTY').substring(0, 80));
  
  if (!command || !command.trim()) {
    return { isMultiStep: false, steps: [], summary: 'Empty command' };
  }
  
  // Call backend AI which will:
  // 1. Check Formula First (if simple pattern → return formula)
  // 2. If not formula → generate workflow (same AI call)
  // Result: 1 AI call total, no duplicate calls
  Logger.log('[TaskChain] Calling backend AI (handles Formula First + workflow generation)');
  
  // Use backend AI to parse the chain
  // Use model passed from UI, fallback to getAgentModel(), then GEMINI
  try {
    var selectedModel = model || getAgentModel() || 'GEMINI';
    var isManagedChain = selectedModel.indexOf('MANAGED:') === 0;
    var managedChainId = isManagedChain ? selectedModel.replace('MANAGED:', '') : null;
    var provider = isManagedChain ? 'GEMINI' : selectedModel;
    Logger.log('[TaskChain] Using provider from UI: ' + provider + (isManagedChain ? ' (managed: ' + managedChainId + ')' : ''));
    
    // DEBUG: Log context before sending to backend
    Logger.log('[TaskChain] Context selectionInfo: ' + (context && context.selectionInfo ? 'present' : 'MISSING!'));
    
    // Log enhanced context fields (from buildEnhancedContext)
    Logger.log('[TaskChain] Enhanced context fields:');
    Logger.log('[TaskChain]   - dataRange (enhanced): ' + (context && context.dataRange));
    Logger.log('[TaskChain]   - dataStartRow (enhanced): ' + (context && context.dataStartRow));
    Logger.log('[TaskChain]   - dataEndRow (enhanced): ' + (context && context.dataEndRow));
    Logger.log('[TaskChain]   - columnsWithData (enhanced): ' + JSON.stringify(context && context.columnsWithData));
    Logger.log('[TaskChain]   - emptyColumns (enhanced): ' + JSON.stringify(context && context.emptyColumns));
    Logger.log('[TaskChain]   - headerRange (enhanced): ' + (context && context.headerRange));
    
    if (context && context.selectionInfo) {
      Logger.log('[TaskChain] selectionInfo (original):');
      Logger.log('[TaskChain]   - columnsWithData: ' + JSON.stringify(context.selectionInfo.columnsWithData));
      Logger.log('[TaskChain]   - dataRange: ' + context.selectionInfo.dataRange);
      Logger.log('[TaskChain]   - dataStartRow: ' + context.selectionInfo.dataStartRow);
      Logger.log('[TaskChain]   - dataEndRow: ' + context.selectionInfo.dataEndRow);
    }
    Logger.log('[TaskChain] Context columnDataRanges: ' + JSON.stringify(Object.keys((context && context.columnDataRanges) || {})));
    Logger.log('[TaskChain] Context sampleData: ' + JSON.stringify(Object.keys((context && context.sampleData) || {})));
    
    // Build payload — managed mode or BYOK
    var payload;
    if (isManagedChain) {
      payload = SecureRequest.buildManagedPayload({
        command: command,
        context: context || {}
      }, managedChainId);
    } else {
      payload = SecureRequest.buildPayload(provider, {
        command: command,
        context: context || {}
      });
    }
    
    Logger.log('[TaskChain] Calling backend with provider: ' + provider);
    
    // Retry logic for rate limit errors (Gemini "Please retry in Xs")
    var maxRetries = 2;
    var response = null;
    var lastError = null;
    
    for (var attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          Logger.log('[TaskChain] Retry attempt ' + attempt + '/' + maxRetries + '...');
        }
        response = ApiClient.post('AGENT_PARSE_CHAIN', payload);
        lastError = null;
        break; // Success - exit retry loop
      } catch (apiErr) {
        lastError = apiErr;
        var errMsg = apiErr.message || String(apiErr);
        // Check if this is a rate limit error that's worth retrying
        var isRateLimit = errMsg.indexOf('retry') !== -1 || 
                          errMsg.indexOf('429') !== -1 || 
                          errMsg.indexOf('quota') !== -1 ||
                          errMsg.indexOf('rate') !== -1 ||
                          errMsg.indexOf('RESOURCE_EXHAUSTED') !== -1;
        
        if (isRateLimit && attempt < maxRetries) {
          // Extract wait time from error message (e.g., "Please retry in 21.6s")
          var waitMatch = errMsg.match(/retry\s+in\s+([\d.]+)s/i);
          var waitSecs = waitMatch ? Math.min(parseFloat(waitMatch[1]), 30) : 5;
          // Add 2s buffer, cap at 30s
          var waitMs = Math.min(Math.ceil((waitSecs + 2) * 1000), 30000);
          Logger.log('[TaskChain] Rate limit hit. Waiting ' + waitMs + 'ms before retry...');
          Utilities.sleep(waitMs);
        } else {
          Logger.log('[TaskChain] Non-retryable error or max retries reached: ' + errMsg.substring(0, 200));
          break; // Don't retry non-rate-limit errors
        }
      }
    }
    
    if (lastError && !response) {
      throw lastError; // Re-throw to hit the outer catch
    }
    
    Logger.log('[TaskChain] Backend response received');
    Logger.log('[TaskChain] Response inputRange: ' + response.inputRange);
    Logger.log('[TaskChain] Response inputColumn: ' + response.inputColumn);
    Logger.log('[TaskChain] Response inputColumns: ' + JSON.stringify(response.inputColumns));
    Logger.log('[TaskChain] Response hasMultipleInputColumns: ' + response.hasMultipleInputColumns);
    Logger.log('[TaskChain] Response step count: ' + (response.steps ? response.steps.length : 0));
    Logger.log('[TaskChain] Full response: ' + JSON.stringify(response).substring(0, 1000));
    return response;
    
  } catch (e) {
    Logger.log('[TaskChain] ❌ API ERROR: ' + e.message);
    Logger.log('[TaskChain] Error stack: ' + (e.stack || 'N/A'));
    
    // Check if this is a rate limit error - tell the user to wait
    var errorMsg = e.message || String(e);
    var isRateLimitError = errorMsg.indexOf('retry') !== -1 || 
                           errorMsg.indexOf('429') !== -1 || 
                           errorMsg.indexOf('quota') !== -1 ||
                           errorMsg.indexOf('RESOURCE_EXHAUSTED') !== -1;
    
    if (isRateLimitError) {
      // Return a clear rate limit message instead of falling back to local parsing
      return {
        isMultiStep: false,
        steps: [],
        needsClarification: true,
        clarification: '⏳ AI model is temporarily rate-limited. Please wait a few seconds and try again.',
        _apiError: errorMsg,
        _usedFallback: false
      };
    }
    
    // For other errors, fall back to local parsing
    var localResult = parseChainLocally(command);
    localResult._apiError = errorMsg;
    localResult._usedFallback = true;
    return localResult;
  }
}

/**
 * Local fallback for chain parsing when API unavailable
 * 
 * @param {string} command - User command
 * @return {Object} Parsed chain
 */
function parseChainLocally(command) {
  var steps = [];
  var lowerCommand = command.toLowerCase();
  
  // Split by common chain words
  var parts = command.split(/,?\s*(?:then|and then|after that|next|finally)\s*/i);
  
  if (parts.length < 2) {
    return { isMultiStep: false, steps: [], summary: 'Single task' };
  }
  
  // Create steps from parts
  parts.forEach(function(part, index) {
    var trimmed = part.trim();
    if (trimmed.length > 0) {
      steps.push({
        id: 'step_' + (index + 1),
        order: index + 1,
        action: detectAction(trimmed),
        description: trimmed,
        dependsOn: index > 0 ? 'step_' + index : null,
        usesResultOf: index > 0 ? 'step_' + index : null,
        prompt: trimmed + '\n\nInput: {{input}}\n\nOutput:'
      });
    }
  });
  
  return {
    isMultiStep: steps.length > 1,
    steps: steps,
    summary: 'Multi-step task: ' + steps.map(function(s) { return s.action; }).join(' → '),
    estimatedTime: '~' + (steps.length * 2) + ' minutes'
  };
}

/**
 * Detect the action type from a command part
 * 
 * @param {string} text - Command text
 * @return {string} Action type
 */
function detectAction(text) {
  var lower = text.toLowerCase();
  if (lower.indexOf('translat') !== -1) return 'translate';
  if (lower.indexOf('summar') !== -1) return 'summarize';
  if (lower.indexOf('classif') !== -1 || lower.indexOf('categoriz') !== -1) return 'classify';
  if (lower.indexOf('extract') !== -1) return 'extract';
  if (lower.indexOf('clean') !== -1) return 'clean';
  if (lower.indexOf('generat') !== -1) return 'generate';
  if (lower.indexOf('analyz') !== -1) return 'analyze';
  return 'process';
}

// ============================================
// SUGGESTIONS API
// ============================================

/**
 * Fetch suggestions from backend API
 * Called from frontend via google.script.run
 * 
 * @param {Object} context - Chain completion context
 * @param {string} [modelFromUI] - Model selected in the UI dropdown (passed from frontend)
 * @return {Object} { suggestions: [], source: string }
 */
function fetchChainSuggestionsFromBackend(context, modelFromUI) {
  Logger.log('[Suggestions] fetchChainSuggestionsFromBackend called');
  
  try {
    // Use model from UI if provided, otherwise fall back to stored preference
    var selectedSuggModel = modelFromUI || getAgentModel() || 'GEMINI';
    var isManagedSuggChain = selectedSuggModel.indexOf('MANAGED:') === 0;
    var managedSuggChainId = isManagedSuggChain ? selectedSuggModel.replace('MANAGED:', '') : null;
    var provider = isManagedSuggChain ? 'GEMINI' : selectedSuggModel;
    Logger.log('[Suggestions] Using provider: ' + provider + (isManagedSuggChain ? ' (managed)' : ''));
    
    var payload;
    if (isManagedSuggChain) {
      payload = SecureRequest.buildManagedPayload(context, managedSuggChainId);
    } else {
      payload = SecureRequest.buildPayload(provider, context);
    }
    
    var response = ApiClient.post('AGENT_SUGGESTIONS', payload);
    
    Logger.log('[Suggestions] Backend response: ' + JSON.stringify(response).substring(0, 500));
    
    return response;
    
  } catch (e) {
    Logger.log('[Suggestions] API error: ' + e.message);
    
    // Return fallback suggestions
    return {
      suggestions: [
        { icon: '📊', title: 'Summarize findings', command: 'Summarize key insights from the data', reason: 'Get a quick overview' },
        { icon: '🔍', title: 'Analyze patterns', command: 'What patterns exist in this data?', reason: 'Discover hidden insights' },
        { icon: '📈', title: 'Create chart', command: 'Create a chart to visualize the results', reason: 'Visual representation' }
      ],
      source: 'fallback'
    };
  }
}

// ============================================
// MODULE LOADED
// ============================================

Logger.log('📦 AgentTaskChain_Parse module loaded');
