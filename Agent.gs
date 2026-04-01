/**
 * @file Agent.gs
 * @version 2.3.0
 * @updated 2026-01-11
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.3.0 (2026-01-11): Added getJobStreamUrl for SSE real-time job updates
 * - 2.2.0 (2026-01-11): Table selection detection with empty column analysis
 * - 2.1.0 (2026-01-11): Local fallback for custom commands when backend unavailable
 * - 2.0.0 (2026-01-10): Added AI parsing, cost estimation, validation via backend
 * - 1.0.0 (2026-01-10): Initial agent backend with preferences & history
 * 
 * ============================================
 * AGENT.gs - AI Agent Backend
 * ============================================
 * 
 * Handles agent preferences and task history:
 * - Model selection (persisted)
 * - Recent task history (last 10)
 * - Agent execution
 * - AI-powered command parsing (via backend)
 * - Cost estimation (via backend)
 * - Plan validation (via backend)
 * 
 * Storage:
 * - UserProperties: Preferences, recent tasks
 * - Supabase: Full job results (via Jobs.gs)
 */

// ============================================
// CONVERSATION PERSISTENCE
// ============================================

/**
 * Load conversation from backend
 * Pain point solved: "Users often lose context with Gemini"
 * 
 * @return {Object} { hasConversation, isStale, conversation }
 */
function loadConversation() {
  try {
    var userId = getUserId();
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    var url = Config.getApiUrl('AGENT_CONVERSATION') + 
              '?userId=' + encodeURIComponent(userId) + 
              '&spreadsheetId=' + encodeURIComponent(spreadsheetId);
    
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    var result = JSON.parse(response.getContentText());
    
    if (result.error) {
      Logger.log('Failed to load conversation: ' + result.error);
      return { hasConversation: false, conversation: null };
    }
    
    return result;
  } catch (e) {
    Logger.log('Error loading conversation: ' + e.message);
    return { hasConversation: false, conversation: null };
  }
}

/**
 * Save conversation to backend
 * Called after each message exchange
 * 
 * @param {Array} messages - Conversation messages
 * @param {Object} context - Current context snapshot
 * @param {Object} currentTask - Current task state (for multi-step)
 * @return {Object} { success, conversationId }
 */
function saveConversation(messages, context, currentTask) {
  try {
    var userId = getUserId();
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var sheetName = SpreadsheetApp.getActiveSheet().getName();
    
    var response = UrlFetchApp.fetch(Config.getApiUrl('AGENT_CONVERSATION'), {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        userId: userId,
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        messages: messages,
        context: context || {},
        currentTask: currentTask || null
      }),
      muteHttpExceptions: true
    });
    
    var result = JSON.parse(response.getContentText());
    
    if (result.error) {
      Logger.log('Failed to save conversation: ' + result.error);
      return { success: false };
    }
    
    return result;
  } catch (e) {
    Logger.log('Error saving conversation: ' + e.message);
    return { success: false };
  }
}

/**
 * Clear conversation (start fresh)
 * 
 * @return {Object} { success }
 */
function clearConversation() {
  try {
    var userId = getUserId();
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    var url = Config.getApiUrl('AGENT_CONVERSATION') + 
              '?userId=' + encodeURIComponent(userId) + 
              '&spreadsheetId=' + encodeURIComponent(spreadsheetId);
    
    var response = UrlFetchApp.fetch(url, {
      method: 'DELETE',
      muteHttpExceptions: true
    });
    
    var result = JSON.parse(response.getContentText());
    return result;
  } catch (e) {
    Logger.log('Error clearing conversation: ' + e.message);
    return { success: false };
  }
}

// ============================================
// AGENT PREFERENCES
// ============================================

/**
 * Get agent preferences
 * @return {Object} { model, recentTasks }
 */
function getAgentPreferences() {
  var props = PropertiesService.getUserProperties();
  
  return {
    model: props.getProperty('agentModel') || 'GEMINI',
    recentTasks: JSON.parse(props.getProperty('agentRecentTasks') || '[]')
  };
}

/**
 * Save agent model preference
 * @param {string} model - Model ID (CHATGPT, CLAUDE, GROQ, GEMINI)
 */
function setAgentModel(model) {
  // Accept both BYOK providers and managed model IDs (MANAGED:xxx)
  var validBYOK = ['CHATGPT', 'CLAUDE', 'GROQ', 'GEMINI'];
  var isManaged = model.indexOf('MANAGED:') === 0;
  
  if (!isManaged && validBYOK.indexOf(model) === -1) {
    throw new Error('Invalid model: ' + model);
  }
  
  PropertiesService.getUserProperties().setProperty('agentModel', model);
  return { success: true, model: model };
}

/**
 * Get agent model preference
 * @return {string} Current model (e.g., 'GEMINI' or 'MANAGED:gpt-5-mini')
 */
function getAgentModel() {
  return PropertiesService.getUserProperties().getProperty('agentModel') || 'GEMINI';
}

/**
 * Save a generic agent preference
 * @param {string} key - Preference key
 * @param {*} value - Preference value
 */
function setAgentPreference(key, value) {
  PropertiesService.getUserProperties().setProperty('agent_' + key, JSON.stringify(value));
  return { success: true };
}

/**
 * Get a generic agent preference
 * @param {string} key - Preference key
 * @return {*} Preference value or null
 */
function getAgentPreference(key) {
  var val = PropertiesService.getUserProperties().getProperty('agent_' + key);
  if (val === null) return null;
  try { return JSON.parse(val); } catch(e) { return val; }
}

/**
 * Check if current model selection is in managed mode
 * @return {boolean} True if using managed AI
 */
function isModelManaged() {
  var model = getAgentModel();
  return model.indexOf('MANAGED:') === 0;
}

/**
 * Get the managed model ID from the model selection
 * @return {string|null} Managed model ID (e.g., 'gpt-5-mini') or null
 */
function getManagedModelId() {
  var model = getAgentModel();
  if (model.indexOf('MANAGED:') === 0) {
    return model.replace('MANAGED:', '');
  }
  return null;
}

/**
 * Get the user's API key for a specific provider
 * @param {string} provider - Provider name (CHATGPT, CLAUDE, GEMINI, GROQ)
 * @return {string|null} API key or null if not set
 */
function getUserApiKey(provider) {
  try {
    var settings = getUserSettings();
    if (settings && settings[provider] && settings[provider].apiKey) {
      return settings[provider].apiKey;
    }
    return null;
  } catch (e) {
    Logger.log('Error getting API key for ' + provider + ': ' + e.message);
    return null;
  }
}

// ============================================
// SSE STREAM URL
// ============================================

/**
 * Get the SSE stream URL for real-time job updates
 * This URL is used by the frontend to open an EventSource connection
 * 
 * @param {Array<string>} jobIds - Array of job IDs to monitor
 * @return {Object} { url, supportsSSE }
 */
function getJobStreamUrl(jobIds) {
  var userId = getUserId();
  var baseUrl = Config.getBaseUrl();
  var streamPath = API_ENDPOINTS.JOBS_STREAM || '/api/jobs/stream';
  
  // Build SSE URL with job IDs
  var url = baseUrl + streamPath + '?jobIds=' + jobIds.join(',') + '&userId=' + encodeURIComponent(userId);
  
  return {
    url: url,
    supportsSSE: true,
    fallbackPollingMs: 2000  // If SSE fails, poll every 2s
  };
}

// ============================================
// TASK HISTORY
// ============================================

/**
 * Add task to recent history (keeps last 10)
 * @param {Object} task - { command, plan, status, timestamp, jobId? }
 */
function addTaskToHistory(task) {
  var props = PropertiesService.getUserProperties();
  var history = JSON.parse(props.getProperty('agentRecentTasks') || '[]');
  
  // Add new task at beginning
  history.unshift({
    id: Utilities.getUuid(),
    command: task.command,
    taskType: task.taskType || 'custom',
    inputRange: task.inputRange,
    outputColumns: task.outputColumns,
    rowCount: task.rowCount,
    model: task.model,
    status: task.status || 'completed',
    jobId: task.jobId || null,
    timestamp: new Date().toISOString()
  });
  
  // Keep only last 10
  if (history.length > 10) {
    history = history.slice(0, 10);
  }
  
  props.setProperty('agentRecentTasks', JSON.stringify(history));
  return history[0];
}

/**
 * Get recent task history
 * @param {number} limit - Max tasks to return (default 10)
 * @return {Array} Recent tasks
 */
function getTaskHistory(limit) {
  limit = limit || 10;
  var history = JSON.parse(
    PropertiesService.getUserProperties().getProperty('agentRecentTasks') || '[]'
  );
  return history.slice(0, limit);
}

/**
 * Clear task history
 */
function clearTaskHistory() {
  PropertiesService.getUserProperties().deleteProperty('agentRecentTasks');
  return { success: true };
}

// ============================================
// AGENT EXECUTION
// ============================================

/**
 * Execute agent plan
 * Creates jobs for each output column
 * @param {Object} plan - Parsed plan from frontend
 * @return {Object} Execution result with job IDs
 */
function executeAgentPlan(plan) {
  Logger.log('[executeAgentPlan] === STARTING ===');
  Logger.log('[executeAgentPlan] plan is null: ' + (plan === null));
  Logger.log('[executeAgentPlan] plan type: ' + typeof plan);
  
  if (!plan || !plan.inputRange || !plan.outputColumns || !plan.prompt) {
    Logger.log('[executeAgentPlan] ❌ Invalid plan! Missing fields:');
    Logger.log('  plan: ' + (plan ? 'exists' : 'null'));
    Logger.log('  inputRange: ' + (plan ? plan.inputRange : 'N/A'));
    Logger.log('  outputColumns: ' + (plan ? JSON.stringify(plan.outputColumns) : 'N/A'));
    Logger.log('  prompt: ' + (plan && plan.prompt ? plan.prompt.substring(0, 50) + '...' : 'N/A'));
    throw new Error('Invalid plan: missing required fields');
  }
  
  // Validate prompt is not empty
  if (!plan.prompt.trim()) {
    throw new Error('Prompt cannot be empty');
  }
  
  Logger.log('[executeAgentPlan] Plan validated. Key values:');
  Logger.log('  inputRange: ' + plan.inputRange);
  Logger.log('  inputColumns: ' + JSON.stringify(plan.inputColumns));
  Logger.log('  hasMultipleInputColumns: ' + plan.hasMultipleInputColumns);
  
  // Get input data - handle both single and multi-column input
  var inputData;
  var inputColumnHeaders = plan.inputColumnHeaders || {};
  var hasMultipleInputColumns = plan.hasMultipleInputColumns && plan.inputColumns && plan.inputColumns.length > 1;
  
  Logger.log('[executeAgentPlan] Computed hasMultipleInputColumns: ' + hasMultipleInputColumns);
  
  // Log plan parameters for debugging
  Logger.log('📋 Plan parameters:');
  Logger.log('  inputRange: ' + plan.inputRange);
  Logger.log('  inputColumn: ' + plan.inputColumn);
  Logger.log('  inputColumns: ' + JSON.stringify(plan.inputColumns));
  Logger.log('  hasMultipleInputColumns: ' + plan.hasMultipleInputColumns);
  Logger.log('  inputColumnHeaders: ' + JSON.stringify(inputColumnHeaders));
  
  if (hasMultipleInputColumns) {
    // MULTI-COLUMN INPUT: Get data from all input columns and structure it
    Logger.log('📊 Using MULTI-column input: ' + plan.inputColumns.join(', '));
    try {
      inputData = getMultiColumnInputData(plan.inputRange, plan.inputColumns, inputColumnHeaders);
      Logger.log('📊 Multi-column data retrieved: ' + (inputData ? inputData.length : 0) + ' rows');
    } catch (multiColError) {
      Logger.log('❌ getMultiColumnInputData failed: ' + multiColError.message);
      Logger.log('Falling back to single-column extraction');
      inputData = getRangeValues(plan.inputRange);
    }
  } else {
    // Single column input (original behavior)
    Logger.log('📊 Using SINGLE-column input: ' + plan.inputRange);
    inputData = getRangeValues(plan.inputRange);
  }
  
  Logger.log('📊 inputData length after extraction: ' + (inputData ? inputData.length : 'null'));
  
  if (!inputData || inputData.length === 0) {
    throw new Error('No data in input range');
  }
  
  // Filter empty values
  inputData = inputData.filter(function(v) {
    return v !== null && v !== undefined && String(v).trim() !== '';
  });
  
  if (inputData.length === 0) {
    throw new Error('All cells in input range are empty');
  }
  
  Logger.log('📊 Input data prepared: ' + inputData.length + ' rows' + (hasMultipleInputColumns ? ' (multi-column)' : ''));
  
  // Log first input for debugging
  if (inputData.length > 0) {
    var firstInput = String(inputData[0]).substring(0, 200);
    Logger.log('📊 First input sample:\n' + firstInput + (String(inputData[0]).length > 200 ? '...' : ''));
  }
  
  // Use agent's preferred model or plan's model
  var model = plan.model || getAgentModel();
  
  // Create jobs for each output column
  // Add small delay between jobs to avoid rate limiting
  var jobs = [];
  var columnPrompts = plan.columnPrompts || {};
  
  // Create all jobs (removed delay - backend handles rate limiting)
  Logger.log('🗺️ Creating jobs for columns: ' + plan.outputColumns.join(', '));
  Logger.log('Column prompts available: ' + Object.keys(columnPrompts).join(', '));
  
  // CRITICAL: Detect multi-aspect task (one job, multiple output columns)
  var isMultiAspect = plan.outputColumns.length > 1 && plan.outputFormat && plan.outputFormat.indexOf('|') !== -1;
  
  if (isMultiAspect) {
    // Multi-aspect: Create ONE job that returns combined output (e.g., "Positive | Negative | Neutral | Positive")
    Logger.log('[Multi-Aspect] Creating ONE job for ' + plan.outputColumns.length + ' aspects: ' + plan.outputColumns.join(', '));
    Logger.log('[Multi-Aspect] Output format: ' + plan.outputFormat);
    
    try {
      var job = createBulkJob(inputData, plan.prompt, model);
      Logger.log('[Multi-Aspect] Job created: ' + job.id);
      
      // Store job info with multi-aspect flag
      jobs.push({
        columns: plan.outputColumns,  // Array of target columns
        jobId: job.id,
        totalRows: job.totalRows,
        status: 'queued',
        isMultiAspect: true,
        outputFormat: plan.outputFormat  // Need this to split results
      });
    } catch (e) {
      Logger.log('[Multi-Aspect] Job creation failed: ' + e.message);
      jobs.push({
        columns: plan.outputColumns,
        jobId: null,
        error: e.message,
        status: 'failed',
        isMultiAspect: true
      });
    }
  } else {
    // Standard: Create separate job for each output column
    for (var idx = 0; idx < plan.outputColumns.length; idx++) {
      var col = plan.outputColumns[idx];
      
      // Get column-specific prompt if available (for multi-language translation, etc.)
      var colPrompt = columnPrompts[col] || plan.prompt;
      
      try {
        // Log the header/task type from the prompt
        var headerMatch = colPrompt.match(/Generate "([^"]+)"/);
        var taskHeader = headerMatch ? headerMatch[1] : 'unknown';
        Logger.log('Creating job for column ' + col + ' (task: "' + taskHeader + '") with prompt: ' + colPrompt.substring(0, 80) + '...');
        
        var job = createBulkJob(inputData, colPrompt, model);
        
        Logger.log('Job created for column ' + col + ': ' + job.id);
        
        jobs.push({
          column: col,
          jobId: job.id,
          totalRows: job.totalRows,
          status: 'queued'
        });
      } catch (e) {
        Logger.log('Failed to create job for column ' + col + ': ' + e.message);
        jobs.push({
          column: col,
          jobId: null,
          error: e.message,
          status: 'failed'
        });
      }
    }
  }
  
  // Save to task history
  addTaskToHistory({
    command: plan.originalCommand || plan.summary,
    taskType: plan.taskType,
    inputRange: plan.inputRange,
    outputColumns: plan.outputColumns,
    rowCount: inputData.length,
    model: model,
    status: jobs.every(function(j) { return j.status !== 'failed'; }) ? 'running' : 'partial',
    jobId: jobs.length === 1 ? jobs[0].jobId : null
  });
  
  // Trigger worker via lightweight endpoint (non-blocking)
  triggerWorkerAsync(jobs.map(function(j) { return j.jobId; }).filter(Boolean));
  
  return {
    success: true,
    jobs: jobs,
    totalRows: inputData.length,
    model: model
  };
}

/**
 * Trigger the job worker via lightweight async endpoint
 * This returns immediately - actual processing happens in background
 */
function triggerWorkerAsync(jobIds) {
  try {
    // Use the trigger endpoint which returns immediately
    var triggerUrl = Config.getApiUrl('JOBS_TRIGGER');
    
    UrlFetchApp.fetch(triggerUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ jobIds: jobIds }),
      muteHttpExceptions: true,
      headers: {
        'Cache-Control': 'no-cache'
      }
    });
    
    Logger.log('Worker trigger sent for ' + jobIds.length + ' jobs');
  } catch (e) {
    // Silently ignore - cron will pick up jobs eventually
    Logger.log('Worker trigger failed (non-critical): ' + e.message);
  }
}

// ============================================
// MULTI-LANGUAGE TRANSLATION HELPER
// ============================================

/**
 * Create prompts for multi-language translation
 * @param {Array<string>} languages - Target languages
 * @param {Array<string>} outputColumns - Output column letters
 * @return {Object} Column-to-prompt mapping
 */
function createTranslationPrompts(languages, outputColumns) {
  var prompts = {};
  
  languages.forEach(function(lang, idx) {
    var col = outputColumns[idx];
    if (col) {
      prompts[col] = 'Translate to ' + lang + '. Keep brand names and proper nouns unchanged. Return only the translation:\n\n{{input}}';
    }
  });
  
  return prompts;
}

// ============================================
// AVAILABLE MODELS
// ============================================

/**
 * Get available models for agent based on user's configured API keys
 * Only shows providers where user has entered an API key
 * @return {Array} Models with pricing info
 */
function getAgentModels() {
  // Model metadata with defaults - these are fallback display names
  // The actual model names come from the backend fetchModels() call
  var modelMeta = {
    GEMINI: {
      icon: '🌐',
      fallbackName: 'Gemini 2.5 Flash',
      description: 'Cheapest & great quality',
      price: '$0.08/MTok',
      order: 1
    },
    GROQ: {
      icon: '⚡',
      fallbackName: 'Llama 3.3 70B',
      description: 'Fastest responses',
      price: '$0.59/MTok',
      order: 2
    },
    CHATGPT: {
      icon: '🧠',
      fallbackName: 'GPT-5 Mini',
      description: 'Balanced performance',
      price: '$0.25/MTok',
      order: 3
    },
    CLAUDE: {
      icon: '🤖',
      fallbackName: 'Claude Haiku 4.5',
      description: 'Best for complex tasks',
      price: '$1.00/MTok',
      order: 4
    }
  };
  
  try {
    // Get user settings to check which API keys are configured
    var settings = getUserSettings();
    
    // Get all available models from backend for proper naming
    var allModels = [];
    try {
      allModels = fetchModels() || [];
    } catch (e) {
      // If fetchModels fails, we'll use default names
      console.log('Could not fetch model names, using defaults');
    }
    
    // Build available models list based on configured API keys
    var availableModels = [];
    var providers = Config.getSupportedProviders();
    
    providers.forEach(function(provider) {
      var providerSettings = settings[provider];
      
      // Only include if user has configured an API key
      if (providerSettings && providerSettings.apiKey && providerSettings.apiKey.trim() !== '') {
        var meta = modelMeta[provider] || { icon: '⚙️', fallbackName: provider, description: '', price: 'N/A', order: 99 };
        
        // Try to get the model display name (priority order):
        // 1. User's selected model from settings (if they chose one)
        // 2. First available model from backend for this provider
        // 3. Fallback static name
        var modelName = meta.fallbackName;
        
        if (providerSettings.defaultModel && providerSettings.defaultModel.trim() !== '') {
          // User selected a specific model - use its display name
          var selectedModel = allModels.find(function(m) { 
            return m.name === providerSettings.defaultModel && m.llm === provider; 
          });
          if (selectedModel && selectedModel.display_name) {
            modelName = selectedModel.display_name;
          } else {
            // Model not found in current list (deprecated) — use fallback name
            // instead of showing raw model ID like "gpt-3.5-turbo"
            modelName = meta.fallbackName;
          }
        } else if (allModels.length > 0) {
          // No user selection - try to get first available model for this provider
          var firstProviderModel = allModels.find(function(m) { return m.llm === provider; });
          if (firstProviderModel && firstProviderModel.display_name) {
            modelName = firstProviderModel.display_name;
          }
        }
        
        availableModels.push({
          id: provider,
          name: modelName,
          icon: meta.icon,
          description: meta.description,
          price: meta.price,
          recommended: provider === 'GEMINI',
          order: meta.order
        });
      }
    });
    
    // Sort by order
    availableModels.sort(function(a, b) { return a.order - b.order; });
    
    // If no models configured, return a message model
    if (availableModels.length === 0) {
      return [{
        id: 'NONE',
        name: 'No API keys configured',
        icon: '⚠️',
        description: 'Go to Settings to add API keys',
        price: '',
        recommended: false
      }];
    }
    
    return availableModels;
    
  } catch (e) {
    console.error('Error getting agent models:', e);
    // Fallback to show all models if settings fail
    return [
      { id: 'GEMINI', name: 'Gemini', icon: '🌐', description: 'Cheapest & great quality', price: '$0.08/MTok', recommended: true },
      { id: 'GROQ', name: 'Groq', icon: '⚡', description: 'Fastest responses', price: '$0.59/MTok', recommended: false },
      { id: 'CHATGPT', name: 'ChatGPT', icon: '🧠', description: 'Balanced performance', price: '$0.25/MTok', recommended: false },
      { id: 'CLAUDE', name: 'Claude', icon: '🤖', description: 'Best for complex tasks', price: '$1.00/MTok', recommended: false }
    ];
  }
}

/**
 * Get available managed AI models for the user's tier.
 * These are models that use platform API keys (no user key needed).
 * 
 * @return {Array} Managed models with metadata for the model selector
 */
function getManagedModels() {
  // Managed model metadata (matches backend MANAGED_MODEL_REGISTRY)
  var managedModelMeta = {
    'gpt-5-mini':               { icon: '🧠', name: 'GPT-5 Mini',          tier: 'mini', price: 'Free',  order: 1 },
    'gemini-2.5-flash':         { icon: '🌐', name: 'Gemini 2.5 Flash',    tier: 'mini', price: 'Free',  order: 2 },
    'claude-haiku-4-5':         { icon: '🤖', name: 'Claude Haiku 4.5',    tier: 'mini', price: 'Free',  order: 3 },
    'meta-llama/llama-4-scout-17b-16e-instruct':  { icon: '⚡', name: 'Llama 4 Scout',  tier: 'mini', price: 'Free',  order: 4 },
    'gpt-5':                    { icon: '🧠', name: 'GPT-5',               tier: 'mid',  price: 'Pro',   order: 5 },
    'claude-sonnet-4-5':        { icon: '🤖', name: 'Claude Sonnet 4.5',   tier: 'mid',  price: 'Pro',   order: 6 },
    'gemini-2.5-pro':           { icon: '🌐', name: 'Gemini 2.5 Pro',      tier: 'mid',  price: 'Pro',   order: 7 },
    'meta-llama/llama-4-maverick-17b-128e-instruct': { icon: '⚡', name: 'Llama 4 Maverick', tier: 'mid', price: 'Pro', order: 8 }
  };
  
  try {
    // Get user subscription to determine tier
    var sub = getUserSubscription();
    var tier = (sub && sub.tier) || 'free';
    
    // Build model list based on tier
    var miniModels = ['gpt-5-mini', 'gemini-2.5-flash', 'claude-haiku-4-5', 'meta-llama/llama-4-scout-17b-16e-instruct'];
    var midModels = ['gpt-5', 'claude-sonnet-4-5', 'gemini-2.5-pro', 'meta-llama/llama-4-maverick-17b-128e-instruct'];
    
    var allowedModels = miniModels.slice(); // Copy mini models (always available)
    if (tier === 'pro' || tier === 'legacy') {
      allowedModels = allowedModels.concat(midModels);
    }
    
    return allowedModels.map(function(modelId) {
      var meta = managedModelMeta[modelId] || { icon: '⚙️', name: modelId, tier: 'mini', price: 'Free', order: 99 };
      return {
        id: 'MANAGED:' + modelId,   // Prefix to distinguish from BYOK models
        modelId: modelId,
        name: meta.name,
        icon: meta.icon,
        description: meta.tier === 'mini' ? 'Free AI' : 'Pro credits',
        price: meta.price,
        isManaged: true,
        recommended: modelId === 'gemini-2.5-flash',
        order: meta.order
      };
    });
    
  } catch (e) {
    console.error('Error getting managed models:', e);
    // Fallback: show mini models only
    return [
      { id: 'MANAGED:gpt-5-mini', modelId: 'gpt-5-mini', name: 'GPT-5 Mini', icon: '🧠', price: 'Free', isManaged: true, recommended: false, order: 1 },
      { id: 'MANAGED:gemini-2.5-flash', modelId: 'gemini-2.5-flash', name: 'Gemini 2.5 Flash', icon: '🌐', price: 'Free', isManaged: true, recommended: true, order: 2 }
    ];
  }
}

// ============================================
// AI-POWERED COMMAND PARSING (Backend)
// ============================================

/**
 * Parse a natural language command using AI
 * Falls back to backend when frontend regex fails
 * Enriched with context for smarter parsing
 * @param {string} command - User's natural language command
 * @param {Object} context - Optional context (selectedRange, headers, columnDataRanges)
 * @return {Object} Parsed plan or clarification request
 */
function parseCommandWithAI(command, context) {
  if (!command || !command.trim()) {
    throw new Error('Command cannot be empty');
  }
  
  // Enrich context if not provided
  if (!context || Object.keys(context).length === 0) {
    context = getAgentContext();
  }
  
  // Build enhanced context description for AI
  var contextDescription = buildContextDescription(context);
  
  // Get user's selected model — may be BYOK (e.g., 'GEMINI') or managed (e.g., 'MANAGED:gpt-5-mini')
  var selectedModel = context.provider || getAgentModel() || 'GEMINI';
  var isManaged = selectedModel.indexOf('MANAGED:') === 0;
  var managedModelId = isManaged ? selectedModel.replace('MANAGED:', '') : null;
  var provider = isManaged ? 'GEMINI' : selectedModel; // Provider is irrelevant for managed, but needed for payload shape
  
  // Build payload — managed mode skips API key encryption, BYOK uses user's encrypted key
  var payload;
  if (isManaged) {
    payload = SecureRequest.buildManagedPayload({
      command: command,
      context: context,
      contextDescription: contextDescription
    }, managedModelId);
  } else {
    payload = SecureRequest.buildPayload(provider, {
      command: command,
      context: context,
      contextDescription: contextDescription
    });
  }
  
  try {
    var response = ApiClient.post('AGENT_PARSE', payload);
    
    return response;
  } catch (e) {
    Logger.log('AI parsing failed: ' + e.message);
    
    // For custom commands, try to create a simple plan locally instead of failing
    var columnDataRanges = context.columnDataRanges || {};
    var firstDataCol = null;
    var firstRange = null;
    
    for (var col in columnDataRanges) {
      if (columnDataRanges[col].hasData) {
        // Check if this column is mentioned in the command
        if (command.toLowerCase().includes('column ' + col.toLowerCase()) ||
            command.toLowerCase().includes(' ' + col.toLowerCase() + ' ') ||
            command.toLowerCase().includes(' ' + col.toLowerCase() + ',')) {
          firstDataCol = col;
          firstRange = columnDataRanges[col].range;
          break;
        }
        // Fallback to first column with data
        if (!firstDataCol) {
          firstDataCol = col;
          firstRange = columnDataRanges[col].range;
        }
      }
    }
    
    // If we have data and the command looks actionable, create a local plan
    if (firstDataCol && firstRange && command.trim().length > 10) {
      var outputCol = String.fromCharCode(firstDataCol.charCodeAt(0) + 1);
      var rowCount = columnDataRanges[firstDataCol].rowCount || 10;
      
      return {
        success: true,
        plan: {
          taskType: 'custom',
          inputRange: firstRange,
          inputColumn: firstDataCol,
          outputColumns: [outputCol],
          prompt: command + '\n\nData: {{input}}\n\nResult:',
          summary: 'I\'ll process ' + firstRange + ' and write results to column ' + outputCol + '.',
          confidence: 'medium',
          rowCount: rowCount
        }
      };
    }
    
    // Generate smart suggestions as last resort
    var suggestions = generateSmartSuggestions(command, context);
    
    return {
      success: false,
      error: 'AI parsing unavailable: ' + e.message,
      clarification: {
        question: "I couldn't parse this command. Try being more specific:",
        suggestions: suggestions
      }
    };
  }
}

/**
 * Build a human-readable context description for AI parsing
 * Helps AI understand the spreadsheet structure
 * @param {Object} context - The context object
 * @return {string} Human-readable description
 */
function buildContextDescription(context) {
  var lines = [];
  
  // Sheet info
  if (context.sheetName) {
    lines.push('Sheet: ' + context.sheetName);
  }
  
  // Column headers and sample data
  if (context.headers && context.headers.length > 0) {
    lines.push('');
    lines.push('Columns with data:');
    
    context.headers.forEach(function(h) {
      var col = h.column;
      var header = h.name;
      var dataInfo = context.columnDataRanges ? context.columnDataRanges[col] : null;
      var samples = context.sampleData ? context.sampleData[col] : [];
      
      var line = '  Column ' + col + ' (header: "' + header + '")';
      
      if (dataInfo && dataInfo.hasData) {
        line += ' - ' + dataInfo.rowCount + ' rows of data';
        if (samples && samples.length > 0) {
          line += ', samples: [' + samples.slice(0, 2).map(function(s) { 
            return '"' + s.substring(0, 50) + (s.length > 50 ? '...' : '') + '"'; 
          }).join(', ') + ']';
        }
      }
      
      lines.push(line);
    });
  }
  
  // Selected range
  if (context.selectedRange) {
    lines.push('');
    lines.push('User has selected: ' + context.selectedRange);
  }
  
  // Hint about headers potentially being prompts or languages
  lines.push('');
  lines.push('NOTE: Headers might be:');
  lines.push('  - Language names (for translation tasks)');
  lines.push('  - Custom prompts/instructions (for per-column processing)');
  lines.push('  - Just column labels');
  
  return lines.join('\n');
}

/**
 * Generate smart suggestions based on context
 * Uses available column data to create relevant examples
 * @param {string} command - Original command
 * @param {Object} context - Sheet context
 * @return {Array<string>} Suggestion strings
 */
function generateSmartSuggestions(command, context) {
  var suggestions = [];
  var lower = command.toLowerCase();
  var columnDataRanges = context.columnDataRanges || {};
  var headers = context.headerRow || {};
  
  // Find first column with data
  var firstDataCol = 'A';
  var firstDataRange = 'A2:A50';
  for (var col in columnDataRanges) {
    if (columnDataRanges[col].hasData) {
      firstDataCol = col;
      firstDataRange = columnDataRanges[col].range;
      break;
    }
  }
  
  // Get next column for output
  var outputCol = String.fromCharCode(firstDataCol.charCodeAt(0) + 1);
  
  // Generate task-specific suggestions
  if (lower.includes('translate')) {
    suggestions.push('Translate column ' + firstDataCol + ' to Spanish in ' + outputCol);
    suggestions.push('Translate column ' + firstDataCol + ' to German, French in ' + outputCol + ', ' + String.fromCharCode(outputCol.charCodeAt(0) + 1));
    suggestions.push('Translate ' + firstDataRange + ' to Korean in column ' + outputCol);
  } else if (lower.includes('summarize') || lower.includes('summary')) {
    suggestions.push('Summarize column ' + firstDataCol + ' to ' + outputCol);
    suggestions.push('Summarize ' + firstDataRange + ' in column ' + outputCol);
  } else if (lower.includes('extract')) {
    suggestions.push('Extract emails from column ' + firstDataCol + ' to ' + outputCol);
    suggestions.push('Extract names from ' + firstDataRange + ' to column ' + outputCol);
  } else if (lower.includes('classify') || lower.includes('categorize')) {
    suggestions.push('Classify column ' + firstDataCol + ' as Positive/Negative in ' + outputCol);
    suggestions.push('Categorize ' + firstDataRange + ' into categories in column ' + outputCol);
  } else {
    // Generic suggestions
    suggestions.push('Translate column ' + firstDataCol + ' to Spanish in ' + outputCol);
    suggestions.push('Summarize column ' + firstDataCol + ' to ' + outputCol);
    suggestions.push('Extract emails from column ' + firstDataCol + ' to ' + outputCol);
  }
  
  return suggestions;
}

// ============================================
// COST ESTIMATION (Backend)
// ============================================

/**
 * Fetch AI-driven classification options for column headers
 * Uses AI to analyze headers and suggest appropriate options (e.g., Yes/No, Good/Bad/Neutral)
 * 
 * @param {Array<Object>} columnHeaders - Array of { column, header } objects
 * @return {Object} { success, columnOptions, language }
 */
function fetchClassificationOptions(columnHeaders) {
  if (!columnHeaders || !Array.isArray(columnHeaders) || columnHeaders.length === 0) {
    Logger.log('No column headers provided for classification options');
    return { success: false, columnOptions: {} };
  }
  
  // Get user's selected model — may be BYOK or managed
  var selectedModel = getAgentModel() || 'GEMINI';
  var isManagedModel = selectedModel.indexOf('MANAGED:') === 0;
  var managedId = isManagedModel ? selectedModel.replace('MANAGED:', '') : null;
  var provider = isManagedModel ? 'GEMINI' : selectedModel;
  
  // Build payload — managed or BYOK
  var payload;
  if (isManagedModel) {
    payload = SecureRequest.buildManagedPayload({ columns: columnHeaders }, managedId);
  } else {
    payload = SecureRequest.buildPayload(provider, { columns: columnHeaders });
  }
  
  try {
    Logger.log('[fetchClassificationOptions] Analyzing ' + columnHeaders.length + ' columns');
    
    var response = ApiClient.post('CLASSIFY_OPTIONS', payload);
    
    if (response && response.success) {
      Logger.log('[fetchClassificationOptions] Success: ' + JSON.stringify(response.columnOptions));
      return response;
    }
    
    Logger.log('[fetchClassificationOptions] Failed: ' + JSON.stringify(response));
    return { success: false, columnOptions: {} };
    
  } catch (e) {
    Logger.log('[fetchClassificationOptions] Error: ' + e.message);
    return { success: false, columnOptions: {}, error: e.message };
  }
}

/**
 * Estimate cost for an operation before execution
 * @param {number} rowCount - Number of rows to process
 * @param {string} taskType - Type of task (translate, summarize, etc.)
 * @param {string} model - Model to use
 * @param {string} sampleText - Optional sample text for better estimation
 * @return {Object} Cost estimation with formatted values
 */
function estimateAgentCost(rowCount, taskType, model, sampleText) {
  if (!rowCount || rowCount < 1) {
    throw new Error('Row count must be positive');
  }
  
  try {
    var response = ApiClient.post('AGENT_ESTIMATE', {
      rowCount: rowCount,
      taskType: taskType || 'custom',
      model: model || getAgentModel(),
      sampleText: sampleText || null
    });
    
    return response;
  } catch (e) {
    Logger.log('Cost estimation failed: ' + e.message);
    // Return a rough estimate as fallback
    return {
      rowCount: rowCount,
      model: model || 'GEMINI',
      cost: {
        total: rowCount * 0.0001, // ~$0.0001 per row rough estimate
        formatted: '<$0.01'
      },
      time: {
        formatted: '~' + Math.ceil(rowCount * 0.5) + 's'
      },
      error: 'Estimation unavailable'
    };
  }
}

// ============================================
// PLAN VALIDATION (Backend)
// ============================================

/**
 * Validate a plan before execution
 * Checks for potential issues and warns user
 * @param {Object} plan - The plan to validate
 * @param {Object} context - Context info (rowCount, outputHasData)
 * @return {Object} Validation result with warnings and errors
 */
function validateAgentPlan(plan, context) {
  if (!plan) {
    return {
      valid: false,
      errors: ['No plan provided'],
      warnings: [],
      suggestions: []
    };
  }
  
  try {
    var response = ApiClient.post('AGENT_VALIDATE', {
      plan: plan,
      context: context || {}
    });
    
    return response;
  } catch (e) {
    Logger.log('Validation failed: ' + e.message);
    // Do basic client-side validation as fallback
    var result = {
      valid: true,
      errors: [],
      warnings: [],
      suggestions: []
    };
    
    if (!plan.prompt || !plan.prompt.trim()) {
      result.valid = false;
      result.errors.push('Prompt cannot be empty');
    }
    
    if (!plan.inputRange) {
      result.valid = false;
      result.errors.push('Input range is required');
    }
    
    if (!plan.outputColumns || plan.outputColumns.length === 0) {
      result.valid = false;
      result.errors.push('Output column is required');
    }
    
    return result;
  }
}

// ============================================
// CONTEXT HELPERS
// ============================================

/**
 * Get current context for agent commands
 * Includes selected range, headers, data ranges, and sample data for AI understanding
 * @return {Object} Context information
 */
function getAgentContext() {
  try {
    // Force a fresh reference to the spreadsheet and active sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var selection = sheet.getActiveRange();
    
    // Debug logging to verify correct sheet
    Logger.log('[getAgentContext] Sheet: ' + sheet.getName());
    Logger.log('[getAgentContext] Selection: ' + (selection ? selection.getA1Notation() : 'none'));
    
    // Get headers from row 1
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = [];
    var headerRow = {};
    
    if (lastCol > 0) {
      var headerRange = sheet.getRange(1, 1, 1, lastCol);
      var headerValues = headerRange.getValues()[0];
      
      for (var i = 0; i < headerValues.length; i++) {
        var h = headerValues[i];
        if (h && String(h).trim()) {
          var colLetter = String.fromCharCode(65 + i); // A, B, C...
          headers.push({
            column: colLetter,
            name: String(h).trim()
          });
          headerRow[colLetter] = String(h).trim();
        }
      }
    }
    
    // Detect data ranges for each column (where content exists)
    var columnDataRanges = {};
    var sampleData = {}; // Sample data for AI context
    
    if (lastCol > 0 && lastRow > 1) {
      for (var col = 1; col <= Math.min(lastCol, 26); col++) { // Limit to A-Z
        var colLetter = String.fromCharCode(64 + col);
        var dataInfo = detectColumnDataRange(sheet, col, lastRow);
        if (dataInfo.hasData) {
          columnDataRanges[colLetter] = dataInfo;
          
          // Get sample data (first 3 non-empty values) for AI context
          sampleData[colLetter] = getSampleData(sheet, col, dataInfo.startRow, 3);
        }
      }
    }
    
    // Analyze selection for table detection
    var selectionInfo = null;
    var selectionSource = 'none'; // 'explicit' | 'auto' | 'none'
    
    if (selection) {
      var selRange = selection.getA1Notation();
      var selStartCol = selection.getColumn();
      var selEndCol = selection.getLastColumn();
      var selStartRow = selection.getRow();
      var selEndRow = selection.getLastRow();
      var numCols = selEndCol - selStartCol + 1;
      var numRows = selEndRow - selStartRow + 1;
      
      // Determine if this is an EXPLICIT selection (user intentionally selected a range)
      // vs a single cell click (should trigger auto-detect)
      var isExplicitSelection = (numCols > 1 || numRows > 1);
      
      // Check if this looks like a table (multiple columns or multiple rows)
      if (isExplicitSelection) {
        selectionSource = 'explicit';
        var selectedHeaders = [];
        var columnsWithData = [];
        var emptyColumns = [];
        
        // SMART HEADER DETECTION: Check if first row of selection is headers or data
        // Check if the row ABOVE the selection has headers (if selection doesn't start at row 1)
        var selFirstRowValues = sheet.getRange(selStartRow, selStartCol, 1, numCols).getValues()[0];
        var rowAboveHasHeaders = false;
        var selectionFirstRowIsHeader = true; // Default assumption
        
        // Check if first row of selection is empty - if so, scan down to find actual headers
        var firstRowIsEmpty = selFirstRowValues.every(function(v) { 
          return String(v || '').trim().length === 0; 
        });
        
        // If first row is empty, find the first non-empty row within selection
        var actualFirstNonEmptyRow = selStartRow;
        var actualFirstRowValues = selFirstRowValues;
        
        if (firstRowIsEmpty) {
          Logger.log('Header detection: First row of selection is empty, scanning for headers...');
          for (var scanRow = selStartRow; scanRow <= selEndRow; scanRow++) {
            var rowValues = sheet.getRange(scanRow, selStartCol, 1, numCols).getValues()[0];
            var hasContent = rowValues.some(function(v) { return String(v || '').trim().length > 0; });
            if (hasContent) {
              actualFirstNonEmptyRow = scanRow;
              actualFirstRowValues = rowValues;
              Logger.log('Header detection: Found first non-empty row at ' + scanRow);
              break;
            }
          }
        }
        
        if (selStartRow > 1 && !firstRowIsEmpty) {
          var rowAbove = sheet.getRange(selStartRow - 1, selStartCol, 1, numCols).getValues()[0];
          var aboveHasText = rowAbove.some(function(v) { 
            var s = String(v || '').trim();
            return s.length > 0 && !s.startsWith('http') && !/^\d+(\.\d+)?$/.test(s);
          });
          if (aboveHasText) {
            rowAboveHasHeaders = true;
            // If row above has headers, first row of selection is likely DATA
            selectionFirstRowIsHeader = false;
            Logger.log('Header detection: Row ' + (selStartRow - 1) + ' above selection has headers');
          }
        }
        
        // Also check if first row of selection looks like DATA (URLs, numbers, etc.)
        if (selectionFirstRowIsHeader && !firstRowIsEmpty) {
          var firstRowLooksLikeData = actualFirstRowValues.some(function(v) {
            var s = String(v || '').trim();
            // Check for URL patterns, pure numbers, dates, etc.
            return s.startsWith('http') || s.startsWith('www.') || 
                   /^\d{4}-\d{2}-\d{2}/.test(s) || // ISO date
                   /^\d+(\.\d+)?$/.test(s); // Pure number
          });
          if (firstRowLooksLikeData) {
            selectionFirstRowIsHeader = false;
            Logger.log('Header detection: First row of selection contains data-like values');
          }
        }
        
        // Determine actual header row and data start row
        // If first rows were empty, use the first non-empty row as the header
        var actualHeaderRow;
        var actualDataStartRow;
        
        if (firstRowIsEmpty) {
          // First non-empty row is likely the header
          actualHeaderRow = actualFirstNonEmptyRow;
          actualDataStartRow = actualFirstNonEmptyRow + 1;
          Logger.log('Header detection: Using row ' + actualHeaderRow + ' as header (skipped empty rows)');
        } else if (selectionFirstRowIsHeader) {
          actualHeaderRow = selStartRow;
          actualDataStartRow = selStartRow + 1;
        } else {
          actualHeaderRow = selStartRow > 1 ? selStartRow - 1 : null;
          actualDataStartRow = selStartRow;
        }
        
        // Get headers from the correct row
        var selHeaders;
        if (firstRowIsEmpty && actualHeaderRow === actualFirstNonEmptyRow) {
          selHeaders = actualFirstRowValues;
        } else if (selectionFirstRowIsHeader) {
          selHeaders = selFirstRowValues;
        } else if (actualHeaderRow) {
          selHeaders = sheet.getRange(actualHeaderRow, selStartCol, 1, numCols).getValues()[0];
        } else {
          selHeaders = []; // No headers found
        }
        
        Logger.log('Header detection result: selectionFirstRowIsHeader=' + selectionFirstRowIsHeader + 
                   ', headerRow=' + actualHeaderRow + ', dataStartRow=' + actualDataStartRow);
        
        // PROMPT ROW DETECTION: Check for instruction rows above the header row
        // These are rows with long text (>40 chars) that describe what to do for each column
        var columnPrompts = {};
        var promptRowNum = null;
        
        if (actualHeaderRow && actualHeaderRow > 1) {
          // Search up to 5 rows above the header for potential prompt rows
          var searchStartRow = Math.max(1, actualHeaderRow - 5);
          
          for (var searchRow = actualHeaderRow - 1; searchRow >= searchStartRow; searchRow--) {
            var potentialPromptRow = sheet.getRange(searchRow, selStartCol, 1, numCols).getValues()[0];
            var hasLongText = false;
            var promptCount = 0;
            
            // Check if this row has instruction-like content (long text)
            for (var pi = 0; pi < potentialPromptRow.length; pi++) {
              var cellText = String(potentialPromptRow[pi] || '').trim();
              if (cellText.length > 40) { // Long text = likely an instruction/prompt
                hasLongText = true;
                promptCount++;
              }
            }
            
            // If multiple cells have long text, this is likely a prompt row
            if (hasLongText && promptCount >= 1) {
              promptRowNum = searchRow;
              Logger.log('Prompt row detected at row ' + searchRow + ' with ' + promptCount + ' prompts');
              
              // Store prompts for each column
              for (var pi = 0; pi < potentialPromptRow.length; pi++) {
                var cellText = String(potentialPromptRow[pi] || '').trim();
                if (cellText.length > 40) {
                  var promptColLetter = String.fromCharCode(64 + selStartCol + pi);
                  columnPrompts[promptColLetter] = cellText;
                  Logger.log('  Column ' + promptColLetter + ' prompt: ' + cellText.substring(0, 50) + '...');
                }
              }
              break; // Use the first (closest to header) prompt row found
            }
          }
        }
        
        for (var c = selStartCol; c <= selEndCol; c++) {
          var colLetter = String.fromCharCode(64 + c);
          var colIndex = c - selStartCol;
          // Use header from detected header row, fallback to sheet's row 1
          var header = String(selHeaders[colIndex] || '').trim() || headerRow[colLetter] || null;
          
          selectedHeaders.push({
            column: colLetter,
            name: header,
            customPrompt: columnPrompts[colLetter] || null // Include custom prompt if found
          });
          
          // Check if column has data WITHIN THE SELECTED DATA RANGE
          var hasDataInSelection = false;
          if (selEndRow >= actualDataStartRow) {
            var dataRowCount = selEndRow - actualDataStartRow + 1;
            var selDataRange = sheet.getRange(actualDataStartRow, c, dataRowCount, 1);
            var selDataValues = selDataRange.getValues();
            for (var r = 0; r < selDataValues.length; r++) {
              var cellVal = selDataValues[r][0];
              if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== '') {
                hasDataInSelection = true;
                break;
              }
            }
          }
          
          if (hasDataInSelection) {
            columnsWithData.push(colLetter);
          } else {
            emptyColumns.push({ column: colLetter, header: header });
          }
        }
        
        // Calculate the data range within the selection
        var selDataStartRow = actualDataStartRow;
        var selDataRowCount = selEndRow - actualDataStartRow + 1;
        
        // Detect data validation rules on empty columns (output columns)
        // This helps constrain AI outputs to valid values
        var columnValidation = {};
        emptyColumns.forEach(function(emptyCol) {
          var validation = getColumnValidation(emptyCol.column, selDataStartRow, selDataRowCount);
          if (validation) {
            columnValidation[emptyCol.column] = validation;
            Logger.log('  Validation on ' + emptyCol.column + ': ' + JSON.stringify(validation));
          }
        });
        
        // Debug logging for table detection
        Logger.log('Table selection detected: ' + selRange);
        Logger.log('  Columns with data: ' + columnsWithData.join(', '));
        Logger.log('  Empty columns: ' + emptyColumns.map(function(c) { return c.column; }).join(', '));
        Logger.log('  Source column: ' + (columnsWithData.length > 0 ? columnsWithData[0] : 'none'));
        Logger.log('  Data rows: ' + selDataStartRow + ' to ' + selEndRow + ' (' + selDataRowCount + ' rows)');
        Logger.log('  Columns with validation: ' + Object.keys(columnValidation).join(', '));
        Logger.log('  Custom prompts: ' + Object.keys(columnPrompts).join(', '));
        
        selectionInfo = {
          range: selRange,
          isTable: true,
          isExplicit: true, // User explicitly selected this range
          selectionSource: 'explicit',
          numRows: selDataRowCount, // Data rows only (excluding header)
          numCols: numCols,
          headers: selectedHeaders,
          columnsWithData: columnsWithData,
          emptyColumns: emptyColumns, // Columns that need to be filled
          sourceColumn: columnsWithData.length > 0 ? columnsWithData[0] : null,
          // Include the actual data range for the source column within the selection
          dataStartRow: selDataStartRow,
          dataEndRow: selEndRow,
          dataRowCount: selDataRowCount,
          // Data validation rules on output columns (for constraining AI output)
          columnValidation: columnValidation,
          // Custom prompts from prompt row above headers
          columnPrompts: Object.keys(columnPrompts).length > 0 ? columnPrompts : null,
          promptRowNum: promptRowNum
        };
      }
    }
    
    // AUTO-DETECT: Only when user clicked a single cell or didn't select anything meaningful
    // This provides smart defaults - finds the table around the cursor
    var shouldAutoDetect = !selectionInfo && lastCol > 0 && lastRow > 1;
    Logger.log('Checking auto-detect: selectionInfo=' + (selectionInfo ? 'exists (explicit)' : 'null') + ', shouldAutoDetect=' + shouldAutoDetect);
    
    if (shouldAutoDetect) {
      Logger.log('Calling autoDetectDataRegion...');
      var autoDetected = autoDetectDataRegion(sheet, columnDataRanges, headerRow, lastRow, lastCol);
      Logger.log('autoDetectDataRegion returned: ' + (autoDetected ? 'object' : 'null'));
      if (autoDetected) {
        // Mark as auto-detected so frontend knows this wasn't explicit user selection
        autoDetected.isExplicit = false;
        autoDetected.selectionSource = 'auto';
        selectionInfo = autoDetected;
        Logger.log('Auto-detected data region: ' + JSON.stringify({
          range: autoDetected.range,
          columnsWithData: autoDetected.columnsWithData,
          emptyColumns: autoDetected.emptyColumns.map(function(c) { return c.column; }),
          source: 'auto'
        }));
      }
    }
    
    return {
      selectedRange: selection ? selection.getA1Notation() : null,
      selectionInfo: selectionInfo,
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),  // Unique ID to detect sheet changes
      capturedAt: new Date().getTime(),  // Timestamp for staleness detection
      headers: headers,
      headerRow: headerRow,
      columnDataRanges: columnDataRanges,
      sampleData: sampleData,
      lastRow: lastRow,
      lastColumn: lastCol,
      // Debug info
      _debug: {
        autoDetectAttempted: !selectionInfo && lastCol > 0 && lastRow > 1,
        autoDetectResult: selectionInfo ? 'success' : 'null_or_not_attempted'
      }
    };
  } catch (e) {
    Logger.log('Failed to get context: ' + e.message);
    return {
      selectedRange: null,
      sheetName: null,
      headers: [],
      headerRow: {},
      columnDataRanges: {},
      sampleData: {},
      lastRow: 0,
      lastColumn: 0
    };
  }
}

/**
 * Auto-detect the data region when no selection is made
 * Finds headers, data columns, and empty columns automatically
 * 
 * @param {Sheet} sheet - The sheet
 * @param {Object} columnDataRanges - Pre-computed column data ranges
 * @param {Object} headerRow - Pre-computed headers from row 1
 * @param {number} lastRow - Last row with data
 * @param {number} lastCol - Last column with data
 * @return {Object|null} Selection info object or null if no data found
 */
function autoDetectDataRegion(sheet, columnDataRanges, headerRow, lastRow, lastCol) {
  Logger.log('Auto-detect called: lastRow=' + lastRow + ', lastCol=' + lastCol);
  Logger.log('Auto-detect columnDataRanges keys: ' + Object.keys(columnDataRanges).join(','));
  Logger.log('Auto-detect initial headerRow: ' + JSON.stringify(headerRow));
  
  // Find columns with data
  var columnsWithData = [];
  var emptyColumns = [];
  var allHeaders = [];
  var columnPrompts = {}; // Store detected prompts per column
  var promptRowNum = null;
  
  // Determine the header row - check if row 1 has headers, otherwise find first row with content
  var headerRowNum = 1;
  var dataStartRow = 2;
  
  // Helper function to check if a row looks like headers vs prompts
  // Headers are typically short (< 50 chars), prompts are long instructions
  function isHeaderRow(rowValues) {
    var nonEmptyCells = rowValues.filter(function(v) {
      return v !== null && v !== undefined && String(v).trim() !== '';
    });
    if (nonEmptyCells.length === 0) return false;
    
    // Check average length and count of long cells
    var longCellCount = 0;
    var totalLength = 0;
    for (var i = 0; i < nonEmptyCells.length; i++) {
      var cellLen = String(nonEmptyCells[i]).trim().length;
      totalLength += cellLen;
      if (cellLen > 50) longCellCount++;
    }
    
    var avgLength = totalLength / nonEmptyCells.length;
    
    // If most cells are long (>50 chars) or average is high, it's likely prompts not headers
    // Headers are typically short column names like "Brand", "Model", "Type"
    var isPromptRow = longCellCount >= Math.ceil(nonEmptyCells.length / 2) || avgLength > 40;
    
    Logger.log('Row analysis: nonEmpty=' + nonEmptyCells.length + ', longCells=' + longCellCount + ', avgLen=' + avgLength.toFixed(0) + ', isPrompt=' + isPromptRow);
    
    return !isPromptRow;
  }
  
  // Check if row 1 looks like headers (has text content)
  var row1HasContent = Object.keys(headerRow).length > 0;
  var row1Values = row1HasContent ? null : [];
  
  // Check if row 1 content looks like headers (not prompts)
  if (row1HasContent) {
    // Get row 1 values to check if they look like headers
    var row1Range = sheet.getRange(1, 1, 1, lastCol);
    row1Values = row1Range.getValues()[0];
    if (!isHeaderRow(row1Values)) {
      Logger.log('Row 1 has content but looks like prompts, searching for actual headers...');
      row1HasContent = false; // Force to search for real headers
      // Store as potential prompts
      promptRowNum = 1;
      for (var c = 0; c < row1Values.length; c++) {
        var cellText = String(row1Values[c] || '').trim();
        if (cellText.length > 40) {
          var colLetter = String.fromCharCode(65 + c);
          columnPrompts[colLetter] = cellText;
        }
      }
    }
  }
  
  if (!row1HasContent) {
    // Row 1 is empty or has prompts - find the actual header row
    // Search through rows to find one that looks like headers (short text)
    var foundHeaders = false;
    
    for (var r = 1; r <= Math.min(lastRow, 15); r++) {
      var rowRange = sheet.getRange(r, 1, 1, lastCol);
      var rowValues = rowRange.getValues()[0];
      var hasContent = rowValues.some(function(v) {
        return v !== null && v !== undefined && String(v).trim() !== '';
      });
      
      if (!hasContent) continue; // Skip empty rows
      
      // Check if this row looks like headers (short text) vs prompts (long text)
      if (isHeaderRow(rowValues)) {
        // This looks like a header row!
        headerRowNum = r;
        dataStartRow = r + 1;
        // Clear old headerRow and read new headers
        headerRow = {};
        for (var c = 0; c < rowValues.length; c++) {
          if (rowValues[c] && String(rowValues[c]).trim()) {
            var colLetter = String.fromCharCode(65 + c);
            headerRow[colLetter] = String(rowValues[c]).trim();
          }
        }
        foundHeaders = true;
        Logger.log('Found header row at row ' + r + ': ' + JSON.stringify(headerRow));
        break;
      } else {
        // This looks like a prompt row - store prompts if not already stored
        if (!promptRowNum) {
          promptRowNum = r;
          for (var c = 0; c < rowValues.length; c++) {
            var cellText = String(rowValues[c] || '').trim();
            if (cellText.length > 40) {
              var colLetter = String.fromCharCode(65 + c);
              columnPrompts[colLetter] = cellText;
            }
          }
          Logger.log('Detected prompt row at row ' + r + ': ' + Object.keys(columnPrompts).join(','));
        }
      }
    }
    
    if (!foundHeaders) {
      Logger.log('No header row found (only prompts detected)');
      return null;
    }
  }
  
  // If still no headers found, can't auto-detect
  if (Object.keys(headerRow).length === 0) {
    return null;
  }
  
  Logger.log('Final header row: ' + headerRowNum + ', data starts: ' + dataStartRow);
  Logger.log('Column prompts detected: ' + JSON.stringify(Object.keys(columnPrompts)));
  
  // Analyze each column - only check columns that have headers defined
  // Don't extend beyond the actual data/header range
  for (var col = 1; col <= Math.min(lastCol, 26); col++) {
    var colLetter = String.fromCharCode(64 + col);
    var header = headerRow[colLetter] || null;
    var colData = columnDataRanges[colLetter];
    
    allHeaders.push({
      column: colLetter,
      name: header
    });
    
    // Check if column has data at or below the data start row
    // Note: colData.startRow might equal headerRowNum if header was detected as data
    // So we check if there's data that extends past the header row
    var hasData = colData && colData.hasData && colData.endRow >= dataStartRow;
    
    if (hasData) {
      columnsWithData.push(colLetter);
    } else if (header) {
      // Has header but no data - this is a potential output column
      emptyColumns.push({
        column: colLetter,
        header: header
      });
    }
  }
  
  // Debug logging
  Logger.log('Auto-detect: headerRowNum=' + headerRowNum + ', dataStartRow=' + dataStartRow);
  Logger.log('Auto-detect: columnsWithData=' + columnsWithData.join(',') + ', emptyColumns=' + emptyColumns.map(function(c) { return c.column; }).join(','));
  
  // Need at least one column with data
  if (columnsWithData.length === 0) {
    Logger.log('Auto-detect: No columns with data found');
    return null;
  }
  
  // Ensure we have enough empty columns for multi-aspect workflows
  // Add the next N columns after the last data/empty column (even if they don't have headers)
  if (emptyColumns.length < 10 && columnsWithData.length > 0) {
    var lastColCode = emptyColumns.length > 0 
      ? emptyColumns[emptyColumns.length - 1].column.charCodeAt(0)
      : columnsWithData[columnsWithData.length - 1].charCodeAt(0);
    
    // Add up to 10 total empty columns
    var columnsToAdd = 10 - emptyColumns.length;
    for (var i = 1; i <= columnsToAdd; i++) {
      var nextColLetter = String.fromCharCode(lastColCode + i);
      if (nextColLetter <= 'Z' && columnsWithData.indexOf(nextColLetter) === -1) {
        // Don't add if already in emptyColumns
        var alreadyExists = emptyColumns.some(function(ec) { return ec.column === nextColLetter; });
        if (!alreadyExists) {
          emptyColumns.push({
            column: nextColLetter,
            header: headerRow[nextColLetter] || null
          });
        }
      }
    }
    Logger.log('Auto-detect: Ensured ' + emptyColumns.length + ' empty columns for multi-aspect support');
  }
  
  // Determine the data extent
  var firstDataCol = columnsWithData[0];
  var lastDataCol = columnsWithData[columnsWithData.length - 1];
  var firstColData = columnDataRanges[firstDataCol];
  
  // Find the actual data end row (max across all data columns)
  // IMPORTANT: Use dataStartRow (not colData.startRow) to exclude header
  var dataEndRow = dataStartRow;
  columnsWithData.forEach(function(col) {
    var colInfo = columnDataRanges[col];
    if (colInfo && colInfo.endRow > dataEndRow) {
      dataEndRow = colInfo.endRow;
    }
  });
  
  // Calculate row count from dataStartRow (excludes header)
  var dataRowCount = dataEndRow - dataStartRow + 1;
  
  // Debug: Log the range calculation
  Logger.log('Auto-detect range: headerRow=' + headerRowNum + ', dataStart=' + dataStartRow + ', dataEnd=' + dataEndRow + ', count=' + dataRowCount);
  
  // Build the auto-detected range string - use dataStartRow NOT headerRowNum for data range
  // The range should be DATA only, not including header
  // IMPORTANT: Include ALL columns with data, not just the first column
  var dataRangeStr = firstDataCol + dataStartRow + ':' + lastDataCol + dataEndRow;
  
  // Full range for display - only include columns with HEADERS (not suggested output columns without headers)
  // This keeps the display tight while still allowing output to next column
  var lastDisplayCol = lastDataCol;
  var emptyColsWithHeaders = emptyColumns.filter(function(c) { return c.header; });
  if (emptyColsWithHeaders.length > 0) {
    lastDisplayCol = emptyColsWithHeaders[emptyColsWithHeaders.length - 1].column;
  }
  var fullRangeStr = firstDataCol + headerRowNum + ':' + lastDisplayCol + dataEndRow;
  
  // Enrich emptyColumns with prompts if available
  emptyColumns = emptyColumns.map(function(col) {
    return {
      column: col.column,
      header: col.header,
      prompt: columnPrompts[col.column] || null  // Include custom prompt if detected
    };
  });
  
  return {
    range: fullRangeStr,  // Full range for display (includes header)
    dataRange: dataRangeStr,  // Data-only range (excludes header)
    isTable: true,
    isAutoDetected: true, // Flag to indicate this was auto-detected
    numRows: dataRowCount,  // Number of DATA rows (excludes header)
    numCols: columnsWithData.length + emptyColumns.length,
    headers: allHeaders.filter(function(h) { return h.name; }),
    columnsWithData: columnsWithData,
    emptyColumns: emptyColumns,
    sourceColumn: firstDataCol,
    dataStartRow: dataStartRow,  // First row of DATA (after header)
    dataEndRow: dataEndRow,
    dataRowCount: dataRowCount,
    headerRow: headerRowNum,  // The row containing headers
    promptRow: promptRowNum,  // Row containing custom prompts (if any)
    columnPrompts: columnPrompts,  // Custom prompts per column
    columnValidation: {} // Could add validation detection here too
  };
}

/**
 * Get sample data from a column for AI context
 * @param {Sheet} sheet - The sheet
 * @param {number} colNum - Column number (1-indexed)
 * @param {number} startRow - Starting row
 * @param {number} count - Number of samples to get
 * @return {Array<string>} Sample values
 */
function getSampleData(sheet, colNum, startRow, count) {
  try {
    var samples = [];
    var maxRows = Math.min(count * 2, 10); // Check up to 10 rows to find samples
    var range = sheet.getRange(startRow, colNum, maxRows, 1);
    var values = range.getValues();
    
    for (var i = 0; i < values.length && samples.length < count; i++) {
      var val = values[i][0];
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        // Truncate long values for context
        var strVal = String(val).trim();
        if (strVal.length > 100) {
          strVal = strVal.substring(0, 100) + '...';
        }
        samples.push(strVal);
      }
    }
    
    return samples;
  } catch (e) {
    return [];
  }
}

/**
 * Detect the data range for a specific column
 * Finds the first and last row with content (skipping header)
 * @param {Sheet} sheet - The sheet to check
 * @param {number} colNum - Column number (1-indexed)
 * @param {number} lastRow - Last row in sheet
 * @return {Object} { hasData, startRow, endRow, rowCount, range }
 */
function detectColumnDataRange(sheet, colNum, lastRow) {
  try {
    if (lastRow < 2) {
      return { hasData: false };
    }
    
    // Get column values starting from row 2 (skip header)
    var range = sheet.getRange(2, colNum, lastRow - 1, 1);
    var values = range.getValues();
    
    var firstDataRow = -1;
    var lastDataRow = -1;
    
    for (var i = 0; i < values.length; i++) {
      var val = values[i][0];
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        if (firstDataRow === -1) {
          firstDataRow = i + 2; // +2 because we started from row 2
        }
        lastDataRow = i + 2;
      }
    }
    
    if (firstDataRow === -1) {
      return { hasData: false };
    }
    
    var colLetter = String.fromCharCode(64 + colNum);
    return {
      hasData: true,
      startRow: firstDataRow,
      endRow: lastDataRow,
      rowCount: lastDataRow - firstDataRow + 1,
      range: colLetter + firstDataRow + ':' + colLetter + lastDataRow
    };
  } catch (e) {
    Logger.log('Error detecting column data: ' + e.message);
    return { hasData: false };
  }
}

/**
 * Check if output columns have existing data
 * @param {Array<string>} outputColumns - Column letters to check
 * @param {number} startRow - Starting row
 * @param {number} rowCount - Number of rows
 * @return {boolean} True if any output column has data
 */
function checkOutputHasData(outputColumns, startRow, rowCount) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    
    for (var i = 0; i < outputColumns.length; i++) {
      var col = outputColumns[i];
      var colNum = col.charCodeAt(0) - 64; // A=1, B=2, etc.
      var range = sheet.getRange(startRow, colNum, rowCount, 1);
      var values = range.getValues();
      
      // Check if any cell has data
      for (var j = 0; j < values.length; j++) {
        if (values[j][0] !== '' && values[j][0] !== null) {
          return true;
        }
      }
    }
    
    return false;
  } catch (e) {
    Logger.log('Failed to check output data: ' + e.message);
    return false;
  }
}

// ============================================
// FORMULA ALTERNATIVE
// ============================================

/**
 * Apply a spreadsheet formula to a range (alternative to AI processing)
 * This is instant and free - no API calls needed!
 * 
 * @param {string} formulaTemplate - Formula template with {input} placeholder
 * @param {string} inputCol - Input column letter (e.g., 'B')
 * @param {string} outputCol - Output column letter (e.g., 'C')
 * @param {number} startRow - Starting row number
 * @param {number} endRow - Ending row number
 * @return {Object} { success, rowCount }
 */
function applyFormulaToRange(formulaTemplate, inputCol, outputCol, startRow, endRow) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var rowCount = endRow - startRow + 1;
    
    Logger.log('Applying formula to ' + rowCount + ' rows');
    Logger.log('Template: ' + formulaTemplate);
    Logger.log('Input: ' + inputCol + ', Output: ' + outputCol + ', Rows: ' + startRow + '-' + endRow);
    
    // Build formulas for each row
    var formulas = [];
    for (var row = startRow; row <= endRow; row++) {
      // Replace {input} placeholder with actual cell reference
      var cellRef = inputCol + row;
      var formula = formulaTemplate.replace(/\{input\}/gi, cellRef);
      formulas.push([formula]);
    }
    
    // Get output range
    var outputColNum = outputCol.charCodeAt(0) - 64; // A=1, B=2, etc.
    var outputRange = sheet.getRange(startRow, outputColNum, rowCount, 1);
    
    // Apply all formulas at once (efficient!)
    outputRange.setFormulas(formulas);
    
    Logger.log('Formula applied successfully to ' + rowCount + ' rows');
    
    return {
      success: true,
      rowCount: rowCount
    };
  } catch (e) {
    Logger.log('Formula application failed: ' + e.message);
    throw new Error('Failed to apply formula: ' + e.message);
  }
}

// ============================================
// SUBSCRIPTION & USAGE MANAGEMENT
// ============================================

/**
 * Get user's subscription and usage information
 * @return {Object} { tier, remaining, limit, isUnlimited, canProceed, needsUpgrade }
 */
function getUserSubscription() {
  try {
    var userId = getUserId();
    var userEmail = getUserEmail();
    
    var result = ApiClient.post('USAGE_CHECK', { 
      userId: userId,
      userEmail: userEmail
    });
    
    return {
      tier: result.tier || 'free',
      remaining: result.remaining || 0,
      limit: result.limit || 500,
      isUnlimited: result.isUnlimited || false,
      canProceed: result.canProceed !== false,
      needsUpgrade: result.needsUpgrade || false,
      reason: result.reason || null,
      // Managed AI credit info (new)
      managedCredits: result.managedCredits || { canUse: false, used_usd: 0, cap_usd: 0, remaining_usd: 0 },
      managedModels: result.managedModels || []
    };
  } catch (e) {
    Logger.log('Error getting subscription: ' + e);
    // Return default free tier on error
    return { 
      tier: 'free', 
      remaining: 300, 
      limit: 300,
      isUnlimited: false,
      canProceed: true,
      needsUpgrade: false,
      managedCredits: { canUse: true, used_usd: 0, cap_usd: 0.015, remaining_usd: 0.015 },
      managedModels: []
    };
  }
}

/**
 * Create Stripe checkout session for Pro upgrade
 * @return {string|null} Stripe checkout URL or null on error
 */
function createCheckoutSession() {
  try {
    var userId = getUserId();
    var userEmail = getUserEmail();
    
    var result = ApiClient.post('STRIPE_CHECKOUT', {
      userId: userId,
      userEmail: userEmail,
      tier: 'pro'
    });
    
    return result.url || null;
  } catch (e) {
    Logger.log('Error creating checkout session: ' + e);
    return null;
  }
}

/**
 * Get Stripe billing portal URL for subscription management
 * @return {string|null} Stripe portal URL or null on error
 */
function getBillingPortalUrl() {
  try {
    var userId = getUserId();
    
    var result = ApiClient.post('STRIPE_PORTAL', { 
      userId: userId 
    });
    
    return result.url || null;
  } catch (e) {
    Logger.log('Error getting billing portal URL: ' + e);
    return null;
  }
}

// ============================================
// ANALYTICS & EVENT LOGGING
// ============================================

/**
 * Log analytics events to backend
 * Called from frontend with batched events
 * Fire-and-forget - failures are silently ignored
 * 
 * @param {Array} events - Array of analytics events
 * @return {boolean} Success status
 */
function logAnalyticsEvents(events) {
  try {
    if (!events || !Array.isArray(events) || events.length === 0) {
      return true; // Nothing to log
    }
    
    var userId = getUserId();
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    // Add user context to all events
    var enrichedEvents = events.map(function(event) {
      return {
        userId: userId,
        spreadsheetId: spreadsheetId,
        sessionId: event.sessionId,
        timestamp: event.timestamp,
        eventType: event.eventType,
        payload: event.payload || {},
        context: event.context || {}
      };
    });
    
    // Send to backend (async, fire-and-forget)
    try {
      ApiClient.post('ANALYTICS_LOG', {
        events: enrichedEvents
      });
    } catch (apiError) {
      // Silent failure - don't disrupt user experience
      Logger.log('Analytics API call failed (silent): ' + apiError.message);
      
      // Fallback: Store in user properties for later sync
      storeAnalyticsLocally(enrichedEvents);
    }
    
    return true;
  } catch (e) {
    // Never throw - analytics should never break the app
    Logger.log('Analytics logging failed (silent): ' + e.message);
    return false;
  }
}

/**
 * Store analytics locally as fallback when API is unavailable
 * Will be synced later when connection is restored
 * 
 * @param {Array} events - Events to store
 */
function storeAnalyticsLocally(events) {
  try {
    var props = PropertiesService.getUserProperties();
    var existing = JSON.parse(props.getProperty('pendingAnalytics') || '[]');
    
    // Add new events, keep only last 100 to avoid quota issues
    var combined = existing.concat(events).slice(-100);
    
    props.setProperty('pendingAnalytics', JSON.stringify(combined));
  } catch (e) {
    // Silent failure
    Logger.log('Local analytics storage failed: ' + e.message);
  }
}

/**
 * Sync pending local analytics to backend
 * Called periodically or when connection is restored
 * 
 * @return {number} Number of events synced
 */
function syncPendingAnalytics() {
  try {
    var props = PropertiesService.getUserProperties();
    var pending = JSON.parse(props.getProperty('pendingAnalytics') || '[]');
    
    if (pending.length === 0) return 0;
    
    // Try to send to backend
    ApiClient.post('ANALYTICS_LOG', {
      events: pending
    });
    
    // Clear on success
    props.deleteProperty('pendingAnalytics');
    
    Logger.log('Synced ' + pending.length + ' pending analytics events');
    return pending.length;
  } catch (e) {
    Logger.log('Analytics sync failed: ' + e.message);
    return 0;
  }
}

// ============================================
// DYNAMIC SUGGESTIONS API
// ============================================

/**
 * Call the suggestions API to get contextual next-action suggestions
 * Uses semantic search + LLM fallback for intelligent recommendations
 * 
 * @param {Object} requestBody - {steps, dataContext, command}
 * @return {Object} {domain, domainLabel, insight, suggestions, source}
 */
function callSuggestionsAPI(requestBody) {
  try {
    // Get user's selected provider — may be BYOK or managed
    var selectedModel = requestBody.provider || getAgentModel() || 'GEMINI';
    var isManagedSugg = selectedModel.indexOf('MANAGED:') === 0;
    var managedSuggId = isManagedSugg ? selectedModel.replace('MANAGED:', '') : null;
    var provider = isManagedSugg ? 'GEMINI' : selectedModel;
    Logger.log('[Suggestions] Provider from request: ' + requestBody.provider + ', resolved: ' + provider + (isManagedSugg ? ' (managed: ' + managedSuggId + ')' : ''));
    
    // Build payload — managed or BYOK
    var payload;
    if (isManagedSugg) {
      payload = SecureRequest.buildManagedPayload(requestBody, managedSuggId);
    } else {
      payload = SecureRequest.buildPayload(provider, requestBody);
    }
    
    Logger.log('[Suggestions] Calling API with provider: ' + provider);
    
    var result = ApiClient.post('AGENT_SUGGESTIONS', payload);
    
    if (result.error) {
      Logger.log('[Suggestions] API error: ' + result.error);
      return null;
    }
    
    Logger.log('[Suggestions] Got ' + (result.suggestions?.length || 0) + ' suggestions from ' + (result.source || 'unknown'));
    return result;
  } catch (e) {
    Logger.log('[Suggestions] API call failed: ' + e.message);
    return null;
  }
}

/**
 * Record a successful suggestion usage for learning
 * Called when user clicks a suggestion and the task completes successfully
 * 
 * @param {Object} data - {workflowId, suggestion: {title, command}, success}
 * @return {Object} {success: boolean}
 */
function recordSuggestionSuccess(data) {
  try {
    if (!data.workflowId || !data.suggestion) {
      Logger.log('[Suggestions] Missing workflowId or suggestion data');
      return { success: false };
    }
    
    Logger.log('[Suggestions] Recording success for workflow: ' + data.workflowId);
    
    var result = ApiClient.put('AGENT_SUGGESTIONS', data);
    
    return { success: !result.error };
  } catch (e) {
    Logger.log('[Suggestions] Record failed: ' + e.message);
    return { success: false };
  }
}
