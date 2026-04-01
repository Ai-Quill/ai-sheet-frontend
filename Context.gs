/**
 * @file Context.gs
 * @version 2.0.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.0.0 (2026-01-10): Initial context engineering implementation
 * 
 * ============================================
 * CONTEXT.gs - Context Engineering
 * ============================================
 * 
 * Implements task type inference for optimized AI responses.
 * Automatically detects user intent and applies appropriate
 * system prompts on the backend.
 * 
 * Task Types:
 * - EXTRACT: Pull specific data from text
 * - SUMMARIZE: Condense information
 * - CLASSIFY: Categorize content
 * - TRANSLATE: Language conversion
 * - CODE: Generate formulas/scripts
 * - ANALYZE: Data analysis & insights
 * - FORMAT: Clean/structure data
 * - GENERAL: Default for other tasks
 */

// ============================================
// TASK TYPE DEFINITIONS
// ============================================

/**
 * Available task types
 * @enum {string}
 */
var TASK_TYPES = {
  EXTRACT: 'EXTRACT',
  SUMMARIZE: 'SUMMARIZE',
  CLASSIFY: 'CLASSIFY',
  TRANSLATE: 'TRANSLATE',
  CODE: 'CODE',
  ANALYZE: 'ANALYZE',
  FORMAT: 'FORMAT',
  GENERAL: 'GENERAL'
};

/**
 * Keywords that trigger specific task types
 * Order matters - first match wins
 */
var TASK_KEYWORDS = {
  EXTRACT: ['extract', 'parse', 'get the', 'find the', 'pull out', 'what is the', 'return only'],
  SUMMARIZE: ['summarize', 'summary', 'tldr', 'brief', 'shorten', 'condense', 'main points'],
  CLASSIFY: ['classify', 'categorize', 'category', 'which type', 'is this', 'label'],
  TRANSLATE: ['translate', 'in spanish', 'in french', 'in german', 'in chinese', 'in japanese', 'to english', 'convert to'],
  CODE: ['formula', 'code', 'script', 'function', 'regex', 'sql', 'query'],
  ANALYZE: ['analyze', 'analysis', 'trend', 'insight', 'pattern', 'compare', 'correlation'],
  FORMAT: ['format', 'clean', 'fix', 'reformat', 'standardize', 'normalize', 'proper case']
};

// ============================================
// INFERENCE FUNCTIONS
// ============================================

/**
 * Infer task type from user prompt
 * Uses keyword matching with priority ordering
 * 
 * @param {string} input - User's prompt
 * @return {string} Detected task type
 */
function inferTaskType(input) {
  if (!input || typeof input !== 'string') {
    return TASK_TYPES.GENERAL;
  }
  
  var lower = input.toLowerCase();
  
  // Check each task type in priority order
  var taskOrder = ['EXTRACT', 'SUMMARIZE', 'CLASSIFY', 'TRANSLATE', 'CODE', 'ANALYZE', 'FORMAT'];
  
  for (var i = 0; i < taskOrder.length; i++) {
    var taskType = taskOrder[i];
    var keywords = TASK_KEYWORDS[taskType];
    
    for (var j = 0; j < keywords.length; j++) {
      if (lower.includes(keywords[j])) {
        if (Config.isDebug()) {
          Logger.log('Task type inferred: ' + taskType + ' (keyword: ' + keywords[j] + ')');
        }
        return taskType;
      }
    }
  }
  
  return TASK_TYPES.GENERAL;
}

/**
 * Get task type description for display
 * 
 * @param {string} taskType - Task type key
 * @return {Object} Description with label and icon
 */
function getTaskTypeInfo(taskType) {
  var info = {
    EXTRACT: { label: 'Extract', icon: '📤', description: 'Pull specific data from text' },
    SUMMARIZE: { label: 'Summarize', icon: '📝', description: 'Condense information' },
    CLASSIFY: { label: 'Classify', icon: '🏷️', description: 'Categorize content' },
    TRANSLATE: { label: 'Translate', icon: '🌐', description: 'Language conversion' },
    CODE: { label: 'Code', icon: '💻', description: 'Generate formulas/scripts' },
    ANALYZE: { label: 'Analyze', icon: '📊', description: 'Data analysis & insights' },
    FORMAT: { label: 'Format', icon: '✨', description: 'Clean/structure data' },
    GENERAL: { label: 'General', icon: '💬', description: 'General AI response' }
  };
  
  return info[taskType] || info.GENERAL;
}

/**
 * Check if a task type benefits from structured output
 * 
 * @param {string} taskType - Task type key
 * @return {boolean} Whether to request structured output
 */
function shouldUseStructuredOutput(taskType) {
  var structuredTypes = ['EXTRACT', 'CLASSIFY', 'CODE', 'FORMAT'];
  return structuredTypes.includes(taskType);
}

/**
 * Get optimal model for a task type (cost/quality balance)
 * 
 * @param {string} taskType - Task type key
 * @return {Object} Recommended models by priority
 */
function getRecommendedModels(taskType) {
  var recommendations = {
    EXTRACT: {
      fast: 'GEMINI',      // Gemini for simple extractions
      quality: 'CHATGPT'   // GPT for complex parsing
    },
    SUMMARIZE: {
      fast: 'GEMINI',
      quality: 'CLAUDE'    // Claude excels at summarization
    },
    CLASSIFY: {
      fast: 'GROQ',        // Groq is fast for classification
      quality: 'CHATGPT'
    },
    TRANSLATE: {
      fast: 'GEMINI',
      quality: 'CLAUDE'
    },
    CODE: {
      fast: 'GROQ',
      quality: 'CLAUDE'    // Claude 4.5 is best for code
    },
    ANALYZE: {
      fast: 'GEMINI',
      quality: 'CHATGPT'
    },
    FORMAT: {
      fast: 'GEMINI',
      quality: 'CHATGPT'
    },
    GENERAL: {
      fast: 'GEMINI',
      quality: 'CHATGPT'
    }
  };
  
  return recommendations[taskType] || recommendations.GENERAL;
}

// ============================================
// ENHANCED AI QUERY (with context)
// ============================================

/**
 * AI Query with automatic task type inference
 * Wrapper around AIQuery that adds context engineering
 * 
 * @param {string} model - Provider name
 * @param {string} input - User prompt
 * @param {string} [imageUrl] - Optional image URL
 * @param {string} [specificModel] - Optional specific model
 * @param {string} [taskType] - Optional explicit task type (auto-detected if not provided)
 * @return {string} AI response
 */
function AIQueryWithContext(model, input, imageUrl, specificModel, taskType) {
  // Infer task type if not provided
  if (!taskType) {
    taskType = inferTaskType(input);
  }
  
  // Use default model if not specified
  if (!specificModel) {
    specificModel = Config.getDefaultModel(model);
  }
  
  // Build enhanced payload using SecureRequest (centralized API key handling)
  var payload = SecureRequest.buildPayloadWithUser(model, {
    model: model,
    input: input,
    specificModel: specificModel,
    imageUrl: imageUrl || null,
    taskType: taskType  // Backend will apply appropriate system prompt
  });
  
  try {
    if (Config.isDebug()) {
      Logger.log('AIQueryWithContext: ' + model + ' → ' + specificModel + ' [' + taskType + ']');
    }
    
    var result = ApiClient.post('QUERY', payload);
    
    // Track credit usage
    if (result.creditsUsed) {
      logCreditUsage(result.creditsUsed);
    }
    
    return result.result;
    
  } catch (error) {
    Logger.log('AIQueryWithContext Error: ' + error.toString());
    return 'Error: ' + error.message;
  }
}
