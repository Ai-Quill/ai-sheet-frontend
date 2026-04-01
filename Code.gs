/**
 * @file Code.gs
 * @version 2.2.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.2.0 (2026-01-10): Added include() helper for modular HTML
 * - 2.1.0 (2026-01-10): Input validation for empty prompts
 * - 2.0.0 (2026-01-09): Initial modular architecture
 * 
 * ============================================
 * CODE.gs - Main Entry Point
 * ============================================
 * 
 * Core AISheeter functionality:
 * - Menu and sidebar initialization
 * - AI query functions (ChatGPT, Claude, Groq, Gemini)
 * - Image generation (DALL-E)
 * - Contact form
 * 
 * Dependencies:
 * - Config.gs    → Environment configuration
 * - ApiClient.gs → HTTP request utilities
 * - Crypto.gs    → Encryption/decryption
 * - User.gs      → User management & settings
 * - Prompts.gs   → Saved prompts CRUD
 */

// ============================================
// MENU & SIDEBAR
// ============================================

/**
 * Initialize menu when spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('AISheeter - Any LLM in One Google Sheets™')
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Manage Saved Prompts', 'openPromptManager')
    .addToUi();
  
  // Auto-show sidebar on open
  showSidebar();
}

/**
 * Display the main sidebar
 * Uses HTML templates with includes for modular architecture
 * Wide mode uses modeless dialog (can be any size)
 * Narrow mode uses standard sidebar
 */
function showSidebar() {
  var template = HtmlService.createTemplateFromFile('Sidebar');
  var isWideMode = getSidebarWidthPreference();
  
  var html = template.evaluate()
    .setTitle('AI Sheet - Any LLM');
  
  if (isWideMode) {
    // Wide mode: Use modeless dialog positioned on the right
    html.setWidth(500).setHeight(700);
    SpreadsheetApp.getUi().showModelessDialog(html, 'AI Sheet - Any LLM');
  } else {
    // Narrow mode: Use standard sidebar
    SpreadsheetApp.getUi().showSidebar(html);
  }
}

/**
 * Get user's sidebar width preference
 * @return {boolean} true for wide mode (dialog), false for narrow (sidebar)
 */
function getSidebarWidthPreference() {
  var props = PropertiesService.getUserProperties();
  var pref = props.getProperty('SIDEBAR_WIDE_MODE');
  // Default to narrow mode (sidebar) for now - more stable
  return pref === 'true';
}

/**
 * Toggle sidebar width and refresh
 * @return {boolean} New wide mode state
 */
function toggleSidebarWidth() {
  var props = PropertiesService.getUserProperties();
  var currentWide = getSidebarWidthPreference();
  var newWide = !currentWide;
  props.setProperty('SIDEBAR_WIDE_MODE', String(newWide));
  
  // Refresh with new mode
  showSidebar();
  
  return newWide;
}

/**
 * Include helper for modular HTML
 * 
 * Usage in HTML: <?!= include('Sidebar_Styles') ?>
 * 
 * This enables splitting large HTML files into manageable modules:
 * - Sidebar_Styles.html   → CSS styles
 * - Sidebar_Bulk.html     → Bulk Agent functionality
 * - Sidebar_Settings.html → Settings forms
 * - Sidebar_Utils.html    → Utilities (toast, menu, helpers)
 * 
 * @param {string} filename - HTML file to include (without .html extension)
 * @return {string} The HTML content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// AI QUERY - CORE FUNCTION
// ============================================

/**
 * Core AI query function
 * Routes requests to the appropriate provider via backend API
 * Includes context engineering with automatic task type inference
 * 
 * @param {string} model - Provider name (CHATGPT, CLAUDE, GROQ, GEMINI)
 * @param {string} input - User prompt
 * @param {string} [imageUrl] - Optional image URL for vision models
 * @param {string} [specificModel] - Optional specific model ID
 * @param {string} [taskType] - Optional task type (auto-detected if not provided)
 * @return {string} AI response or error message
 */
function AIQuery(model, input, imageUrl, specificModel, taskType) {
  // Validate input - prevent empty API calls
  if (!input || (typeof input === 'string' && input.trim() === '')) {
    return '';  // Return empty for empty input (no error, just skip)
  }
  
  // Use default model if not specified
  if (!specificModel) {
    specificModel = Config.getDefaultModel(model);
  }
  
  // Infer task type for context engineering (if not provided)
  if (!taskType) {
    taskType = inferTaskType(input);
  }
  
  // Build request payload using SecureRequest (centralized API key handling)
  var payload = SecureRequest.buildPayloadWithUser(model, {
    model: model,
    input: input,
    specificModel: specificModel,
    imageUrl: imageUrl || null,
    taskType: taskType  // Backend applies appropriate system prompt
  });
      
      try {
    // Log in debug mode
    if (Config.isDebug()) {
      Logger.log('AIQuery: ' + model + ' → ' + specificModel + ' [' + taskType + ']');
    }
    
    var result = ApiClient.post('QUERY', payload);
    
    // Track credit usage locally
    if (result.creditsUsed) {
        logCreditUsage(result.creditsUsed);
    }
    
        return result.result;
    
      } catch (error) {
    Logger.log('AIQuery Error: ' + error.toString());
    return 'Error: ' + error.message;
      }
    }

// ============================================
// CUSTOM FUNCTIONS (Sheet Formulas)
// ============================================

    /**
 * Query ChatGPT (OpenAI)
 * @param {string} prompt - The input prompt
 * @param {string} [imageUrl] - Optional image URL for vision
 * @param {string} [model] - Specific model (default: gpt-5-mini)
 * @return {string} AI response
     * @customfunction
     */
    function ChatGPT(prompt, imageUrl, model) {
      return AIQuery('CHATGPT', prompt, imageUrl, model);
    }

/**
 * Query Claude (Anthropic)
 * @param {string} prompt - The input prompt
 * @param {string} [imageUrl] - Optional image URL for vision
 * @param {string} [model] - Specific model (default: claude-haiku-4-5)
 * @return {string} AI response
 * @customfunction
 */
function Claude(prompt, imageUrl, model) {
  return AIQuery('CLAUDE', prompt, imageUrl, model);
}

/**
 * Query Groq (Fast inference)
 * @param {string} prompt - The input prompt
 * @param {string} [imageUrl] - Optional image URL for vision
 * @param {string} [model] - Specific model (default: meta-llama/llama-4-scout-17b-16e-instruct)
 * @return {string} AI response
 * @customfunction
 */
function Groq(prompt, imageUrl, model) {
  return AIQuery('GROQ', prompt, imageUrl, model);
}

/**
 * Query Gemini (Google)
 * @param {string} prompt - The input prompt
 * @param {string} [imageUrl] - Optional image URL for vision
 * @param {string} [model] - Specific model (default: gemini-2.5-flash)
 * @return {string} AI response
 * @customfunction
 */
function Gemini(prompt, imageUrl, model) {
  return AIQuery('GEMINI', prompt, imageUrl, model);
}

// ============================================
// IMAGE GENERATION
// ============================================

/**
 * Generate an image using AI
 * @param {string} provider - Provider (DALLE, GEMINI)
 * @param {string} prompt - Image description
 * @param {string} model - Specific model
 * @return {string} Generated image URL
 * @throws {Error} If generation fails
 */
function generateImage(provider, prompt, model) {
  // Validate input - prevent empty API calls
  if (!prompt || (typeof prompt === 'string' && prompt.trim() === '')) {
    return '';  // Return empty for empty input
  }
  
  // Map image model to API provider for key lookup
  var apiProvider;
  if (model === 'DALLE') {
    apiProvider = 'CHATGPT';  // DALL-E uses OpenAI API key
  } else if (model === 'GEMINI') {
    apiProvider = 'GEMINI';
  } else {
    throw new Error('Unsupported model for image generation: ' + model);
  }
  
  // Build payload using SecureRequest (centralized API key handling)
  var payload = SecureRequest.buildPayload(apiProvider, {
    model: provider,
    prompt: prompt,
    userId: getUserId(),
    specificModel: model
  });

  try {
    var result = ApiClient.post('GENERATE_IMAGE', payload);
    return result.imageUrl;
  } catch (error) {
    console.error('Error generating image:', error);
    throw new Error('Failed to generate image: ' + error.message);
  }
}

/**
 * Generate image with DALL-E (OpenAI)
 * @param {string} prompt - Image description
 * @return {string} Generated image URL
 * @customfunction
 */
function DALLE(prompt) {
  return generateImage('DALLE', prompt, 'DALLE');
}

// ============================================
// MODELS API
// ============================================

/**
 * Fetch available models from backend
 * Used by Settings panel to populate dropdowns
 * @return {Array<Object>} List of available models
 * @throws {Error} If request fails
 */
function fetchModels() {
  try {
    return ApiClient.get('MODELS');
  } catch (error) {
    console.error('Error fetching models:', error);
    throw new Error('Failed to fetch models: ' + error.message);
  }
}

// ============================================
// FORMULA-FIRST EVALUATION
// ============================================

/**
 * Check if a task can be solved with a native formula
 * This is called BEFORE creating AI jobs - the "formula-first" approach
 * 
 * @param {Object} context - Task context
 * @param {string} context.command - User's original command
 * @param {string} context.taskType - Detected task type
 * @param {string} context.inputColumn - Input column letter
 * @param {string} context.outputColumn - Output column letter
 * @param {string} context.inputRange - Input data range
 * @param {string} context.categories - Categories for classification
 * @return {Object} { canUseFormula: boolean, formula?: string, explanation: string }
 */
/**
 * Check if task can be solved with formula
 * Uses unified context system for FULL data context
 * 
 * @param {Object} context - Task context from frontend
 * @return {Object} Formula evaluation result
 */
function checkFormulaFirst(context) {
  Logger.log('[Code] checkFormulaFirst called with unified context approach');
  
  try {
    // If we don't have rich context, build it
    if (!context.columnDetails && !context.dataSummary) {
      Logger.log('[Code] Enriching context with unified context builder...');
      var unifiedContext = buildUnifiedContext({ command: context.command });
      
      // Merge unified context with provided context
      context.allColumns = context.allColumns || unifiedContext.inputColumns;
      context.columnDetails = unifiedContext.dataColumns;
      context.dataSummary = unifiedContext.dataSummary;
      context.columnHeaders = context.columnHeaders || unifiedContext.headers;
      context.dataStartRow = context.dataStartRow || unifiedContext.dataStartRow;
      context.dataEndRow = context.dataEndRow || unifiedContext.endRow;
      context.isMultiColumn = unifiedContext.inputColumns.length > 1;
      
      Logger.log('[Code] Enriched context with ' + unifiedContext.inputColumns.length + ' columns');
    }
    
    // GENERIC APPROACH: Let AI evaluate with FULL context
    // No hardcoded rules - AI sees all data and decides the best approach
    var result = evaluateFormulaFirst(context);
    
    Logger.log('[Code] Formula evaluation result: ' + JSON.stringify(result));
    return result;
    
  } catch (e) {
    Logger.log('[Code] checkFormulaFirst error: ' + e.message);
    return {
      canUseFormula: false,
      explanation: 'Error during formula evaluation: ' + e.message
    };
  }
}

/**
 * Get unified context for frontend
 * This is the primary way to get full data context
 */
function getUnifiedSelectionContext(command) {
  return buildUnifiedContext({ command: command });
}

// ============================================
// CONTACT FORM
// ============================================

/**
 * Submit contact form to backend
 * @param {string} name - Sender name
 * @param {string} email - Sender email
 * @param {string} message - Message content
 * @return {boolean} Success status
 * @throws {Error} If submission fails
 */
function submitContactForm(name, email, message) {
  var payload = {
    name: name,
    email: email,
    message: message
  };
  
  try {
    ApiClient.post('CONTACT', payload);
    return true;
  } catch (error) {
    console.error('Error submitting contact form:', error);
    throw new Error('Failed to submit contact form: ' + error.message);
  }
}

// ============================================
// DEPRECATED FUNCTIONS
// ============================================

/**
 * @deprecated Stratico API is no longer supported (January 2026)
 * Use ChatGPT(), Claude(), Groq(), or Gemini() instead.
 * @customfunction
 */
function Stratico(prompt, imageUrl, model) {
  return 'DEPRECATED: Stratico is no longer supported. Please use =ChatGPT(), =Claude(), =Groq(), or =Gemini() instead.';
}

/**
 * @deprecated Stratico API is no longer supported (January 2026)
 * Use DALLE() instead for image generation.
 * @customfunction
 */
function StraticoImage(prompt, model) {
  return 'DEPRECATED: StraticoImage is no longer supported. Please use =DALLE() instead.';
}
