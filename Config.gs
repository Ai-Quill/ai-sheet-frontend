/**
 * @file Config.gs
 * @version 2.3.0
 * @updated 2026-01-11
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.3.0 (2026-01-11): Added JOBS_STREAM endpoint for SSE real-time updates
 * - 2.2.0 (2026-01-10): Added Agent endpoints (parse, estimate, validate)
 * - 2.1.0 (2026-01-10): Added GET_USER_STATUS endpoint
 * - 2.0.0 (2026-01-09): Initial centralized config
 * 
 * ============================================
 * CONFIG.gs - Environment Configuration
 * ============================================
 * 
 * Centralized configuration for all environment settings.
 * Switch between LOCAL and PRODUCTION by changing CURRENT_ENV.
 * 
 * Usage:
 *   var url = Config.getBaseUrl() + '/api/query';
 *   var env = Config.getEnv();
 */

// ============================================
// ENVIRONMENT SETTINGS
// ============================================

/**
 * Environment type
 * @enum {string}
 */
var ENV_TYPE = {
  LOCAL: 'LOCAL',
  PRODUCTION: 'PRODUCTION'
};

/**
 * Current environment - CHANGE THIS TO SWITCH ENVIRONMENTS
 * @type {string}
 */
var CURRENT_ENV = ENV_TYPE.PRODUCTION;

/**
 * Environment-specific configuration
 */
var ENV_CONFIG = {
  LOCAL: {
    baseUrl: 'http://localhost:3000',
    debug: true,
    logRequests: true,
    logResponses: true
  },
  PRODUCTION: {
    baseUrl: 'https://aisheet.vercel.app',
    debug: false,
    logRequests: false,
    logResponses: false
  }
};

// ============================================
// API ENDPOINTS
// ============================================

/**
 * API endpoint paths (relative to base URL)
 */
var API_ENDPOINTS = {
  // Core AI
  QUERY: '/api/query',
  GENERATE_IMAGE: '/api/generate-image',
  MODELS: '/api/models',
  
  // User management
  GET_OR_CREATE_USER: '/api/get-or-create-user',
  GET_USER_SETTINGS: '/api/get-user-settings',
  SAVE_ALL_SETTINGS: '/api/save-all-settings',
  
  // Prompts
  PROMPTS: '/api/prompts',
  
  // Jobs (bulk processing)
  JOBS: '/api/jobs',
  JOBS_STREAM: '/api/jobs/stream',  // SSE real-time updates
  JOBS_TRIGGER: '/api/jobs/trigger',  // Lightweight worker trigger
  
  // Agent (AI-powered command processing)
  AGENT_PARSE: '/api/agent/parse',
  AGENT_PARSE_CHAIN: '/api/agent/parse-chain',  // Multi-step task parsing
  AGENT_ESTIMATE: '/api/agent/estimate',
  AGENT_VALIDATE: '/api/agent/validate',
  CLASSIFY_OPTIONS: '/api/agent/classify-options',
  AGENT_CONVERSATION: '/api/agent/conversation',  // Conversation persistence
  AGENT_SUGGESTIONS: '/api/agent/suggestions',    // Dynamic suggestions API
  AGENT_LEARNING: '/api/agent/learning',          // Self-learning feedback & patterns
  
  // Subscription & Usage
  USAGE_CHECK: '/api/usage/check',
  STRIPE_CHECKOUT: '/api/stripe/checkout',
  STRIPE_PORTAL: '/api/stripe/portal',
  
  // Analytics
  ANALYTICS_LOG: '/api/analytics/log',
  
  // Contact
  CONTACT: '/api/contact'
};

// ============================================
// MODEL DEFAULTS
// ============================================

/**
 * Default models per provider - February 2026 cost-optimized
 * See docs/research/models.md for full pricing table
 * Groq: Llama 4 Scout supports structured outputs (JSON schema mode)
 */
var DEFAULT_MODELS = {
  CHATGPT: 'gpt-5-mini',        // $0.25/$2.00 per MTok - best value
  CLAUDE: 'claude-haiku-4-5',   // $1.00/$5.00 per MTok - fastest Claude
  GROQ: 'meta-llama/llama-4-scout-17b-16e-instruct', // $0.11/$0.34 per MTok - supports structured outputs
  GEMINI: 'gemini-2.5-flash'    // $0.075/$0.30 per MTok - cheapest
};

/**
 * Supported AI providers
 */
var SUPPORTED_PROVIDERS = ['CHATGPT', 'CLAUDE', 'GROQ', 'GEMINI'];

// ============================================
// CONFIG OBJECT
// ============================================

/**
 * Configuration accessor object
 * Provides methods to access configuration values
 */
var Config = {
  
  /**
   * Get current environment
   * @return {string} Current environment name
   */
  getEnv: function() {
    return CURRENT_ENV;
  },
  
  /**
   * Check if running in production
   * @return {boolean}
   */
  isProduction: function() {
    return CURRENT_ENV === ENV_TYPE.PRODUCTION;
  },
  
  /**
   * Check if running locally
   * @return {boolean}
   */
  isLocal: function() {
    return CURRENT_ENV === ENV_TYPE.LOCAL;
  },
  
  /**
   * Get base URL for current environment
   * @return {string} Base URL (no trailing slash)
   */
  getBaseUrl: function() {
    return ENV_CONFIG[CURRENT_ENV].baseUrl;
  },
  
  /**
   * Get full API URL for an endpoint
   * @param {string} endpoint - Endpoint key from API_ENDPOINTS
   * @return {string} Full URL
   */
  getApiUrl: function(endpoint) {
    var path = API_ENDPOINTS[endpoint];
    if (!path) {
      throw new Error('Unknown endpoint: ' + endpoint);
    }
    return this.getBaseUrl() + path;
  },
  
  /**
   * Check if debug mode is enabled
   * @return {boolean}
   */
  isDebug: function() {
    return ENV_CONFIG[CURRENT_ENV].debug;
  },
  
  /**
   * Check if request logging is enabled
   * @return {boolean}
   */
  shouldLogRequests: function() {
    return ENV_CONFIG[CURRENT_ENV].logRequests;
  },
  
  /**
   * Check if response logging is enabled
   * @return {boolean}
   */
  shouldLogResponses: function() {
    return ENV_CONFIG[CURRENT_ENV].logResponses;
  },
  
  /**
   * Get default model for a provider
   * @param {string} provider - Provider name (CHATGPT, CLAUDE, etc.)
   * @return {string|null} Default model or null if unknown provider
   */
  getDefaultModel: function(provider) {
    return DEFAULT_MODELS[provider] || null;
  },
  
  /**
   * Get all supported providers
   * @return {Array<string>} List of provider names
   */
  getSupportedProviders: function() {
    return SUPPORTED_PROVIDERS.slice(); // Return copy to prevent mutation
  },
  
  /**
   * Log configuration summary (for debugging)
   */
  logConfig: function() {
    Logger.log('=== AISheeter Configuration ===');
    Logger.log('Environment: ' + this.getEnv());
    Logger.log('Base URL: ' + this.getBaseUrl());
    Logger.log('Debug Mode: ' + this.isDebug());
    Logger.log('================================');
  }
};

// ============================================
// INITIALIZATION
// ============================================

/**
 * Log configuration on script load (only in debug mode)
 */
(function() {
  if (Config.isDebug()) {
    Config.logConfig();
  }
})();
