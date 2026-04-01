/**
 * @file SecureRequest.gs
 * @version 1.0.0
 * @created 2026-01-20
 * 
 * ============================================
 * SECURE REQUEST BUILDER
 * ============================================
 * 
 * Centralized service for building API request payloads
 * with properly encrypted API keys.
 * 
 * This ensures:
 * - API keys are always encrypted before transit
 * - Consistent payload structure across all API calls
 * - No decrypted keys are ever sent over the network
 * 
 * Dependencies:
 * - Crypto.gs     → Encryption utilities
 * - User.gs       → User settings (API keys)
 * - Config.gs     → Configuration
 * 
 * Usage:
 *   // Simple: just get encrypted key for a provider
 *   var encryptedKey = SecureRequest.getEncryptedKey('CHATGPT');
 *   
 *   // Build complete payload with encrypted key
 *   var payload = SecureRequest.buildPayload('GEMINI', {
 *     input: 'Hello world',
 *     specificModel: 'gemini-2.5-flash'
 *   });
 */

/**
 * SecureRequest - Centralized API key handling for outbound requests
 */
var SecureRequest = {
  
  /**
   * Get the encrypted API key for a specific provider
   * 
   * @param {string} provider - Provider name (CHATGPT, CLAUDE, GROQ, GEMINI, STRATICO)
   * @return {string} Encrypted API key ready for transit
   * @throws {Error} If no API key configured for provider
   */
  getEncryptedKey: function(provider) {
    var settings = getUserSettings();
    
    if (!settings || !settings[provider]) {
      throw new Error('No settings found for provider: ' + provider);
    }
    
    var apiKey = settings[provider].apiKey;
    
    if (!apiKey) {
      throw new Error('No API key configured for ' + provider + '. Please configure in Settings.');
    }
    
    // Use ensureEncrypted to handle cache returning encrypted keys
    return Crypto.ensureEncrypted(apiKey);
  },
  
  /**
   * Build a complete API payload with encrypted API key
   * 
   * Merges provided data with provider and encrypted key.
   * The encrypted key is named 'encryptedApiKey' to match backend expectations.
   * 
   * @param {string} provider - Provider name (CHATGPT, CLAUDE, GROQ, GEMINI, STRATICO)
   * @param {Object} data - Request data to include in payload
   * @return {Object} Complete payload with provider and encryptedApiKey
   * @throws {Error} If no API key configured for provider
   * 
   * @example
   * // Returns: { input: 'Hello', provider: 'GEMINI', encryptedApiKey: 'U2F...' }
   * SecureRequest.buildPayload('GEMINI', { input: 'Hello' });
   */
  buildPayload: function(provider, data) {
    var encryptedKey = this.getEncryptedKey(provider);
    
    // Merge data with auth fields
    return Object.assign({}, data, {
      provider: provider,
      encryptedApiKey: encryptedKey
    });
  },
  
  /**
   * Build payload with additional user context (email, userId)
   * 
   * For routes that need user identification in addition to auth.
   * 
   * @param {string} provider - Provider name
   * @param {Object} data - Request data
   * @return {Object} Payload with provider, encryptedApiKey, userEmail, userId
   */
  buildPayloadWithUser: function(provider, data) {
    var payload = this.buildPayload(provider, data);
    
    return Object.assign(payload, {
      userEmail: getUserEmail(),
      userId: getUserId()
    });
  },
  
  /**
   * Build payload for model-specific requests
   * 
   * Includes specific model selection if not provided.
   * 
   * @param {string} provider - Provider name
   * @param {Object} data - Request data
   * @param {string} [specificModel] - Optional specific model, uses default if not provided
   * @return {Object} Payload with specificModel included
   */
  buildPayloadWithModel: function(provider, data, specificModel) {
    var model = specificModel || Config.getDefaultModel(provider);
    
    return this.buildPayload(provider, Object.assign({}, data, {
      specificModel: model
    }));
  },
  
  /**
   * Validate that a provider has a configured API key
   * 
   * @param {string} provider - Provider name to check
   * @return {boolean} True if API key exists for provider
   */
  hasApiKey: function(provider) {
    try {
      var settings = getUserSettings();
      return !!(settings && settings[provider] && settings[provider].apiKey);
    } catch (e) {
      return false;
    }
  },
  
  /**
   * Get list of providers with configured API keys
   * 
   * @return {string[]} Array of provider names with API keys
   */
  getConfiguredProviders: function() {
    var settings = getUserSettings();
    var providers = ['CHATGPT', 'CLAUDE', 'GROQ', 'GEMINI', 'STRATICO'];
    
    return providers.filter(function(p) {
      return settings && settings[p] && settings[p].apiKey;
    });
  },
  
  /**
   * Build a payload for managed AI mode (no user API key required)
   * 
   * Uses platform-owned API keys. The backend will check managed credit balance
   * and route through pooled keys.
   * 
   * @param {Object} data - Request data to include in payload
   * @param {string} [specificModel] - Managed model ID (e.g., 'gpt-5-mini', 'gemini-2.5-flash')
   * @return {Object} Payload with managedMode flag and user identity
   * 
   * @example
   * // Returns: { input: 'Hello', managedMode: true, userEmail: '...', userId: '...', specificModel: 'gpt-5-mini' }
   * SecureRequest.buildManagedPayload({ input: 'Hello' }, 'gpt-5-mini');
   */
  buildManagedPayload: function(data, specificModel) {
    return Object.assign({}, data, {
      managedMode: true,
      userEmail: getUserEmail(),
      userId: getUserId(),
      specificModel: specificModel || undefined
    });
  },
  
  /**
   * Build a payload that works for either managed or BYOK mode.
   * 
   * If isManaged is true and a managed model is specified, uses managed mode.
   * Otherwise falls back to BYOK with the user's encrypted API key.
   * 
   * @param {string} provider - Provider name (for BYOK mode)
   * @param {Object} data - Request data
   * @param {boolean} isManaged - Whether to use managed mode
   * @param {string} [managedModelId] - Specific managed model ID
   * @return {Object} Payload ready for the backend
   */
  buildSmartPayload: function(provider, data, isManaged, managedModelId) {
    if (isManaged) {
      return this.buildManagedPayload(data, managedModelId);
    }
    return this.buildPayload(provider, data);
  }
};
