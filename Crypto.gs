/**
 * @file Crypto.gs
 * @version 2.0.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.0.0 (2026-01-10): Initial modular extraction
 * 
 * ============================================
 * CRYPTO.gs - Encryption Utilities
 * ============================================
 * 
 * Handles API key encryption/decryption using AES.
 * Uses cCryptoGS library for cryptographic operations.
 * 
 * Usage:
 *   var encrypted = Crypto.encrypt(apiKey);
 *   var decrypted = Crypto.decrypt(encryptedKey);
 */

/**
 * CryptoJS library reference
 * Loaded from cCryptoGS Apps Script library
 */
var CryptoJS = cCryptoGS.CryptoJS;

/**
 * Crypto utilities object
 */
var Crypto = {
  
  /**
   * Get or create encryption salt
   * Salt is stored in Script Properties (shared across all users)
   * @return {string} Encryption salt
   * @private
   */
  _getSalt: function() {
    var salt = PropertiesService.getScriptProperties().getProperty('ENCRYPTION_SALT');
    
    if (!salt) {
      salt = Utilities.getUuid();
      PropertiesService.getScriptProperties().setProperty('ENCRYPTION_SALT', salt);
      Logger.log('New encryption salt generated');
    }
    
    return salt;
  },
  
  /**
   * Encrypt a string (typically an API key)
   * @param {string} plaintext - String to encrypt
   * @return {string} Encrypted string (base64)
   */
  encrypt: function(plaintext) {
    if (!plaintext) {
      return '';
    }
    
    var salt = this._getSalt();
    return CryptoJS.AES.encrypt(plaintext, salt).toString();
  },
  
  /**
   * Decrypt an encrypted string
   * @param {string} ciphertext - Encrypted string (base64)
   * @return {string} Decrypted string
   */
  decrypt: function(ciphertext) {
    if (!ciphertext) {
      return '';
    }
    
    var salt = this._getSalt();
    return CryptoJS.AES.decrypt(ciphertext, salt).toString(CryptoJS.enc.Utf8);
  },
  
  /**
   * Regenerate encryption salt
   * WARNING: This will invalidate all previously encrypted API keys!
   * Only use during initial setup or if salt is compromised.
   * @return {string} New salt
   */
  regenerateSalt: function() {
    var salt = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('ENCRYPTION_SALT', salt);
    Logger.log('Encryption salt regenerated: ' + salt);
    return salt;
  },
  
  /**
   * Check if salt exists
   * @return {boolean}
   */
  hasSalt: function() {
    return !!PropertiesService.getScriptProperties().getProperty('ENCRYPTION_SALT');
  },
  
  /**
   * Check if a string is already encrypted
   * Encrypted strings from CryptoJS AES start with "U2FsdGVkX1" (Base64 of "Salted__")
   * @param {string} text - String to check
   * @return {boolean} True if already encrypted
   */
  isEncrypted: function(text) {
    return text && text.indexOf('U2F') === 0;
  },
  
  /**
   * Ensure a string is encrypted for transit to backend
   * If already encrypted, returns as-is. If decrypted, encrypts it.
   * This handles the case where getUserSettings() cache returns encrypted keys.
   * @param {string} text - API key (may be encrypted or decrypted)
   * @return {string|null} Encrypted string or null if input is falsy
   */
  ensureEncrypted: function(text) {
    if (!text) return null;
    if (this.isEncrypted(text)) {
      return text;  // Already encrypted
    }
    return this.encrypt(text);  // Encrypt for transit
  }
};

// ============================================
// LEGACY ALIASES (for backward compatibility)
// ============================================

/**
 * @deprecated Use Crypto.encrypt() instead
 */
function encryptApiKey(apiKey) {
  return Crypto.encrypt(apiKey);
}

/**
 * @deprecated Use Crypto.decrypt() instead
 */
function decryptApiKey(encryptedApiKey) {
  return Crypto.decrypt(encryptedApiKey);
}

/**
 * @deprecated Use Crypto.regenerateSalt() instead
 */
function setEncryptionSalt() {
  return Crypto.regenerateSalt();
}

/**
 * @deprecated Use Crypto.regenerateSalt() instead
 */
function generateAndSetEncryptionSalt() {
  return Crypto.regenerateSalt();
}
