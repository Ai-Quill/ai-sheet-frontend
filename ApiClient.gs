/**
 * @file ApiClient.gs
 * @version 2.0.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.0.0 (2026-01-10): Initial centralized HTTP client
 * 
 * ============================================
 * APICLIENT.gs - HTTP Request Utilities
 * ============================================
 * 
 * Centralized HTTP client for all API calls.
 * Handles logging, error formatting, and common patterns.
 * 
 * Usage:
 *   var result = ApiClient.post('QUERY', payload);
 *   var data = ApiClient.get('MODELS');
 */

/**
 * API Client object
 * Provides methods for making HTTP requests to the backend
 */
var ApiClient = {
  
  /**
   * Make a GET request
   * @param {string} endpoint - Endpoint key from API_ENDPOINTS
   * @param {Object} [params] - Optional query parameters
   * @return {Object} Parsed JSON response
   * @throws {Error} If request fails
   */
  get: function(endpoint, params) {
    var url = Config.getApiUrl(endpoint);
    
    // Add query parameters
    if (params && Object.keys(params).length > 0) {
      var queryString = Object.keys(params)
        .filter(function(key) { return params[key] !== null && params[key] !== undefined; })
        .map(function(key) { 
          return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]); 
        })
        .join('&');
      
      if (queryString) {
        url += '?' + queryString;
      }
    }
    
    return this._request(url, {
      method: 'get',
      muteHttpExceptions: true
    });
  },
  
  /**
   * Make a POST request
   * @param {string} endpoint - Endpoint key from API_ENDPOINTS
   * @param {Object} payload - Request body
   * @return {Object} Parsed JSON response
   * @throws {Error} If request fails
   */
  post: function(endpoint, payload) {
    var url = Config.getApiUrl(endpoint);
    
    return this._request(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  },
  
  /**
   * Make a PUT request
   * @param {string} endpoint - Endpoint key from API_ENDPOINTS
   * @param {Object} payload - Request body
   * @return {Object} Parsed JSON response
   * @throws {Error} If request fails
   */
  put: function(endpoint, payload) {
    var url = Config.getApiUrl(endpoint);
    
    return this._request(url, {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  },
  
  /**
   * Make a DELETE request
   * @param {string} endpoint - Endpoint key from API_ENDPOINTS
   * @param {Object} [params] - Optional query parameters
   * @return {Object} Parsed JSON response
   * @throws {Error} If request fails
   */
  delete: function(endpoint, params) {
    var url = Config.getApiUrl(endpoint);
    
    // Add query parameters
    if (params && Object.keys(params).length > 0) {
      var queryString = Object.keys(params)
        .filter(function(key) { return params[key] !== null && params[key] !== undefined; })
        .map(function(key) { 
          return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]); 
        })
        .join('&');
      
      if (queryString) {
        url += '?' + queryString;
      }
    }
    
    return this._request(url, {
      method: 'delete',
      muteHttpExceptions: true
    });
  },
  
  /**
   * Internal request handler
   * @param {string} url - Full URL
   * @param {Object} options - UrlFetchApp options
   * @return {Object} Parsed JSON response
   * @throws {Error} If request fails
   * @private
   */
  _request: function(url, options) {
    // Log request in debug mode
    if (Config.shouldLogRequests()) {
      Logger.log('API Request: ' + options.method.toUpperCase() + ' ' + url);
      if (options.payload) {
        Logger.log('Payload: ' + options.payload);
      }
    }
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      var responseText = response.getContentText();
      var responseCode = response.getResponseCode();
      
      // Log response in debug mode
      if (Config.shouldLogResponses()) {
        Logger.log('Response (' + responseCode + '): ' + responseText);
      }
      
      // Parse JSON
      var result;
      try {
        result = JSON.parse(responseText);
      } catch (parseError) {
        throw new Error('Invalid JSON response: ' + responseText.substring(0, 200));
      }
      
      // Check for API error
      if (result.error) {
        throw new Error(result.error);
      }
      
      return result;
      
    } catch (error) {
      // Log error
      Logger.log('API Error: ' + error.toString());
      
      // Re-throw with context
      throw new Error('API request failed: ' + error.message);
    }
  },
  
  /**
   * Build URL with query parameters (utility method)
   * @param {string} endpoint - Endpoint key
   * @param {Object} params - Query parameters
   * @return {string} Full URL with query string
   */
  buildUrl: function(endpoint, params) {
    var url = Config.getApiUrl(endpoint);
    
    if (params && Object.keys(params).length > 0) {
      var queryString = Object.keys(params)
        .filter(function(key) { return params[key] !== null && params[key] !== undefined; })
        .map(function(key) { 
          return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]); 
        })
        .join('&');
      
      if (queryString) {
        url += '?' + queryString;
      }
    }
    
    return url;
  }
};
