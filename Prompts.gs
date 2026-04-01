/**
 * @file Prompts.gs
 * @version 2.0.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.0.0 (2026-01-10): Initial modular extraction
 * 
 * ============================================
 * PROMPTS.gs - Saved Prompt Management
 * ============================================
 * 
 * Handles CRUD operations for user-saved prompts.
 * 
 * Usage (from Sidebar/PromptManager):
 *   var prompts = getSavedPrompts();
 *   savePrompt(name, prompt, variables);
 *   updatePrompt(id, name, prompt, variables);
 *   deletePrompt(id);
 */

/**
 * Get all saved prompts for current user
 * @return {Array<Object>} Array of prompt objects
 * @throws {Error} If request fails
 */
function getSavedPrompts() {
  var userId = getUserId();
  
  try {
    var result = ApiClient.get('PROMPTS', { user_id: userId });
    return result;
  } catch (error) {
    console.error('Error fetching saved prompts:', error);
    throw new Error('Failed to fetch saved prompts: ' + error.message);
  }
}

/**
 * Save a new prompt
 * @param {string} name - Prompt name/title
 * @param {string} prompt - Prompt template text
 * @param {string} variables - Comma-separated variable names
 * @return {Object} Created prompt data
 * @throws {Error} If request fails
 */
function savePrompt(name, prompt, variables) {
  var userId = getUserId();
  
  if (Config.isDebug()) {
    console.log('Saving prompt for user:', userId);
  }
  
  try {
    var payload = {
      name: name,
      prompt: prompt,
      variables: variables,
      user_id: userId
    };
    
    var result = ApiClient.post('PROMPTS', payload);
    return result.data;
  } catch (error) {
    console.error('Error saving prompt:', error);
    throw new Error('Failed to save prompt: ' + error.message);
  }
}

/**
 * Update an existing prompt
 * @param {string} id - Prompt ID
 * @param {string} name - Prompt name/title
 * @param {string} prompt - Prompt template text
 * @param {string} variables - Comma-separated variable names
 * @return {Object} Updated prompt data
 * @throws {Error} If request fails
 */
function updatePrompt(id, name, prompt, variables) {
  var userId = getUserId();
  
  try {
    var payload = {
      id: id,
      name: name,
      prompt: prompt,
      variables: variables,
      user_id: userId
    };
    
    var result = ApiClient.put('PROMPTS', payload);
    return result.data;
  } catch (error) {
    console.error('Error updating prompt:', error);
    throw new Error('Failed to update prompt: ' + error.message);
  }
}

/**
 * Delete a prompt
 * @param {string} id - Prompt ID
 * @return {string} Success message
 * @throws {Error} If request fails
 */
function deletePrompt(id) {
  var userId = getUserId();
  
  try {
    var result = ApiClient.delete('PROMPTS', { id: id, user_id: userId });
    return result.message;
  } catch (error) {
    console.error('Error deleting prompt:', error);
    throw new Error('Failed to delete prompt: ' + error.message);
  }
}

/**
 * Open the Prompt Manager dialog
 * Called from menu or sidebar
 */
function openPromptManager() {
  var html = HtmlService.createHtmlOutputFromFile('PromptManager')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Prompt Manager');
}
