var CryptoJS = cCryptoGS.CryptoJS;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AISheeter - Any LLM in One Google Sheets™')
    .addItem('Show Sidebar', 'showSidebar')
    // Remove the line below if it exists
    // .addItem('Test API Connection', 'testAPIConnection')
    .addToUi();
Sheeter
  // Automatically show the sidebar when the spreadsheet is opened
  showSidebar();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI Sheet - Any LLM')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getUserEmail() {
  var email = PropertiesService.getUserProperties().getProperty('userEmail');
  if (!email) {
    email = Session.getActiveUser().getEmail();
    setUserEmail(email);
  }
  return email;
}

function setUserEmail(email) {
  PropertiesService.getUserProperties().setProperty('userEmail', email);
}

function logCreditUsage(creditsUsed) {
  var userProperties = PropertiesService.getUserProperties();
  var logs = userProperties.getProperty('creditUsageLogs') || '';
  logs += `Credits used: ${creditsUsed.toFixed(4)}\n`;
  userProperties.setProperty('creditUsageLogs', logs);
}

function AIQuery(model, input, specificModel) {
  var userEmail = getUserEmail();
  var userSettings = getUserSettings();
  var encryptedApiKey = encryptApiKey(userSettings[model].apiKey);
  if (!specificModel) {
    switch(model) {
      case 'CHATGPT':
        specificModel = 'gpt-4o';
        break;
      case 'CLAUDE':
        specificModel = 'claude-3-sonnet-20240229';
        break;
      case 'GROQ':
        specificModel = 'llama-3.1-8b-instant';
        break;
      case 'GEMINI':
        specificModel = 'gemini-1.5-flash';
        break;
      default:
        // Optional: handle unknown model or leave specificModel as undefined
        break;
    }
  }
  var url = 'https://aisheet.vercel.app/api/query';
  var payload = {
    model: model,
    input: input,
    userEmail: userEmail,
    specificModel: specificModel,
    encryptedApiKey: encryptedApiKey  // This is already encrypted
  };
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    Logger.log('Request URL: ' + url);
    Logger.log('Request Payload: ' + JSON.stringify(payload));
    
    var response = UrlFetchApp.fetch(url, options);
    var responseText = response.getContentText();
    Logger.log('Response: ' + responseText);
    
    var result = JSON.parse(responseText);
    if (result.error) {
      return "Error: " + result.error;
    }
    logCreditUsage(result.creditsUsed);
    return result.result;
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return "Error: " + error.toString();
  }
}

/**
 * Custom function to call ChatGPT model
 * @param {string} input - The input prompt for the model
 * @param {string} model - The specific model to use (optional)
 * @return {string} - The result from the AI model
 * @customfunction
 */
function ChatGPT(input, model) {
  return AIQuery('CHATGPT', input, model);
}

/**
 * Custom function to call Claude model
 * @param {string} input - The input prompt for the model
 * @param {string} model - The specific model to use (optional)
 * @return {string} - The result from the AI model
 * @customfunction
 */
function Claude(input, model) {
  return AIQuery('CLAUDE', input, model);
}

/**
 * Custom function to call Groq model
 * @param {string} input - The input prompt for the model
 * @param {string} model - The specific model to use (optional)
 * @return {string} - The result from the AI model
 * @customfunction
 */
function Groq(input, model) {
  return AIQuery('GROQ', input, model);
}

/**
 * Custom function to call Gemini model
 * @param {string} input - The input prompt for the model
 * @param {string} model - The specific model to use (optional)
 * @return {string} - The result from the AI model
 * @customfunction
 */
function Gemini(input, model) {
  return AIQuery('GEMINI', input, model);
}

// Update the saveAllSettings function
function saveAllSettings(settings) {
  try {
    console.log('Received settings:', JSON.stringify(settings));
    
    // Validate settings
    if (!settings || typeof settings !== 'object') {
      throw new Error('Invalid settings object');
    }

    const models = ['CHATGPT', 'CLAUDE', 'GROQ', 'GEMINI'];
    models.forEach(model => {
      if (!settings[model] || typeof settings[model] !== 'object') {
        throw new Error(`Invalid settings for ${model}`);
      }
      if (typeof settings[model].apiKey !== 'string') {
        throw new Error(`Invalid API key for ${model}`);
      }
      if (typeof settings[model].defaultModel !== 'string') {
        throw new Error(`Invalid default model for ${model}`);
      }
      // Encrypt the API key
      settings[model].apiKey = encryptApiKey(settings[model].apiKey);
    });

    // Get the user's email
    var userEmail = getUserEmail();
    
    // Prepare the payload in the format the server expects
    var payload = {
      userEmail: userEmail,
      settings: settings
    };

    // Save settings
    var url = 'https://aisheet.vercel.app/api/save-all-settings';
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(result.error);
    }

    console.log('Settings saved successfully');
    return 'Settings saved successfully';
  } catch (error) {
    console.error('Error in saveAllSettings:', error);
    throw new Error('Failed to save settings: ' + error.message);
  }
}

// Update the getUserSettings function
function getUserSettings() {
  var userEmail = getUserEmail();
  
  var url = 'https://aisheet.vercel.app/api/get-user-settings?userEmail=' + encodeURIComponent(userEmail);
  
  try {
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());
    var settings = result.settings;

    // Decrypt the API keys to show to the user
    Object.keys(settings).forEach(model => {
      if (settings[model].apiKey) {
        settings[model].apiKey = decryptApiKey(settings[model].apiKey);
      }
    });

    return settings;
  } catch (error) {
    throw new Error("Error fetching user settings: " + error.toString());
  }
}

function getCreditUsageLogs() {
  return PropertiesService.getUserProperties().getProperty('creditUsageLogs') || '';
}

function saveApiKey(apiKey) {
    // Logic to save the API key
    // This might involve saving to a user properties, or sending to a server
    PropertiesService.getUserProperties().setProperty('apiKey', apiKey);
}

function fetchModels() {
  var url = 'https://aisheet.vercel.app/api/models';
  try {
    var response = UrlFetchApp.fetch(url);
    var models = JSON.parse(response.getContentText());
    console.log('Fetched models:', models);  // Add this line for debugging
    return models;
  } catch (error) {
    console.error('Error fetching models:', error);
    throw error;  // Rethrow the error so it can be caught by the failure handler
  }
}

// Update the encryptApiKey function
function encryptApiKey(apiKey) {
  var salt = PropertiesService.getScriptProperties().getProperty('ENCRYPTION_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('ENCRYPTION_SALT', salt);
  }
  return CryptoJS.AES.encrypt(apiKey, salt).toString();
}

// Update the decryptApiKey function
function decryptApiKey(encryptedApiKey) {
  var salt = PropertiesService.getScriptProperties().getProperty('ENCRYPTION_SALT');
  return CryptoJS.AES.decrypt(encryptedApiKey, salt).toString(CryptoJS.enc.Utf8);
}

function setEncryptionSalt() {
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ENCRYPTION_SALT', salt);
  Logger.log('Encryption salt set: ' + salt);
}

function generateAndSetEncryptionSalt() {
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ENCRYPTION_SALT', salt);
  Logger.log('New encryption salt generated and set: ' + salt);
  return salt;
}