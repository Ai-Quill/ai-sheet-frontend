function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Assistant')
    .addItem('Manage API Keys', 'showApiKeyManager')
    .addToUi();
}

function showApiKeyManager() {
  var html = HtmlService.createHtmlOutputFromFile('ApiKeyManager')
    .setWidth(600)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage API Keys');
}

function AIQuery(model, input) {
  var userId = PropertiesService.getUserProperties().getProperty('USER_ID');
  if (!userId) {
    return "Please set up your user ID in the AI Assistant menu.";
  }
  
  var url = 'https://your-vercel-deployment-url.vercel.app/api/query';
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({
      model: model,
      input: input,
      userId: userId
    })
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    return result.result;
  } catch (error) {
    return "Error: " + error.toString();
  }
}

function ChatGPT(input) { return AIQuery('CHATGPT', input); }
function Claude(input) { return AIQuery('CLAUDE', input); }
function Groq(input) { return AIQuery('GROQ', input); }
function Gemini(input) { return AIQuery('GEMINI', input); }

function saveApiKey(model, key) {
  PropertiesService.getUserProperties().setProperty(model + '_API_KEY', key);
  return "API key saved for " + model;
}

// Add this new function to save the user ID
function saveUserId(userId) {
  PropertiesService.getUserProperties().setProperty('USER_ID', userId);
  return "User ID saved";
}