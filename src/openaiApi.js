/**
 * Configuration for OpenAI API
 * Replace 'your-openai-api-key' with your actual OpenAI API key
 */
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

/**
 * Gets the current settings from Script Properties
 */
function getSettings() {
  const properties = PropertiesService.getUserProperties();
  const settings = properties.getProperty('settings');
  return settings ? JSON.parse(settings) : {
    baseUrl: 'https://api.openai.com/v1/chat/completions',
    model: 'gpt-3.5-turbo',
    temperature: 0.7,
    maxTokens: 150
  };
}

/**
 * Calls OpenAI's API with the given prompt
 * @param {string} prompt - The user's input prompt
 * @returns {string} The generated text response
 */
function callOpenAI(prompt) {
  try {
    const settings = getSettings();
    const headers = {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    };
    
    const payload = {
      model: settings.model,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: settings.temperature,
      max_tokens: settings.maxTokens
    };

    const options = {
      method: 'POST',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(settings.baseUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      console.error('Error:', responseText);
      throw new Error(`API request failed with status ${responseCode}`);
    }

    const data = JSON.parse(responseText);
    return data.choices[0].message.content.trim();
    
  } catch (error) {
    console.error('Error calling OpenAI API:', error);
    return `Error: ${error.message}`;
  }
}

/**
 * Tests the OpenAI API connection
 * @returns {boolean} True if connection is successful
 */
function testOpenAIConnection() {
  try {
    const response = callOpenAI("Hello, please respond with 'Connection successful'");
    console.log('API Test Response:', response);
    return true;
  } catch (error) {
    console.error('API Test Failed:', error);
    return false;
  }
}
