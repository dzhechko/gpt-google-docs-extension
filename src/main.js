/**
 * This function runs automatically when the document is opened
 * It creates the custom menu in Google Docs
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('DZ GPT Помощник')
    .addItem('Показать панель ассистента', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Операции с текстом')
      .addItem('Улучшить текст', 'enhanceSelectedText')
      .addItem('Сделать краткое содержание', 'summarizeSelectedText')
      .addItem('Исправить грамматику', 'fixGrammar'))
    .addSeparator()
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

/**
 * Shows the sidebar with the GPT interface
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui')
    .setTitle('GPT Assistant')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Gets the selected text from the document
 * @returns {string} The selected text or empty string if no selection
 */
function getSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Пожалуйста, выделите текст перед использованием.');
    return '';
  }
  
  const elements = selection.getRangeElements();
  return elements.map(element => {
    const text = element.getElement().asText();
    const startOffset = element.isPartial() ? element.getStartOffset() : 0;
    const endOffset = element.isPartial() ? 
      element.getEndOffsetInclusive() + 1 : 
      text.getText().length;
    return text.getText().substring(startOffset, endOffset);
  }).join(' ');
}

/**
 * Inserts text after the selected text or at cursor position
 * @param {string} text - The text to insert
 * @param {boolean} preserveSelection - Whether to keep the selected text
 */
function insertText(text, preserveSelection = false) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getRangeElements();
    // Get the last element of the selection
    const lastElement = elements[elements.length - 1];
    const textElement = lastElement.getElement().editAsText();
    
    // Calculate insertion point after the selection
    const insertionOffset = lastElement.isPartial() ? 
      lastElement.getEndOffsetInclusive() + 1 : 
      textElement.getText().length;
    
    // Insert two newlines and then the new text
    textElement.insertText(insertionOffset, '\n\n' + text);
    
  } else {
    const cursor = doc.getCursor();
    if (cursor) {
      const element = cursor.getElement();
      const position = cursor.getOffset();
      element.asText().insertText(position, text);
    } else {
      doc.getBody().appendParagraph(text);
    }
  }
}

/**
 * Enhances the selected text using GPT
 */
function enhanceSelectedText() {
  const selectedText = getSelectedText();
  if (!selectedText) return;
  
  const prompt = `Please enhance the following text while maintaining its core meaning. Make it more professional and engaging: "${selectedText}"`;
  const enhancedText = callOpenAI(prompt);
  insertText(enhancedText);
}

/**
 * Summarizes the selected text using GPT and adds summary below
 */
function summarizeSelectedText() {
  const selectedText = getSelectedText();
  if (!selectedText) return;
  
  const prompt = `Please provide a concise summary of the following text: "${selectedText}"`;
  const summary = '📝 Summary:\n' + callOpenAI(prompt);
  insertText(summary, true);
}

/**
 * Fixes grammar and style in the selected text using GPT
 */
function fixGrammar() {
  const selectedText = getSelectedText();
  if (!selectedText) return;
  
  const prompt = `Please fix any grammar and style issues in the following text: "${selectedText}"`;
  const correctedText = callOpenAI(prompt);
  insertText(correctedText);
}

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
 * Saves settings to Script Properties
 */
function saveSettings(settings) {
  // Validate settings
  if (!settings.baseUrl) throw new Error('URL API обязателен');
  if (!settings.model) throw new Error('Модель обязательна');
  if (settings.temperature < 0 || settings.temperature > 1) 
    throw new Error('Temperature должен быть от 0 до 1');
  if (settings.maxTokens < 150) 
    throw new Error('Max Tokens должен быть не менее 150');

  PropertiesService.getUserProperties()
    .setProperty('settings', JSON.stringify(settings));
}

/**
 * Shows the settings dialog
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(450)
    .setTitle('Настройки');
  DocumentApp.getUi().showModalDialog(html, 'Настройки');
}
