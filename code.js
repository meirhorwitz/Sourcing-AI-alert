// =================================================================================
// --- MASTER CONFIGURATION ---
// =================================================================================

// --- GENERAL CONFIGURATION ---
const OPENAI_API_KEY_PROPERTY_NAME = 'OPENAI_API_KEY';
const TARGET_SHEET_NAME = "Sheet1";
const QUERY_SHEET_NAME = "cool"; // BigQuery data source sheet

// Default recipients - will be overridden by settings
const DEFAULT_TARGET_EMAIL_FOR_NOTIFICATIONS = ["meir.horwitz@fiverr.com"];
const DEFAULT_CONDITIONAL_SELLER_RECIPIENT = "meir.horwitz@fiverr.com";

// --- OpenAI API CONFIGURATION ---
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// --- BATCH AI PROCESSING COLUMN CONFIGURATION (uses header names) ---
const AI_PROMPT_COLUMN_NAME = "prompt_for_ai";
const AI_INPUT_CONVO_COLUMN_NAME = "conversation_messages";
const AI_INPUT_TIME_COLUMN_NAME = "current_time_info";
const AI_OUTPUT_STATUS_COLUMN_NAME = "status";
const AI_OUTPUT_ATTENTION_COLUMN_NAME = "need attention?";
const AI_OUTPUT_WHY_COLUMN_NAME = "why";
const AI_OUTPUT_LAST_MSG_COLUMN_NAME = "last_message_summary";
const AI_PROCESSING_STATUS_COLUMN_NAME = "AI Processing Status";
const AI_LAST_PROCESSED_TIME_COLUMN_NAME = "Last Processed Time Info";
const LAST_MESSAGE_COUNT_COLUMN_NAME = "Last Message Count";

// --- EMAIL NOTIFICATION CONFIGURATION ---
const PROCESSED_CONVERSATIONS_PROPERTY_KEY = 'PROCESSED_ATTENTION_CONVERSATIONS';

// =================================================================================
// --- SETTINGS MANAGEMENT ---
// =================================================================================
function getSettings() {
  const props = PropertiesService.getScriptProperties();
  const savedSettings = props.getProperty('SCRIPT_SETTINGS');
  
  // Always get the API key from the property, not from saved settings
  const apiKey = props.getProperty(OPENAI_API_KEY_PROPERTY_NAME) || '';
  
  const defaultSettings = {
    model: 'gpt-4o-mini',
    temperature: 0.2,
    maxTokens: 300,
    prompt: getDefaultPrompt(),
    emailRecipients: DEFAULT_TARGET_EMAIL_FOR_NOTIFICATIONS,
    conditionalRecipient: DEFAULT_CONDITIONAL_SELLER_RECIPIENT,
    notificationsEnabled: true,
    triggers: {
      dataUpdate: { enabled: true, time: '10:00' },
      aiProcessing: { enabled: true, interval: 15 }
    }
  };
  
  if (savedSettings) {
    try {
      const parsed = JSON.parse(savedSettings);
      // Don't include openaiApiKey from saved settings
      delete parsed.openaiApiKey;
      return { ...defaultSettings, ...parsed, openaiApiKey: apiKey };
    } catch (e) {
      return { ...defaultSettings, openaiApiKey: apiKey };
    }
  }
  
  return { ...defaultSettings, openaiApiKey: apiKey };
}

function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  
  // Remove API key from settings before saving (it's managed separately)
  const settingsToSave = { ...settings };
  delete settingsToSave.openaiApiKey;
  
  props.setProperty('SCRIPT_SETTINGS', JSON.stringify(settingsToSave));
  
  // Update triggers based on settings
  updateTriggers(settings.triggers);
}

function getDefaultPrompt() {
  return `Analyze this conversation and determine:
1. The current status of the conversation
2. Whether it needs attention (yes/no)
3. Why it needs attention (if applicable)
4. A brief summary of the last message

Format your response as:
Status: [status]
Needs attention: [yes/no]
Last message: [summary]
Why: [reason if needs attention]`;
}

function getAvailableOpenAIModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY_NAME);
  if (!apiKey) return [];
  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
      headers: { Authorization: 'Bearer ' + apiKey }
    });
    const data = JSON.parse(response.getContentText());
    return data.data.map(m => m.id);
  } catch (e) {
    Logger.log('Failed to fetch models: ' + e);
    return [];
  }
}

// =================================================================================
// --- DATA SYNC FROM BIGQUERY SHEET ---
// =================================================================================
function syncDataFromBigQuery() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("Sync skipped: Another instance is already running.");
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(QUERY_SHEET_NAME);
    const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sourceSheet || !targetSheet) {
      Logger.log("Error: Required sheets not found.");
      return;
    }
    
    // =======================================================
    // --- NEW LINES TO REFRESH THE BIGQUERY CONNECTED SHEET ---
    Logger.log("Starting refresh of the Connected Sheet...");
    sourceSheet.refreshData(); 
    SpreadsheetApp.flush(); // Waits for the refresh to complete
    Logger.log("Refresh complete. Reading data...");
    // =======================================================

    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length < 2) {
      Logger.log("Sync skipped: The 'cool' sheet is empty after refresh.");
      return;
    }
    
    const sourceHeaders = sourceData[0];
    
    let targetData = targetSheet.getDataRange().getValues();
    let targetHeaders = targetData[0] || [];
    
    const aiColumns = [
      AI_PROMPT_COLUMN_NAME,
      AI_OUTPUT_STATUS_COLUMN_NAME,
      AI_OUTPUT_ATTENTION_COLUMN_NAME,
      AI_OUTPUT_WHY_COLUMN_NAME,
      AI_OUTPUT_LAST_MSG_COLUMN_NAME,
      AI_PROCESSING_STATUS_COLUMN_NAME,
      AI_LAST_PROCESSED_TIME_COLUMN_NAME,
      LAST_MESSAGE_COUNT_COLUMN_NAME
    ];
    
    let completeHeaders = [...sourceHeaders];
    aiColumns.forEach(col => {
      if (!sourceHeaders.includes(col) && !completeHeaders.includes(col)) {
        completeHeaders.push(col);
      }
    });
    
    const existingConversations = new Map();
    const conversationIdIndex = targetHeaders.indexOf('conversation_id');
    
    if (conversationIdIndex !== -1 && targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const convId = targetData[i][conversationIdIndex];
        if (convId) {
          existingConversations.set(String(convId), {
            rowIndex: i,
            data: targetData[i]
          });
        }
      }
    }
    
    const newData = [completeHeaders];
    const conversationsToProcess = [];
    
    for (let i = 1; i < sourceData.length; i++) {
      const sourceRow = sourceData[i];
      const convId = sourceRow[sourceHeaders.indexOf('conversation_id')];
      const currentMessageCount = sourceRow[sourceHeaders.indexOf('inbox_message_count')] || 0;
      
      let newRow = new Array(completeHeaders.length).fill('');
      
      sourceHeaders.forEach((header, idx) => {
        const targetIdx = completeHeaders.indexOf(header);
        if (targetIdx !== -1) {
          newRow[targetIdx] = sourceRow[idx];
        }
      });
      
      const existing = existingConversations.get(String(convId));
      
      if (existing) {
        aiColumns.forEach(col => {
          const sourceIdx = targetHeaders.indexOf(col);
          const targetIdx = completeHeaders.indexOf(col);
          if (sourceIdx !== -1 && targetIdx !== -1 && existing.data[sourceIdx]) {
            newRow[targetIdx] = existing.data[sourceIdx];
          }
        });
        
        const lastMessageCountIdx = targetHeaders.indexOf(LAST_MESSAGE_COUNT_COLUMN_NAME);
        const lastMessageCount = lastMessageCountIdx !== -1 ? existing.data[lastMessageCountIdx] : 0;
        
        if (currentMessageCount !== lastMessageCount) {
          conversationsToProcess.push({
            conversationId: convId,
            rowIndex: newData.length
          });
          
          const newLastMsgCountIdx = completeHeaders.indexOf(LAST_MESSAGE_COUNT_COLUMN_NAME);
          if (newLastMsgCountIdx !== -1) {
            newRow[newLastMsgCountIdx] = currentMessageCount;
          }
        }
      } else {
        conversationsToProcess.push({
          conversationId: convId,
          rowIndex: newData.length
        });
        
        const newLastMsgCountIdx = completeHeaders.indexOf(LAST_MESSAGE_COUNT_COLUMN_NAME);
        if (newLastMsgCountIdx !== -1) {
          newRow[newLastMsgCountIdx] = currentMessageCount;
        }
      }
      
      const promptIdx = completeHeaders.indexOf(AI_PROMPT_COLUMN_NAME);
      if (promptIdx !== -1 && !newRow[promptIdx]) {
        newRow[promptIdx] = getSettings().prompt;
      }
      
      newData.push(newRow);
    }
    
    // Only clear the sheet if there is new data to write
    if (newData.length > 1) { 
      targetSheet.clear();
      targetSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      Logger.log(`Sync successful. Updated Sheet1 with ${newData.length - 1} rows of data.`);
    } else {
      Logger.log("Sync skipped: No new data was processed from the 'cool' sheet, so 'Sheet1' was left untouched.");
    }
    
    PropertiesService.getScriptProperties().setProperty('CONVERSATIONS_TO_PROCESS', JSON.stringify(conversationsToProcess));
    
  } catch (e) {
    Logger.log("Error in syncDataFromBigQuery: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// =================================================================================
// --- SMART AI PROCESSING (Only New/Changed) ---
// =================================================================================
function processNewAndChangedConversations() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("AI processing skipped: Another instance is already running.");
    return;
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const toProcessJson = props.getProperty('CONVERSATIONS_TO_PROCESS');
    
    if (!toProcessJson) {
      Logger.log("No conversations marked for processing.");
      return;
    }
    
    const conversationsToProcess = JSON.parse(toProcessJson);
    if (conversationsToProcess.length === 0) {
      Logger.log("No conversations to process.");
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) return;
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0].map(h => String(h).trim().toLowerCase());
    
    const colIdx = {
      prompt: headers.indexOf(AI_PROMPT_COLUMN_NAME.toLowerCase()),
      convoInput: headers.indexOf(AI_INPUT_CONVO_COLUMN_NAME.toLowerCase()),
      timeInput: headers.indexOf(AI_INPUT_TIME_COLUMN_NAME.toLowerCase()),
      statusOut: headers.indexOf(AI_OUTPUT_STATUS_COLUMN_NAME.toLowerCase()),
      attentionOut: headers.indexOf(AI_OUTPUT_ATTENTION_COLUMN_NAME.toLowerCase()),
      whyOut: headers.indexOf(AI_OUTPUT_WHY_COLUMN_NAME.toLowerCase()),
      lastMsgOut: headers.indexOf(AI_OUTPUT_LAST_MSG_COLUMN_NAME.toLowerCase()),
      processingStatus: headers.indexOf(AI_PROCESSING_STATUS_COLUMN_NAME.toLowerCase()),
      lastProcessedTime: headers.indexOf(AI_LAST_PROCESSED_TIME_COLUMN_NAME.toLowerCase()),
    };
    
    let processedCount = 0;
    
    for (const item of conversationsToProcess) {
      const rowIndex = item.rowIndex - 1; // Adjust for 0-based index
      if (rowIndex >= values.length) continue;
      
      const rowData = values[rowIndex];
      const prompt = rowData[colIdx.prompt];
      const conversation = rowData[colIdx.convoInput];
      
      if (!prompt || !conversation) continue;
      
      Logger.log(`Processing conversation ${item.conversationId} at row ${item.rowIndex}`);
      
      // Update status to processing
      sheet.getRange(item.rowIndex, colIdx.processingStatus + 1).setValue("Processing...");
      
      const aiResult = getAiAnalysisForRow(rowData, colIdx);
      
      // Update results
      if (aiResult.error) {
        sheet.getRange(item.rowIndex, colIdx.processingStatus + 1).setValue(`Error: ${aiResult.error}`);
      } else {
        const updateRanges = [
          { row: item.rowIndex, col: colIdx.statusOut + 1, value: aiResult.status },
          { row: item.rowIndex, col: colIdx.attentionOut + 1, value: aiResult.needsAttention },
          { row: item.rowIndex, col: colIdx.whyOut + 1, value: aiResult.why },
          { row: item.rowIndex, col: colIdx.lastMsgOut + 1, value: aiResult.lastMessage },
          { row: item.rowIndex, col: colIdx.processingStatus + 1, value: aiResult.processingStatus },
          { row: item.rowIndex, col: colIdx.lastProcessedTime + 1, value: rowData[colIdx.timeInput] }
        ];
        
        updateRanges.forEach(update => {
          if (update.col > 0) {
            sheet.getRange(update.row, update.col).setValue(update.value);
          }
        });
      }
      
      processedCount++;
      
      // Add small delay to avoid rate limits
      Utilities.sleep(100);
    }
    
    // Clear the processing queue
    props.deleteProperty('CONVERSATIONS_TO_PROCESS');
    
    Logger.log(`AI processing complete. Processed ${processedCount} conversations.`);
    
  } catch (e) {
    Logger.log("Error in processNewAndChangedConversations: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// =================================================================================
// --- AI ANALYSIS FUNCTION ---
// =================================================================================
function getAiAnalysisForRow(rowData, colIdx) {
  const settings = getSettings();
  const apiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY_NAME);
  
  if (!apiKey) return { error: "API key not set in script properties" };
  
  const fullPromptContent = `${rowData[colIdx.prompt]}\n\n${rowData[colIdx.convoInput]}\n\n${rowData[colIdx.timeInput]}`;
  
  const payload = {
    model: settings.model,
    messages: [{ role: "user", content: fullPromptContent }],
    max_tokens: settings.maxTokens,
    temperature: settings.temperature
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode !== 200) return { error: `API Error ${responseCode}: ${responseBody}` };
    
    const jsonResponse = JSON.parse(responseBody);
    const rawResponse = jsonResponse.choices?.[0]?.message?.content.trim() || "";
    
    if (!rawResponse) return { error: "Empty AI response" };
    
    const responseLines = rawResponse.split('\n').filter(line => line.trim() !== '');
    const aiStatus = responseLines[0]?.replace(/^Status:\s*/i, '').trim() || "N/A";
    const needsAttention = responseLines[1]?.replace(/^Needs attention:\s*/i, '').trim() || "N/A";
    const lastMessage = responseLines[2]?.replace(/^Last message:\s*/i, '').trim() || "N/A";
    const why = responseLines[3]?.replace(/^Why:\s*/i, '').trim() || "N/A";
    
    let processingStatus = "Processed";
    let warning = null;
    
    if (needsAttention.toLowerCase() !== 'yes' && needsAttention.toLowerCase() !== 'no') {
      processingStatus = "Processed with Warning";
      warning = "AI response for 'need attention?' was not 'yes' or 'no'";
    }
    
    return {
      error: null,
      warning,
      rawResponse,
      status: aiStatus,
      needsAttention,
      why,
      lastMessage,
      processingStatus
    };
  } catch (e) {
    return { error: "Script execution error: " + e.message };
  }
}

// =================================================================================
// --- EMAIL NOTIFICATION FUNCTIONS ---
// =================================================================================
function sendManualNotifications(conversationIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet not found" };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = {};
  
  ['conversation_id', 'buyer_name', 'seller_name', 'conversation_link', 'status', 'why', 'expert_email'].forEach(col => {
    colIdx[col] = headers.indexOf(col);
  });
  
  const alerts = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const convId = row[colIdx.conversation_id];
    
    if (conversationIds.includes(convId)) {
      alerts.push({
        conversationId: convId,
        buyerName: row[colIdx.buyer_name] || "Unknown Buyer",
        sellerName: row[colIdx.seller_name] || "Unknown Seller",
        conversationLink: row[colIdx.conversation_link] || "#",
        aiStatus: row[colIdx.status] || "Attention Needed",
        whyNeedsAttention: row[colIdx.why] || "Manual notification requested",
        expertEmail: row[colIdx.expert_email]
      });
    }
  }
  
  if (alerts.length === 0) {
    return { success: false, message: "No conversations found" };
  }

  const settings = getSettings();
  if (!settings.notificationsEnabled) {
    return { success: false, message: "Notifications are disabled" };
  }

  const recipients = settings.emailRecipients.join(',');
  const subject = `Manual Alert: ${alerts.length} conversation${alerts.length > 1 ? 's' : ''} need attention`;

  const grouped = {};
  alerts.forEach(a => {
    const email = a.expertEmail || recipients;
    if (!grouped[email]) grouped[email] = [];
    grouped[email].push(a);
  });

  Object.keys(grouped).forEach(email => {
    const htmlBody = buildManualNotificationHtml(grouped[email]);
    MailApp.sendEmail({
      to: email,
      cc: recipients,
      subject: subject,
      htmlBody: htmlBody
    });
  });

  return { success: true, message: `Notification sent for ${alerts.length} conversations` };
}

function buildManualNotificationHtml(alerts) {
  let html = `
    <html>
    <head>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; }
        .container { max-width: 700px; margin: 20px auto; }
        .alert { padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 20px; background: #fff; }
        .alert h3 { color: #c82333; margin-top: 0; }
        .button { display: inline-block; background: #007bff; color: white; padding: 12px 28px; text-decoration: none; border-radius: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Manual Alert Notification</h1>
        <p>The following conversations were manually flagged for attention:</p>
  `;
  
  alerts.forEach((alert, index) => {
    html += `
      <div class="alert">
        <h3>${index + 1}. ${alert.aiStatus} (Buyer: ${alert.buyerName})</h3>
        <p><strong>Reason:</strong> ${alert.whyNeedsAttention}</p>
        <p><strong>Seller:</strong> ${alert.sellerName}</p>
        <a href="${alert.conversationLink}" class="button">Open Conversation</a>
      </div>
    `;
  });
  
  html += `
      </div>
    </body>
    </html>
  `;
  
  return html;
}

// =================================================================================
// --- EXPORT FUNCTIONS ---
// =================================================================================
function exportConversations(conversationIds, format = 'csv') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const convIdIdx = headers.indexOf('conversation_id');
  
  const exportData = [headers];
  
  for (let i = 1; i < data.length; i++) {
    if (conversationIds.includes(data[i][convIdIdx])) {
      exportData.push(data[i]);
    }
  }
  
  if (format === 'csv') {
    return exportData.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
  } else if (format === 'json') {
    const jsonData = [];
    for (let i = 1; i < exportData.length; i++) {
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = exportData[i][idx];
      });
      jsonData.push(row);
    }
    return JSON.stringify(jsonData, null, 2);
  }
  
  return null;
}

// =================================================================================
// --- TRIGGER MANAGEMENT ---
// =================================================================================
function updateTriggers(triggerSettings) {
  // Remove all existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  const timezone = Session.getScriptTimeZone();
  
  // Data update trigger (daily at specified time)
  if (triggerSettings.dataUpdate.enabled) {
    const hour = parseInt(triggerSettings.dataUpdate.time.split(':')[0]);
    ScriptApp.newTrigger('runDailyDataUpdate')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .inTimezone(timezone)
      .create();
  }
  
  // AI processing trigger (every N minutes)
  if (triggerSettings.aiProcessing.enabled) {
    ScriptApp.newTrigger('processNewAndChangedConversations')
      .timeBased()
      .everyMinutes(triggerSettings.aiProcessing.interval)
      .create();
  }
}

// =================================================================================
// --- SCHEDULED FUNCTIONS ---
// =================================================================================
function runDailyDataUpdate() {
  Logger.log("Running daily data update...");
  syncDataFromBigQuery();
}

// =================================================================================
// --- WEB APP FUNCTIONS ---
// =================================================================================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sourcing Alerts Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadConversations() {
  console.log('Starting to load conversations...');
  document.getElementById('conversationsList').innerHTML = '<div class="loading"><div class="spinner"></div></div>';

  google.script.run
    .withSuccessHandler(function(response) {
      // The server now returns an object: { data: [], error: '...' }
      if (response && response.error) {
        console.error('Failed to load conversations:', response.error);
        showToast('Error: ' + response.error, 'error');
        conversations = [];
      } else if (response && response.data) {
        console.log(`Successfully loaded ${response.data.length} conversations.`);
        conversations = response.data || [];
      } else {
        console.error('Received an invalid response from the server.');
        showToast('An unknown error occurred while loading data.', 'error');
        conversations = [];
      }
      renderConversationsList();
    })
    .withFailureHandler(function(error) {
      console.error('Script execution failed:', error);
      showToast('Failed to communicate with server: ' + error.message, 'error');
      conversations = [];
      renderConversationsList();
    })
    .getConversations();
}

// =================================================================================
// --- MANUAL EXECUTION FUNCTIONS ---
// =================================================================================
function runManualDataSync() {
  syncDataFromBigQuery();
  return { success: true, message: "Data sync completed successfully" };
}

function runManualAiProcessing() {
  processSheetWithAI();
  return { success: true, message: "AI processing completed successfully" };
}

// Test function to check sheet data
function testGetSheetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      return { 
        error: `Sheet "${TARGET_SHEET_NAME}" not found`, 
        sheets: ss.getSheets().map(s => s.getName()),
        hasData: false
      };
    }
    
    const data = sheet.getDataRange().getValues();
    return {
      sheetName: TARGET_SHEET_NAME,
      totalRows: data.length,
      headers: data[0] || [],
      firstDataRow: data[1] || [],
      hasData: data.length > 1,
      error: null
    };
  } catch (e) {
    console.error('Error in testGetSheetData:', e);
    return {
      error: e.toString(),
      hasData: false,
      totalRows: 0,
      headers: [],
      firstDataRow: []
    };
  }
}

function processSheetWithAI() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("Execution skipped: Another instance of the AI script is already running.");
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Error: Sheet "${TARGET_SHEET_NAME}" not found.`);
      return;
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      Logger.log("No data to process.");
      return;
    }
    
    const conversationsToProcess = [];
    
    // Mark all rows for processing
    for (let i = 2; i <= values.length; i++) {
      conversationsToProcess.push({
        conversationId: values[i-1][0], // Assuming conversation_id is first column
        rowIndex: i
      });
    }
    
    // Store conversations to process
    PropertiesService.getScriptProperties().setProperty('CONVERSATIONS_TO_PROCESS', JSON.stringify(conversationsToProcess));
    
    // Process them
    processNewAndChangedConversations();
    
  } catch (e) {
    Logger.log("Error in processSheetWithAI: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// Email processing function with daily digest
function processSheetForEmailNotifications() {
  const settings = getSettings();
  const recipients = settings.emailRecipients;
  if (!settings.notificationsEnabled) {
    Logger.log("Notifications are disabled.");
    return;
  }
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("Email processing skipped: Another instance is running.");
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const colIdx = {};
    
    ['conversation_id', 'need attention?', 'buyer_name', 'seller_name', 'expert_email', 
     'conversation_link', 'status', 'why', 'current_time_info'].forEach(col => {
      colIdx[col] = headers.indexOf(col);
    });
    
    const alerts = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const needsAttention = row[colIdx['need attention?']]?.toString().toLowerCase();
      
      if (needsAttention === 'yes') {
        alerts.push({
          conversationId: row[colIdx.conversation_id],
          buyerName: row[colIdx.buyer_name] || "Unknown Buyer",
          sellerName: row[colIdx.seller_name] || "Unknown Seller",
          conversationLink: row[colIdx.conversation_link] || "#",
          aiStatus: row[colIdx.status] || "Attention Needed",
          whyNeedsAttention: row[colIdx.why] || "No reason provided",
          expertEmail: row[colIdx.expert_email],
          timeInfo: row[colIdx.current_time_info]
        });
      }
    }
    
    if (alerts.length === 0) {
      Logger.log("No alerts to send.");
      return;
    }
    
    // Send digest grouped by expert
    const subject = `Daily Sourcing Digest: ${alerts.length} conversation${alerts.length > 1 ? 's' : ''} need attention`;
    const grouped = {};
    alerts.forEach(a => {
      const email = a.expertEmail || '';
      if (!grouped[email]) grouped[email] = [];
      grouped[email].push(a);
    });

    Object.keys(grouped).forEach(email => {
      const htmlBody = buildDigestHtml(grouped[email]);
      const to = email || recipients.join(',');
      MailApp.sendEmail({
        to: to,
        cc: recipients.join(','),
        subject: subject,
        htmlBody: htmlBody
      });
    });

    Logger.log(`Sent digest with ${alerts.length} alerts to ${Object.keys(grouped).length} experts.`);
    
  } catch (e) {
    Logger.log("Error in processSheetForEmailNotifications: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

function buildDigestHtml(alerts) {
  let html = `
    <html>
    <head>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; }
        .container { max-width: 700px; margin: 20px auto; }
        .alert { padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 20px; background: #fff; }
        .alert h3 { color: #c82333; margin-top: 0; }
        .button { display: inline-block; background: #007bff; color: white; padding: 12px 28px; text-decoration: none; border-radius: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Daily Sourcing Alerts Digest</h1>
  `;
  
  alerts.forEach((alert, index) => {
    html += `
      <div class="alert">
        <h3>${index + 1}. ${alert.aiStatus} (Buyer: ${alert.buyerName})</h3>
        <p><strong>Reason:</strong> ${alert.whyNeedsAttention}</p>
        <p><strong>Seller:</strong> ${alert.sellerName}</p>
        <a href="${alert.conversationLink}" class="button">Open Conversation</a>
      </div>
    `;
  });
  
  html += `
      </div>
    </body>
    </html>
  `;
  
  return html;
}
// =================================================================================
// --- DEBUG & UTILITY FUNCTIONS (Add these to your script) ---
// =================================================================================

/**
 * Provides a simple connection test for the web app.
 */
function testConnection() {
  try {
    // Attempt a simple Apps Script service call
    const email = Session.getActiveUser().getEmail();
    return { success: true, message: `Connection successful. Logged in as ${email}` };
  } catch (e) {
    return { success: false, message: `Connection failed: ${e.message}` };
  }
}

/**
 * Gathers and returns metadata about the spreadsheet and its sheets.
 */
function getSheetInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().map(s => ({
      name: s.getName(),
      lastRow: s.getLastRow(),
      lastCol: s.getLastColumn()
    }));
    
    return {
      spreadsheetName: ss.getName(),
      targetSheet: TARGET_SHEET_NAME,
      querySheet: QUERY_SHEET_NAME,
      sheets: sheets,
      error: null
    };
  } catch (e) {
    return { error: `Failed to get sheet info: ${e.toString()}` };
  }
}
// =================================================================================
// --- WEB APP DATA RETRIEVAL FUNCTIONS ---
// =================================================================================

/**
 * Retrieves all conversations from the target sheet for the web app
 * @returns {Object} Object containing data array or error message
 */
function getConversations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      console.error(`Sheet "${TARGET_SHEET_NAME}" not found`);
      return { data: [], error: `Sheet "${TARGET_SHEET_NAME}" not found` };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Keep the limit for now to ensure the function completes
    //if (data.length > 50) data.length = 50; 
    
    if (data.length < 2) {
      console.log('No data rows found in sheet');
      return { data: [], error: null };
    }
    
    const headers = data[0];
    const conversations = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const conversation = {};
      
      headers.forEach((header, index) => {
        // Ensure all values are serializable (convert dates to strings)
        let cellValue = row[index];
        if (cellValue instanceof Date) {
          conversation[header] = cellValue.toISOString();
        } else {
          conversation[header] = cellValue;
        }
      });
      
      conversations.push(conversation);
    }
    
    // --- REPLACE THE OLD LOG WITH THIS ---
    // This will try to serialize the object. If it fails, the error will appear in your Apps Script logs.
    try {
      Logger.log('Attempting to serialize ' + conversations.length + ' conversations...');
      Logger.log(JSON.stringify(conversations));
      Logger.log('Serialization successful.');
    } catch(e) {
      Logger.log('SERIALIZATION FAILED: ' + e.message);
    }
    // ---------------------------------
    
    return { data: conversations, error: null };
    
  } catch (error) {
    console.error('Error in getConversations:', error);
    return { data: [], error: error.toString() };
  }
}
