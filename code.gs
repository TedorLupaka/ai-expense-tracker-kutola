// ==========================================
// SECTION 1 — BACKEND PARAMETER INITIALIZATION
// ==========================================

function setupBackendProperties() {
  var props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('llamaParseApiKey') === null) {
    props.setProperty('llamaParseApiKey', 'YOUR_LLAMA_PARSE_API_KEY');
  }
  if (props.getProperty('driveFolderId') === null) {
    props.setProperty('driveFolderId', '');
  }
  if (props.getProperty('systemState') === null) {
    props.setProperty('systemState', 'OFF');
  }
  if (props.getProperty('maxReceipts') === null) {
    props.setProperty('maxReceipts', '10');
  }
  if (props.getProperty('defaultCurrency') === null) {
    props.setProperty('defaultCurrency', 'ZMW');
  }
  if (props.getProperty('autoCategorization') === null) {
    props.setProperty('autoCategorization', 'ON');
  }
  if (props.getProperty('processedFiles') === null) {
    props.setProperty('processedFiles', '{}');
  }
  
  checkAndCreateTrackingSheet();
}

// ==========================================
// SECTION 2 — WEB APP ENDPOINT
// ==========================================

function doGet(e) {
  setupBackendProperties();
  var props = PropertiesService.getScriptProperties();
  
  var template = HtmlService.createTemplateFromFile('Index');
  
  template.spreadsheetId = props.getProperty('spreadsheetId') || '';
  template.systemState = props.getProperty('systemState') || 'OFF';
  template.maxReceipts = props.getProperty('maxReceipts') || '10';
  template.driveFolderId = props.getProperty('driveFolderId') || '';
  template.defaultCurrency = props.getProperty('defaultCurrency') || 'ZMW';
  template.autoCategorization = props.getProperty('autoCategorization') || 'ON';
  
  return template.evaluate()
    .setTitle('AI Expense Tracker Control Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// SECTION 3 — SELF-HEALING AUTOMATION
// ==========================================

function checkAndCreateTrackingSheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('spreadsheetId');
  var ss = null;
  
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = null;
    }
  }
  
  if (!ss) {
    ss = SpreadsheetApp.create('AI Expense Tracker Records');
    props.setProperty('spreadsheetId', ss.getId());
  }
  
  var expenseSheet = ss.getSheetByName('Expenses');
  if (!expenseSheet) {
    var sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0 && sheets[0].getLastColumn() === 0) {
      expenseSheet = sheets[0];
      expenseSheet.setName('Expenses');
    } else {
      expenseSheet = ss.insertSheet('Expenses');
    }
  }
  
  var expHeaders = ['Date', 'Vendor', 'Amount', 'Currency', 'Category', 'Description', 'File Name', 'File ID', 'Status', 'Processed At'];
  if (expenseSheet.getLastRow() === 0) {
    expenseSheet.appendRow(expHeaders);
    expenseSheet.getRange(1, 1, 1, expHeaders.length)
      .setFontWeight('bold')
      .setBackground('#F1F3F4');
  }
  
  var logSheet = ss.getSheetByName('Logs');
  if (!logSheet) {
    logSheet = ss.insertSheet('Logs');
  }
  
  var logHeaders = ['Timestamp', 'File Name', 'Status', 'Message', 'Execution Type'];
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(logHeaders);
    logSheet.getRange(1, 1, 1, logHeaders.length)
      .setFontWeight('bold')
      .setBackground('#F1F3F4');
  }
  
  return ss.getId();
}

// Helper function for quick appending to logs
function internalLog(fileName, status, message, executionType) {
  try {
    var props = PropertiesService.getScriptProperties();
    var ssId = props.getProperty('spreadsheetId');
    if (ssId) {
      var ss = SpreadsheetApp.openById(ssId);
      var logSheet = ss.getSheetByName('Logs');
      if (logSheet) {
        logSheet.appendRow([new Date(), fileName, status, message, executionType || 'SYSTEM']);
      }
    }
  } catch (e) {
    // Fail-safe to avoid blocking execution
  }
}

// ==========================================
// SECTION 4 — FRONTEND AJAX GATEWAYS
// ==========================================

function updateSystemConfiguration(state) {
  setupBackendProperties();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('systemState', state);
  
  if (state === 'ON') {
    startHighFrequencyTrigger();
    internalLog('SYSTEM', 'SUCCESS', 'Automated engine activated.', 'SYSTEM');
  } else {
    stopHighFrequencyTrigger();
    internalLog('SYSTEM', 'SUCCESS', 'Automated engine deactivated.', 'SYSTEM');
  }
  return { success: true, systemState: state };
}

function executeManualRunNow() {
  return runExpenseProcessor('MANUAL_TRIGGER');
}

function executeRetryAllNow() {
  PropertiesService.getScriptProperties().deleteProperty('processedFiles');
  return runExpenseProcessor('MANUAL_RETRY_ALL');
}

function resetSpreadsheetData() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('spreadsheetId');
  if (ssId) {
    try {
      var ss = SpreadsheetApp.openById(ssId);
      
      var expenseSheet = ss.getSheetByName('Expenses');
      if (expenseSheet) {
        var lastRow = expenseSheet.getLastRow();
        if (lastRow > 1) {
          expenseSheet.getRange(2, 1, lastRow - 1, expenseSheet.getLastColumn()).clearContent();
        }
      }
      
      var logSheet = ss.getSheetByName('Logs');
      if (logSheet) {
        var lastRow = logSheet.getLastRow();
        if (lastRow > 1) {
          logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).clearContent();
        }
      }
    } catch (e) {
      return { success: false, reason: e.toString() };
    }
  }
  
  props.deleteProperty('processedFiles');
  internalLog('SYSTEM', 'SUCCESS', 'All spreadsheet data and logs were manually purged.', 'SYSTEM');
  
  return { success: true };
}

function fetchRealTimeAuditLogs() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('spreadsheetId');
  var logs = [];
  if (!ssId) return logs;
  
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var logSheet = ss.getSheetByName('Logs');
    if (logSheet) {
      var lastRow = logSheet.getLastRow();
      if (lastRow > 1) {
        var startRow = Math.max(2, lastRow - 14);
        var numRows = lastRow - startRow + 1;
        var values = logSheet.getRange(startRow, 1, numRows, 5).getValues();
        
        for (var i = values.length - 1; i >= 0; i--) {
          var dateVal = values[i][0];
          var formattedDate = (dateVal instanceof Date) ? dateVal.toISOString().replace('T', ' ').substring(0, 19) : dateVal.toString();
          logs.push({
            timestamp: formattedDate,
            fileName: values[i][1],
            status: values[i][2],
            message: values[i][3],
            executionType: values[i][4]
          });
        }
      }
    }
  } catch (e) {
    // Graceful return of empty array on exception
  }
  return logs;
}

function fetchStats() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('spreadsheetId');
  var folderId = props.getProperty('driveFolderId');
  
  var total = 0;
  var successful = 0;
  var failed = 0;
  var skipped = 0;
  var pending = 0;
  var totalAmount = 0;
  
  if (ssId) {
    try {
      var ss = SpreadsheetApp.openById(ssId);
      var expenseSheet = ss.getSheetByName('Expenses');
      if (expenseSheet) {
        var data = expenseSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          total++;
          var stat = data[i][8];
          if (stat === 'SUCCESS') {
            successful++;
            var amt = parseFloat(data[i][2]);
            if (!isNaN(amt)) {
              totalAmount += amt;
            }
          }
          else if (stat === 'FAILED') failed++;
          else if (stat === 'SKIPPED') skipped++;
        }
      }
    } catch (e) {}
  }
  
  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var query = "mimeType = 'application/pdf' or mimeType = 'image/jpeg' or mimeType = 'image/png' or mimeType = 'image/webp'";
      var files = folder.searchFiles(query);
      var processedStr = props.getProperty('processedFiles') || '{}';
      var processed = JSON.parse(processedStr);
      
      while (files.hasNext()) {
        var file = files.next();
        if (!processed[file.getId()]) {
          pending++;
        }
      }
    } catch (e) {}
  }
  
  return {
    totalReceipts: total,
    successful: successful,
    failed: failed,
    skipped: skipped,
    pending: pending,
    totalAmount: totalAmount,
    currency: props.getProperty('defaultCurrency') || 'ZMW'
  };
}

function getSettings() {
  setupBackendProperties();
  var props = PropertiesService.getScriptProperties();
  return {
    driveFolderId: props.getProperty('driveFolderId') || '',
    maxReceipts: props.getProperty('maxReceipts') || '10',
    defaultCurrency: props.getProperty('defaultCurrency') || 'ZMW',
    autoCategorization: props.getProperty('autoCategorization') || 'ON',
    llamaParseApiKey: props.getProperty('llamaParseApiKey') || 'YOUR_LLAMA_PARSE_API_KEY'
  };
}

function saveSettings(settings) {
  setupBackendProperties();
  var props = PropertiesService.getScriptProperties();
  
  if (settings.driveFolderId !== undefined) props.setProperty('driveFolderId', settings.driveFolderId);
  if (settings.maxReceipts !== undefined) props.setProperty('maxReceipts', settings.maxReceipts);
  if (settings.defaultCurrency !== undefined) props.setProperty('defaultCurrency', settings.defaultCurrency);
  if (settings.autoCategorization !== undefined) props.setProperty('autoCategorization', settings.autoCategorization);
  if (settings.llamaParseApiKey !== undefined) props.setProperty('llamaParseApiKey', settings.llamaParseApiKey);
  
  internalLog('SYSTEM', 'SUCCESS', 'System configuration updated.', 'SYSTEM');
  return { success: true };
}

// ==========================================
// SECTION 5 — CRITICAL EXECUTION ENGINE
// ==========================================

function runExpenseProcessor(executionType) {
  setupBackendProperties();
  var props = PropertiesService.getScriptProperties();
  var state = props.getProperty('systemState');
  var execMode = executionType || 'AUTOMATED_TRIGGER';
  
  if (state !== 'ON' && execMode !== 'MANUAL_TRIGGER') {
    return { status: 'SKIPPED', reason: 'Processor is turned OFF' };
  }
  
  var folderId = props.getProperty('driveFolderId');
  if (!folderId) {
    internalLog('SYSTEM', 'FAILED', 'No Drive folder configuration provided.', execMode);
    return { status: 'FAILED', reason: 'No Drive folder configured' };
  }
  
  var folder = null;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    internalLog('SYSTEM', 'FAILED', 'Target folder ID is missing or inaccessible.', execMode);
    return { status: 'FAILED', reason: 'Invalid Drive folder' };
  }
  
  var maxReceipts = parseInt(props.getProperty('maxReceipts') || '10', 10);
  var query = "mimeType = 'application/pdf' or mimeType = 'image/jpeg' or mimeType = 'image/png' or mimeType = 'image/webp'";
  var files = folder.searchFiles(query);
  var processedCount = 0;
  
  var ss = SpreadsheetApp.openById(props.getProperty('spreadsheetId'));
  var expenseSheet = ss.getSheetByName('Expenses');
  
  while (files.hasNext() && processedCount < maxReceipts) {
    var file = files.next();
    var fileId = file.getId();
    var fileName = file.getName();
    
    if (isFileAlreadyProcessed(fileId)) {
      continue;
    }
    
    try {
      var record = callLlamaParse(file);
      
      if (record) {
        var isValid = validateExpenseRecord(record);
        var statusStr = isValid ? 'SUCCESS' : 'FAILED';
        var msg = isValid ? 'Parsing extracted successfully.' : 'Validation failure. Output: ' + (record._rawOutput ? record._rawOutput.substring(0, 200) : JSON.stringify(record).substring(0, 200));
        
        expenseSheet.appendRow([
          record.date || '',
          record.vendor || '',
          record.amount !== null ? record.amount : '',
          record.currency || props.getProperty('defaultCurrency'),
          record.category || '',
          record.description || '',
          fileName,
          fileId,
          statusStr,
          new Date()
        ]);
        
        internalLog(fileName, statusStr, msg, execMode);
        markFileAsProcessed(fileId);
        processedCount++;
      } else {
        expenseSheet.appendRow([
          '', '', '', '', '', '', fileName, fileId, 'FAILED', new Date()
        ]);
        internalLog(fileName, 'FAILED', 'LlamaParse processing returned an empty payload text.', execMode);
        markFileAsProcessed(fileId);
        processedCount++;
      }
    } catch (err) {
      expenseSheet.appendRow([
        '', '', '', '', '', '', fileName, fileId, 'FAILED', new Date()
      ]);
      internalLog(fileName, 'FAILED', 'Exception: ' + err.toString(), execMode);
      markFileAsProcessed(fileId);
      processedCount++;
    }
  }
  
  return { status: 'COMPLETED', processedCount: processedCount };
}

// ==========================================
// SECTION 6 — RECEIPT VALIDATION
// ==========================================

function validateExpenseRecord(record) {
  if (!record) return false;
  if (record.amount === null || record.amount === undefined || record.amount === '') return false;
  if (!record.vendor || record.vendor === '') return false;
  if (!record.date || record.date === '') return false;
  
  var timestamp = Date.parse(record.date);
  if (isNaN(timestamp)) return false;
  
  return true;
}

function sanitizeReceiptText(text) {
  if (!text) return '';
  // Force to string in case the API returned an object/array
  var str = (typeof text === 'string') ? text : JSON.stringify(text);
  var cleaned = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '');
  cleaned = cleaned.replace(/\s+/g, ' ');
  return cleaned.trim();
}

// ==========================================
// SECTION 7 — FILE MANAGEMENT
// ==========================================

function isFileAlreadyProcessed(fileId) {
  var props = PropertiesService.getScriptProperties();
  var processedStr = props.getProperty('processedFiles') || '{}';
  var processed = JSON.parse(processedStr);
  return !!processed[fileId];
}

function markFileAsProcessed(fileId) {
  var props = PropertiesService.getScriptProperties();
  var processedStr = props.getProperty('processedFiles') || '{}';
  var processed = JSON.parse(processedStr);
  
  var now = new Date().getTime();
  processed[fileId] = now;
  
  var thirtyDaysAgo = now - (30 * 24 * 60 * 60 * 1000);
  var cleanProcessed = {};
  
  for (var id in processed) {
    if (processed.hasOwnProperty(id)) {
      if (processed[id] >= thirtyDaysAgo) {
        cleanProcessed[id] = processed[id];
      }
    }
  }
  
  props.setProperty('processedFiles', JSON.stringify(cleanProcessed));
}

// ==========================================
// SECTION 8 — LLAMAPARSE ENGINE
// ==========================================

function callLlamaParse(file) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('llamaParseApiKey');
  
  if (!apiKey || apiKey === 'YOUR_LLAMA_PARSE_API_KEY') {
    throw new Error('LlamaParse execution halted: missing authentication credential.');
  }
  
  var blob = file.getBlob();
  var instructions = 'You are a receipt data extraction bot. You MUST return ONLY a valid JSON object containing the financial information. Do NOT return standard markdown text, headers, or explanations. Just the JSON. Schema: {"date": "YYYY-MM-DD", "vendor": "string", "amount": 0.0, "currency": "string", "category": "string", "description": "string"}. Populate null if missing.';
  
  // 1. Upload file and provide parsing instructions directly (v1 API supports this natively)
  var filePayload = {
    'file': blob,
    'parsing_instruction': instructions
  };
  
  var fileUploadOptions = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': filePayload,
    'muteHttpExceptions': true
  };
  var fileResponse = UrlFetchApp.fetch('https://api.cloud.llamaindex.ai/api/parsing/upload', fileUploadOptions);
  var fileResponseCode = fileResponse.getResponseCode();
  if (fileResponseCode < 200 || fileResponseCode >= 300) {
    throw new Error('LlamaParse v1 file upload failed (' + fileResponseCode + '): ' + fileResponse.getContentText());
  }
  var fileData = JSON.parse(fileResponse.getContentText());
  var jobId = fileData.id;

  // 2. Poll for result (v1 API)
  var statusUrl = 'https://api.cloud.llamaindex.ai/api/parsing/job/' + jobId;
  var checkOptions = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Accept': 'application/json'
    },
    'muteHttpExceptions': true
  };
  
  var jobState = 'PENDING';
  var attempt = 0;
  var maxAttempts = 15;
  var finalResultData = null;
  var lastRawResponse = '';
  
  while ((jobState === 'PENDING' || jobState === 'RUNNING') && attempt < maxAttempts) {
    var sleepDuration = Math.pow(2, attempt) * 1000;
    if (sleepDuration > 10000) sleepDuration = 10000;
    Utilities.sleep(sleepDuration);
    
    var statusResponse = UrlFetchApp.fetch(statusUrl, checkOptions);
    var statusCode = statusResponse.getResponseCode();
    if (statusCode === 200 || statusCode === 201) {
      lastRawResponse = statusResponse.getContentText();
      var statusData = JSON.parse(lastRawResponse);
      
      var resolvedState = statusData.status || statusData.state;
      if (resolvedState) {
        jobState = resolvedState.toUpperCase();
      }
      if (jobState === 'COMPLETED' || jobState === 'SUCCESS') {
        jobState = 'COMPLETED';
        finalResultData = statusData;
      }
    }
    attempt++;
  }
  
  if (jobState !== 'COMPLETED') {
    throw new Error('LlamaParse timeout. State: ' + jobState + '. Response: ' + lastRawResponse.substring(0, 300));
  }
  
  var resultUrl = 'https://api.cloud.llamaindex.ai/api/parsing/job/' + jobId + '/result/markdown';
  var resultResponse = UrlFetchApp.fetch(resultUrl, checkOptions);
  if (resultResponse.getResponseCode() === 200) {
    try {
      finalResultData = JSON.parse(resultResponse.getContentText());
    } catch(e) {
      finalResultData = { markdown: resultResponse.getContentText() };
    }
  }
  
  var resultData = finalResultData;
  var combinedText = '';
  
  if (resultData && resultData.markdown && typeof resultData.markdown === 'string') {
    combinedText = resultData.markdown;
  } else if (resultData && resultData.markdown && resultData.markdown.pages) {
    for (var i = 0; i < resultData.markdown.pages.length; i++) {
      combinedText += (resultData.markdown.pages[i].markdown || '') + '\n';
    }
  } else if (resultData && resultData.pages) {
    for (var i = 0; i < resultData.pages.length; i++) {
      combinedText += (resultData.pages[i].text || resultData.pages[i].markdown || '') + '\n';
    }
  } else if (resultData && resultData.text) {
    combinedText = resultData.text;
  } else {
    combinedText = JSON.stringify(resultData);
  }
  
  combinedText = sanitizeReceiptText(combinedText);
  
  var firstBrace = combinedText.indexOf('{');
  var lastBrace = combinedText.lastIndexOf('}');
  
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    var possibleJson = combinedText.substring(firstBrace, lastBrace + 1);
    try {
      var extractedJson = JSON.parse(possibleJson);
      return {
        date: extractedJson.date || null,
        vendor: extractedJson.vendor || null,
        amount: extractedJson.amount !== undefined ? extractedJson.amount : null,
        currency: extractedJson.currency || null,
        category: extractedJson.category || null,
        description: extractedJson.description || null,
        _rawOutput: combinedText
      };
    } catch (err) {
      try {
        var fallbackJson = JSON.parse(combinedText);
        fallbackJson._rawOutput = combinedText;
        return fallbackJson;
      } catch (e) {
        return null;
      }
    }
  }
  
  // If no braces found, but we have text, log it as an error
  return { _rawOutput: combinedText };
}

// ==========================================
// SECTION 9 — TRIGGER MANAGER
// ==========================================

function startHighFrequencyTrigger() {
  stopHighFrequencyTrigger();
  ScriptApp.newTrigger('runExpenseProcessor')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function stopHighFrequencyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runExpenseProcessor') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}