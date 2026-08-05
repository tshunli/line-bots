// ==================== 基礎設定 ====================
// 改為唯讀的設定包，避免全域直接執行 PropertiesService
function getConfigs() {
  const PROPS = PropertiesService.getScriptProperties();
  return {
    LINE_ACCESS_TOKEN: PROPS.getProperty('LINE_ACCESS_TOKEN'),
    GEMINI_API_KEY: PROPS.getProperty('GEMINI_API_KEY'),
    SHEET_ID: PROPS.getProperty('SHEET_ID')
  };
}
// ==================================================

// 接收 LINE Webhook 的主要入口
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    
    // 安全檢查：確保有事件才處理
    if (!json.events || json.events.length === 0) return;

    // 巡檢所有事件，避免漏掉訊息
    for (const event of json.events) {
      handleLineEvent(event);
    }
  } catch (error) {
    writeText(['ERROR', 'doPost', 'SYSTEM', 'post(e)_catch', "系統嚴重崩潰: " + error.toString()], "log");
  }
}

// 處理單一 LINE 事件
function handleLineEvent(event) {
  const configs = getConfigs();
  const replyToken = event.replyToken;
  const eventType = event.type;
  const userId = event.source.userId;
  
  let analysis = null;
  let msg = [];

  // 1. 判定事件類型
  if (eventType === 'message' && event.message.type === 'text') {
    const userMessage = event.message.text;
    if (userMessage === '確定') return; // 跳過關鍵字

    analysis = textGemini(userMessage, configs.GEMINI_API_KEY);
    if (analysis) analysis.clientID = userId; // 安全寫入
    
    msg = ['INFO', 'handleLineEvent', userId, 'receive_msg', '收到使用者的文字訊息'];
    writeText(msg, 'log');

  } else if (eventType === 'message' && event.message.type === 'image') {
    const messageId = event.message.id;
    const getBase = getImage(messageId, configs.LINE_ACCESS_TOKEN);
    
    analysis = imgGemini(getBase, configs.GEMINI_API_KEY);
    if (analysis) analysis.clientID = userId; // 安全寫入
    
    msg = ['INFO', 'handleLineEvent', userId, 'receive_msg', '收到使用者的圖片訊息'];
    writeText(msg, 'log');

  } else if (eventType === 'postback') {
    const postbackData = event.postback.data; 
    const confirmT = checkConfirmData(userId, postbackData, configs.SHEET_ID);
    
    if (confirmT === '1') {
      replyMsg(replyToken, '已收到，已設定排程。', configs.LINE_ACCESS_TOKEN);
      msg = ['INFO', 'handleLineEvent', userId, 'receive_postback', '使用者已確定提醒'];
    } else if (confirmT === '0') {
      replyMsg(replyToken, '已取消，請重新設定。', configs.LINE_ACCESS_TOKEN);
      msg = ['INFO', 'handleLineEvent', userId, 'receive_postback', '使用者點選取消提醒'];
    } else if (confirmT === '2') {
      replyMsg(replyToken, '請勿點選舊資料的確認模板', configs.LINE_ACCESS_TOKEN);
      msg = ['WARN', 'handleLineEvent', userId, 'receive_postback', '使用者點選舊按鈕，可忽略'];
    } else {
      replyMsg(replyToken, '無法理解', configs.LINE_ACCESS_TOKEN);
      msg = ['ERROR', 'handleLineEvent', userId, 'receive_postback', 'checkConfirm非自定義內容:' + confirmT];
    }
    writeText(msg, 'log');
    return;

  } else {
    replyMsg(replyToken, '請輸入文字或圖片訊息', configs.LINE_ACCESS_TOKEN);
    msg = ['WARN', 'handleLineEvent', userId, 'receive_otherType', `收到 ${eventType} 訊息，已忽略。`];
    writeText(msg, 'log');
    return;
  }

  // 2. 處理 Gemini 分析結果
  if (analysis) {
    if (analysis.intent === 'remind' && analysis.remind_data.time && analysis.remind_data.content) {
      const task = [
        analysis.clientID,
        analysis.remind_data.time,
        analysis.remind_data.content,
        '未發送'
      ];
      writeText(task, "task");
      replyConfirm(replyToken, `⏰ 時間：${analysis.remind_data.time}\n📝 內容：${analysis.remind_data.content}`, configs.LINE_ACCESS_TOKEN);
      msg = ['INFO', 'handleLineEvent', userId, 'send_confirmMsg', `發送提醒資料確認給使用者。`];
    } else if (analysis.intent === 'unknown') {
      replyMsg(replyToken, '非提醒事項', configs.LINE_ACCESS_TOKEN);
      msg = ['INFO', 'handleLineEvent', userId, 'parse_fail', `Gemini分析為 unknown`];
    } else {
      replyMsg(replyToken, '無法判讀內容', configs.LINE_ACCESS_TOKEN);
      msg = ['WARN', 'handleLineEvent', userId, 'parse_fail', `Gemini回傳結構異常`];
    }
  } else {
    replyMsg(replyToken, '解析失敗，請重新輸入。', configs.LINE_ACCESS_TOKEN);
    msg = ['WARN', 'handleLineEvent', userId, 'parse_fail', `Gemini API 回傳 null`];
  }
  writeText(msg, "log");
}

// 分析確認按鈕按下的時機
function checkConfirmData(clientID, clientMessage, sheetId) {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('task');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return '2'; // 沒資料防呆
    
    const lastColumn = sheet.getLastColumn();
    const sheetData = sheet.getSheetValues(1, 1, lastRow, lastColumn);
    let returnData;
    
    for (let i = lastRow - 1; i >= 0; i--) {
        // 欄B(索引1)符合使用者ID 且 欄F(索引5)是空白
        if (sheetData[i][1] == clientID && (sheetData[i][5] == "" || sheetData[i][5] === undefined)) {
            if (clientMessage == "confirm=yes") {
                sheet.getRange(i + 1, 6).setValue(new Date()); 
                returnData = '1'; 
            } else if (clientMessage == "confirm=no") {
                sheet.deleteRow(i + 1); 
                returnData = '0'; 
            }
            return returnData;
        }
    }
    return '2';    
}

// 獲取 LINE 圖片並轉成 base64
function getImage(messageId, accessToken) {
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const options = {
    "headers": { "Authorization": "Bearer " + accessToken },
    'method': 'get',
    'muteHttpExceptions': true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob();
  return {
    base64Data: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType()
  };
}

// 定時檢查 Sheet 並發送 LINE 通知
function checkAndSendReminders() {
  const configs = getConfigs();
  const sheet = SpreadsheetApp.openById(configs.SHEET_ID).getSheetByName('task');
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const targetUserId = row[1]; // B欄：使用者的 LINE ID【關鍵修正：從這邊拿 ID】
    const reminderTime = new Date(row[2]); // C欄：提醒時間
    const content = row[3];      // D欄：內容
    const status = row[4];       // E欄：狀態

    if (status === '未發送' && now >= reminderTime) {
      const pushmsg = `【提醒通知！】\n${content}`;
      
      // 傳入正確的 targetUserId
      if (pushMsg(targetUserId, pushmsg, configs.LINE_ACCESS_TOKEN)) {
        sheet.getRange(i + 1, 5).setValue('已發送');
        writeText(['INFO', 'checkAndSendReminders', 'SYSTEM', 'Update_Status', `成功發送提醒給 ${targetUserId}`], "log");
      }
    }
  }
}

function requestLineApi(apiname, payload, accessToken, funcName) {
  const url = `https://api.line.me/v2/bot/message/${apiname}`;
  const options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) {
      writeText(['INFO', `${funcName}`, 'SYSTEM', `${funcName}_success`, '推播成功 '], "log");
      return res.getResponseCode() === 200;
    } else {
      writeText(['ERROR', `${funcName}`, 'SYSTEM', `${funcName}_fail`,`status code: ${res.getResponseCode()}`], "log");
    }
  } catch(e) {
    writeText(['ERROR', `${funcName}`, 'SYSTEM', `${funcName}_fail`, '推播失敗: ' + e.toString()], "log");
    return false;
  }
}

// 主動發送 LINE 訊息（已修正參數，支援多用戶）
function pushMsg(toUserId, text, accessToken) {
  const payload = {
    to: toUserId,
    messages: [{ type: 'text', text: text }]
  };

  return requestLineApi(`push`, payload, accessToken, `push`);
}

// 發送確認模板
function replyConfirm(replyToken, text, accessToken) {
  const payload = {
    replyToken: replyToken,
    messages: [{
      type: 'template',
      altText: '請確認資料是否正確',
      template: {
        type: 'confirm',
        text: `請確認資料是否正確\n${text}`,
        actions: [
          { type: 'postback', label: '確定', data: 'confirm=yes', text: '確定' },
          { type: 'postback', label: '取消', data: 'confirm=no', text: '取消' }
        ]
      }
    }]
  };
  return requestLineApi(`reply`, payload, accessToken, `Confirm`);
}

// 發送一般 Reply 訊息
function replyMsg(replyToken, text, accessToken) {
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  };
  return requestLineApi(`reply`, payload, accessToken, `reply`);
}

function requestGeminiApi(parts, apiKey, funcName) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`;
  const payload = {
    "contents": [{ "parts": parts }],
    "generationConfig": { "responseMimeType": "application/json" }
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resData = JSON.parse(response.getContentText());
    const resultText = resData.candidates[0].content.parts[0].text.trim();
    writeText(['INFO', funcName, 'SYSTEM', 'GeminiApi_success', '分析成功'], "log");
    return JSON.parse(resultText);
  } catch (e) {
    writeText(['ERROR', funcName, 'SYSTEM', 'GeminiApi_fail', '分析失敗: ' + e.toString()], "log");
    return null;
  }
}

// 文字分析
function textGemini(userMessage, apiKey) {
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

  const prompt =
`你是一個精準的提醒事件解析助手。請分析使用者的輸入，判斷其是否包含提醒意圖，並嚴格以 JSON 格式輸出。
【基準時間】
現在的時間是：${nowStr}
（請以此時間為基準，推算如「明天」、「下週一」、「5分鐘後」等相對時間。）
【處理邏輯】
1. 意圖判斷 (intent)：
- 若使用者明確要求在某個時間點（或一段時間後）提醒某件事，"intent"為 "remind"。
- 若訊息不包含時間、不屬於提醒、或語意不明，"intent" 為"unknown"。

2. 時間推算 (time)：
- 當 intent 為 "remind" 時，請將時間格式化為 "YYYY-MM-DD HH:mm:ss"。
- 若使用者只說了時間（例如：15:00），日期請自動帶入基準時間的當天或下一個合理的未來時間。
- 當 intent 為 "unknown" 時，"time" 請填寫為空字串 ""。

3. 內容擷取 (content)：
-擷取需要被提醒的核心核心事件，去除「記得叫我」、「提醒我」等冗詞。
-當 intent 為 "unknown" 時，"content" 請填寫為空字串 ""。

【輸出限制】
-只能回傳純 JSON 格式的字串，絕對不要包含任何 Markdown 標籤，也不要包含任何前言、解釋或後續說明。

【輸出 JSON 結構】
{"intent": "remind" 或 "unknown","remind_data": {"time": "YYYY-MM-DD HH:mm:ss","content": "提醒內容"}}
【使用者輸入】"${userMessage}"`;

  return requestGeminiApi([{ "text": prompt }], apiKey, 'textGemini');
}

// 圖片分析
function imgGemini(imageData, apiKey) {
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

  const prompt =
`現在的基準時間是：${nowStr}。
你是一個視覺分析助手。請仔細觀看這張圖片（可能是行事曆、通知截圖），判斷使用者的意圖。
【分類規則】
若圖片是行事曆、活動海報、會議通知截圖等：意圖為 "remind"。請抓取活動時間（換算成 YYYY-MM-DD HH:mm:ss）與活動主旨。
請嚴格以純 JSON 格式回傳，不要包含 markdown 標籤：
{"intent": "remind" 或 "unknown","remind_data": { "time": "YYYY-MM-DD HH:mm:ss", "content": "..." },}`;

  const parts = [
    { "text": prompt },
    { "inlineData": { "mimeType": imageData.mimeType, "data": imageData.base64Data } }
  ];
  return requestGeminiApi(parts, apiKey, 'imgGemini');
}

// 紀錄（共用函式）
function writeText(message, sheetName) {
  try {
    const configs = getConfigs();
    const sheet = SpreadsheetApp.openById(configs.SHEET_ID).getSheetByName(sheetName);
    let rowData = [new Date()];
    if (Array.isArray(message)) {
      rowData.push(...message);
    } else if (typeof message === 'object' && message !== null) {
      rowData.push(JSON.stringify(message));
    } else {
      rowData.push(message);
    }
    sheet.appendRow(rowData);
  } catch (e) {
    console.error("無法寫入 sheet: " + e.toString());
  }
}