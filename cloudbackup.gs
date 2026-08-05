// 填入你的相關資訊
const PROPS = PropertiesService.getScriptProperties();
const LINE_ACCESS_TOKEN = PROPS.getProperty('LINE_ACCESS_TOKEN');
const SHEET_ID = PROPS.getProperty('SHEET_ID');
const LINE_USER_ID = PROPS.getProperty('LINE_USER_ID');
const DRIVE_FOLDER_ID = PROPS.getProperty('DRIVE_FOLDER_ID');

function doPost(e) {
  try {
    // 解析 LINE 傳來的 Webhook 資料
    const jsonData = JSON.parse(e.postData.contents);
    const events = jsonData.events;

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const replyToken = event.replyToken;
      const userID = event.source.userId;

      // 1. 權限檢查：若非設定的私人使用者，直接拒絕並跳過
      if (userID !== LINE_USER_ID) {
        writeText(`有其他使用者傳送訊息`, userID);
        replyMessage(replyToken, `該機器人僅限私人使用`);
        continue; // 跳過本次迴圈，處理下一個 event
      }

      // 2. 安全檢查：確保事件類型是「訊息」，否則不處理（防止 follow 等事件造成程式噴錯）
      if (event.type !== 'message') {
        continue; 
      }

      const msgType = event.message.type;
      const messageId = event.message.id;
      
      // 產生時間戳記（用於建立自訂檔名）
      const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");

      // 3. 使用 switch 判斷訊息類型，呼叫統一的備份函式
      switch (msgType) {
        case 'image':
          writeText(`收到圖片`, userID);
          saveLineContent(messageId, `LINE_${timestamp}.jpg`, replyToken, '圖片');
          break;

        case 'video':
          writeText(`收到影片`, userID);
          saveLineContent(messageId, `LINE_${timestamp}.mp4`, replyToken, '影片');
          break;

        case 'audio':
          writeText(`收到音訊`, userID);
          saveLineContent(messageId, `LINE_${timestamp}.m4a`, replyToken, '音訊');
          break;

        case 'file':
          // 檔案格式比較特殊，LINE 會直接提供原始檔名
          const originalFileName = event.message.fileName;
          writeText(`收到檔案: ${originalFileName}`, userID);
          saveLineContent(messageId, originalFileName, replyToken, '檔案');
          break;

        default:
          replyMessage(replyToken, '目前僅支援圖片、影片、音訊與檔案備份');
          break;
      }
    }
  } catch (error) {
    Logger.log('錯誤原因: ' + error.toString());
  }
}

/**
 * 統一處理 LINE 多媒體檔案下載與儲存的共用函式
 * @param {string} messageId - LINE 訊息 ID
 * @param {string} folderId - Google Drive 資料夾 ID
 * @param {string} fileName - 要儲存的檔案名稱
 * @param {string} replyToken - LINE 回覆 Token
 * @param {string} typeLabel - 記錄與回覆文字用的標籤 (如：圖片、影片)
 */
function saveLineContent(messageId, fileName, replyToken, typeLabel) {
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  
  // 呼叫 LINE API 下載二進位檔案
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${LINE_ACCESS_TOKEN}`
    },
    muteHttpExceptions: true // 防止 401 或 404 等錯誤直接讓整個 GAS 專案中斷
  });

  const statusCode = response.getResponseCode();
  
  if (statusCode === 200) {
    // 取得檔案的 Blob 並設定檔名
    const blob = response.getBlob();
    blob.setName(fileName);

    // 儲存到指定雲端硬碟資料夾
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl(); // 取得雲端連結

    // 紀錄操作到 Google Sheet
    writeText(`已備份${typeLabel} ${fileUrl}`, `system`);

    // 回覆使用者備份成功
    replyMessage(replyToken, `${typeLabel}已成功備份至 Google Drive！`);
  } else {
    writeText(`備份${typeLabel}失敗，status code: ${statusCode}`, `system`);
    replyMessage(replyToken, `${typeLabel}備份失敗，請稍後再試`);
  }
}

// 回覆 LINE 訊息的函式
function replyMessage(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${LINE_ACCESS_TOKEN}`
    },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: text }]
    })
  });
}

// 紀錄（共用函式）
function writeText(text, userID) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('工作表1');
    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, userID, text]);
  } catch (e) {
    console.error("無法寫入 sheet: " + e.toString());
  }
}