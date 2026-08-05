# line-bots 聊天機器人

> ### **簡介**：結合Google Apps Script + LINE BOT + Google試算表的聊天機器人。  

### 目的  
* cloud backup： LINE上的圖片、檔案等都有時間限制，若將資料轉送聊天機器人即可備份至雲端硬碟上，避免過期而無法使用也不必下載至手機上佔據儲存空間。
* remind： 使用者傳送訊息到聊天機器人後會交給AI分辨提醒時間與內容，會在指定時間發送通知訊息給使用者。

### 使用方法
1. 申請並登入 [LINE Developers](https://developers.line.biz/en/)，建立聊天機器人。
2. 取得userID、Channel access token，請勿隨意公布資料。
3. 在Google雲端硬碟上新增GAS和試算表，試算表是用來當作資料庫使用。
4. Google試算表SHEET_ID從網址上取得，  
    例如試算表網址為`https://docs.google.com/spreadsheets/d/ **此為試算表ID** /edit?gid=0#gid=0`
5. 在GAS專案設定裡面的指令碼屬性將userID、Channel access token、SHEET_ID等資料設定好，這些資料都極為重要請勿隨意公開洩露。  
    依據專案還需要的ID取得方法：
    * DRIVE_FOLDER_ID： 此為Google雲端硬碟資料夾ID，從該資料夾網址中取得，例如`https://drive.google.com/drive/folders/ **資料夾ID** `。
    * GEMINI_API_KEY： 此為Gemini API KEY，從Google AI Studio建立APIkey。
6. 將程式碼複製貼上到GAS的編輯器當中後存檔，請先嘗試使用編輯器的執行功能執行一次，取得GAS專案的存取權，確認授權後右上角選擇新增部署作業，類型選擇**網頁應用程式**，誰可以存取選擇**所有人**，即可取得網址。
7. 將取得的網址貼到聊天機器人的Webhook URL即可使用聊天機器人。  
    * 聊天機器人有一個Messaging API分頁，請在當中設定Webhook URL。

### 個人說明
專案建立有使用AI（Gemini）編寫程式碼，此為個人練習的專案內容。