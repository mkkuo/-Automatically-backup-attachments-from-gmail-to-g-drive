自動化 Gmail 附件轉存至 Google 雲端硬碟這個 Google Apps Script 腳本能自動從指定的 Gmail 信箱中，根據郵件主題（日報、週報、月報）篩選出帶有附件的郵件，並將附件自動轉存至 Google 雲端硬碟中的指定資料夾。主要功能自動化處理：每日、每週或每月自動執行，無需手動操作。智能分類：根據郵件主題將附件分類至「日報」、「週報」或「月報」資料夾。動態資料夾：根據附件檔名中的日期資訊（例如 __d__2025_09_18、__w__2025_09_14_to_20、__m__2025_08_01_to_31），自動建立對應的日期子資料夾，確保檔案井然有序。避免重複：已處理的郵件會被自動加上星號，避免腳本重複轉存相同的附件。安裝與設定建立 Google Apps Script 專案前往 Google Apps Script 並點擊「新增專案」。複製程式碼將以下完整程式碼複製並貼到 Apps Script 編輯器中。設定變數在程式碼頂部的 const SENDER_EMAIL 和 const ROOT_FOLDER_ID 變數中，填入您的寄件者信箱和雲端硬碟根資料夾 ID。const SENDER_EMAIL = "sender@example.com"; 
const ROOT_FOLDER_ID = "請在此貼上您的資料夾 ID"; 
設定自動化觸發條件點擊 Apps Script 編輯器左側的 觸發條件 (時鐘圖示)。點擊 新增觸發條件，並依照您的需求設定以下三個排程：函數名稱事件來源時間驅動類型說明runDailyJob時間驅動每日計時器每日自動執行，處理日報郵件runWeeklyJob時間驅動每週計時器每週自動執行，處理週報郵件runMonthlyJob時間驅動每月計時器每月自動執行，處理月報郵件程式碼/**
 * 這是 Google Apps Script，用於自動從 Gmail 讀取郵件，
 * 並根據郵件主題（日報/週報/月報）將附件儲存至指定的 Google Drive 資料夾。
 *
 * 設定步驟：
 * 1. 在 Google Drive 建立一個用於儲存附件的根資料夾。
 * 2. 複製該資料夾的 ID（從網址列取得）。
 * 3. 在下方的 SENDER_EMAIL 和 ROOT_FOLDER_ID 變數中，填入您自己的資訊。
 * 4. 儲存此檔案。
 * 5. 點擊「執行」>「執行函數」選擇 runDailyJob，首次執行時會要求授權。
 * 6. 點擊左側的「觸發條件」（時鐘圖示），新增觸發條件。
 * - 選擇要執行的函數：runDailyJob
 * - 選擇部署：Head
 * - 選擇事件來源：時間驅動
 * - 選擇時間驅動程式類型：每日計時器
 * - 選擇時間：選一個您希望每天執行的時間，例如「上午 12 點到上午 1 點」
 * 7. 再次新增觸發條件，設定每週執行：
 * - 選擇要執行的函數：runWeeklyJob
 * - 選擇部署：Head
 * - 選擇事件來源：時間驅動
 * - 選擇時間驅動程式類型：每週計時器
 * - 選擇星期幾：選擇您希望執行的日期
 * - 選擇時間：選擇您希望執行的時間
 * 8. 再次新增觸發條件，設定每月執行：
 * - 選擇要執行的函數：runMonthlyJob
 * - 選擇部署：Head
 * - 選擇事件來源：時間驅動
 * - 選擇時間驅動程式類型：每月計時器
 * - 選擇當月日期：選擇您希望執行的日期
 * - 選擇時間：選擇您希望執行的時間
 */

// ============== 請在下方填入您的設定 ==============
// 指定的寄件者信箱
const SENDER_EMAIL = "sender@example.com"; 

// 儲存附件的根資料夾 ID
// 例如：[https://drive.google.com/drive/folders/xxxxxxxxxxxxxxxxxxxxxxxxxxxxx](https://drive.google.com/drive/folders/xxxxxxxxxxxxxxxxxxxxxxxxxxxxx)
// ID 就是網址最後那串字串
const ROOT_FOLDER_ID = "請在此貼上您的資料夾 ID"; 

// ============== 程式碼開始，一般情況下無需修改 ==============

/**
 * 每日執行日報郵件處理。
 */
function runDailyJob() {
  processEmails("日報", "daily");
}

/**
 * 每週執行週報郵件處理。
 */
function runWeeklyJob() {
  processEmails("週報", "weekly");
}

/**
 * 每月執行月報郵件處理。
 */
function runMonthlyJob() {
  processEmails("月報", "monthly");
}

/**
 * 主要處理函數，根據主題和標籤處理郵件。
 * @param {string} subjectKeyword 郵件主題的關鍵字（例如："日報" 或 "週報"）。
 * @param {string} dateType 創建資料夾的日期類型（"daily" 或 "weekly" 或 "monthly"）。
 */
function processEmails(subjectKeyword, dateType) {
  const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  
  // 建立 Gmail 搜尋字串
  const searchQuery = `from:${SENDER_EMAIL} subject:${subjectKeyword} has:attachment`;
  Logger.log(`搜尋條件: ${searchQuery}`);
  
  const threads = GmailApp.search(searchQuery, 0, 100); // 最多處理 100 封信件
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // 檢查是否已處理過，避免重複儲存
      if (!message.isStarred()) {
        const attachments = message.getAttachments();
        
        if (attachments.length > 0) {
          Logger.log(`正在處理郵件：${message.getSubject()}`);
          
          attachments.forEach(attachment => {
            const fileName = attachment.getName();
            let extractedDateString = null;
            
            // 根據郵件類型從檔名中提取日期字串
            if (dateType === "daily") {
              const dateMatch = fileName.match(/__d__(\d{4}_\d{2}_\d{2})/);
              if (dateMatch) {
                extractedDateString = dateMatch[1];
              }
            } else if (dateType === "weekly") {
              // 匹配週報檔名格式，例如: "__w__2025_09_14_to_20"
              const dateMatch = fileName.match(/__w__(\d{4}_\d{2}_\d{2}_to_\d{2})/);
              if (dateMatch) {
                extractedDateString = dateMatch[1];
              }
            } else if (dateType === "monthly") {
              // 匹配月報檔名格式，例如: "__m__2025_08_01_to_31"
              const dateMatch = fileName.match(/__m__(\d{4}_\d{2})/);
              if (dateMatch) {
                extractedDateString = dateMatch[1];
              }
            }

            const targetFolder = getOrCreateFolder(rootFolder, subjectKeyword, dateType, extractedDateString);
            
            Logger.log(`正在儲存附件：${fileName}`);
            targetFolder.createFile(attachment);
          });
          
          // 為已處理的郵件加上星號，避免下次重複處理
          message.star(); 
        }
      }
    });
  });
  
  Logger.log(`${subjectKeyword} 處理完成。`);
}

/**
 * 根據日期和類型取得或建立子資料夾。
 * @param {GoogleAppsScript.Drive.Folder} parentFolder 父資料夾。
 * @param {string} subFolderName 子資料夾的名稱（例如："日報"）。
 * @param {string} dateType 日期類型（"daily" 或 "weekly" 或 "monthly"）。
 * @param {string} extractedDateString 可選。從檔名中提取的日期字串。
 * @return {GoogleAppsScript.Drive.Folder} 儲存附件的目標資料夾。
 */
function getOrCreateFolder(parentFolder, subFolderName, dateType, extractedDateString) {
  let dateFolderName = "";
  
  // 如果從檔名中提取到了日期字串，則使用它來建立資料夾名稱
  if (extractedDateString) {
    if (dateType === "daily") {
      dateFolderName = extractedDateString.replace(/_/g, '-');
    } else if (dateType === "weekly") {
      // 處理週報格式，例如: "2025_09_14_to_20" 轉為 "2025-09-14_09-20"
      const parts = extractedDateString.split('_');
      dateFolderName = `${parts[0]}-${parts[1]}-${parts[2]}_${parts[1]}-${parts[4]}`;
    } else if (dateType === "monthly") {
      // 處理月報格式，例如: "2025_08" 轉為 "2025-08_月報"
      dateFolderName = extractedDateString.replace(/_/g, '-') + '_月報';
    }
  } else {
    // 否則，使用當前執行日期作為後備方案
    const today = new Date();
    if (dateType === "daily") {
      dateFolderName = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
    } else if (dateType === "weekly") {
      // 取得當週的第一天 (星期日)
      const sunday = new Date(today.setDate(today.getDate() - today.getDay()));
      dateFolderName = Utilities.formatDate(sunday, "GMT+8", "yyyy-MM-dd") + " 週";
    } else if (dateType === "monthly") {
      dateFolderName = Utilities.formatDate(today, "GMT+8", "yyyy-MM") + "_月報";
    }
  }
  
  // 先檢查是否存在「日報」、「週報」或「月報」資料夾
  const folders = parentFolder.getFoldersByName(subFolderName);
  let subFolder;
  if (folders.hasNext()) {
    subFolder = folders.next();
  } else {
    subFolder = parentFolder.createFolder(subFolderName);
    Logger.log(`已建立新資料夾：${subFolderName}`);
  }
  
  // 在子資料夾下，檢查是否存在日期資料夾
  const dateFolders = subFolder.getFoldersByName(dateFolderName);
  let finalFolder;
  if (dateFolders.hasNext()) {
    finalFolder = dateFolders.next();
  } else {
    finalFolder = subFolder.createFolder(dateFolderName);
    Logger.log(`已建立新日期資料夾：${dateFolderName}`);
  }
  
  return finalFolder;
}
