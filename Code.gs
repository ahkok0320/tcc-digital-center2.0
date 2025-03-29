/*****************************************************
 * Google Apps Script - 伺服器端
 *****************************************************/

/**
 * 初始化系統配置
 * 只在首次部署時執行
 */
function initializeConfig() {
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // 檢查配置是否已存在，若不存在則設置預設值
  if (!scriptProperties.getProperty('PARENT_FOLDER_ID')) {
    // 這裡設置預設值，實際部署時需要替換
    scriptProperties.setProperty('PARENT_FOLDER_ID', 'YOUR_PARENT_FOLDER_ID');
    scriptProperties.setProperty('STAT_SHEET_ID', 'YOUR_STAT_SHEET_ID');
    scriptProperties.setProperty('STAT_SHEET_NAME', '工作表1');
  }
}

/**
 * 獲取配置值
 */
function getConfigValue(key) {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(key);
}

/**
 * 主函式：處理 GET 請求，回傳前端 HTML
 */
function doGet() {
  // 確保配置已初始化
  initializeConfig();
  
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('【昶青數位資料中心APP】')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // 使用 IFRAME 模式
}

/**
 * 取得母資料夾底下所有子資料夾清單
 */
function getSubfolders() {
  var parentFolderId = getConfigValue('PARENT_FOLDER_ID');
  
  // 驗證用戶權限
  if (!validateUserAccess()) {
    return { error: "您沒有權限訪問此資源" };
  }
  
  try {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var folders = parentFolder.getFolders();
    
    var result = [];
    while (folders.hasNext()) {
      var folder = folders.next();
      result.push({
        id: folder.getId(),
        name: folder.getName()
      });
    }
    return result; // 回傳陣列 (JSON 格式)
  } catch (e) {
    Logger.log("獲取子資料夾時出錯: " + e.toString());
    return { error: "獲取資料夾列表失敗" };
  }
}

/**
 * 傳入資料夾 ID 和關鍵字，回傳符合關鍵字的檔案清單
 */
function getFiles(folderId, keyword) {
  // 驗證用戶權限
  if (!validateUserAccess()) {
    return { error: "您沒有權限訪問此資源" };
  }
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    
    var result = [];
    var lowerKeyword = keyword ? keyword.toLowerCase() : '';
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      if (!lowerKeyword || fileName.toLowerCase().indexOf(lowerKeyword) !== -1) {
        result.push({
          id: file.getId(),
          name: fileName,
          mimeType: file.getMimeType()
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log("獲取檔案時出錯: " + e.toString());
    return { error: "獲取檔案列表失敗" };
  }
}

/**
 * 只回傳檔案的預覽連結
 */
function getFilePreviewUrl(fileId) {
  // 驗證用戶權限
  if (!validateUserAccess()) {
    return { error: "您沒有權限訪問此資源" };
  }
  
  try {
    var file = DriveApp.getFileById(fileId);
    var mimeType = file.getMimeType();
    var previewUrl;

    if (mimeType === MimeType.PDF) {
      previewUrl = "https://drive.google.com/file/d/" + fileId + "/preview";
    } else if (
      mimeType === MimeType.GOOGLE_DOCS ||
      mimeType === MimeType.GOOGLE_SHEETS ||
      mimeType === MimeType.GOOGLE_SLIDES
    ) {
      previewUrl = file.getUrl(); // Google 原生檔案直接使用其 URL
    } else {
      previewUrl = file.getUrl(); // 其他檔案類型，提供下載連結
    }
    
    return {
      name: file.getName(),
      url: previewUrl,
      mimeType: mimeType
    };
  } catch (e) {
    Logger.log("獲取檔案預覽連結時出錯: " + e.toString());
    return { error: "獲取檔案預覽失敗" };
  }
}

/**
 * 將使用者檢索紀錄寫入 Google Sheet
 */
function logRetrieval(vendorName, fileName) {
  // 驗證用戶權限
  if (!validateUserAccess()) {
    return { error: "您沒有權限執行此操作" };
  }
  
  try {
    var sheetId = getConfigValue('STAT_SHEET_ID');
    var sheetName = getConfigValue('STAT_SHEET_NAME');
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    
    var timestamp = new Date();
    var user = Session.getActiveUser().getEmail();
    
    // A欄：登錄時間, B欄：廠商, C欄：品規書檔名, D欄：使用者
    sheet.appendRow([timestamp, vendorName, fileName, user]);
    return { success: true };
  } catch (e) {
    Logger.log("記錄檢索時出錯: " + e.toString());
    return { error: "記錄使用記錄失敗" };
  }
}

/**
 * 只回傳「近兩個月」的品規書使用統計資料（廠商數量統計、品規書數量統計）。
 * 不回傳任何可執行程式碼，只回傳 JSON，前端自行繪圖。
 */
function getStatisticsDataOnly() {
  // 驗證用戶權限
  if (!validateUserAccess()) {
    return { error: "您沒有權限訪問此資源" };
  }
  
  try {
    var sheetId = getConfigValue('STAT_SHEET_ID');
    var sheetName = getConfigValue('STAT_SHEET_NAME');
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues(); // 取得所有資料

    if (data.length <= 1) {
      // 只有表頭或尚無資料
      return {
        vendors: [],
        files: []
      };
    }

    var now = new Date();
    var twoMonthsAgo = new Date();
    twoMonthsAgo.setMonth(now.getMonth() - 2);

    var rows = data.slice(1); // 去掉表頭
    var vendorsCount = {};
    var filesCount = {};

    rows.forEach(function(row) {
      var ts = row[0];       // A欄: 時間
      var vendor = row[1];   // B欄: 廠商
      var file = row[2];     // C欄: 品規書
      var dateObj = new Date(ts);

      if (dateObj >= twoMonthsAgo && dateObj <= now) {
        // 計算廠商次數
        vendorsCount[vendor] = (vendorsCount[vendor] || 0) + 1;
        // 計算品規書次數
        filesCount[file] = (filesCount[file] || 0) + 1;
      }
    });

    // 轉成陣列形式給前端
    // e.g. [["廠商A", 10], ["廠商B", 5]]
    var vendorsArray = Object.keys(vendorsCount).map(function(k) {
      return [k, vendorsCount[k]];
    });
    var filesArray = Object.keys(filesCount).map(function(k) {
      return [k, filesCount[k]];
    });

    return {
      vendors: vendorsArray,
      files: filesArray
    };
  } catch (e) {
    Logger.log("獲取統計數據時出錯: " + e.toString());
    return { error: "獲取統計數據失敗" };
  }
}

/**
 * 驗證使用者是否有權限訪問系統
 * 可以根據實際需求調整驗證邏輯
 */
function validateUserAccess() {
  try {
    var user = Session.getActiveUser().getEmail();
    
    // 這裡可以添加更多的驗證邏輯，例如:
    // 1. 檢查使用者是否在允許清單中
    // 2. 檢查使用者是否屬於特定網域
    // 3. 其他自定義驗證邏輯
    
    // 簡單驗證：使用者必須已登入
    return user && user.length > 0;
  } catch (e) {
    Logger.log("驗證使用者權限時出錯: " + e.toString());
    return false;
  }
}

/**
 * 獲取當前登入的使用者信息
 */
function getCurrentUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    return { email: email };
  } catch (e) {
    return { error: "無法獲取使用者信息" };
  }
}
