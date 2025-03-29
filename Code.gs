/*****************************************************
 * Google Apps Script - 伺服器端 (Code.gs)
 *****************************************************/

/**
 * 主函式：處理 GET 請求，回傳前端 HTML
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('【昶青數位資料中心APP】')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // 使用 IFRAME 模式
}

/**
 * 取得母資料夾底下所有子資料夾清單
 * 母資料夾 ID：請自行修改
 */
function getSubfolders() {
  // 用佔位符取代原先的母資料夾ID
  var parentFolderId = 'YOUR_PARENT_FOLDER_ID'; // 請改成您真正的母資料夾ID
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
}

/**
 * 傳入資料夾 ID 和關鍵字，回傳符合關鍵字的檔案清單
 */
function getFiles(folderId, keyword) {
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
}

/**
 * 只回傳檔案的預覽連結
 */
function getFilePreviewUrl(fileId) {
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
}

/**
 * 將使用者檢索紀錄寫入 Google Sheet
 * Google Sheet ID請自行修改
 * Sheet 名稱請自行修改
 */
function logRetrieval(vendorName, fileName) {
  var sheetId = 'YOUR_GOOGLE_SHEET_ID'; // <--請改成您真正的 Google Sheet ID
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('工作表1');
  
  var timestamp = new Date();
  // A欄：登錄時間, B欄：廠商, C欄：品規書檔名
  sheet.appendRow([timestamp, vendorName, fileName]);
}

/**
 * 只回傳「近兩個月」的品規書使用統計資料（廠商數量統計、品規書數量統計）。
 * 不回傳任何可執行程式碼，只回傳 JSON，前端自行繪圖。
 */
function getStatisticsDataOnly() {
  var sheetId = 'YOUR_GOOGLE_SHEET_ID'; // <--請改成您真正的 Google Sheet ID
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('工作表1');
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
}
