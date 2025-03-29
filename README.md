# 昶青數位資料中心APP

這是一個使用 Google Apps Script 開發的網頁應用程式，用於查詢和瀏覽 Google Drive 中特定資料夾的檔案。

## 安裝步驟

1. 在 Google Apps Script 中創建新專案
2. 複製 `Code.gs` 和 `index.html` 到專案中
3. 在 `Code.gs` 的 `initializeConfig` 函數中填入您的 Google Drive 資料夾 ID 和 Google Sheet ID
4. 部署為網頁應用程式

## 安全注意事項

- 確保 Google Drive 資料夾和 Google Sheet 有適當的權限設置
- 此應用程式應僅供內部使用
