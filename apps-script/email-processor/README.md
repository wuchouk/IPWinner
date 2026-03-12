# Email Auto-Processor V2

Google Apps Script 自動下載並歸檔 IP Winner 轉寄的信件。

## 功能

- 自動搜尋 `ip@ipwinner.com` 收發的信件（預設 48 小時內）
- 從主旨提取慧盈案號（三層提取：慧盈案號標記 → Ref 欄位 → Fallback regex）
- 依案號類型分類：專利（P/M/D/A/C）、商標（T/B/W）、未分類
- 附件下載至 Google Drive，自動建立 `類型/案號/` 資料夾結構
- 信件內容存為 `.eml` 檔案
- Gmail 標籤標示處理狀態（已下載/專利、已下載/商標、已下載/未分類、已跳過）
- Message ID 層級去重（透過 Google Sheets 追蹤）
- 可設定過濾 Template（自動回覆信等）

## 運行環境

- Google Apps Script（綁定在 Google Sheets `Email自動整理-設定檔`）
- 執行帳號：peter@ipwinner.com

## 設定檔（Google Sheets）

| 工作表 | 用途 |
|--------|------|
| 過濾Template | 信件內容比對，符合的信件會被跳過 |
| 執行紀錄 | 每次執行的統計數據 |
| 已處理信件 | Message ID 追蹤，避免重複處理 |

## 版本

目前版本：**v2.1.0**（2026-03-12）

詳細 changelog 請見 `.gs` 檔案頂部註解。
