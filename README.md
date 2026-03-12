# IP Winner 工具箱

IP Winner 智慧財產權管理的內部自動化工具集。

## 工具總覽

### 1. Streamlit Web App（根目錄）

部署於 Streamlit Cloud 的網頁工具，包含兩個功能：

| 功能 | 說明 |
|------|------|
| 📋 合併檔案 | 合併多個商標監控資料庫（Markify、摩知輪、Comp）的報告為統一 Excel |
| 📥 下載公開說明書 | 輸入專利號碼，自動從台灣專利公開資訊 API / Google Patents 下載 PDF |

- 版本：v14
- 部署說明：[部署說明.md](部署說明.md)

### 2. Email Auto-Processor（`apps-script/email-processor/`）

Google Apps Script，自動下載並歸檔 ip@ipwinner.com 轉寄的信件。

| 功能 | 說明 |
|------|------|
| 信件分類 | 從主旨提取慧盈案號，依類型分為專利/商標/未分類 |
| 附件下載 | 自動存入 Google Drive（類型/案號 資料夾結構） |
| Gmail 標籤 | 標示處理狀態（已下載/專利、已下載/商標 等） |
| 去重機制 | Message ID 層級追蹤，避免重複處理 |

- 版本：v2.1.0
- 詳細說明：[apps-script/email-processor/README.md](apps-script/email-processor/README.md)

## 技術架構

```
IPWinner/
├── streamlit_app.py              ← Streamlit Web App（Streamlit Cloud）
├── requirements.txt
├── 部署說明.md
└── apps-script/
    └── email-processor/
        ├── README.md
        └── IPWinner-EmailProcessor-V2.gs   ← Google Apps Script
```
