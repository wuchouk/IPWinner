# IP Winner 工具箱

IP Winner 智慧財產權事務所的內部自動化工具集。

## 工具總覽

### 1. Email Processor V3（[`email-processor/`](email-processor/)）

Google Apps Script 智慧郵件分類與歸檔系統，結合規則引擎 + Gemini LLM。

| 功能 | 說明 |
|------|------|
| 自動分類 | 8 種收發碼（FC/TC/FA/TA/FG/TG/FX/TX）+ 案件類別（專利/商標/未分類） |
| 智慧命名 | LLM 產生語義化檔名（前綴引導 + 自由摘要，含截止日與事項碼） |
| 自動歸檔 | Drive 資料夾結構（類別 → 案號 → EML + 附件） |
| 回饋學習 | Gmail 標籤修正 / Sheet 修正名稱 / Sheet 直填收發碼 → Few-shot 注入 |

- 版本：V3
- 技術：Google Apps Script + Gmail + Drive + Sheets + Gemini 3.0 Flash
- 規則引擎準確率：98.8%，AI 月費約 NT$160
- 詳細文件：[產品計畫](email-processor/docs/product-plan.md) / [工程架構](email-processor/docs/engineering-plan.md)

### 2. 合併工具 + 說明書下載（[`merge-tool/`](merge-tool/)）

部署於 Streamlit Cloud 的網頁工具。

| 功能 | 說明 |
|------|------|
| 合併檔案 | 合併多個商標監控資料庫（Markify、摩知輪、Comp）的報告為統一 Excel |
| 下載公開說明書 | 輸入專利號碼，自動從台灣專利公開資訊 API / Google Patents 下載 PDF |

- 版本：v14
- 部署說明：[merge-tool/部署說明.md](merge-tool/部署說明.md)

## 目錄結構

```
IPWinner/
├── README.md
├── LICENSE
├── logo.jpg
│
├── email-processor/                 ← Email Processor V3
│   ├── Code.gs                      ← 主程式（~2,672 行，單檔部署）
│   ├── docs/
│   │   ├── product-plan.md          ← 產品計畫
│   │   └── engineering-plan.md      ← 工程架構
│   ├── TODO.md                      ← 任務追蹤
│   ├── email-processor-development-principles.md
│   ├── V3-分類規則與LLM-Prompt設計.md
│   ├── V3-分類測試與成本效益分析.md
│   ├── V3-分類比對表-TUS1+PUS5+PTW1.xlsx
│   ├── process_download_area.py     ← Python 輔助腳本（V1/V2 遺留）
│   └── SESSION-2026-03-13.md
│
└── merge-tool/                      ← Streamlit 合併工具
    ├── streamlit_app.py
    ├── requirements.txt
    └── 部署說明.md
```
