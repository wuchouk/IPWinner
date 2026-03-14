# IP Winner Email Processor V3 — 工程架構文件

> 版本：1.0
> 最後更新：2026-03-14
> 程式碼：`apps-script/Code.gs`（~2,672 行，單檔部署）

---

## 一、系統架構概覽

```
┌─────────────────────────────────────────────────────┐
│                  Google Apps Script                   │
│                    (Code.gs)                          │
│                                                       │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐           │
│  │  Setup &  │  │  Menu &  │  │ Scheduler│           │
│  │  Init     │  │  UI      │  │ & Trigger│           │
│  └──────────┘  └──────────┘  └──────────┘           │
│                                                       │
│  ┌───────────────────────────────────────────────┐   │
│  │            Core Processing Pipeline            │   │
│  │                                                │   │
│  │  Gmail Search → Preprocessing → LLM Classify  │   │
│  │  → Source ID → Naming → Drive Archive → Label │   │
│  └───────────────────────────────────────────────┘   │
│                                                       │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐           │
│  │ Feedback │  │  Prompt  │  │  Stats   │           │
│  │ Learning │  │  Mgmt    │  │ & Log    │           │
│  └──────────┘  └──────────┘  └──────────┘           │
└──────────┬──────────┬──────────┬──────────┬──────────┘
           │          │          │          │
      ┌────▼───┐ ┌───▼───┐ ┌───▼──┐ ┌────▼─────┐
      │ Gmail  │ │ Drive │ │Sheet │ │ Gemini   │
      │ API    │ │ API   │ │ API  │ │ 3.0 Flash│
      └────────┘ └───────┘ └──────┘ └──────────┘
```

### 部署模式
- **Standalone Script**：不綁定特定 Spreadsheet
- `SpreadsheetApp.create()` + Drive 搜尋，可多租戶共用
- 透過 `setupAll()` 一鍵初始化所有資源

### 外部依賴
| 服務 | 用途 | API |
|------|------|-----|
| Gmail | 讀取郵件、管理標籤 | GmailApp |
| Google Drive | EML/附件存檔 | DriveApp |
| Google Sheets | 設定、紀錄、Sender 名單 | SpreadsheetApp |
| Google Docs | LLM Prompt 匯出 | DocumentApp |
| Gemini 3.0 Flash | 語義分類 + 命名 | UrlFetchApp (REST) |

---

## 二、資料流

### 主處理流程

```
Gmail Inbox
    │
    ▼
[1] Gmail Search（query: newer_than:7d 等）
    │  排除已處理 messageId（Sheet 比對）
    ▼
[2] Preprocessing（規則引擎）
    ├── 抽取案號（CASE_NUMBER_REGEX + word boundary）
    ├── 判斷 F/T（sender ∈ OWN_DOMAINS?）
    ├── 查 Sender 名單 → role = C/A/G/X
    ├── 去除 RE:/Fwd: 前綴
    ├── 抽取附件名稱列表
    └── 提取 HTML highlights（粗體/上色文字）
    │
    ▼
[3] LLM Batch Classify（UrlFetchApp.fetchAll）
    ├── 每批最多 20 封同時送出
    ├── Input: structured email info + 高頻模板 + 近期修正
    ├── Output: JSON（收發碼、語義名、類別、信心、inferred_role）
    └── _repairJson() 修復截斷回應
    │
    ▼
[4] Source Identification（僅 role=X）
    ├── confidence ≥ 0.6 + inferred_role → 改具體碼 + AI/自動辨識來源
    ├── confidence < 0.6 → 保留 FX/TX + AI/未知來源
    └── 完全無法判斷 → AI/已跳過
    │
    ▼
[5] Naming（組合檔名）
    └── {yyyyMMdd}-{收發碼}-{案號標記}-{語義名}.eml
    │
    ▼
[6] Drive Archive（逐封處理，單線程）
    ├── 定位/建立案號資料夾
    ├── 重複偵測（同名 + 同大小 ±100 bytes → skip）
    ├── getRawContent() → createFile()
    ├── 附件下載（F 方向 + FX）
    └── 多案號 → 複製到每個案號資料夾
    │
    ▼
[7] Label & Record
    ├── Gmail: 加收發碼 + 類別 + 狀態標籤
    └── Sheet: 寫入處理紀錄（20 欄）
```

### 回饋流程

```
人工操作                         系統偵測（runFeedback）
───────                         ──────────────────

[Part 1] Gmail 標籤修正
移除 AI/自動辨識來源        →   偵測標籤移除
修改 AI/FX → AI/FA         →   比對原始收發碼 vs 當前標籤
                            →   更新 Sender 名單（_addSender 含去重）
                            →   更新 Sheet（來源確認狀態、修正來源）

[Part 2] Sheet 修正名稱
填入「修正後名稱」欄        →   偵測欄位有值
                            →   組出新 baseName
                            →   Drive EML/附件改名（跨多案號資料夾）
                            →   更新修正來源（+ name_change）

[Part 3] Sheet 直填收發碼
填入「最終收發碼」欄        →   偵測欄位有值
                            →   用收件人/寄件人查對應 email
                            →   寫入 Sender 名單
                            →   更新修正來源（+ sheet_edit）
```

---

## 三、核心資料結構

### CONFIG 常數

```javascript
CONFIG = {
  PROJECT_FOLDER_NAME: 'Email自動整理v2',
  SPREADSHEET_NAME: 'Email自動整理v2-設定檔',
  GEMINI_MODEL: 'gemini-3-flash-preview',
  BATCH_SIZE: 20,
  CONFIDENCE_AUTO: 0.8,      // 自動處理閾值
  CONFIDENCE_INFER: 0.6,     // 來源推斷閾值
  CONFIDENCE_LOW: 0.5,       // 低信心閾值
  BODY_SNIPPET_LENGTH: 1500,
  TIMEOUT_SAFETY_MS: 25 * 60 * 1000,  // 25 分鐘安全閾值
  CASE_NUMBER_REGEX: /(?<![A-Za-z0-9])[A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*(?![A-Za-z0-9])/g,
  OWN_DOMAINS: ['ipwinner.com', 'ipwinner.com.tw'],
  SEND_RECEIVE_CODES: ['FC', 'TC', 'FA', 'TA', 'FG', 'TG', 'FX', 'TX'],
  CASE_CATEGORIES: ['專利', '商標', '未分類'],
}
```

### Google Sheets 結構

**Sheet 1: Sender 名單**
| 欄位 | 說明 |
|------|------|
| Email 或 Domain | `@bskb.com` 或 `tanaka@gmail.com` |
| 角色（C/A/G） | 客戶/代理人/政府 |
| 名稱備註 | 如 BSKB - 美國代理人 |

查詢優先序：email 精確匹配 → domain 匹配。公共 domain（gmail.com 等）強制用完整 email。

**Sheet 2: 處理紀錄（20 欄）**
```
 0: messageId        10: AI案件類別
 1: 日期              11: 來源確認狀態（na/pending/confirmed/corrected）
 2: 原始標題          12: 最終收發碼
 3: sender           13: 修正後名稱
 4: AI收發碼          14: 修正原因
 5: AI推斷角色        15: 修正時間
 6: 歸檔案號          16: 修正來源（tag_change/name_change/sheet_edit）
 7: 內文案號          17: 重試次數
 8: AI語義名          18: Input Tokens
 9: AI信心            19: Output Tokens
```

**Sheet 3: 分類規則**（8 分類 29 條，setupAll 自動寫入）

**Sheet 4: 設定**（信心閾值、batch 大小、checkpoint 等）

### Drive 資料夾結構

```
Email自動整理v2/
├── 專利/
│   ├── BRIT21002PUS5/
│   │   ├── 20260314-FA-BRIT21002PUS5-送件報告-(ROA1).eml
│   │   └── 20260314-FA-BRIT21002PUS5-送件報告-(ROA1)-附件1.pdf
│   └── BRIT24001DUS1/
├── 商標/
│   ├── KOIT20004TCN7/
│   └── BRIT25001TJP1/
├── 未分類/
│   ├── BRIT/          ← 有客戶碼但無完整案號
│   └── 無案號/        ← 完全無案號
└── LLM Prompt 文件    ← Google Doc
```

### Gmail 標籤樹

```
AI/
├── FC, TC, FA, TA, FG, TG, FX, TX    ← 收發碼（互斥）
├── 專利, 商標, 未分類                   ← 案件類別（互斥）
├── 多案號, 無案號                       ← 案號狀態（0-1）
├── 待確認                               ← 信心 < 0.8
├── 自動辨識來源                         ← LLM 推斷成功，待人確認
├── 未知來源                             ← LLM 無法推斷
├── 已跳過                               ← 完全無法處理
├── 附件下載錯誤                         ← 附件失敗，自動重試
└── 處理失敗                             ← 系統錯誤，自動重試
```

---

## 四、狀態機

### 信件處理狀態

```
                    ┌─────────┐
                    │ 未處理   │
                    └────┬────┘
                         │ processEmails / trialRun
                         ▼
              ┌─────────────────────┐
              │ 規則引擎前處理       │
              │ (案號/方向/角色)      │
              └──────────┬──────────┘
                         │
              ┌──────────▼──────────┐
              │ LLM 分類             │
              └──────────┬──────────┘
                         │
          ┌──────────────┼──────────────┐
          ▼              ▼              ▼
    conf ≥ 0.8     0.5 ≤ conf < 0.8   conf < 0.5
    ┌────────┐     ┌────────────┐    ┌──────────┐
    │ 自動處理│     │ AI/待確認  │    │ AI/待確認│
    │        │     │ (LLM命名)  │    │ (原標題) │
    └───┬────┘     └─────┬──────┘    └────┬─────┘
        │                │                │
        ▼                ▼                ▼
    ┌────────────────────────────────────────┐
    │        Drive 歸檔 + Gmail 標籤         │
    └────────────────────┬───────────────────┘
                         │
                         ▼
                    ┌─────────┐
                    │ 已處理   │ (Sheet 紀錄)
                    └────┬────┘
                         │ runFeedback
                         ▼
              ┌─────────────────────┐
              │ 回饋修正（可選）     │
              │ 標籤/名稱/收發碼    │
              └─────────────────────┘
```

### Sender 來源確認狀態機

```
                ┌──────┐
                │  na  │ ← role 已知（在 Sender 名單中）
                └──────┘

sender 不在名單 + LLM 推斷成功
                ┌─────────┐
                │ pending │ ← AI/自動辨識來源
                └────┬────┘
                     │
          ┌──────────┼──────────┐
          ▼                     ▼
    ┌───────────┐        ┌───────────┐
    │ confirmed │        │ corrected │
    │ (人確認)   │        │ (人修正)   │
    └───────────┘        └───────────┘
    移除標籤 →            改收發碼標籤 →
    Sender 加入名單       Sender 以修正角色加入名單
```

---

## 五、效能特性

### LLM 呼叫（最大優化點）
- `UrlFetchApp.fetchAll()` 批次並行：20 封 ~6 秒
- 逐封呼叫同量約 60 秒（10x 差異）

### Drive 操作（瓶頸）
- Apps Script 單線程，無法並行
- 資料夾定位：1-3 秒/首次
- getRawContent()：0.2 秒（一般）～ 1.7 秒（25MB）
- createFile()：1.2 秒（一般）～ 3.4 秒（25MB）
- 10 封信 Drive 總計：30-40 秒

### Timeout 防護
- Google Workspace：30 分鐘執行限制
- 安全閾值：25 分鐘
- 每批 20 封寫入 Sheet，更新 checkpoint
- 接近閾值時自動排下一個 trigger

---

## 六、錯誤處理

### 重試機制
| 錯誤類型 | 處理方式 | 上限 |
|---------|---------|------|
| API 呼叫失敗 | 自動重試 | 3 次 |
| 附件下載失敗 | 標記 AI/附件下載錯誤 + 重試 | 3 次 |
| JSON 解析失敗 | _repairJson() 修復 | 1 次 |
| 系統錯誤 | 標記 AI/處理失敗 + 重試 | 3 次 |

### 重跑保護
- 同檔名 + 同大小（±100 bytes）→ 跳過不重建
- MessageId 去重：Sheet 已有的 messageId 不重複處理

### 已知 Bug 模式（避雷）
| Bug | 根因 | 已修復 |
|-----|------|--------|
| 假案號從 base64 匹配 | regex 無邊界 | ✅ lookbehind/lookahead |
| T 方向全變 TX | 查寄件人而非收件人 | ✅ 改查外部收件人 |
| 類型碼誤判 | regex search 先匹到客戶碼 | ✅ 用固定位置 index 9 |
| 多案號建 19 個資料夾 | 內文案號全觸發 | ✅ 只看主旨 + LLM |
| 重跑產生重複檔案 | 無偵測 | ✅ 同名同大小 skip |
| 公共 email 加為客戶 | @gmail.com 代表整個 gmail | ✅ PUBLIC_DOMAINS |
| T 方向回饋學到自己 | 用寄件人學習 | ✅ 用收件人 |
| Drive 改名日期偏移 | Sheet Date 時區轉換 | ✅ getSpreadsheetTimeZone() |
| Sender 名單重複 | 無去重 | ✅ 寫入前掃描 |
| 回授互相覆蓋 | Part 2 覆蓋 Part 1 | ✅ indexOf + 串接 |
| TC 用過期期限 | LLM 不知信件日期 | ✅ 傳入 email_date |

---

## 七、LLM Prompt 架構

### System Prompt 結構
```
[角色定義] → IP Winner 智財事務所 email 分類助手
[輸入格式] → subject/direction/role/sender/recipients/case_numbers/
              body_snippet/attachment_names/email_date
[第一步] → 確認收發碼 + 未知來源角色推斷
[第二步] → 產生語義檔名（前綴引導 + 自由摘要 + 截止日）
[第三步] → 期限過期判斷（email_date 比對）
[第四步] → 案件類別判斷
[第五步] → 案號狀態判斷
[高頻模板] → %%TEMPLATES%%（動態注入 80 個）
[近期修正] → %%CORRECTIONS%%（動態注入最近 20 筆）
[輸出格式] → JSON schema
```

### Few-shot 學習循環
```
人工修正 → Sheet 處理紀錄 → 最近 20 筆注入 prompt
    ↑                                          │
    └──── LLM 命名改善 ◄──────────────────────┘
```

### 每週整理
- Trigger: 每週一 9:00（weeklyConsolidate）
- 流程：收集 correction log → LLM 歸納模式 → 寫入 Prompt Doc
- 人工審核後決定是否併入 SYSTEM_PROMPT

---

## 八、測試策略

### 現有測試機制
| 模式 | 函式 | 用途 |
|------|------|------|
| 快速驗證 | `trialRunSmall()` | 10 封，檢查基本功能 |
| 完整測試 | `trialRun()` | 50 封，Phase 1 驗收 |
| 單封分析 | `testSingleEmail()` | 詳細 log，不寫入 Sheet |

### Phase 1 驗收 Checklist
- [ ] 語義名正確（特別是期限日期）
- [ ] 收發碼判定正確
- [ ] 歸檔案號 vs 內文案號正確區分
- [ ] Drive 資料夾結構和檔名正確
- [ ] 修正名稱 → Drive 改名（時區 bug 已修）
- [ ] Sender 去重（同 sender 多封信）
- [ ] 匯出 LLM Prompt 文件成功
- [ ] 整理學習紀錄功能正常

### Log 設計
- 每封信獨立 log（失敗不影響其他封）
- Drive 操作分段計時
- Gemini API 並行時間 vs Drive 逐封時間分開顯示
- Token 用量 per-email 追蹤（Sheet col 18-19）

---

## 九、安全性考量

| 項目 | 處理方式 |
|------|---------|
| API Key | Script Properties 存儲，不在程式碼中 |
| OAuth Scopes | 最小權限原則（gmail.modify + drive + sheets + external_request） |
| 資料隔離 | 所有資料在客戶自己的 Google Workspace 內 |
| PII | 郵件內容不持久化（只存 snippet 在 Sheet，EML 在 Drive） |

---

## 十、SaaS 產品化路線

### 需拆分的模組
```
apps-script/
├── config.gs          ← CONFIG + 可自訂設定
├── preprocessing.gs   ← 案號提取、方向判斷、角色查詢
├── llm.gs             ← Gemini API 呼叫、prompt 管理
├── drive.gs           ← Drive 歸檔、改名、重複偵測
├── sheet.gs           ← Sheet 讀寫、回饋偵測
├── feedback.gs        ← 三通道回饋邏輯
├── labels.gs          ← Gmail 標籤管理
├── ui.gs              ← 選單、對話框
└── main.gs            ← setupAll、processEmails 入口
```

### 需可設定化的項目
- 案號 regex 格式
- 類型碼位置和分類映射
- 語義名前綴列表（per 收發碼）
- 代理人帳單代碼對照
- OWN_DOMAINS / PUBLIC_DOMAINS / GOV_DOMAINS
- 信心閾值
- Drive 資料夾命名規則

---

*文件維護：每次 QA 完成後更新架構變更與測試結果*
