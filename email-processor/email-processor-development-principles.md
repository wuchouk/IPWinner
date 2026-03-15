# IP Winner Email Processor — 開發原則與學習紀錄

> 從 V1 → V3 的實戰經驗，為未來 SaaS 產品化奠定基礎

---

## 一、架構原則

### 1.1 單檔部署 vs 模組化
- V3 選擇單檔 Code.gs（~1800 行），方便客戶「貼上即用」
- SaaS 化時應拆成模組：preprocessing / llm / drive / sheet / feedback
- 但 Apps Script 無原生 import，需用 clasp + TypeScript 或改用 Cloud Functions

### 1.2 Standalone vs Container-bound
- 用 `SpreadsheetApp.create()` + Drive 搜尋，不依賴 `getActiveSpreadsheet()`
- 好處：Script 與 Sheet 解耦，可多租戶共用同一份 Script
- SaaS 化時：每個客戶一組 Sheet + Drive 資料夾，Script 統一管理

### 1.3 自動初始化
- `setupAll()` 一鍵建立 Sheet（4 tabs）、Drive 資料夾、Gmail 標籤
- 客戶不需手動建任何結構，降低上手門檻
- SaaS 化時：onboarding flow 自動化這一步

---

## 二、Gmail 處理原則

### 2.1 訊息去重用 Message ID，不用 Label 排除
- **V1 教訓**：`-label:AI/已處理` 會排除同 thread 的新回覆
- **正解**：Sheet 記錄已處理的 Message ID，搜尋時不加 label 排除條件
- Gmail thread model 特性：label 加在 thread 上，不是個別 message

### 2.2 Gmail Label 有同步延遲
- 加完 label 後，Gmail UI 可能要幾秒到幾分鐘才顯示
- 這是 Gmail API 的已知快取行為，無法從 Apps Script 端解決
- 不要依賴 label 做即時邏輯判斷

### 2.3 getRawContent() 的效能特性
- 大部分信件 0.2 秒左右，但大附件信（25MB）可能 1.7 秒
- 無法並行（Apps Script 單線程限制）
- 同一封信多案號時，取一次 rawContent 重用，不要重複呼叫

---

## 三、案號系統設計

### 3.1 案號結構（IP Winner 格式）
```
[4碼客戶號][2碼年份][3碼序號][1碼類型][2碼國碼][選填分案號]
例：BRIT25710PUS1
     BRIT = 客戶碼
     25 = 年份
     710 = 序號
     P = 專利（P/M/D/A/C=專利, T/B/W=商標）
     US = 國碼
     1 = 分案號
```

### 3.2 類型碼分類規則
- **必須用固定位置（index 9）取類型碼**，不能用 regex search
- V3 bug：`BRIT` 的 T 被 regex 先匹配到，誤判為商標
- 分類：PMDAC → 專利 | TBW → 商標 | 混合或無 → 未分類

### 3.3 案號來源區分（核心設計決策）
- **主旨案號**：決定資料夾歸檔、多案號、專利/商標分類
- **內文案號**：僅記錄在 Sheet，不自動觸發歸檔
- 原因：內文常引用參考案號（如異議依據的先前商標），不應全部歸檔
- **LLM 判斷 filing_case_numbers**：由 Gemini 分析哪些是真正要歸檔的案號

### 3.4 多案號判定
- 主旨 2+ 個不同案號 → 多案號
- 主旨 1 個案號 + 「等」→ 多案號（去內文找其他案號）
- 「～」「~」「、」不需額外判斷（主旨有這些符號時 regex 自然會抓到 2+ 案號）
- 內文有再多案號也不算多案號（除非 LLM 判定需要歸檔）

### 3.5 無案號信件的歸檔
- 主旨無案號但內文有 → `未分類/{客戶碼前4碼}/`（如 `未分類/BRIT/`）
- 完全無案號 → `未分類/無案號/`
- Tag 上「無案號」

---

## 四、收發碼系統

### 4.1 八碼設計
- 方向（F/T）× 角色（C/A/G/X）= FC/TC/FA/TA/FG/TG/FX/TX
- F = From（收到）、T = To（寄出）
- C = Client（客戶）、A = Agent（代理人）、G = Government（政府）、X = Unknown

### 4.2 T 方向的角色判定
- **V3 bug**：寄出信檢查寄件人（ipwinner.com）→ 永遠是 X → 全變 TX
- **正解**：T 方向要看「第一個外部收件人」的角色，不是寄件人
- 回饋學習也一樣：T 方向用收件人作為學習對象

### 4.3 Sender 名單設計
- 私人 domain（如 bskb.com）→ 用 @domain 代表整間公司
- 公共 email（gmail.com 等）→ 用完整 email（如 john@gmail.com）
- OWN_DOMAINS（ipwinner.com）→ 永遠不加入 Sender 名單

---

## 五、Drive 存檔原則

### 5.1 檔名規則
- EML：`{日期}-{收發碼}-{案號標記}-{AI語義名}.eml`
- 附件：`{日期}-{收發碼}-{案號標記}-{AI語義名}-附件N.{副檔名}`
- 多案號：每個案號資料夾都存一份實體 EML（不用捷徑）

### 5.2 重跑保護
- 同檔名 + 同大小（±100 bytes）→ 跳過不重建
- 不需要手動刪檔就能安全重跑
- 掃描資料夾現有檔案一次，EML 和附件共用

### 5.3 TA/TC/TG 附件規則
- 寄出的信只存 EML，不存附件
- 原因：附件是我方自己寄出的，本地已有原始檔

### 5.4 Drive 效能瓶頸
- 資料夾定位：1-3 秒（第一次最慢）
- getRawContent：0.2 秒（一般）～ 1.7 秒（25MB）
- createFile：1.2-1.5 秒（一般）～ 3.4 秒（25MB）
- **無法並行**：Apps Script 單線程，Drive API 也不支援 fetchAll
- 10 封信 Drive 總計約 30-40 秒，這是正常速度

---

## 六、LLM（Gemini）整合原則

### 6.1 並行呼叫
- `UrlFetchApp.fetchAll()` 一次送出多封信的 API 請求
- 10 封信約 6 秒完成（vs 逐封呼叫約 30 秒）
- 是目前最大的效能優化點

### 6.2 語義檔名的期限判斷
- 信件常有多個日期（官方期限、我方期限、申請日等）
- 規則：用「行動期限」不是「背景日期」
  - TA → 我方要求代理人的期限
  - TC → 我方要求客戶的期限
  - FA → 代理人要求我方的期限
- HTML highlight 提取：`_extractHighlights()` 抓取被加粗/上色的文字給 LLM 判斷
- V3.1 改進：結構化期限提取 — LLM 輸出 `dates_found`（所有日期+分類為 4 種 type），程式碼 `_selectDeadline()` 根據收發碼自動選日期
  - 解決 Gemini Flash 無法正確區分官方期限 vs 我方期限的問題
  - 4 種 type：`official_deadline`、`our_request`、`counterpart_eta`、`background`
  - `eml_filename` 不含日期，由程式碼附加

### 6.3 filing_case_numbers（歸檔案號）
- LLM 判斷信件實際應歸檔的案號，區分「主要處理的案號」和「順便引用的案號」
- 範例：異議申請主案 KOIS23004BGB5，引用先前商標 KOIS23004TGB1 → 只歸主案
- Fallback：LLM 沒回傳時用主旨案號

### 6.4 JSON 修復
- Gemini 有時會截斷回應（reasoning 太長）
- `_repairJson()` 嘗試修復常見的截斷模式
- 設定 `responseMimeType: 'application/json'` 可降低格式錯誤機率

### 6.5 Prompt 設計要點
- reasoning 限制 30 字以內（避免截斷）
- 語義名 25 字以內
- 用具體範例比抽象規則有效
- 近期人工修正紀錄（%%CORRECTIONS%%）是持續改善的關鍵

---

## 七、回饋學習系統

### 7.1 自動辨識來源 → 人工確認 → 學習
- LLM 推斷 sender 角色（信心 ≥ 0.6）→ 標記「自動辨識來源」
- 人類確認後寫入 Sender 名單
- 下次遇到同 sender 就不需要 LLM 推斷

### 7.2 回饋偵測機制
- 監控被移除「自動辨識來源」label 的信件
- Gmail label 有頁面快取問題：停在收件匣頁面偵測不到，要切換頁面或重新整理

---

## 八、智慧附件過濾（V1 教訓）

### 8.1 簽名檔圖片過濾
- < 5KB 的圖片 → 跳過
- 檔名含 image00/logo/banner/signature → 跳過
- 用 `getAttachments({ includeInlineImages: false })` 排除基本 inline 圖

### 8.2 進階判斷
- 如果 HTML body 的主文區域有 `cid:` 引用 → 保留 inline 圖（可能是截圖）
- 排除 gmail_quote 和 gmail_signature 區域內的圖片

---

## 九、SaaS 產品化建議

### 9.1 多租戶架構
- 每個客戶需要自訂：OWN_DOMAINS、Sender名單、案號格式、分類規則
- 「分類規則」tab 讓客戶可以 review 和修改商業邏輯
- 規則應從 Sheet 讀取而非寫死在程式碼

### 9.2 案號格式可設定化
- 目前 regex 寫死 IP Winner 的格式
- SaaS 需支援不同客戶的案號格式（如 4 碼 vs 5 碼客戶號）
- 類型碼位置和分類映射也要可設定

### 9.3 Prompt 模板化
- 語義檔名的前綴列表（FC/TC/FA 等）應可自訂
- 代理人帳單代碼對照表應從 Sheet 讀取
- 每個客戶的分類偏好不同

### 9.4 效能考量
- Apps Script 有 6 分鐘執行時限 → 大量信件需分批 + trigger
- 考慮 Cloud Functions + Cloud Tasks 做非同步處理
- Drive API quota：每 100 秒 12,000 次請求（通常不是瓶頸）

### 9.5 計費指標
- Token 用量追蹤（已實作 per-email tracking）
- Drive 儲存空間
- 處理封數

---

## 十、常見 Bug 模式（避雷清單）

| Bug | 根因 | 修法 |
|---|---|---|
| 假案號從 base64 匹配 | regex 無邊界檢查 | 加 lookbehind/lookahead |
| T 方向全變 TX | 檢查寄件人（自己）而非收件人 | T 方向看外部收件人 |
| 分類錯誤（專利→商標） | regex search 先匹配到客戶碼裡的 T | 用固定位置 index 9 |
| AI 語義名沒用到 | _determineFinalResult 沒傳 emlFilename | 加入 result 物件 |
| 多案號建 19 個資料夾 | 內文案號全部觸發歸檔 | 只看主旨 + LLM 判斷 |
| 重跑產生重複檔案 | 無重複偵測 | 同名+同大小 skip |
| 公共 email 加為客戶 | @gmail.com 代表整個 gmail | PUBLIC_DOMAINS 用完整 email |
| 回饋學加到自己 | T 方向用寄件人學習 | T 方向用收件人 |
| 信心顯示 1% | LLM 回 0-1 scale，Logger 沒 ×100 | Math.round(conf × 100) |
| 期限用錯日期 | 沒區分官方期限 vs 行動期限 | Prompt 加期限選擇規則 |
| 回授改名找不到 Drive 檔案 | Sheet Date 經 Sheets 時區轉換後日期偏移 | 用 `ss.getSpreadsheetTimeZone()` 讀回日期 |
| Sender 名單重複寫入 | `_addSender` 無去重檢查 | 寫入前掃描名單，已存在就更新 |
| 回授互相覆蓋修正來源 | Part 1 寫完後 Part 2 檢查空值就跳過 | 改用 `indexOf('name')` 檢查 + `+` 串接 |
| TA 委託信語義名漏期限 | Prompt 對「何時加期限」不夠強制 | 截止日改為「必加」規則 + TA 委託範例 |
| OA1 在語義名出現兩次 | 前綴和事項碼重複 | 加去重規則：二選一精簡 |
| TC 信用了已過期的回覆期限 | LLM 不知道信件日期、無法判斷過期 | 傳入 email_date + 第三步過期判斷邏輯 |
| OA-(OA1) 語義重複 | 描述文字和事項碼都出現 OA | 描述改用中文「審查意見」「答辯」，OA 只留在括號 |
| TA 期限選錯（官方 vs 我方） | Gemini Flash 無法遵循 TA 期限規則 | P2 結構化期限提取：LLM 只分類，程式碼選日期 |

---

## 十一、測試策略

### 11.1 試跑機制
- `trialRunSmall(10封)` → 快速驗證
- `trialRun(50封)` → 完整 Phase 1 測試
- `testSingleEmail()` → 單封信詳細分析（不寫入 Sheet）

### 11.2 Log 設計
- 每封信獨立 log（失敗不影響其他封）
- Drive 操作分段計時（資料夾定位/掃描/getRawContent/createFile）
- Gemini API 並行時間 vs Drive 逐封時間分開顯示

### 11.3 回饋測試
- 刪除 label → 切換頁面 → 跑回饋偵測 → 檢查 Sender 名單
- Gmail label 同步有延遲，需要等待或重新整理

---

---

## 十二、回饋學習系統 V2（2026-03-13 新增）

### 12.1 三段式回饋偵測（runFeedback）
- **Part 1**：Sender 角色 — 偵測 Gmail label 變化（移除「自動辨識來源」→ 確認/修正）
- **Part 2**：語義名改名 — 偵測 Sheet「修正後名稱」欄 → Drive EML/附件改名
- **Part 3**：Sheet 直填收發碼 — 偵測「最終收發碼」欄 → 寫入 Sender 名單
- 三段獨立運作，修正來源用 `+` 串接（如 `tag_change+name_change`）

### 12.2 修正來源追蹤
- 每段回饋獨立檢查是否已執行（用 `indexOf` 而非空值判斷）
- Part 2 重新讀取 freshData（Part 1 可能已更新修正來源欄）
- 防止互相覆蓋的關鍵

### 12.3 Drive 改名的精準比對
- 從 Sheet 組出完整 baseName（日期+收發碼+案號標記+語義名）
- 多案號時組出多個 baseName（每個歸檔案號一個）
- 搜尋所有分類資料夾下的案號子資料夾
- **時區陷阱**：讀回 Sheet Date 物件時必須用 `ss.getSpreadsheetTimeZone()`，不能用 `Asia/Taipei`

### 12.4 Sender 去重
- `_addSender` 寫入前掃描 Sender 名單第一欄
- 已存在 + 角色不同 → 更新；角色相同 → 跳過；不存在 → 新增

### 12.5 歸檔案號 vs 內文案號
- Sheet col 6「歸檔案號」：實際建資料夾的案號（用於 Drive 改名比對）
- Sheet col 7「內文案號」：所有偵測到的案號（僅供參考）
- LLM 的 `filing_case_numbers` 決定歸檔案號

---

## 十三、LLM Prompt 管理（2026-03-13 新增）

### 13.1 SYSTEM_PROMPT 是唯一真實來源
- Code.gs 裡的 `SYSTEM_PROMPT` 常數是 LLM 行為的唯一真實來源
- 「匯出 LLM Prompt 文件」產生 Google Doc 唯讀副本（在專案資料夾內）

### 13.2 期限選擇規則（三步驟）
1. 辨識所有日期：官方期限 / 我方期限 / 背景日期
2. 根據收發碼選正確期限（TA 用我方要求的，絕不用官方的）
3. 用 email_date 檢查是否過期 → 過期改用官方期限 + 語氣改「通知官方期限」

### 13.3 截止日是強制規則
- 只要信件有行動期限就**必須**加在語義名中
- TA 委託信的 "prepare draft by"、"for review by" 都是期限
- 常被 LLM 省略，需要在 prompt 中明確列出 TA 委託範例

### 13.4 括號事項碼去重
- 前綴和括號事項碼不可重複同一資訊
- 例：「提醒回覆指示-OA分析-(OA1)」→ 精簡為「提醒回覆指示-(OA1)」或「提醒回覆指示-OA1分析」
- 描述文字避免英文縮寫「OA」，改用中文「審查意見」「答辯」
- 「OA」只出現在括號事項碼中如 -(OA1)、-(ROA1)

### 13.5 學習紀錄自動整理
- Few-shot：最近 20 筆修正紀錄（含原因）注入每次 LLM 呼叫
- 每週一自動執行 `weeklyConsolidate`：LLM 歸納修正紀錄 → 寫入 Prompt Doc
- 歸納出的規則需人工確認後加入 SYSTEM_PROMPT

### 13.6 修正原因的價值
- Sheet「修正原因」欄的文字直接出現在 few-shot 範例中
- 例：`（原因：TA委託信「for review by 3/17」是我方要求代理人的期限，不可省略）`
- 原因越具體，LLM 學到的規則越精準

---

## 十四、分類規則 Sheet（2026-03-13 新增）

### 14.1 設計目的
- 讓非技術人員（客戶/同事）能閱讀和理解分類邏輯
- 8 個色碼分類，29 條規則，每組前有合併儲存格的標題行
- `setupAll()` 時自動寫入，每次重跑會覆蓋

### 14.2 八大分類
| 類別 | 規則ID | 說明 |
|---|---|---|
| 案號結構 | C01-C03 | regex、固定位置取類型碼、邊界檢查 |
| 專利/商標分類 | R01-R04 | 類型碼→分類、混合→未分類 |
| 收發碼判定 | S01-S05 | F/T判定、角色判定、Sender名單 |
| 資料夾歸檔 | F01-F05 | 有案號/無案號/多案號的歸檔策略 |
| 檔名規則 | N01-N04 | EML/附件檔名格式、語義名、多案號標記 |
| 附件處理 | A01-A02 | 寄出信不存附件、小圖片過濾 |
| 重跑保護 | D01-D02 | 同名同大小跳過、Message ID 去重 |
| LLM 回饋學習 | L01-L06 | 三種回饋方式、修正原因、few-shot learning |
| LLM 語義名規則 | L07 | OA 縮寫去重：描述文字改用中文 |
| LLM 期限選擇 | L08-L10 | 結構化期限提取：LLM 分類 + 程式碼選日期 |

---

## 十五、Sheet 欄位結構（2026-03-14 P2 更新）

處理紀錄 Sheet 22 欄（0-based）：
```
 0: messageId        10: AI案件類別
 1: 日期              11: 來源確認狀態
 2: 原始標題          12: 資料夾連結
 3: sender           13: 最終收發碼
 4: AI收發碼          14: 修正後名稱
 5: AI推斷角色        15: 修正原因
 6: 歸檔案號          16: 修正時間
 7: 內文案號          17: 修正來源
 8: AI語義名          18: 重試次數
 9: AI信心            19: Input Tokens
                      20: Output Tokens
                      21: dates_found
```

---

*最後更新：2026-03-13（晚間 Session 2）*
*版本：V3*
