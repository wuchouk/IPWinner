# TODO

## Active

### Phase 1 驗收：50 封測試
- [ ] 用 `trialRun(50)` 跑 50 封信
- [ ] 檢查語義名是否正確（特別是期限日期）
- [ ] 檢查收發碼判定
- [ ] 檢查歸檔案號 vs 內文案號
- [ ] 檢查 Drive 資料夾結構和檔名

### 回饋機制完整測試
- [ ] 測試「修正後名稱」+ Drive 改名（確認時區 Bug 已修）
- [ ] 測試 Sender 去重（同一 sender 多封信）
- [ ] 測試「匯出 LLM Prompt 文件」確認 Doc 在專案資料夾
- [ ] 測試「整理學習紀錄」確認 LLM 歸納功能

### 持續改善
- [ ] 根據 50 封測試結果調整 SYSTEM_PROMPT
- [ ] 把歸納出的規則正式合併進 SYSTEM_PROMPT

### Phase 2 開發（待定義）
- [ ] 確認 Phase 2 需求範圍
- [ ] 附件重新命名（D/A/F 版本管理）
- [ ] 精確子類型碼自動判斷

---

## Completed

### 2026-03-14
- [x] 結構化期限提取 — `_selectDeadline()` + `_extractDatesFromText()`，LLM 分類日期、程式碼選期限
- [x] `_extractMainBody()` — 從 HTML 切割主文，排除引用和簽名檔內容
- [x] Thread 上下文 — `_getThreadSubjects()` + thread_context 傳入 LLM
- [x] Thread 事項碼對齊 — `_alignThreadEventCodes()` 後處理，同 thread 取最高 OA/ROA 序號
- [x] Drive 資料夾快取 — `_createFolderCache()` 避免重複查詢同一資料夾
- [x] LLM 回應解析重構 — `_parseGeminiResponse()` + 單封重試（最多 5 次）
- [x] JSON 修復強化 — `_repairJson()` 新增陣列截斷和通用未閉合括號修復
- [x] SYSTEM_PROMPT 大幅更新 — OA/ROA 判斷規則、OA 縮寫去重、dates_found 輸出、correction_applied、廣告前綴、FA/FC 前綴精確語義
- [x] 分類規則 Sheet 新增 L07-L10（語義名規則 + 期限選擇規則）
- [x] Gmail 標籤 — `resetAllAILabels()` + 確保父標籤先建立
- [x] 自動續行機制 — `_scheduleOnce()` / `_continueProcessing()` / `_retryFailedEmails()`
- [x] Sheet 欄位新增 — col 12 資料夾連結、col 21 dates_found，全部索引偏移
- [x] 多案號多行紀錄 — 每案號一行，各自對應資料夾連結
- [x] consolidateLearning 自動寫入分類規則 Sheet + 標記 consolidated
- [x] CONFIG 調整 — MAX_TOKENS 4096、BODY_SNIPPET_LENGTH 10000、MAX_RETRY 5
- [x] 逾時安全檢查 — 批次迴圈內加 timeout check
- [x] 建立 `docs/product-plan.md`（CEO Product Plan）
- [x] 建立 `docs/engineering-plan.md`（Engineering Architecture）
- [x] 推送 GitHub（初始 commit）
- [x] 建立 `TODO.md`

### 2026-03-13
- [x] 分類規則 Sheet 美化（8 色碼分類 + 標題行）
- [x] LLM Prompt 五大 Bug 修復（TA 期限/OA 去重/過期判斷/語氣/日期不一致）
- [x] email_date 傳入 LLM
- [x] 截止日改為強制規則
- [x] Drive 改名時區 Bug 修復（getSpreadsheetTimeZone）
- [x] Sender 名單去重（_addSender 寫入前掃描）
- [x] LLM Prompt 文件管理系統（exportPromptDoc + consolidateLearning + 週排程）
- [x] Debug Log 強化
