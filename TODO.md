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
