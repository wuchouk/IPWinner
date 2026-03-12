// ============================================================
// IP Winner Email Auto-Processor V2
// Google Apps Script — 每日自動下載並歸檔信件
//
// 版本: v2.1.0
// 更新日期: 2026-03-12
//
// --- CHANGELOG ---
// v2.1.0 (2026-03-12)
//   - 修正 shouldSkip() 對回覆信的誤判：加入 stripQuotedText()
//     在比對 template 前先移除引用文字
//   - 修正 sanitizeFileName()：保留 RE:/FW:/Fwd:/[EXTERNAL] 前綴
//   - 搜尋結果按時間排序（最新優先），確保新信優先處理
//   - 錯誤 log 加入 stack trace
//   - 新增 debugMissingEmails() 追蹤遺失信件
//   - 新增 debugSpecificEmails() 以關鍵字追蹤特定信件
//
// v2.0.0 (2026-03-12)
//   - 訊息層級追蹤（取代 V1 的 Thread 標籤去重）
//   - 附件獨立下載 + 同資料夾去重
//   - 類型 → 完整案號 兩層資料夾
//   - 三層案號提取：慧盈案號 → Ref+格式驗證 → Fallback regex
//   - 多案號信件存進所有對應資料夾
//   - 檔名加入時間戳避免同日衝突
//   - Gmail 標籤僅供視覺辨識（專利/商標/未分類/已跳過）
// ============================================================

// ==================== 設定區 ====================
var CONFIG = {
  TARGET_EMAIL: 'ip@ipwinner.com',
  SEARCH_HOURS: 48,
  PROJECT_FOLDER_NAME: 'Email自動整理',
  DOWNLOAD_FOLDER_NAME: '下載區',
  PATENT_FOLDER_NAME: '專利',
  TRADEMARK_FOLDER_NAME: '商標',
  UNCATEGORIZED_FOLDER_NAME: '未分類',
  CONFIG_SHEET_NAME: 'Email自動整理-設定檔',
  PROCESSED_SHEET_NAME: '已處理信件',
  CLEANUP_DAYS: 30,  // 自動清理超過 N 天的已處理 ID
  // Gmail 標籤（僅供視覺辨識，不做去重）
  LABEL_PARENT: '已下載',
  LABEL_PATENT: '已下載/專利',
  LABEL_TRADEMARK: '已下載/商標',
  LABEL_UNCATEGORIZED: '已下載/未分類',
  LABEL_SKIPPED: '已跳過',
};

// 案號 regex：4碼客戶號 + 5碼年份序號 + 類型碼 + 2碼國碼 + 可選流水號 + 可選「等」
var CASE_REGEX = /[A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*等?/g;
var CASE_REGEX_SINGLE = /[A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*等?/;

// 專利類型碼（放進「專利」資料夾）
var PATENT_TYPE_CODES = 'PMDAC';
// 商標類型碼（放進「商標」資料夾）
var TRADEMARK_TYPE_CODES = 'TBW';


// ==================== 主程式 ====================

/**
 * 主入口：每天排程執行（處理全部信件）
 */
function processEmails() {
  processEmailsWithLimit(0);  // 0 = 不限制
}

/**
 * 測試用：只處理前 100 封未處理信件
 * 在 Apps Script 編輯器直接選這個函式來跑
 */
function processEmailsTest() {
  processEmailsWithLimit(100);
}

/**
 * 核心處理邏輯
 * @param {number} maxMessages 最多處理幾封（0 = 不限制）
 */
function processEmailsWithLimit(maxMessages) {
  var startTime = new Date();
  var stats = { total: 0, skipped: 0, patent: 0, trademark: 0, uncategorized: 0, errors: 0 };
  var limitReached = false;

  try {
    // 1. 從 Google Sheets 讀取過濾 template
    var templates = getFilterTemplates();

    // 2. 取得已處理的 message ID 集合
    var processedIds = getProcessedMessageIds();

    // 3. 取得資料夾結構
    var folders = getOrCreateFolders();

    // 3.5 初始化 Gmail 標籤（僅視覺辨識用）
    var labels = getOrCreateLabels();

    // 4. 搜尋相關信件
    var threads = searchThreads();
    Logger.log('找到 ' + threads.length + ' 個信件串');

    // 4.5 按最新訊息時間排序（新 → 舊），確保最近的信優先處理
    threads.sort(function(a, b) {
      return b.getLastMessageDate().getTime() - a.getLastMessageDate().getTime();
    });
    Logger.log('已按時間排序（最新優先）');

    if (maxMessages > 0) {
      Logger.log('測試模式：最多處理 ' + maxMessages + ' 封');
    }

    // 5. 逐封處理（只處理未處理過的 message）
    var newProcessedIds = [];
    var dedupMap = {};  // 附件去重 map：{ folderId_origName_size: true }

    for (var t = 0; t < threads.length; t++) {
      var messages = threads[t].getMessages();
      for (var m = 0; m < messages.length; m++) {
        // 測試模式：達到上限就停
        if (maxMessages > 0 && stats.total >= maxMessages) {
          limitReached = true;
          break;
        }

        var msgId = messages[m].getId();

        // 已處理過 → 跳過（不計入 total）
        if (processedIds[msgId]) continue;

        try {
          stats.total++;
          var result = processSingleMessage(messages[m], templates, folders, dedupMap, labels);

          if (result === 'skipped') {
            stats.skipped++;
          } else if (result === 'patent') {
            stats.patent++;
          } else if (result === 'trademark') {
            stats.trademark++;
          } else if (result === 'uncategorized') {
            stats.uncategorized++;
          } else if (result === 'multi') {
            stats.patent++;
            stats.trademark++;
          }

        } catch (e) {
          stats.errors++;
          Logger.log('處理信件失敗: ' + messages[m].getSubject());
          Logger.log('  Error: ' + e.message);
          Logger.log('  Stack: ' + (e.stack || '(無 stack trace)'));
        }

        // 不管成功失敗都標記為已處理，避免重複出錯
        newProcessedIds.push({ id: msgId, date: new Date() });
      }
      if (limitReached) break;
    }

    // 6. 批次寫入新的已處理 ID
    if (newProcessedIds.length > 0) {
      saveProcessedMessageIds(newProcessedIds);
    }

    // 7. 清理過期的已處理 ID（測試模式跳過，避免刪到東西）
    if (!maxMessages) {
      cleanupOldProcessedIds();
    }

    // 8. 寫入執行紀錄
    var statusMsg = '成功';
    if (limitReached) statusMsg = '測試完成（限 ' + maxMessages + ' 封）';
    writeLog(startTime, stats, statusMsg);
    Logger.log('處理完成: ' + JSON.stringify(stats));

  } catch (e) {
    writeLog(startTime, stats, '失敗: ' + e.message);
    Logger.log('執行失敗: ' + e.message);
  }
}


// ==================== 信件處理 ====================

/**
 * 處理單封信件
 * @returns {string} 'skipped' | 'patent' | 'trademark' | 'uncategorized' | 'multi'
 */
function processSingleMessage(message, templates, folders, dedupMap, labels) {
  var from = message.getFrom();
  var to = message.getTo();
  var cc = message.getCc() || '';
  var subject = message.getSubject();
  var body = message.getPlainBody();
  var date = message.getDate();

  // --- 過濾 template 信件 ---
  if (shouldSkip(from, body, templates)) {
    Logger.log('跳過 template 信: ' + subject);
    labels.skipped.addToThread(message.getThread());
    return 'skipped';
  }

  // --- 判斷 TX / FX ---
  var direction = getDirection(from);

  // --- 從主旨提取案號（可能多個）---
  var caseNumbers = extractCaseNumbers(subject);

  // --- 日期時間戳 ---
  var dateTimeStr = formatDateTime(date);

  // --- 取得信件原始內容 ---
  var rawContent = message.getRawContent();

  // --- 取得附件 ---
  var attachments = message.getAttachments();

  // --- 如果沒有案號 → 存到未分類 ---
  if (caseNumbers.length === 0) {
    var cleanSubject = sanitizeFileName(subject);
    var emlName = dateTimeStr + '-' + direction + '-未知案號-' + cleanSubject + '.eml';
    folders.uncategorized.createFile(emlName, rawContent, 'message/rfc822');
    Logger.log('已儲存（未分類）: ' + emlName);

    // 附件也存到未分類
    saveAttachments(attachments, dateTimeStr, direction, '未知案號', cleanSubject, folders.uncategorized, dedupMap);

    // 加上 Gmail 標籤
    labels.uncategorized.addToThread(message.getThread());

    return 'uncategorized';
  }

  // --- 有案號 → 逐個案號建資料夾並儲存 ---
  var categories = {};

  for (var i = 0; i < caseNumbers.length; i++) {
    var caseNum = caseNumbers[i];
    var folderName = caseNum.replace(/等$/, '');  // 資料夾名不含「等」
    var typeCode = folderName.charAt(9);          // 第 10 碼

    // 決定上層資料夾
    var parentFolder;
    if (PATENT_TYPE_CODES.indexOf(typeCode) !== -1) {
      parentFolder = folders.patent;
      categories['patent'] = true;
    } else if (TRADEMARK_TYPE_CODES.indexOf(typeCode) !== -1) {
      parentFolder = folders.trademark;
      categories['trademark'] = true;
    } else {
      parentFolder = folders.uncategorized;
      categories['uncategorized'] = true;
    }

    // 取得或建立案號資料夾
    var caseFolder = getOrCreateFolder(parentFolder, folderName);

    // 組合檔名
    var cleanSubject = sanitizeFileName(subject);
    var emlName = dateTimeStr + '-' + direction + '-' + caseNum + '-' + cleanSubject + '.eml';

    // 儲存 .eml
    caseFolder.createFile(emlName, rawContent, 'message/rfc822');
    Logger.log('已儲存: ' + emlName + ' → ' + parentFolder.getName() + '/' + folderName);

    // 儲存附件（含 inline 去重）
    saveAttachments(attachments, dateTimeStr, direction, caseNum, cleanSubject, caseFolder, dedupMap);
  }

  // 加上 Gmail 標籤（根據分類結果）
  var thread = message.getThread();
  if (categories['patent']) labels.patent.addToThread(thread);
  if (categories['trademark']) labels.trademark.addToThread(thread);
  if (categories['uncategorized']) labels.uncategorized.addToThread(thread);

  // 回傳分類結果
  var catKeys = Object.keys(categories);
  if (catKeys.length > 1) return 'multi';
  return catKeys[0] || 'uncategorized';
}

/**
 * 儲存附件（含 inline 圖片去重）
 * dedupMap: 執行期間的全域去重 map，key 為 folderId_origName_size
 */
function saveAttachments(attachments, dateTimeStr, direction, caseNum, cleanSubject, targetFolder, dedupMap) {
  if (!attachments || attachments.length === 0) return;

  var folderId = targetFolder.getId();
  var count = 0;

  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    var attName = att.getName();
    var attSize = att.getSize();

    // --- Inline 圖片去重 ---
    // 策略 1：執行期間的記憶體 map（同一次執行內，同資料夾、同原始檔名、同大小 → 跳過）
    var dedupKey = folderId + '_' + attName + '_' + attSize;
    if (dedupMap[dedupKey]) {
      Logger.log('跳過重複附件（記憶體去重）: ' + attName + ' (' + attSize + ' bytes)');
      continue;
    }

    // 策略 2：跨執行去重 — 對小檔案（< 20KB，通常是簽名圖片），檢查資料夾內是否有同大小的檔案
    if (attSize < 20480 && hasSameSizeFile(targetFolder, attSize)) {
      Logger.log('跳過重複附件（大小去重）: ' + attName + ' (' + attSize + ' bytes)');
      dedupMap[dedupKey] = true;
      continue;
    }

    // 標記為已處理
    dedupMap[dedupKey] = true;
    count++;

    // 取得原始副檔名
    var ext = '';
    var dotIndex = attName.lastIndexOf('.');
    if (dotIndex !== -1) {
      ext = attName.substring(dotIndex);  // 包含「.」
    }

    // 組合附件檔名
    var attachFileName = dateTimeStr + '-' + direction + '-' + caseNum + '-' + cleanSubject + '-Attachment' + count + ext;

    targetFolder.createFile(att.copyBlob().setName(attachFileName));
    Logger.log('已儲存附件: ' + attachFileName);
  }
}

/**
 * 檢查資料夾內是否有相同大小的檔案（用於小檔案的跨執行去重）
 */
function hasSameSizeFile(folder, fileSize) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    if (files.next().getSize() === fileSize) {
      return true;
    }
  }
  return false;
}


// ==================== 案號提取 ====================

/**
 * 三層策略提取案號（回傳陣列，支援多案號）
 *
 * 第 1 層：「慧盈案號」標記（最可靠）
 * 第 2 層：Your Ref / Our Ref 標記 + 格式驗證
 * 第 3 層：Fallback — regex 掃描整個主旨
 */
function extractCaseNumbers(subject) {
  // 先清理 Gmail 可能加上的前綴
  subject = subject
    .replace(/^\[From:\s*[^\]]*\]\s*/gi, '')
    .replace(/^收件匣\n?/g, '');

  var results;

  // 第 1 層：慧盈案號
  results = extractFromHuiYingMarker(subject);
  if (results.length > 0) {
    Logger.log('案號來源：慧盈案號 → ' + results.join(', '));
    return results;
  }

  // 第 2 層：Your Ref / Our Ref + 格式驗證
  results = extractFromRefMarkers(subject);
  if (results.length > 0) {
    Logger.log('案號來源：Ref 標記 → ' + results.join(', '));
    return results;
  }

  // 第 3 層：Fallback regex
  results = extractFallback(subject);
  if (results.length > 0) {
    Logger.log('案號來源：Fallback → ' + results.join(', '));
    return results;
  }

  Logger.log('未找到案號: ' + subject);
  return [];
}

/**
 * 第 1 層：從「慧盈案號」標記後面提取
 * 支援：慧盈案號：XXX、YYY  或  慧盈案號：XXX等
 */
function extractFromHuiYingMarker(subject) {
  var marker = /慧盈案號[：:]\s*/;
  var match = subject.match(marker);
  if (!match) return [];

  // 取得標記後面的文字
  var afterMarker = subject.substring(match.index + match[0].length);

  // 提取所有符合格式的案號
  var caseNums = [];
  var re = new RegExp(CASE_REGEX.source, 'g');
  var m;
  while ((m = re.exec(afterMarker)) !== null) {
    caseNums.push(m[0]);
    // 只往後看一小段（避免抓到 Ref 標記之後的對方案號）
    if (re.lastIndex > 100) break;
  }
  return caseNums;
}

/**
 * 第 2 層：從 Your Ref / Our Ref 後面提取，用格式驗證（不靠方向判斷）
 */
function extractFromRefMarkers(subject) {
  // 匹配 Your Ref / Our Ref 各種變體
  var refPattern = /(?:Your|Our)\s*Ref\.?\s*[：:]\s*/gi;
  var caseNums = [];
  var refMatch;

  while ((refMatch = refPattern.exec(subject)) !== null) {
    var afterRef = subject.substring(refMatch.index + refMatch[0].length);
    // 取 Ref 後面第一段文字（到分號或空格為止）
    var chunk = afterRef.match(/^[^\s;；,、)）]+/);
    if (chunk) {
      var candidate = chunk[0];
      // 用 IP Winner 案號格式驗證
      if (CASE_REGEX_SINGLE.test(candidate)) {
        caseNums.push(candidate.match(CASE_REGEX_SINGLE)[0]);
      }
    }
  }

  // 去重
  return uniqueArray(caseNums);
}

/**
 * 第 3 層：Fallback — regex 掃描整個主旨
 */
function extractFallback(subject) {
  var re = new RegExp(CASE_REGEX.source, 'g');
  var results = [];
  var m;
  while ((m = re.exec(subject)) !== null) {
    results.push(m[0]);
  }
  return uniqueArray(results);
}


// ==================== 過濾 & 方向判斷 ====================

/**
 * 判斷是否應該跳過（template 信件）
 */
function shouldSkip(from, body, templates) {
  if (from.toLowerCase().indexOf(CONFIG.TARGET_EMAIL.toLowerCase()) === -1) {
    return false;
  }
  if (!body || templates.length === 0) return false;

  // 只比對「本文」，去掉引用（quoted）的部分，避免回覆信裡引用的 template 造成誤判
  var originalBody = stripQuotedText(body);
  var bodyClean = originalBody.replace(/\s+/g, '');

  for (var i = 0; i < templates.length; i++) {
    var tpl = templates[i];
    if (!tpl) continue;
    var tplClean = tpl.replace(/\s+/g, '');
    if (tplClean.length > 0 && bodyClean.indexOf(tplClean) !== -1) {
      Logger.log('符合過濾範本，跳過: ' + from);
      return true;
    }
  }
  return false;
}

/**
 * 去除 email body 中被引用的文字（只保留本文）
 * 處理三種常見格式：
 *   1. Gmail 引用：On ... wrote: 後面的 > 開頭行
 *   2. Outlook 引用：---------- Forwarded message / *From:* ... 之後
 *   3. 中文轉寄：发件人：/ 寄件者：之後
 */
function stripQuotedText(body) {
  var lines = body.split('\n');
  var result = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var trimmed = line.trim();

    // Gmail 風格：「On ... wrote:」→ 從這行開始截斷（可能換行）
    if (/^On .+ wrote:\s*$/.test(trimmed)) break;
    if (trimmed === 'wrote:') break;

    // Outlook 風格：分隔線
    if (/^-{5,}/.test(trimmed) && /forwarded|原始|轉寄/i.test(trimmed)) break;

    // Outlook 風格：引用 header（*From:* 或 From: 開頭，後面帶 email）
    if (/^\*?From:\*?\s*.+@/.test(trimmed)) break;

    // 中文轉寄 header
    if (/^(发件人|寄件者)[：:]/.test(trimmed)) break;

    // 引用行（> 開頭）→ 跳過但不截斷（可能只是部分引用）
    if (/^>/.test(trimmed)) continue;

    result.push(line);
  }

  return result.join('\n');
}

/**
 * 判斷 TX（寄出）或 FX（收到）
 */
function getDirection(from) {
  if (from.toLowerCase().indexOf(CONFIG.TARGET_EMAIL.toLowerCase()) !== -1) {
    return 'TX';
  }
  return 'FX';
}


// ==================== 已處理信件追蹤 ====================

/**
 * 從 Google Sheets 讀取已處理的 message ID（回傳 { id: true } 的 map）
 */
function getProcessedMessageIds() {
  var ss = getConfigSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) map[data[i][0]] = true;
  }
  return map;
}

/**
 * 批次寫入新的已處理 message ID
 */
function saveProcessedMessageIds(entries) {
  var ss = getConfigSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PROCESSED_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Message ID', '處理時間']]);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#f0f0f0');
  }

  var rows = [];
  for (var i = 0; i < entries.length; i++) {
    rows.push([
      entries[i].id,
      Utilities.formatDate(entries[i].date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'),
    ]);
  }
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 2).setValues(rows);
}

/**
 * 清理超過 N 天的已處理 ID
 */
function cleanupOldProcessedIds() {
  var ss = getConfigSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - CONFIG.CLEANUP_DAYS);

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var dateStr = data[i][1];
    if (dateStr) {
      var rowDate = new Date(dateStr);
      if (rowDate < cutoff) {
        rowsToDelete.push(i + 2);  // +2 因為標題列 + 0-based index
      }
    }
  }

  // 從後往前刪，避免 row index 位移
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  if (rowsToDelete.length > 0) {
    Logger.log('已清理 ' + rowsToDelete.length + ' 筆過期的已處理 ID');
  }
}


// ==================== 工具函式 ====================

/**
 * 搜尋相關信件（不再用標籤過濾，改用 message ID 追蹤）
 */
function searchThreads() {
  var hoursAgo = new Date();
  hoursAgo.setHours(hoursAgo.getHours() - CONFIG.SEARCH_HOURS);
  var afterDate = formatDateForSearch(hoursAgo);

  var query = '(from:' + CONFIG.TARGET_EMAIL + ' OR to:' + CONFIG.TARGET_EMAIL + ') after:' + afterDate;
  Logger.log('搜尋條件: ' + query);

  return GmailApp.search(query);
}

/**
 * 日期時間格式化：yyyymmdd_HHmm（使用台灣時區）
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, 'Asia/Taipei', 'yyyyMMdd_HHmm');
}

/**
 * 日期格式化（Gmail 搜尋用）：yyyy/mm/dd
 */
function formatDateForSearch(date) {
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  return y + '/' + m + '/' + d;
}

/**
 * 清理檔名中不合法的字元
 */
function sanitizeFileName(name) {
  return name
    .replace(/^\[From:\s*[^\]]*\]\s*/gi, '')  // 去掉 Gmail 加的 [From: "xxx" <email>] 前綴
    .replace(/^收件匣\n?/g, '')                // 去掉「收件匣」前綴
    // RE:/FW:/Fwd:/[EXTERNAL] 全部保留，不再移除
    .replace(/[\/\\:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 150);  // 限制主旨部分為 150 字元（前綴約 50 字元 + 150 + 副檔名 < 255）
}

// ==================== Gmail 標籤 ====================

/**
 * 取得或建立單一 Gmail 標籤
 */
function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log('已建立標籤: ' + labelName);
  }
  return label;
}

/**
 * 初始化所有所需的 Gmail 標籤
 * @returns {object} { patent, trademark, uncategorized }
 */
function getOrCreateLabels() {
  // 先確保父標籤存在
  getOrCreateLabel(CONFIG.LABEL_PARENT);
  return {
    patent: getOrCreateLabel(CONFIG.LABEL_PATENT),
    trademark: getOrCreateLabel(CONFIG.LABEL_TRADEMARK),
    uncategorized: getOrCreateLabel(CONFIG.LABEL_UNCATEGORIZED),
    skipped: getOrCreateLabel(CONFIG.LABEL_SKIPPED),
  };
}


/**
 * 取得或建立資料夾
 */
function getOrCreateFolder(parent, name) {
  var folder;
  if (parent) {
    var iter = parent.getFoldersByName(name);
    folder = iter.hasNext() ? iter.next() : parent.createFolder(name);
  } else {
    var iter = DriveApp.getFoldersByName(name);
    folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
  }
  return folder;
}

/**
 * 取得或建立 Google Drive 資料夾結構
 */
function getOrCreateFolders() {
  var project = getOrCreateFolder(null, CONFIG.PROJECT_FOLDER_NAME);
  var download = getOrCreateFolder(project, CONFIG.DOWNLOAD_FOLDER_NAME);
  var patent = getOrCreateFolder(download, CONFIG.PATENT_FOLDER_NAME);
  var trademark = getOrCreateFolder(download, CONFIG.TRADEMARK_FOLDER_NAME);
  var uncategorized = getOrCreateFolder(download, CONFIG.UNCATEGORIZED_FOLDER_NAME);
  return {
    project: project,
    download: download,
    patent: patent,
    trademark: trademark,
    uncategorized: uncategorized,
  };
}

/**
 * 陣列去重
 */
function uniqueArray(arr) {
  var seen = {};
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    if (!seen[arr[i]]) {
      seen[arr[i]] = true;
      result.push(arr[i]);
    }
  }
  return result;
}


// ==================== Google Sheets 整合 ====================

/**
 * 取得設定試算表
 */
function getConfigSpreadsheet() {
  var projectFolder = getOrCreateFolder(null, CONFIG.PROJECT_FOLDER_NAME);
  var files = projectFolder.getFilesByName(CONFIG.CONFIG_SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return createConfigSpreadsheet(projectFolder);
}

/**
 * 建立設定試算表
 */
function createConfigSpreadsheet(projectFolder) {
  var ss = SpreadsheetApp.create(CONFIG.CONFIG_SHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  projectFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // --- 過濾 Template 工作表 ---
  var templateSheet = ss.getSheetByName('Sheet1');
  templateSheet.setName('過濾Template');
  templateSheet.getRange('A1').setValue('信件內容關鍵字');
  templateSheet.getRange('B1').setValue('備註');
  templateSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f0f0f0');
  templateSheet.getRange('A2').setValue(
    'Dear Colleagues,We acknowledge safe receipt of this email, thanks.Sincerely Yours,IP Winner'
  );
  templateSheet.getRange('B2').setValue('自動回覆確認信');
  templateSheet.autoResizeColumns(1, 2);

  // --- 執行紀錄 工作表 ---
  var logSheet = ss.insertSheet('執行紀錄');
  var headers = ['執行時間', '處理封數', '跳過封數', '專利', '商標', '未分類', '錯誤', '狀態'];
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange('1:1').setFontWeight('bold').setBackground('#f0f0f0');
  logSheet.autoResizeColumns(1, headers.length);

  // --- 已處理信件 工作表 ---
  var processedSheet = ss.insertSheet(CONFIG.PROCESSED_SHEET_NAME);
  processedSheet.getRange(1, 1, 1, 2).setValues([['Message ID', '處理時間']]);
  processedSheet.getRange('1:1').setFontWeight('bold').setBackground('#f0f0f0');
  processedSheet.autoResizeColumns(1, 2);

  Logger.log('已建立設定試算表: ' + ss.getUrl());
  return ss;
}

/**
 * 從 Google Sheets 讀取過濾 template
 */
function getFilterTemplates() {
  var ss = getConfigSpreadsheet();
  var sheet = ss.getSheetByName('過濾Template');

  // 如果「過濾Template」工作表不存在，自動建立（相容舊版試算表）
  if (!sheet) {
    sheet = ss.insertSheet('過濾Template');
    sheet.getRange('A1').setValue('信件內容關鍵字');
    sheet.getRange('B1').setValue('備註');
    sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.getRange('A2').setValue(
      'Dear Colleagues,We acknowledge safe receipt of this email, thanks.Sincerely Yours,IP Winner'
    );
    sheet.getRange('B2').setValue('自動回覆確認信');
    sheet.autoResizeColumns(1, 2);
    Logger.log('自動建立「過濾Template」工作表');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) result.push(data[i][0]);
  }
  return result;
}

/**
 * 寫入執行紀錄
 */
function writeLog(startTime, stats, status) {
  var ss = getConfigSpreadsheet();
  var sheet = ss.getSheetByName('執行紀錄');

  // 如果「執行紀錄」工作表不存在，自動建立（相容舊版試算表）
  if (!sheet) {
    sheet = ss.insertSheet('執行紀錄');
    var headers = ['執行時間', '處理封數', '跳過封數', '專利', '商標', '未分類', '錯誤', '狀態'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.autoResizeColumns(1, headers.length);
    Logger.log('自動建立「執行紀錄」工作表');
  }

  var row = [
    Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'),
    stats.total - stats.skipped,
    stats.skipped,
    stats.patent,
    stats.trademark,
    stats.uncategorized,
    stats.errors,
    status,
  ];
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
}


// ==================== 排程設定 ====================

/**
 * 首次執行：初始化設定檔 + 建立排程
 */
function setup() {
  var folders = getOrCreateFolders();
  Logger.log('專案資料夾: ' + folders.project.getUrl());
  Logger.log('下載區: ' + folders.download.getUrl());
  Logger.log('專利: ' + folders.patent.getUrl());
  Logger.log('商標: ' + folders.trademark.getUrl());
  Logger.log('未分類: ' + folders.uncategorized.getUrl());

  var ss = getConfigSpreadsheet();
  Logger.log('設定試算表: ' + ss.getUrl());

  createTrigger();
  Logger.log('=== V2 初始化完成 ===');
}

/**
 * 獨立建立排程
 */
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('目前有 ' + triggers.length + ' 個觸發條件');

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('已刪除舊的 processEmails 觸發條件');
    }
  }

  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  var newTriggers = ScriptApp.getProjectTriggers();
  Logger.log('建立完成，目前有 ' + newTriggers.length + ' 個觸發條件');
  for (var i = 0; i < newTriggers.length; i++) {
    Logger.log('  - ' + newTriggers[i].getHandlerFunction() + ' (' + newTriggers[i].getEventType() + ')');
  }
}


// ==================== Debug 工具 ====================

/**
 * 除錯用：測試案號提取邏輯
 */
function debugCaseExtraction() {
  var testSubjects = [
    '【供歸檔】進度通知-專利暫時案送件完成 (貴方案號：N/A；慧盈案號：25A125703PUS0)(PRIVILEGED & CONFIDENTIAL)',
    'Re: [EXTERNAL] Order Letter of Trademark Filing (Your Ref.: N/A ; Our Ref.: KOIT20004TUS8); BSKB Ref. No. 6024-0651US1',
    'RE: 委託商標新案 (貴方案號：NT-20190-SA等；慧盈案號：KOIT23001TSA2等)',
    'Re: 【請回覆】確認商標類別 (商標名稱：FIFTYLAN；慧盈案號：KOIT20004TCN7、KOIT20004TCN8)',
    '(URGENT) Order Letter of Responding to Office Action (Your Ref.: 25TW1TMAP35601 ; Our Ref.: KOIS25002TEM1)',
    'RE: Re: REMINDER: Reporting Approval Notice; (Your Ref: KOIS23007TCA1) (S&B Ref: /94891572)',
    'Re: TRON FUTURE; Your Ref.: TRON19001TUS; BSKB Ref.: 6024-0361US1',
    'CPA Global Renewal Notice',
  ];

  for (var i = 0; i < testSubjects.length; i++) {
    var result = extractCaseNumbers(testSubjects[i]);
    Logger.log('---');
    Logger.log('主旨: ' + testSubjects[i]);
    Logger.log('案號: ' + (result.length > 0 ? result.join(', ') : '(無)'));
  }
}

/**
 * 除錯用：追蹤特定案號的信件為何沒被處理
 * 在 Apps Script 編輯器中直接執行，查看 Logger 輸出
 */
function debugMissingEmails() {
  var keyword = 'KOIS23004WWW';  // ← 改成你要查的案號關鍵字

  Logger.log('========== 追蹤遺失信件 ==========');

  // 1. 用更寬的條件搜尋（不限 48 小時）
  var query1 = 'subject:' + keyword;
  var threads1 = GmailApp.search(query1, 0, 20);
  Logger.log('\n【搜尋 1】subject:' + keyword + ' → 找到 ' + threads1.length + ' 個 thread');

  for (var t = 0; t < threads1.length; t++) {
    var msgs = threads1[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var msg = msgs[m];
      Logger.log('  Thread ' + t + ' / Msg ' + m + ':');
      Logger.log('    Subject: ' + msg.getSubject());
      Logger.log('    From: ' + msg.getFrom());
      Logger.log('    To: ' + msg.getTo());
      Logger.log('    Date: ' + msg.getDate());
      Logger.log('    ID: ' + msg.getId());
    }
  }

  // 2. 用目前 script 的搜尋條件搜（含 after 限制）
  var hoursAgo = new Date();
  hoursAgo.setHours(hoursAgo.getHours() - CONFIG.SEARCH_HOURS);
  var afterDate = formatDateForSearch(hoursAgo);
  var query2 = '(from:' + CONFIG.TARGET_EMAIL + ' OR to:' + CONFIG.TARGET_EMAIL + ') after:' + afterDate;
  var threads2 = GmailApp.search(query2);
  Logger.log('\n【搜尋 2】' + query2 + ' → 找到 ' + threads2.length + ' 個 thread');

  // 看 keyword 信件有沒有在裡面
  var found = false;
  for (var t = 0; t < threads2.length; t++) {
    var msgs = threads2[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      if (msgs[m].getSubject().indexOf(keyword) !== -1) {
        Logger.log('  ✓ 有找到! Thread ' + t + ' / Msg ' + m + ': ' + msgs[m].getSubject());
        found = true;
      }
    }
  }
  if (!found) {
    Logger.log('  ✗ 目前搜尋條件找不到含 ' + keyword + ' 的信件！');
    Logger.log('  → 可能原因：信件超過 ' + CONFIG.SEARCH_HOURS + ' 小時（after:' + afterDate + '）');
  }

  // 3. 檢查已處理紀錄
  var processedIds = getProcessedMessageIds();
  Logger.log('\n【已處理紀錄】共 ' + Object.keys(processedIds).length + ' 筆');

  // 4. 對第一個搜到的信跑完整的處理流程（不實際存檔，只看結果）
  if (threads1.length > 0) {
    var templates = getFilterTemplates();
    Logger.log('\n【逐封測試 processSingleMessage 的邏輯】');

    for (var t = 0; t < threads1.length; t++) {
      var msgs = threads1[t].getMessages();
      for (var m = 0; m < msgs.length; m++) {
        var msg = msgs[m];
        var msgId = msg.getId();
        Logger.log('\n  --- Msg: ' + msg.getSubject().substring(0, 80) + ' ---');
        Logger.log('  Message ID: ' + msgId);
        Logger.log('  已處理過? ' + (processedIds[msgId] ? '是 ← 這就是原因！' : '否'));

        var from = msg.getFrom();
        var body = msg.getPlainBody();

        // shouldSkip 測試
        var fromMatch = from.toLowerCase().indexOf(CONFIG.TARGET_EMAIL.toLowerCase()) !== -1;
        Logger.log('  From 含 ip@ipwinner.com? ' + fromMatch + ' (From: ' + from + ')');

        if (fromMatch) {
          var strippedBody = stripQuotedText(body);
          Logger.log('  stripQuotedText 後 body 長度: ' + strippedBody.length + ' (原始: ' + body.length + ')');
          var skipResult = shouldSkip(from, body, templates);
          Logger.log('  shouldSkip 結果: ' + skipResult);
        }

        // 案號提取測試
        var caseNumbers = extractCaseNumbers(msg.getSubject());
        Logger.log('  提取到的案號: ' + (caseNumbers.length > 0 ? caseNumbers.join(', ') : '(無)'));
        Logger.log('  方向: ' + getDirection(from));
      }
    }
  }

  Logger.log('\n========== 追蹤結束 ==========');
}

/**
 * 除錯用：檢查 template 過濾
 */
/**
 * 除錯用：追蹤特定信件是否被處理
 * 用法：直接在 Apps Script 執行此函數，查看 Logger
 */
function debugSpecificEmails() {
  // --- 要追蹤的關鍵字（每封信一個） ---
  var keywords = [
    'KOIS22001BID1',
    'Visit your office in October',
    'KOIS23004WWW1'
  ];

  // 載入已處理紀錄
  var ssId = '149JSnTtEpQyK4Qy_RJnBNIUywJ7A5WwXxVATIj4V3po';
  var ss = SpreadsheetApp.openById(ssId);
  var processedSheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME);
  var processedData = processedSheet.getDataRange().getValues();
  var processedIds = {};
  for (var i = 1; i < processedData.length; i++) {
    processedIds[processedData[i][0]] = {
      date: processedData[i][1],
      subject: processedData[i][2] || ''
    };
  }
  Logger.log('已處理紀錄總數: ' + (processedData.length - 1));

  // 載入 template
  var templates = getFilterTemplates();

  for (var k = 0; k < keywords.length; k++) {
    var kw = keywords[k];
    Logger.log('\n' + '='.repeat(60));
    Logger.log('搜尋關鍵字: ' + kw);
    Logger.log('='.repeat(60));

    // 用主程式相同的搜尋邏輯（from OR to）+ subject 關鍵字
    var query = '(from:' + CONFIG.TARGET_EMAIL + ' OR to:' + CONFIG.TARGET_EMAIL + ') subject:(' + kw + ')';
    Logger.log('查詢: ' + query);
    var threads = GmailApp.search(query, 0, 20);
    Logger.log('找到 threads: ' + threads.length);

    if (threads.length === 0) {
      // 試試不限 subject，用全文搜尋
      query = '(from:' + CONFIG.TARGET_EMAIL + ' OR to:' + CONFIG.TARGET_EMAIL + ') ' + kw;
      Logger.log('Subject 搜尋無結果，改用全文: ' + query);
      threads = GmailApp.search(query, 0, 20);
      Logger.log('全文搜尋找到 threads: ' + threads.length);
    }

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var msgs = thread.getMessages();
      Logger.log('\n--- Thread ' + t + ' (共 ' + msgs.length + ' 封) ---');
      Logger.log('  Thread 主旨: ' + thread.getFirstMessageSubject());

      // 檢查 thread 上的 labels
      var threadLabels = thread.getLabels();
      var labelNames = [];
      for (var lb = 0; lb < threadLabels.length; lb++) {
        labelNames.push(threadLabels[lb].getName());
      }
      Logger.log('  Thread 標籤: ' + (labelNames.length > 0 ? labelNames.join(', ') : '(無)'));

      for (var m = 0; m < msgs.length; m++) {
        var msg = msgs[m];
        var msgId = msg.getId();
        var from = msg.getFrom();
        var subject = msg.getSubject();
        var date = msg.getDate();

        Logger.log('  [Message ' + m + '] ID: ' + msgId);
        Logger.log('    From: ' + from);
        Logger.log('    Subject: ' + subject);
        Logger.log('    Date: ' + date);

        // 是否在已處理紀錄？
        if (processedIds[msgId]) {
          Logger.log('    ✅ 已在處理紀錄中 (處理日期: ' + processedIds[msgId].date + ')');
        } else {
          Logger.log('    ❌ 不在處理紀錄中');
        }

        // shouldSkip 測試
        var body = msg.getPlainBody() || '';
        var strippedBody = stripQuotedText(body);
        var skipResult = shouldSkip(from, strippedBody, templates);
        Logger.log('    shouldSkip: ' + skipResult);

        // 案號提取
        var caseNumbers = extractCaseNumbers(subject);
        Logger.log('    案號: ' + (caseNumbers.length > 0 ? caseNumbers.join(', ') : '(無案號 → 未分類)'));

        // 附件數
        var attachments = msg.getAttachments();
        Logger.log('    附件: ' + attachments.length + ' 個');
      }
    }
  }

  Logger.log('\n' + '='.repeat(60));
  Logger.log('追蹤完成');
  Logger.log('='.repeat(60));
}

function debugTemplate() {
  var threads = GmailApp.search('subject:進度通知 from:ip@ipwinner.com to:ip@ipwinner.com');
  if (threads.length > 0) {
    var msg = threads[0].getMessages()[0];
    Logger.log('=== 純文字內容 ===');
    Logger.log(JSON.stringify(msg.getPlainBody()));

    var templates = getFilterTemplates();
    for (var i = 0; i < templates.length; i++) {
      Logger.log('=== Template ' + i + ' ===');
      Logger.log(JSON.stringify(templates[i]));
    }
  } else {
    Logger.log('找不到該信件');
  }
}
