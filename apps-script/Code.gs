/**
 * ============================================================
 * IP Winner Email Processor V3
 * ============================================================
 * 安裝方式：
 *   1. 到 https://script.google.com 新建專案
 *   2. 把這整份程式碼貼到 Code.gs（覆蓋原本的內容）
 *   3. 點左邊齒輪「專案設定」→ 勾選「在編輯器中顯示 appsscript.json」
 *      → 打開 appsscript.json → 貼上下面這段：
 *
 *   {
 *     "timeZone": "Asia/Taipei",
 *     "dependencies": {},
 *     "exceptionLogging": "STACKDRIVER",
 *     "runtimeVersion": "V8",
 *     "oauthScopes": [
 *       "https://www.googleapis.com/auth/gmail.modify",
 *       "https://www.googleapis.com/auth/gmail.labels",
 *       "https://www.googleapis.com/auth/gmail.readonly",
 *       "https://www.googleapis.com/auth/spreadsheets",
 *       "https://www.googleapis.com/auth/drive",
 *       "https://www.googleapis.com/auth/script.external_request",
 *       "https://www.googleapis.com/auth/script.scriptapp",
 *       "https://www.googleapis.com/auth/documents"
 *     ]
 *   }
 *
 *   4. 在「專案設定」→「Script 屬性」新增：
 *      GEMINI_API_KEY = 你的 Gemini API Key
 *   5. 在編輯器中選擇 setupAll 函式 → 按 ▶ 執行
 *      （會自動建立 Google Sheet、Drive 資料夾、Gmail 標籤）
 *   6. 到自動建立的 Google Sheet，上方會有「📧 Email Processor V3」選單
 *   7. 到「Sender名單」Sheet 填入已知 sender
 *   8. 從選單點「🧪 試跑 → 試跑 50 封」開始測試
 * ============================================================
 */

// ===================== 設定 =====================

const CONFIG = {
  PROJECT_FOLDER_NAME: 'Email自動整理v2',
  SPREADSHEET_NAME: 'Email自動整理v2-設定檔',

  GEMINI_MODEL: 'gemini-3-flash-preview',
  GEMINI_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/',
  GEMINI_MAX_TOKENS: 2048,
  GEMINI_TEMPERATURE: 0.1,

  BATCH_SIZE: 20,
  CONFIDENCE_AUTO: 0.8,
  CONFIDENCE_INFER: 0.6,
  CONFIDENCE_LOW: 0.5,
  BODY_SNIPPET_LENGTH: 1500,
  TIMEOUT_SAFETY_MS: 25 * 60 * 1000,
  MAX_RETRY: 3,

  PROMPT_DOC_NAME: 'LLM Prompt 文件',

  SHEET_NAMES: {
    SENDERS: 'Sender名單',
    LOG: '處理紀錄',
    RULES: '分類規則',
    SETTINGS: '設定',
  },

  LABEL_PREFIX: 'AI',

  // word boundary 防止從 base64 或亂碼中誤匹配假案號
  CASE_NUMBER_REGEX: /(?<![A-Za-z0-9])[A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*(?![A-Za-z0-9])/g,

  OWN_DOMAINS: ['ipwinner.com', 'ipwinner.com.tw'],

  // 公共 email 服務 — 不能用 domain 代表一個客戶，必須用完整 email
  PUBLIC_DOMAINS: [
    'gmail.com', 'googlemail.com', 'hotmail.com', 'outlook.com',
    'live.com', 'msn.com', 'yahoo.com', 'yahoo.com.tw',
    'icloud.com', 'me.com', 'mac.com',
    'aol.com', 'mail.com', 'protonmail.com', 'proton.me',
    'qq.com', '163.com', '126.com', 'sina.com',
    'pchome.com.tw', 'seed.net.tw', 'hinet.net',
  ],

  GOV_DOMAINS: [
    'tipo.gov.tw', 'gov.tw', 'uspto.gov', 'epo.org',
    'wipo.int', 'jpo.go.jp', 'kipo.go.kr', 'cnipa.gov.cn',
    'ipaustralia.gov.au',
  ],

  SEND_RECEIVE_CODES: ['FC', 'TC', 'FA', 'TA', 'FG', 'TG', 'FX', 'TX'],
  CASE_CATEGORIES: ['專利', '商標', '未分類'],
};

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('請先在「專案設定 → Script 屬性」設定 GEMINI_API_KEY');
  return key;
}


// ===================== 自動建立 Spreadsheet & Drive =====================

/**
 * 取得或建立資料夾（V1 同款邏輯）
 * parent = null 時從 Drive 根目錄搜尋
 */
function _getOrCreateFolder(parent, name) {
  if (parent) {
    const iter = parent.getFoldersByName(name);
    return iter.hasNext() ? iter.next() : parent.createFolder(name);
  } else {
    const iter = DriveApp.getFoldersByName(name);
    return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
  }
}

/**
 * 取得專案根資料夾（自動建立）
 */
function _getProjectFolder() {
  return _getOrCreateFolder(null, CONFIG.PROJECT_FOLDER_NAME);
}

/**
 * 取得 Drive 根資料夾 ID（自動建立子資料夾結構）
 */
function _getDriveRootFolder() {
  const projectFolder = _getProjectFolder();
  // 確保三個子資料夾存在
  _getOrCreateFolder(projectFolder, '專利');
  _getOrCreateFolder(projectFolder, '商標');
  _getOrCreateFolder(projectFolder, '未分類');
  return projectFolder;
}

/**
 * 核心：取得設定試算表（自動搜尋 or 建立）
 *
 * 邏輯：
 *   1. 先查 Script Properties 有無 SHEET_ID → 嘗試直接開
 *   2. 沒有 → 在專案資料夾搜尋同名 Sheet
 *   3. 都找不到 → SpreadsheetApp.create() 建新的
 */
function _getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();

  // 嘗試 1：用已存的 ID 開啟
  const savedId = props.getProperty('SHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      Logger.log('SHEET_ID 無效，重新搜尋...');
    }
  }

  // 嘗試 2：在專案資料夾搜尋同名 Sheet
  const projectFolder = _getProjectFolder();
  const files = projectFolder.getFilesByName(CONFIG.SPREADSHEET_NAME);
  if (files.hasNext()) {
    const ss = SpreadsheetApp.open(files.next());
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('找到現有試算表: ' + ss.getUrl());
    return ss;
  }

  // 嘗試 3：建立新的
  return _createSpreadsheet(projectFolder);
}

/**
 * 建立新的試算表，移入專案資料夾
 */
function _createSpreadsheet(projectFolder) {
  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  const file = DriveApp.getFileById(ss.getId());

  // 移到專案資料夾（從根目錄移走）
  projectFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // 儲存 ID
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  Logger.log('已建立新試算表: ' + ss.getUrl());
  return ss;
}


// ===================== 選單 =====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Email Processor V3')
    .addItem('🔧 首次安裝設定', 'setupAll')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 試跑')
      .addItem('試跑 50 封（分類＋下載＋掛標籤）', 'trialRun')
      .addItem('試跑 10 封（快速驗證）', 'trialRunSmall')
      .addItem('試跑單封信（輸入搜尋條件）', 'testSingleEmail'))
    .addItem('▶️ 正式處理 (Batch)', 'processEmails')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 回授與學習')
      .addItem('執行回授偵測', 'runFeedback')
      .addItem('查看學習紀錄', 'showLearningLog')
      .addItem('匯出 LLM Prompt 文件', 'exportPromptDoc')
      .addItem('整理學習紀錄（合併進 Prompt）', 'consolidateLearning'))
    .addSeparator()
    .addItem('📊 處理統計', 'showStats')
    .addSubMenu(ui.createMenu('⏱️ 排程')
      .addItem('安裝每日排程（早上 7-8 點）', 'installTrigger')
      .addItem('移除排程', 'removeTrigger'))
    .addItem('⚙️ 重設 API Key', 'setupApiKey_ui')
    .addToUi();
}


// ===================== 首次安裝 =====================

/**
 * 首次安裝 — 從 Apps Script 編輯器直接執行即可
 *
 * 自動建立：
 *   ✅ Drive「Email自動整理v2」資料夾 + 子資料夾
 *   ✅ Google Sheet（4 張 Tab）移入資料夾
 *   ✅ Gmail AI/ 系列標籤
 *   ✅ 為 Google Sheet 安裝 onOpen 選單觸發器
 *
 * 事前準備：
 *   在「專案設定 → Script 屬性」設好 GEMINI_API_KEY
 */
function setupAll() {
  // 1. 建立 Drive 資料夾結構
  const projectFolder = _getDriveRootFolder();
  Logger.log('📁 Drive 資料夾: ' + projectFolder.getUrl());

  // 2. 取得或建立 Spreadsheet（自動移入資料夾）
  const ss = _getSpreadsheet();

  // 3. 建立四張 Sheet Tab
  _setupSheets(ss);

  // 4. 建立 Gmail 標籤
  _ensureLabels();

  // 5. 安裝 onOpen 觸發器（讓 Sheet 打開時出現選單）
  _installOnOpenTrigger(ss);

  // 6. 安裝每週學習整理排程（每週一 9 點）
  installConsolidationTrigger();

  // 7. 檢查 API Key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  Logger.log('✅ 安裝完成！');
  Logger.log('📊 試算表: ' + ss.getUrl());
  Logger.log('📁 Drive 資料夾: ' + projectFolder.getUrl());
  if (!apiKey) {
    Logger.log('⚠️ 尚未設定 GEMINI_API_KEY！請到「專案設定 → Script 屬性」新增');
  }
  Logger.log('');
  Logger.log('📋 下一步：');
  Logger.log('   1. 打開上方試算表連結');
  Logger.log('   2. 到「Sender名單」填入已知 sender');
  Logger.log('   3. 從選單「📧 Email Processor V3 → 🧪 試跑」開始測試');

  // 如果是從 Sheet 選單呼叫的，顯示 alert
  try {
    const ui = SpreadsheetApp.getUi();
    let msg = '✅ 安裝完成！\n\n' +
      '已建立：\n' +
      '• Drive「Email自動整理v2」資料夾\n' +
      '• Google Sheet（4 張 Tab）\n' +
      '• Gmail AI/ 標籤\n\n';
    if (!apiKey) {
      msg += '⚠️ 尚未設定 API Key！\n請到 Apps Script「專案設定 → Script 屬性」新增 GEMINI_API_KEY\n\n';
    }
    msg += '下一步：\n1. 到「Sender名單」填入已知 sender\n2. 從選單 🧪 試跑 → 試跑 50 封';
    ui.alert(msg);
  } catch (e) {
    // 從編輯器執行時沒有 UI，忽略（已有 Logger 輸出）
  }
}


/**
 * 為 Sheet 安裝 onOpen 觸發器
 * 讓 standalone script 也能在 Sheet 打開時顯示選單
 */
function _installOnOpenTrigger(ss) {
  // 先移除舊的 onOpen 觸發器
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // 安裝新的
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  Logger.log('✅ onOpen 選單觸發器已安裝');
}


/**
 * 設定 API Key（從編輯器直接呼叫）
 * 用法：在編輯器中把 key 改成你的 API Key，然後執行
 */
function setupApiKey(key) {
  if (!key) {
    Logger.log('用法：setupApiKey("你的Gemini_API_Key")');
    Logger.log('或到 Apps Script「專案設定」→「Script 屬性」→ 新增 GEMINI_API_KEY');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key.trim());
  Logger.log('✅ API Key 已設定');
}


/**
 * 從 Sheet 選單重設 API Key（需要 UI）
 */
function setupApiKey_ui() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('設定 API Key', '請輸入 Gemini 3.0 Flash API Key：', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const key = result.getResponseText().trim();
  if (!key) { ui.alert('API Key 不能為空'); return; }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('✅ API Key 已設定');
}


function _setupSheets(ss) {
  // Sheet 1: Sender 名單
  if (!ss.getSheetByName(CONFIG.SHEET_NAMES.SENDERS)) {
    const s = ss.insertSheet(CONFIG.SHEET_NAMES.SENDERS);
    s.appendRow(['Email 或 Domain', '角色（C/A/G）', '名稱備註']);
    s.getRange('1:1').setFontWeight('bold').setBackground('#4a86c8').setFontColor('white');
    s.setColumnWidth(1, 280);
    s.setColumnWidth(2, 120);
    s.setColumnWidth(3, 280);

    // 角色欄下拉選單（B2:B500）
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['C', 'A', 'G'], true)
      .setHelpText('C = 客戶, A = 代理人, G = 政府機關')
      .setAllowInvalid(false)
      .build();
    s.getRange('B2:B500').setDataValidation(roleRule);

    [
      ['@tipo.gov.tw', 'G', 'TIPO 台灣智慧局'],
      ['@tiponet.tipo.gov.tw', 'G', 'TIPO 網路申辦'],
      ['@uspto.gov', 'G', 'USPTO 美國專利局'],
      ['@epo.org', 'G', 'EPO 歐洲專利局'],
      ['@wipo.int', 'G', 'WIPO 世界智財組織'],
      ['@jpo.go.jp', 'G', 'JPO 日本特許廳'],
      ['@naipo.com', 'A', 'NAIP 北美智權（代理人）'],
      ['@cpaglobal.com', 'A', 'CPA Global 年費代繳（代理人）'],
    ].forEach(row => s.appendRow(row));
  }

  // Sheet 2: 處理紀錄
  // 欄位索引（0-based）：
  //  0: messageId  1: 日期  2: 原始標題  3: sender
  //  4: AI收發碼  5: AI推斷角色  6: 歸檔案號  7: 內文案號  8: AI語義名
  //  9: AI信心  10: AI案件類別  11: 來源確認狀態
  //  12: 最終收發碼  13: 修正後名稱  14: 修正原因
  //  15: 修正時間  16: 修正來源  17: 重試次數
  //  18: Input Tokens  19: Output Tokens
  if (!ss.getSheetByName(CONFIG.SHEET_NAMES.LOG)) {
    const s = ss.insertSheet(CONFIG.SHEET_NAMES.LOG);
    s.appendRow([
      'messageId', '日期', '原始標題', 'sender',
      'AI收發碼', 'AI推斷角色', '歸檔案號', '內文案號', 'AI語義名',
      'AI信心', 'AI案件類別', '來源確認狀態',
      '最終收發碼', '修正後名稱', '修正原因',
      '修正時間', '修正來源', '重試次數',
      'Input Tokens', 'Output Tokens',
    ]);
    s.getRange('1:1').setFontWeight('bold').setBackground('#4a86c8').setFontColor('white');
    s.setFrozenRows(1);
    s.hideColumns(1);  // 隱藏 messageId

    s.setColumnWidth(2, 130);   // 日期
    s.setColumnWidth(3, 300);   // 原始標題
    s.setColumnWidth(4, 200);   // sender
    s.setColumnWidth(5, 80);    // AI收發碼
    s.setColumnWidth(6, 80);    // AI推斷角色
    s.setColumnWidth(7, 200);   // 歸檔案號
    s.setColumnWidth(8, 200);   // 內文案號
    s.setColumnWidth(9, 250);   // AI語義名
    s.setColumnWidth(10, 60);   // AI信心
    s.setColumnWidth(11, 80);   // AI案件類別
    s.setColumnWidth(12, 100);  // 來源確認狀態
    s.setColumnWidth(13, 80);   // 最終收發碼
    s.setColumnWidth(14, 300);  // 修正後名稱
    s.setColumnWidth(15, 350);  // 修正原因
  }

  // Sheet 3: 分類規則
  _setupRulesSheet(ss);

  // Sheet 4: 設定
  if (!ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS)) {
    const s = ss.insertSheet(CONFIG.SHEET_NAMES.SETTINGS);
    s.appendRow(['參數', '值', '說明']);
    s.getRange('1:1').setFontWeight('bold').setBackground('#4a86c8').setFontColor('white');
    s.setColumnWidth(1, 200);
    s.setColumnWidth(2, 200);
    s.setColumnWidth(3, 350);
    [
      ['信心閾值', 0.8, '≥此值自動處理，<此值標記待確認'],
      ['角色推斷閾值', 0.6, '≥此值採用LLM推斷的sender角色'],
      ['附件重試上限', 3, '附件下載失敗的最大重試次數'],
      ['batch大小', 20, '正式處理每批封數'],
      ['body截取字數', 1500, '傳給LLM的內文長度'],
      ['lastProcessedTime', '', '（系統自動更新）'],
      ['累計處理數量', 0, '（系統自動更新）'],
    ].forEach(row => s.appendRow(row));
  }

  // 刪除預設的「Sheet1」（如果還在的話）
  try {
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  } catch (e) { /* 如果只剩一張 sheet 刪不掉，沒關係 */ }

  // 再試一次（可能 Sheet1 是唯一一張被鎖定的）
  try {
    const defaultSheet = ss.getSheetByName('工作表1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  } catch (e) { /* 忽略 */ }
}


// ===================== 分類規則 Tab 初始化 =====================

function _setupRulesSheet(ss) {
  let s = ss.getSheetByName(CONFIG.SHEET_NAMES.RULES);
  if (!s) {
    s = ss.insertSheet(CONFIG.SHEET_NAMES.RULES);
  } else {
    // 已存在 → 清除舊資料再重寫（保留 tab）
    s.clear();
  }

  // 欄位結構
  const headers = ['規則ID', '類別', '觸發條件', '動作', '說明/範例'];
  s.appendRow(headers);
  s.getRange('1:1').setFontWeight('bold').setBackground('#4a86c8').setFontColor('white');
  s.setColumnWidth(1, 80);
  s.setColumnWidth(2, 100);
  s.setColumnWidth(3, 350);
  s.setColumnWidth(4, 300);
  s.setColumnWidth(5, 400);

  // ── 規則資料 ──
  const rules = [
    // 案號格式
    ['C01', '案號格式', '案號結構：[4碼客戶號][2碼年份][3碼序號][1碼類型][2碼國碼][選填分案號]',
     '用正規表達式 /[A-Z0-9]{4}\\d{5}[PMDTABCW][A-Z]{2}\\d*/ 提取',
     '例：BRIT25710PUS1 → BRIT=客戶碼, 25=年份, 710=序號, P=專利, US=國碼, 1=分案號'],

    ['C02', '案號格式', '類型碼位於案號第10碼（index 9），固定位置取值',
     '用 charAt(9) 取類型碼，不用 regex search',
     'Bug教訓：BRIT 的 T 會被 regex 先匹配到，誤判為商標'],

    ['C03', '案號格式', '案號前後必須有邊界（非英數字元）',
     '正規表達式加 lookbehind/lookahead 避免從 base64 編碼誤匹配',
     '例：/(?<![A-Za-z0-9])...(?>![A-Za-z0-9])/'],

    // 分類（專利/商標）
    ['R01', '分類', '案號第10碼為 P、M、D、A、C',
     '歸入「專利」資料夾 + 加上「專利」Gmail label',
     'P=Patent, M=Model, D=Design, A=Application, C=Continuation'],

    ['R02', '分類', '案號第10碼為 T、B、W',
     '歸入「商標」資料夾 + 加上「商標」Gmail label',
     'T=Trademark, B=Brand, W=WIPO商標'],

    ['R03', '分類', '同封信歸檔案號含專利+商標類型碼（混合）',
     '歸入「未分類」資料夾',
     '例：同封信有 PMDAC 和 TBW 的案號 → 未分類'],

    ['R04', '分類', '主旨無案號（不論內文有無）',
     '歸入「未分類」資料夾，不加專利/商標 label',
     '因為無法從主旨判斷是專利還是商標'],

    // 收發碼
    ['S01', '收發碼', '寄件人 domain 為自家（ipwinner.com）→ T（寄出）；否則 → F（收到）',
     '設定方向碼 F 或 T',
     'OWN_DOMAINS 可在設定中自訂'],

    ['S02', '收發碼', '收到的信（F方向）：看寄件人 email 在 Sender 名單的角色',
     '決定角色碼 C=客戶 / A=代理人 / G=政府 / X=未知 → 組成 FC/FA/FG/FX',
     '例：bskb.com 在名單標為 Agent → FA'],

    ['S03', '收發碼', '寄出的信（T方向）：看第一個外部收件人在 Sender 名單的角色',
     '決定角色碼 → 組成 TC/TA/TG/TX',
     'Bug教訓：T方向看寄件人（自己）永遠是 X → 全變 TX'],

    ['S04', '收發碼', '寄件人 domain 含政府機構 domain（tipo.gov.tw, uspto.gov 等）',
     '角色 = G（政府）',
     'GOV_DOMAINS 可在設定中擴充'],

    ['S05', '收發碼', 'Sender 名單設計：私人 domain → 用 @domain；公共 email → 用完整 email',
     '避免 @gmail.com 代表所有 Gmail 用戶',
     '例：bskb.com → @bskb.com | john@gmail.com → john@gmail.com'],

    // 歸檔邏輯
    ['F01', '歸檔', '主旨有案號 → 由 LLM 判斷 filing_case_numbers（實際歸檔案號）',
     '依歸檔案號建立資料夾存 EML 和附件',
     'LLM 區分「主要處理的案號」和「順便引用的案號」，只歸主案'],

    ['F02', '歸檔', '主旨無案號但內文有案號',
     '歸入 未分類/{客戶碼前4碼}/ + 標記「無案號」',
     '例：主旨無案號但內文有 BRIT... → 未分類/BRIT/'],

    ['F03', '歸檔', '完全無案號（主旨和內文都沒有）',
     '歸入 未分類/無案號/ + 標記「無案號」',
     ''],

    ['F04', '歸檔', '主旨有 2 個以上不同案號',
     '每個歸檔案號各建一個資料夾存 EML + 標記「多案號」',
     '例：主旨有 BRIT25710PUS1 和 BRIT25711PUS2 → 各自資料夾都存一份 EML'],

    ['F05', '歸檔', '主旨有 1 個案號 +「等」字（如 BRIT25710PUS1等3案）',
     '觸發多案號：去內文找其他案號，由 LLM 判斷歸檔案號，各建資料夾存 EML + 標記「多案號」',
     '「～」「~」「、」不需額外判斷，因為這些情況主旨自然有 2+ 個案號（已被 F04 處理）'],

    // 檔名規則
    ['N01', '檔名', 'EML 檔名格式：{日期}-{收發碼}-{案號標記}-{AI語義名}.eml',
     '日期=yyyyMMdd, 收發碼=FC/TA等, 案號標記=案號或客戶碼或「等N案」',
     '例：20260313-FA-BRIT25710PUS1-異議答辯期限0317.eml'],

    ['N02', '檔名', '附件檔名格式：{EML基礎名}-附件N.{副檔名}',
     'N 從 1 開始，副檔名保留原始附件的副檔名',
     '例：20260313-FA-BRIT25710PUS1-異議答辯期限0317-附件1.pdf'],

    ['N03', '檔名', 'AI 語義名由 LLM 生成，25 字以內，含最關鍵的期限日期',
     '期限選擇規則：TA→我方要求代理人的期限; TC→我方要求客戶的期限; FA→代理人要求我方的期限',
     '用「行動期限」而非「背景日期」（如官方通知日期）'],

    ['N04', '檔名', '多案號的 EML 檔名案號標記帶「等N案」',
     '例：2案 → BRIT25710PUS1等2案',
     '每個資料夾的 EML 都帶「等N案」，方便辨識'],

    // 附件規則
    ['A01', '附件', 'TA/TC/TG（寄出的信）只存 EML，不存附件',
     '因為附件是我方自己寄出的，本地已有原始檔',
     '收到的信（FA/FC/FG/FX）才存附件'],

    ['A02', '附件', '< 5KB 的圖片 → 跳過（通常是簽名檔圖片）',
     '檔名含 image00/logo/banner/signature 的也跳過',
     '另用 includeInlineImages:false 排除基本 inline 圖'],

    // 重跑保護
    ['D01', '重跑保護', '同檔名 + 同大小（±100 bytes）的檔案',
     '跳過不重建（EML 和附件共用此邏輯）',
     '安全重跑：不需手動刪檔就能重跑處理'],

    ['D02', '重跑保護', '訊息去重用 Message ID（記錄在處理紀錄 Sheet）',
     '搜尋 Gmail 時不加 -label 排除條件',
     'Bug教訓：label 加在 thread 上，用 -label 排除會漏掉同 thread 新回覆'],

    // LLM 回饋
    ['L01', 'LLM回饋', 'Sender 角色回授（方法A）：Gmail 移除「自動辨識來源」標籤',
     '執行回授偵測 → 偵測 Gmail label → 寫入 Sender 名單',
     'T 方向的回饋學習用收件人（非寄件人）作為學習對象'],

    ['L02', 'LLM回饋', 'Sender 角色回授（方法B）：在 Sheet「最終收發碼」欄直接填寫正確收發碼',
     '執行回授偵測 → 比對 AI 碼 vs 最終碼 → 寫入 Sender 名單',
     '兩種方法擇一即可，修正來源分別記錄為 tag_change / sheet_code'],

    ['L03', 'LLM回饋', '檔名回授：在 Sheet「修正後名稱」欄填寫正確語義名',
     '執行回授偵測 → 自動改名 Drive 裡的 EML 和附件 + 作為 LLM 未來 few-shot 範例',
     '例：「委託-提出商標異議-期限3/23」→「委託-提出商標異議-期限3/17」'],

    ['L04', 'LLM回饋', '在「修正原因」欄填寫原因（建議填寫）',
     '修正原因會傳給 LLM 作為 few-shot 範例的一部分，幫助 LLM 理解為什麼修改',
     '例：「應採用代理人要求的行動期限，非官方通知期限」'],

    ['L05', 'LLM回饋', '多種回授可同時進行，修正來源用「+」串接',
     '例：tag_change+name_change 代表同時修正了 sender 角色和語義名',
     '各項回授獨立判斷是否已處理，不會互相覆蓋'],

    ['L06', 'LLM回饋', '修正紀錄用於 LLM few-shot learning（最近 20 筆）',
     '每次處理新信件時，把修正紀錄（含原因）注入 system prompt',
     'LLM 看到修正範例 + 原因，能逐漸學到命名偏好和判斷規則'],
  ];

  // 類別標題與顏色設定
  const categoryConfig = {
    '案號格式': { title: '案號結構', bg: '#E8F0FE', titleBg: '#C5DAF0' },
    '分類':     { title: '專利/商標分類', bg: '#FCE8E6', titleBg: '#F5C6C2' },
    '收發碼':   { title: '收發碼判定', bg: '#FEF7E0', titleBg: '#F5E6A8' },
    '歸檔':     { title: '資料夾歸檔', bg: '#E6F4EA', titleBg: '#B7DFC4' },
    '檔名':     { title: '檔名規則', bg: '#F3E8FD', titleBg: '#D8C0F0' },
    '附件':     { title: '附件處理', bg: '#E8F7FE', titleBg: '#B8DFF5' },
    '重跑保護': { title: '重跑保護', bg: '#F1F3F4', titleBg: '#D2D6D9' },
    'LLM回饋': { title: 'LLM 回饋學習', bg: '#FFF8E1', titleBg: '#FFE082' },
  };

  // 將規則依類別分組，保持原順序
  const groups = [];
  let currentCat = null;
  let currentGroup = null;
  for (const rule of rules) {
    if (rule[1] !== currentCat) {
      currentCat = rule[1];
      currentGroup = { category: currentCat, rules: [] };
      groups.push(currentGroup);
    }
    currentGroup.rules.push(rule);
  }

  // 寫入規則（含類別標題行）
  let row = 2; // 從 header 下一行開始
  for (const group of groups) {
    const cfg = categoryConfig[group.category] || { title: group.category, bg: '#FFFFFF', titleBg: '#E0E0E0' };

    // 寫入類別標題行
    s.getRange(row, 1).setValue(cfg.title);
    const titleRange = s.getRange(row, 1, 1, 5);
    titleRange.merge();
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(11);
    titleRange.setBackground(cfg.titleBg);
    titleRange.setVerticalAlignment('middle');
    row++;

    // 寫入該類別的規則
    if (group.rules.length > 0) {
      s.getRange(row, 1, group.rules.length, 5).setValues(group.rules);
      // 規則行上色
      for (let i = 0; i < group.rules.length; i++) {
        s.getRange(row + i, 1, 1, 5).setBackground(cfg.bg);
      }
      // 規則ID 粗體
      s.getRange(row, 1, group.rules.length, 1).setFontWeight('bold');
      row += group.rules.length;
    }
  }

  // 格式美化
  s.setFrozenRows(1);
  s.autoResizeColumn(1);

  Logger.log('✅ 分類規則 tab：已寫入 ' + rules.length + ' 條規則（含 ' + groups.length + ' 個類別標題）');
}


// ===================== Gmail 標籤 =====================

function _ensureLabels() {
  const allLabels = [
    ...CONFIG.SEND_RECEIVE_CODES,
    ...CONFIG.CASE_CATEGORIES,
    '多案號', '無案號',
    '待確認', '自動辨識來源', '未知來源', '已跳過',
    '附件下載錯誤', '處理失敗',
  ];

  allLabels.forEach(name => {
    const fullName = CONFIG.LABEL_PREFIX + '/' + name;
    if (!GmailApp.getUserLabelByName(fullName)) {
      GmailApp.createLabel(fullName);
    }
  });
}

function _getLabel(name) {
  return GmailApp.getUserLabelByName(CONFIG.LABEL_PREFIX + '/' + name);
}


// ===================== 規則引擎 =====================

function _extractEmail(str) {
  if (!str) return '';
  const m = str.match(/<([^>]+)>/);
  return (m ? m[1] : str).toLowerCase().trim();
}

function _extractDomain(email) {
  const p = email.split('@');
  return p.length > 1 ? p[1] : '';
}

function _isGovDomain(domain) {
  return CONFIG.GOV_DOMAINS.some(g => domain === g || domain.endsWith('.' + g));
}

function _cleanSubject(raw) {
  let s = raw || '';
  const original = s;
  let changed = true;
  while (changed) {
    changed = false;
    const before = s;
    s = s.replace(/^(RE|Re|re|FW|Fw|fw|FWD|Fwd|fwd|回覆|轉寄|提醒)\s*[:：_]?\s*/i, '');
    if (s !== before) changed = true;
  }
  // 清除中文方括號標記（【請回覆】【急】等）
  s = s.replace(/【[^】]{0,10}】\s*/g, '');
  return { cleaned: s.trim(), original: original.trim() };
}

function _getDirection(senderEmail) {
  const domain = _extractDomain(senderEmail);
  return CONFIG.OWN_DOMAINS.includes(domain) ? 'T' : 'F';
}

function _getSenderRole(senderEmail, senderMap) {
  const email = senderEmail.toLowerCase().trim();
  const domain = _extractDomain(email);

  if (senderMap.has(email)) return { role: senderMap.get(email).role, source: 'exact' };
  if (senderMap.has('@' + domain)) return { role: senderMap.get('@' + domain).role, source: 'domain' };

  // 子網域匹配（tiponet.tipo.gov.tw → @tipo.gov.tw）
  const parts = domain.split('.');
  for (let i = 1; i < parts.length - 1; i++) {
    const parent = '@' + parts.slice(i).join('.');
    if (senderMap.has(parent)) return { role: senderMap.get(parent).role, source: 'subdomain' };
  }

  if (_isGovDomain(domain)) return { role: 'G', source: 'gov' };
  return { role: 'X', source: 'unknown' };
}

function _getSendReceiveCode(direction, role) {
  const map = {
    'F_C': 'FC', 'F_A': 'FA', 'F_G': 'FG', 'F_X': 'FX',
    'T_C': 'TC', 'T_A': 'TA', 'T_G': 'TG', 'T_X': 'TX',
  };
  return map[direction + '_' + role] || 'FX';
}

/**
 * 當 LLM 沒回傳語義名時的 fallback
 * 從標題提取核心語義，去掉案號、括號資訊等已在檔名其他欄位的冗餘內容
 */
function _fallbackSemanticName(subject) {
  let s = subject || '';
  // 去掉括號內的案號/名稱等冗餘資訊（這些已在檔名其他段）
  s = s.replace(/\([^)]*案號[^)]*\)/g, '');
  s = s.replace(/\([^)]*[A-Z0-9]{4}\d{5}[^)]*\)/g, '');
  s = s.replace(/（[^）]*案號[^）]*）/g, '');
  s = s.replace(/（[^）]*[A-Z0-9]{4}\d{5}[^）]*）/g, '');
  // 去掉 PRIVILEGED & CONFIDENTIAL 等常見尾巴
  s = s.replace(/\(?\s*PRIVILEGED\s*[&＆]\s*CONFIDENTIAL\s*\)?/gi, '');
  // 去掉多餘空格和標點
  s = s.replace(/\s+/g, ' ').replace(/[;；,，]+\s*$/g, '').trim();
  // 限制長度（最多 20 字）
  if (s.length > 20) s = s.substring(0, 20);
  return s || '未命名';
}

function _extractCaseInfo(subject, body, attNames) {
  // 分開提取：主旨案號 vs 內文/附件案號
  const subjectMatches = (subject || '').match(CONFIG.CASE_NUMBER_REGEX) || [];
  const subjectCaseNumbers = [...new Set(subjectMatches)];

  const bodyAttText = [body, ...(attNames || [])].join(' ');
  const bodyMatches = bodyAttText.match(CONFIG.CASE_NUMBER_REGEX) || [];
  const bodyCaseNumbers = [...new Set(bodyMatches)];

  // 全部案號（去重）
  const allCaseNumbers = [...new Set([...subjectCaseNumbers, ...bodyCaseNumbers])];

  // 分類邏輯：用案號第10碼（index 9）判定專利/商標
  //   P, M, D, A, C → 專利
  //   T, B, W → 商標
  //   只看主旨案號；主旨無案號 → 未分類
  let category = '未分類';
  if (subjectCaseNumbers.length > 0) {
    const PATENT_TYPES = 'PMDAC';
    const TRADEMARK_TYPES = 'TBW';
    const typeChars = subjectCaseNumbers.map(cn => {
      // 案號第10碼（index 9）= 類型碼
      return cn.length >= 10 ? cn.charAt(9) : '';
    });
    const hasPatent = typeChars.some(c => PATENT_TYPES.includes(c));
    const hasTrademark = typeChars.some(c => TRADEMARK_TYPES.includes(c));
    if (hasPatent && !hasTrademark) category = '專利';
    else if (hasTrademark && !hasPatent) category = '商標';
  }

  // 客戶碼：案號前4碼（用於主旨無案號時的資料夾名稱）
  let clientCode = null;
  if (allCaseNumbers.length > 0) {
    clientCode = allCaseNumbers[0].substring(0, 4);
  }

  // 案號狀態
  let caseStatus = null;
  if (allCaseNumbers.length === 0) {
    caseStatus = '無案號';
  } else if (subjectCaseNumbers.length > 1) {
    // 只看主旨是否多案號
    caseStatus = '多案號';
  } else if (subjectCaseNumbers.length === 0 && bodyCaseNumbers.length > 0) {
    // 主旨無案號但內文有 → 標記無案號（不算多案號）
    caseStatus = '無案號';
  }
  // 主旨有案號 + 「等」→ 算多案號（如 BRIT25710PUS1等3案）
  // 註：「～」「~」「、」不需要，因為這些情況主旨一定有 2+ 個案號，上面已處理
  if (!caseStatus && subjectCaseNumbers.length > 0 && /等/.test(subject || '')) {
    if ((subject || '').match(/(?<![A-Za-z0-9])[A-Z0-9]{4}\d{5}[PMDTABCW][A-Z]{2}\d*\s*等/)) {
      caseStatus = '多案號';
    }
  }

  return {
    caseNumbers: allCaseNumbers,         // 全部案號（Sheet 紀錄用）
    subjectCaseNumbers: subjectCaseNumbers, // 主旨案號（決定多資料夾用）
    clientCode: clientCode,              // 客戶碼（主旨無案號時的資料夾名）
    caseCategory: category,
    caseStatus,
  };
}

function _preprocessMessage(message, senderMap) {
  const sender = message.getFrom();
  const senderEmail = _extractEmail(sender);
  const { cleaned, original } = _cleanSubject(message.getSubject() || '');
  const direction = _getDirection(senderEmail);

  const recipients = [
    ...(message.getTo() || '').split(','),
    ...(message.getCc() || '').split(','),
  ].map(r => _extractEmail(r)).filter(Boolean);

  // 角色判定邏輯：
  //   收到的信（F）→ 看「寄件人」是誰（客戶/代理人/政府）
  //   寄出的信（T）→ 看「收件人」是誰（我們寄給客戶/代理人/政府）
  let role, roleSource;
  if (direction === 'T') {
    // 寄出：找第一個非自己 domain 的收件人來判角色
    const externalRecipient = recipients.find(r => !CONFIG.OWN_DOMAINS.includes(_extractDomain(r)));
    if (externalRecipient) {
      ({ role, source: roleSource } = _getSenderRole(externalRecipient, senderMap));
    } else {
      role = 'X';
      roleSource = 'no_external_recipient';
    }
  } else {
    // 收到：看寄件人
    ({ role, source: roleSource } = _getSenderRole(senderEmail, senderMap));
  }
  const code = _getSendReceiveCode(direction, role);

  const attachments = _getSmartAttachments(message);
  const attachmentNames = attachments.map(a => a.getName());
  const body = (message.getPlainBody() || '').substring(0, CONFIG.BODY_SNIPPET_LENGTH);
  const caseInfo = _extractCaseInfo(cleaned, body, attachmentNames);

  // 提取 HTML 中被強調的文字（bold/highlight/colored），供 LLM 判斷重點
  const highlights = _extractHighlights(message);

  return {
    messageId: message.getId(),
    date: message.getDate(),
    sender: senderEmail,
    recipients: recipients,
    subject: cleaned,
    originalSubject: original,
    direction: direction,
    role: role,
    sendReceiveCode: code,
    hasAttachments: attachments.length > 0,
    attachmentNames: attachmentNames,
    bodySnippet: body,
    highlights: highlights,
    caseNumbers: caseInfo.caseNumbers,
    subjectCaseNumbers: caseInfo.subjectCaseNumbers,
    clientCode: caseInfo.clientCode,
    caseCategory: caseInfo.caseCategory,
    caseStatus: caseInfo.caseStatus,
    _message: message,
    _attachments: attachments,
  };
}


// ===================== 智慧附件過濾（V1 教訓） =====================

/**
 * 智慧取得附件：過濾 inline 簽名圖片
 *
 * 策略（V1 累積的經驗）：
 * 1. Gmail 信件（有 gmail_quote / gmail_signature 標記）：
 *    分析 HTML 結構，若主文區域有 cid: 引用 → 保留全部（可能含截圖）
 *    若主文區域無 cid: → 過濾所有 inline 圖片
 * 2. Outlook 等非 Gmail 信件（無標記）：
 *    無法判斷主文邊界，直接用 includeInlineImages:false 過濾
 *    正文截圖仍完整保留於 .eml 檔中
 */
function _getSmartAttachments(message) {
  const htmlBody = message.getBody() || '';
  const hasGmailMarkers =
    htmlBody.indexOf('gmail_quote') !== -1 || htmlBody.indexOf('gmail_signature') !== -1;

  if (hasGmailMarkers) {
    // Gmail 信件：用主文邊界偵測
    if (_hasBodyInlineImages(htmlBody)) {
      // 主文含 inline 圖片（可能是截圖），保留全部附件
      return message.getAttachments() || [];
    } else {
      // 主文沒有 inline 圖片，過濾簽名檔圖片
      return message.getAttachments({ includeInlineImages: false }) || [];
    }
  } else {
    // Outlook 等非 Gmail 信件：直接過濾 inline
    return message.getAttachments({ includeInlineImages: false }) || [];
  }
}

/**
 * 檢查 Gmail HTML body 的「主文區域」是否有 inline 圖片引用（cid:）
 * 主文區域 = gmail_quote 和 gmail_signature 標記之前的內容
 */
function _hasBodyInlineImages(htmlBody) {
  if (!htmlBody) return false;

  let mainContent = htmlBody;

  // 找到 gmail_quote 的位置（引用的舊信）
  const quotePos = htmlBody.indexOf('gmail_quote');
  if (quotePos > 0) mainContent = htmlBody.substring(0, quotePos);

  // 再排除 gmail_signature（自己的簽名檔）
  const sigPos = mainContent.indexOf('gmail_signature');
  if (sigPos > 0) mainContent = mainContent.substring(0, sigPos);

  // 檢查主文區域是否有 cid: 引用
  return mainContent.indexOf('cid:') !== -1;
}

/**
 * 從 HTML body 提取被強調的文字（bold、highlight、彩色字）
 * 這些通常是寄件人想強調的重點（如截止日期、重要指示）
 * 回傳去重後的文字陣列，最多 10 項
 */
function _extractHighlights(message) {
  try {
    const html = message.getBody() || '';
    if (!html) return [];

    // 只看主文（排除引用和簽名檔）
    let mainHtml = html;
    const quotePos = html.indexOf('gmail_quote');
    if (quotePos > 0) mainHtml = html.substring(0, quotePos);
    const sigPos = mainHtml.indexOf('gmail_signature');
    if (sigPos > 0) mainHtml = mainHtml.substring(0, sigPos);

    const highlights = [];

    // 匹配 <b>、<strong> 標籤內的文字
    const boldRegex = /<(?:b|strong)(?:\s[^>]*)?>([^<]{3,200})<\/(?:b|strong)>/gi;
    let m;
    while ((m = boldRegex.exec(mainHtml)) !== null) {
      const text = m[1].replace(/<[^>]+>/g, '').replace(/&[^;]+;/g, ' ').trim();
      if (text.length >= 3) highlights.push(text);
    }

    // 匹配有 color 或 background-color 的 font/span（排除黑色和白色）
    const colorRegex = /<(?:font|span)[^>]*(?:color\s*=\s*["'](?!(?:#000|#fff|black|white))([^"']+)["']|background[^:]*:[^;"]*(?!(?:white|transparent)))[^>]*>([^<]{3,200})<\/(?:font|span)>/gi;
    while ((m = colorRegex.exec(mainHtml)) !== null) {
      const text = m[2].replace(/<[^>]+>/g, '').replace(/&[^;]+;/g, ' ').trim();
      if (text.length >= 3) highlights.push(text);
    }

    // 去重，最多 10 項
    return [...new Set(highlights)].slice(0, 10);
  } catch (e) {
    return [];
  }
}


// ===================== Gemini LLM =====================

const SYSTEM_PROMPT = `你是 IP Winner 智財事務所的 email 分類助手。根據以下資訊判斷 email 類型並產生語義檔名。

## 輸入格式
你會收到一封 email 的結構化 JSON：
- subject: 信件標題（已去除 RE:/Fwd: 前綴）
- original_subject: 原始標題
- direction: "F"（收到）或 "T"（寄出）
- role: "C"（客戶）, "A"（代理人）, "G"（政府）, "X"（未知）
  ※ 收到的信（F）：role = 寄件人角色；寄出的信（T）：role = 收件人角色
- sender: 寄件人 email
- recipients: 收件人列表
- case_numbers: 案號列表
- has_attachments: 是否有附件
- attachment_names: 附件檔名列表
- body_snippet: 信件內文前 1500 字
- highlighted_text: 信件中被加粗/上色/標記的重點文字（從 HTML 提取，通常是寄件人想強調的內容如期限、重要指示）
- email_date: 信件寄送日期（ISO格式），用來判斷信中提到的期限是否已過期

## 第一步：確認收發文碼
- F + C → FC | F + A → FA | F + G → FG | F + X → FX
- T + C → TC | T + A → TA | T + G → TG | T + X → TX

### 未知來源角色推斷（role = "X" 時必做）
**代理人（A）線索：** 送件報告、Filing Report、確認承辦、帳單/Invoice/Debit Note、正式商業英文提及 filing/prosecution/registration
**客戶（C）線索：** 簡短回覆（ok/確認/同意）、指示答辯/領證、提及「我們的產品」「我司」、一般企業 domain
**政府（G）線索：** 官方通知/核駁/電子收據、domain 含 gov

能推斷 → inferred_role 填 C/A/G，收發碼改具體碼。無法判斷 → inferred_role 填 null，保留 FX/TX。

## 第二步：產生語義檔名
前綴引導 + 自由摘要，必須從以下前綴選擇：

**FC：** 確認、已簽、提供、回覆、指示、答辯指示、領證指示、年費指示
**TC：** 進度通知-、送核-、待簽-、確認、提醒回覆指示-、OA分析、商標監控、商標核准通知、專利領證通知
**FA：** 轉寄-、送件報告、{代理人}帳單、確認承辦、送核-、來信告知
**TA：** 委託、確認可送件、校稿意見、提供、詢問、告知、結案通知
**FG：** E-Filing Receipt、TIPO電子收據、TIPO電子帳單、線上變更...成功通知

### 括號事項碼：-(P-新), -(OA1), -(ROA1), -(領證), -(領證+Y1-3), -(T-新), -(變更申請人名稱) 等
**去重規則：** 如果前綴本身已包含事項碼資訊，括號中不再重複。
  例：前綴是「OA分析」→ 事項碼寫 -(OA1)，完整為「OA分析-(OA1)」✓
  錯誤範例：「OA分析-(OA1)」中再寫「OA分析-(OA1)-(OA1)」✗
  錯誤範例：前綴用「OA分析」但事項碼也寫入OA → 結果出現「OA分析-(OA分析)」或「提醒回覆指示-OA分析-(OA1)」這種重複，應精簡為「提醒回覆指示-(OA1)」或「提醒回覆指示-OA1分析」二選一

### 代理人帳單代碼：bskb.com→BSKB帳單, naipo.com→NAIP帳單, cpaglobal.com→CPA帳單, atmac.com.au→ATMAC帳單

### 截止日（必加）：只要信件中出現任何行動期限，語義名稱就**必須**包含日期
⚠ 這是強制規則，不可省略。即使是委託信、新案申請、送件指示，只要信中有要求完成/回覆的日期，就要加在語義名最後。
格式：{前綴}-{摘要}-{事項碼}-期限M/DD 或 -建議M/DD前回覆 或 -通知官方期限M/DD
例：「委託-(B-新)-期限3/17」「委託-英國商標新案申請-(T-新)-期限3/17」「轉寄-(OA1)-建議3/15前回覆」
常見被漏掉的情況：
- TA委託信中「please prepare draft by [日期]」「for review by [日期]」→ 這就是期限，必須加
- TC送核信中「請於 [日期] 前回覆」→ 必須加
- 只有信件完全沒提到任何日期時，才不加期限

期限選擇規則（信件常有多個日期，必須選對）：

**第一步：辨識信中的所有日期及其性質**
- 「官方期限」：政府機關設定的最後期限（如 Opposition Deadline、OA response due date、答辯期限）
- 「我方/本所期限」：本所要求對方回覆的期限（如「X/XX前回覆」「by [date]」「for review by [date]」）
- 「背景日期」：申請日、公告日、來信日等非行動性日期

**第二步：根據收發碼選擇正確期限**
- TA（寄給代理人）→ 用「我方要求代理人回覆的期限」，**絕對不是官方期限**
  ⚠ 這是最常犯的錯誤！信件同時出現官方期限和我方要求期限時，必須選我方要求的
  例：信中有「Opposition Deadline: 2026-03-23」（官方）和「for review by 2026-03-17」（我方要求）→ 必須用 3/17
  例：信中有「答辯期限：4/16」（官方）和「請於 3/15 前提供初步意見」（我方要求）→ 必須用 3/15
- TC（寄給客戶）→ 用「我方要求客戶回覆的期限」（通常是「X/XX前回覆」「回覆本所期限」）
- FA（代理人寄來）→ 用「代理人要求我方回覆的期限」
- FC（客戶寄來）→ 用信中提到的行動期限

**第三步：檢查期限是否已過期（用 email_date 比對）**
將選出的期限與 email_date 比較：
- 如果選出的期限 ≥ email_date → 正常使用該期限
- 如果選出的期限 < email_date（已過期）→ 改用「官方期限」，且：
  - 語義名稱中的日期改為官方期限日期
  - 用詞從「建議X/XX前回覆」改為「通知官方期限X/XX」
  例：TC信，回覆本所期限 2/23 已過（email_date=3/13），官方期限 3/23
  → 語義名：「提醒回覆指示-(OA1)-通知官方期限3/23」而非「建議2/23前回覆」
  例：TC信，回覆本所期限 2/16 已過（email_date=3/13），官方期限 4/16
  → 語義名：「提醒回覆指示-(OA1)-通知官方期限4/16」

**日期格式：** 期限在語義名中一律用 M/DD 格式（如 3/17、4/16），不加年份

### 高頻模板參考
%%TEMPLATES%%

### 近期人工修正紀錄
%%CORRECTIONS%%

## 第三步：案件類別（由程式碼根據案號第10碼判定，LLM 不需處理）

## 第四步：判斷歸檔案號
分析信件內容，判斷這封信「實際應該歸檔到哪些案號」：
- 信件主要處理/討論的案號 → 放入 filing_case_numbers
- 只是順便引用/參考的案號 → 不放入
- 如果主旨有案號但內文討論的案號更多 → 以實際內容為準
- 範例：異議申請信主案是 KOIS23004BGB5，內文引用先前商標 KOIS23004TGB1 作為依據 → 只歸 ["KOIS23004BGB5"]
- 範例：PCT 與 US 案對應表列出 19 個案號都是信件主題 → 歸全部 19 個

## 輸出 JSON（只回 JSON，不加其他文字）
{
  "send_receive_code": "FA",
  "eml_filename": "送件報告-(ROA1)",
  "case_category": "專利",
  "case_status": null,
  "filing_case_numbers": ["XXXX00000PXX1"],
  "inferred_role": null,
  "confidence": 0.92,
  "reasoning": "簡短理由"
}

注意：語義名稱用繁體中文、25字以內、不照抄英文標題。
注意：reasoning 請控制在 30 字以內，只寫關鍵判斷依據。
注意：filing_case_numbers 只列實際要歸檔的案號，不含參考引用的案號。`;


function _buildPrompt(corrections, templates) {
  let sysPrompt = SYSTEM_PROMPT;

  if (templates && templates.length > 0) {
    sysPrompt = sysPrompt.replace('%%TEMPLATES%%',
      templates.map(t => '- ' + t[2] + ': ' + t[3]).join('\n'));
  } else {
    sysPrompt = sysPrompt.replace('%%TEMPLATES%%', '（尚無模板資料）');
  }

  if (corrections && corrections.length > 0) {
    sysPrompt = sysPrompt.replace('%%CORRECTIONS%%',
      corrections.map(c => {
        let line = '- 「' + c.subject + '」';
        if (c.aiCode !== c.finalCode) line += ' ' + c.aiCode + '→' + c.finalCode;
        if (c.aiName !== c.finalName) line += ' 「' + c.aiName + '」→「' + c.finalName + '」';
        if (c.reason) line += ' （原因：' + c.reason + '）';
        return line;
      }).join('\n'));
  } else {
    sysPrompt = sysPrompt.replace('%%CORRECTIONS%%', '（尚無修正紀錄）');
  }

  return sysPrompt;
}


/**
 * 單封呼叫 Gemini（備用）
 */
function _callGemini(preprocessed, corrections, templates) {
  const results = _callGeminiBatch([preprocessed], corrections, templates);
  return results[0];
}

/**
 * 批次並行呼叫 Gemini — 用 UrlFetchApp.fetchAll() 一次送多封
 * 大幅加速：10 封從 ~100 秒 → ~15 秒
 */
function _callGeminiBatch(preprocessedList, corrections, templates) {
  const apiKey = getApiKey();
  const endpoint = CONFIG.GEMINI_ENDPOINT + CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;
  const sysPrompt = _buildPrompt(corrections, templates);

  // 組裝所有請求
  const requests = preprocessedList.map(preprocessed => {
    const userPrompt = JSON.stringify({
      subject: preprocessed.subject,
      original_subject: preprocessed.originalSubject,
      direction: preprocessed.direction,
      role: preprocessed.role,
      sender: preprocessed.sender,
      recipients: preprocessed.recipients.slice(0, 5),
      case_numbers: preprocessed.caseNumbers,
      has_attachments: preprocessed.hasAttachments,
      attachment_names: preprocessed.attachmentNames.slice(0, 10),
      body_snippet: preprocessed.bodySnippet,
      highlighted_text: preprocessed.highlights || [],
      email_date: Utilities.formatDate(preprocessed.date, 'Asia/Taipei', 'yyyy-MM-dd'),
    });

    return {
      url: endpoint,
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
        systemInstruction: { parts: [{ text: sysPrompt }] },
        generationConfig: {
          temperature: CONFIG.GEMINI_TEMPERATURE,
          maxOutputTokens: CONFIG.GEMINI_MAX_TOKENS,
          responseMimeType: 'application/json',
        },
      }),
      muteHttpExceptions: true,
    };
  });

  // 一次送出所有請求（並行）
  const responses = UrlFetchApp.fetchAll(requests);

  // 解析所有回應
  return responses.map((resp, idx) => {
    if (resp.getResponseCode() !== 200) {
      Logger.log('Gemini API 錯誤 (#' + idx + '): ' + resp.getResponseCode());
      return {
        sendReceiveCode: null, emlFilename: null, caseCategory: '未分類',
        caseStatus: null, inferredRole: null, confidence: 0,
        reasoning: 'API錯誤:' + resp.getResponseCode(),
        tokenInfo: { inputTokens: 0, outputTokens: 0, totalTokens: 0 },
      };
    }

    try {
      const respJson = JSON.parse(resp.getContentText());

      const usage = respJson.usageMetadata || {};
      const tokenInfo = {
        inputTokens: usage.promptTokenCount || 0,
        outputTokens: usage.candidatesTokenCount || 0,
        totalTokens: (usage.promptTokenCount || 0) + (usage.candidatesTokenCount || 0),
      };

      let text = respJson.candidates[0].content.parts[0].text;

      let r;
      try {
        r = JSON.parse(text);
      } catch (parseErr) {
        r = _repairJson(text);
        if (!r) {
          Logger.log('LLM 回應解析失敗 (#' + idx + '): ' + parseErr.message);
          Logger.log('原始回應: ' + text.substring(0, 500));
          return {
            sendReceiveCode: null, emlFilename: null, caseCategory: '未分類',
            caseStatus: null, inferredRole: null, confidence: 0,
            reasoning: '解析失敗', tokenInfo: tokenInfo,
          };
        }
      }

      return {
        sendReceiveCode: r.send_receive_code || 'FX',
        emlFilename: r.eml_filename || '',
        caseCategory: r.case_category || '未分類',
        caseStatus: r.case_status || null,
        filingCaseNumbers: Array.isArray(r.filing_case_numbers) ? r.filing_case_numbers : [],
        inferredRole: r.inferred_role || null,
        confidence: parseFloat(r.confidence) || 0,
        reasoning: r.reasoning || '',
        tokenInfo: tokenInfo,
      };
    } catch (e) {
      Logger.log('LLM 回應解析失敗 (#' + idx + '): ' + e.message);
      return {
        sendReceiveCode: null, emlFilename: null, caseCategory: '未分類',
        caseStatus: null, inferredRole: null, confidence: 0, reasoning: '解析失敗',
        tokenInfo: { inputTokens: 0, outputTokens: 0, totalTokens: 0 },
      };
    }
  });
}


/**
 * 修復截斷的 JSON — 常見情況是 reasoning 太長被截斷
 * 策略：逐步嘗試修復尾端
 */
function _repairJson(text) {
  // 先嘗試完整的 { ... }
  const match = text.match(/\{[\s\S]*/);
  if (!match) return null;

  let json = match[0];

  // 嘗試 1：原文已經完整
  try { return JSON.parse(json); } catch (e) { /* 繼續 */ }

  // 嘗試 2：字串被截斷 — 關掉未結束的字串，補上 }
  // 例如 ..."reasoning": "寄件者為 Questel，內容為...（被截斷）
  const repairs = [
    json + '"}',           // 缺 " 和 }
    json + '"}\n}',        // 缺 " 和兩個 }
    json + '}',            // 缺 }
  ];

  // 嘗試 3：把截斷的 reasoning 之後的部分砍掉，重建
  // 找到 reasoning 之前已經解析成功的欄位
  const reasoningIdx = json.indexOf('"reasoning"');
  if (reasoningIdx > 0) {
    // 把 reasoning 砍掉，直接給個預設值
    const before = json.substring(0, reasoningIdx);
    repairs.push(before + '"reasoning": "（截斷）"}');
  }

  for (const attempt of repairs) {
    try { return JSON.parse(attempt); } catch (e) { /* 繼續 */ }
  }

  return null;
}


// ===================== Sheet 讀寫 =====================

function _loadSenderMap() {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SENDERS);
  if (!sheet) return new Map();
  const data = sheet.getDataRange().getValues();
  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim().toLowerCase();
    const role = String(data[i][1]).trim().toUpperCase();
    if (key && role) map.set(key, { role });
  }
  return map;
}

function _getProcessedIds() {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
  if (!sheet) return new Set();
  const data = sheet.getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const aiName = String(data[i][8] || '');  // col 8: AI語義名
    // 失敗的紀錄不算已處理，下次會自動重試
    if (id && !aiName.startsWith('[失敗]')) ids.add(id);
  }
  return ids;
}

function _getRecentCorrections(limit) {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const corrections = [];
  for (let i = data.length - 1; i >= 1 && corrections.length < limit; i--) {
    const finalCode = String(data[i][12] || '').trim();       // col 12: 最終收發碼
    const correctedName = String(data[i][13] || '').trim();   // col 13: 修正後名稱
    const correctionReason = String(data[i][14] || '').trim(); // col 14: 修正原因
    if (finalCode || correctedName) {
      corrections.push({
        subject: data[i][2],
        aiCode: data[i][4],
        aiName: data[i][8],    // col 8: AI語義名
        finalCode: finalCode || data[i][4],
        finalName: correctedName || data[i][8],
        reason: correctionReason,
      });
    }
  }
  return corrections;
}

function _loadTemplates() {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RULES);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

function _appendLogRecords(records) {
  if (!records || records.length === 0) return;
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
  const rows = records.map(r => [
    r.messageId, r.date, r.originalSubject, r.sender,
    r.aiCode, r.inferredRole,
    r.filingCaseNumbers || '',   // col 6: 歸檔案號（實際存檔用的）
    r.allCaseNumbers || '',      // col 7: 內文案號（全部偵測到的）
    r.aiSemanticName,
    r.confidence, r.caseCategory, r.sourceStatus,
    '', '', '', '', '', 0,
    r.inputTokens || 0, r.outputTokens || 0,
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function _setSetting(key, value) {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function _addSender(emailOrDomain, role, note) {
  const ss = _getSpreadsheet();
  const senderSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SENDERS);
  const key = emailOrDomain.toLowerCase();
  const newRole = role.toUpperCase();

  // 檢查是否已存在（比對第 1 欄的 email/domain）
  const data = senderSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {  // 跳過 header
    if (String(data[i][0]).toLowerCase() === key) {
      // 已存在 → 更新角色和備註（如果角色不同才更新）
      if (String(data[i][1]).toUpperCase() !== newRole) {
        senderSheet.getRange(i + 1, 2).setValue(newRole);
        senderSheet.getRange(i + 1, 3).setValue(note || '');
        Logger.log('  📝 Sender 更新: ' + key + ' ' + data[i][1] + '→' + newRole);
      }
      return;
    }
  }

  // 不存在 → 新增
  senderSheet.appendRow([key, newRole, note || '']);
  Logger.log('  📝 Sender 新增: ' + key + ' = ' + newRole);
}


// ===================== Drive 管理 =====================

function _saveEmailToDrive(preprocessed, llmResult) {
  const result = { emlFileId: null, errors: [], timing: {} };

  try {
    const t0 = Date.now();
    const rootFolder = _getDriveRootFolder();

    const date = Utilities.formatDate(preprocessed.date, 'Asia/Taipei', 'yyyyMMdd');
    const code = llmResult.sendReceiveCode || preprocessed.sendReceiveCode;

    // 歸檔案號邏輯（由 _determineFinalResult 決定）：
    //   LLM 判定有歸檔案號 → 用第一個案號建資料夾（多個建多資料夾）
    //   主旨無案號但內文有 → 用客戶碼（如 BRIT）→ 未分類/BRIT
    //   完全沒案號 → '無案號' → 未分類/無案號
    const filingCases = llmResult.filingCaseNumbers || [];
    let folderCaseNum;  // 主資料夾名稱
    let caseLabel;      // EML 檔名中的案號部分
    if (filingCases.length > 0) {
      folderCaseNum = filingCases[0];
      caseLabel = folderCaseNum;
      if (filingCases.length > 1) {
        caseLabel += '等' + filingCases.length + '案';
      }
    } else if (preprocessed.clientCode) {
      // 主旨無案號但內文有 → 用客戶碼作資料夾
      folderCaseNum = preprocessed.clientCode;
      caseLabel = preprocessed.clientCode;
    } else {
      folderCaseNum = '無案號';
      caseLabel = '無案號';
    }

    const semantic = llmResult.emlFilename || _fallbackSemanticName(preprocessed.subject);
    const baseName = date + '-' + code + '-' + caseLabel + '-' + semantic;

    // 分類：由 _determineFinalResult 已根據歸檔案號判定
    const category = llmResult.caseCategory || '未分類';
    const catFolder = _getOrCreateFolder(rootFolder, category);
    const targetFolder = _getOrCreateFolder(catFolder, folderCaseNum);

    const tFolder = Date.now();
    Logger.log('    📂 資料夾定位: ' + ((tFolder - t0) / 1000).toFixed(2) + 's');

    // 掃描資料夾現有檔案（一次讀取，供 EML 和附件共用）
    const existingFiles = _scanFolderFiles(targetFolder);
    const tScan = Date.now();
    Logger.log('    🔍 掃描現有檔案: ' + ((tScan - tFolder) / 1000).toFixed(2) + 's (' + Object.keys(existingFiles).length + ' 個)');

    // 取一次 EML 原始內容，後面主案號和多案號都重用（省掉重複 API 呼叫）
    let rawContent = null;
    try {
      const tRawStart = Date.now();
      rawContent = GmailApp.getMessageById(preprocessed.messageId).getRawContent();
      const tRawEnd = Date.now();
      const rawSizeKB = (rawContent.length / 1024).toFixed(1);
      Logger.log('    📥 getRawContent: ' + ((tRawEnd - tRawStart) / 1000).toFixed(2) + 's (' + rawSizeKB + ' KB)');

      // 重跑保護：同名 + 同大小 → 跳過不重建
      const existingId = _findExistingFile(existingFiles, baseName, '.eml', rawContent.length);
      if (existingId) {
        result.emlFileId = existingId;
        Logger.log('    ⏭️ EML 已存在，跳過: ' + baseName + '.eml');
      } else {
        const filename = _getUniqueFileName(existingFiles, baseName, '.eml');
        const blob = Utilities.newBlob(rawContent, 'message/rfc822', filename);
        const tCreateStart = Date.now();
        const created = targetFolder.createFile(blob);
        const tCreateEnd = Date.now();
        Logger.log('    💾 createFile(EML): ' + ((tCreateEnd - tCreateStart) / 1000).toFixed(2) + 's');
        result.emlFileId = created.getId();
        existingFiles[filename] = { size: rawContent.length, id: created.getId() };
      }
    } catch (emlErr) {
      result.errors.push('EML 下載失敗: ' + emlErr.message);
    }

    // 附件（TA/TC/TG 寄出的信不存附件，只存 EML）
    // 附件命名規則：{baseName}-附件1.pdf, {baseName}-附件2.xlsx, ...
    const skipAttachments = ['TA', 'TC', 'TG'].includes(code);
    if (!skipAttachments && preprocessed._attachments.length > 0) {
      Logger.log('    📎 附件數量: ' + preprocessed._attachments.length);
      let attIdx = 0;
      for (const att of preprocessed._attachments) {
        try {
          attIdx++;
          const origName = att.getName();
          // 取原始副檔名
          const dotPos = origName.lastIndexOf('.');
          const ext = dotPos > 0 ? origName.substring(dotPos) : '';
          const attBaseName = baseName + '-附件' + attIdx;

          const attBlob = att.copyBlob();
          const attSize = attBlob.getBytes().length;

          // 重跑保護：同名 + 同大小附件跳過
          const existingAtt = _findExistingFile(existingFiles, attBaseName, ext, attSize);
          if (existingAtt) {
            Logger.log('    ⏭️ 附件已存在，跳過: ' + attBaseName + ext);
          } else {
            const attFilename = _getUniqueFileName(existingFiles, attBaseName, ext);
            attBlob.setName(attFilename);
            const tAttStart = Date.now();
            targetFolder.createFile(attBlob);
            const tAttEnd = Date.now();
            Logger.log('    💾 createFile(附件): ' + ((tAttEnd - tAttStart) / 1000).toFixed(2) + 's → ' + attFilename.substring(0, 60));
            existingFiles[attFilename] = { size: attSize, id: '' };
          }
        }
        catch (e) { result.errors.push('附件「' + att.getName() + '」失敗'); }
      }
    }

    // 多案號：LLM 判定的歸檔案號 > 1 個才建多資料夾
    if (filingCases.length > 1 && rawContent) {
      Logger.log('    📋 多案號處理（LLM歸檔）: ' + (filingCases.length - 1) + ' 個副案號');
      for (let i = 1; i < filingCases.length; i++) {
        try {
          const secCaseNum = filingCases[i];
          const tSecStart = Date.now();
          const secFolder = _getOrCreateFolder(catFolder, secCaseNum);

          let secCaseLabel = secCaseNum;
          secCaseLabel += '等' + filingCases.length + '案';
          const secBaseName = date + '-' + code + '-' + secCaseLabel + '-' + semantic;

          const secExisting = _scanFolderFiles(secFolder);
          const secExistingId = _findExistingFile(secExisting, secBaseName, '.eml', rawContent.length);

          if (secExistingId) {
            Logger.log('    ⏭️ 多案號 EML 已存在，跳過: ' + secCaseNum);
          } else {
            const secFilename = _getUniqueFileName(secExisting, secBaseName, '.eml');
            const secBlob = Utilities.newBlob(rawContent, 'message/rfc822', secFilename);
            secFolder.createFile(secBlob);
            const tSecEnd = Date.now();
            Logger.log('    💾 多案號 EML(' + secCaseNum + '): ' + ((tSecEnd - tSecStart) / 1000).toFixed(2) + 's');
          }
        } catch (e) { result.errors.push('多案號 EML(' + filingCases[i] + ') 失敗'); }
      }
    }

    const tTotal = Date.now();
    Logger.log('    ⏱️ Drive 總計: ' + ((tTotal - t0) / 1000).toFixed(2) + 's');
  } catch (e) {
    result.errors.push('儲存失敗: ' + e.message);
  }

  return result;
}

/**
 * 掃描資料夾內的檔案，回傳 { name → { size, id } } 對照表
 * 供去重和跳過已存在檔案使用
 */
function _scanFolderFiles(targetFolder) {
  const map = {};
  const files = targetFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    map[f.getName()] = { size: f.getSize(), id: f.getId() };
  }
  return map;
}

/**
 * 同檔名去重（V1 教訓）
 * 如果已有同名檔案 → 加 (1), (2), ...
 */
function _getUniqueFileName(existingFiles, baseName, ext) {
  let candidate = baseName + ext;
  if (!existingFiles[candidate]) return candidate;

  for (let n = 1; n <= 999; n++) {
    candidate = baseName + '(' + n + ')' + ext;
    if (!existingFiles[candidate]) return candidate;
  }

  return baseName + '(' + new Date().getTime() + ')' + ext;
}

/**
 * 檢查資料夾裡是否已有同名 + 同大小的檔案（代表重跑時的重複檔）
 * 有 → 回傳該檔案的 ID（跳過不重建）
 * 沒有 → 回傳 null
 */
function _findExistingFile(existingFiles, baseName, ext, blobSize) {
  const candidate = baseName + ext;
  const existing = existingFiles[candidate];
  if (existing && Math.abs(existing.size - blobSize) < 100) {
    // 同名且大小差距 < 100 bytes（EML 可能有微小時間差）→ 視為同一檔案
    return existing.id;
  }
  return null;
}

// ===================== 核心處理 =====================

function _determineFinalResult(preprocessed, llmResult) {
  const hasSubjectCase = preprocessed.subjectCaseNumbers && preprocessed.subjectCaseNumbers.length > 0;

  // 歸檔案號邏輯：
  //   主旨有案號 → 優先用 LLM 的 filingCaseNumbers，fallback 用主旨案號
  //   主旨無案號 → 不歸案號資料夾（用客戶碼或無案號）
  let filingCaseNumbers = [];
  if (hasSubjectCase) {
    const llmFiling = llmResult.filingCaseNumbers || [];
    filingCaseNumbers = llmFiling.length > 0 ? llmFiling : preprocessed.subjectCaseNumbers;
  }

  // 分類邏輯：用歸檔案號的第10碼判定（不用主旨案號，因為 LLM 可能修正）
  let finalCategory = '未分類';
  if (filingCaseNumbers.length > 0) {
    const PATENT_TYPES = 'PMDAC';
    const TRADEMARK_TYPES = 'TBW';
    const typeChars = filingCaseNumbers.map(cn => cn.length >= 10 ? cn.charAt(9) : '');
    const hasPatent = typeChars.some(c => PATENT_TYPES.includes(c));
    const hasTrademark = typeChars.some(c => TRADEMARK_TYPES.includes(c));
    if (hasPatent && !hasTrademark) finalCategory = '專利';
    else if (hasTrademark && !hasPatent) finalCategory = '商標';
  }

  // 多案號狀態：看歸檔案號數量
  let finalCaseStatus = preprocessed.caseStatus;
  if (filingCaseNumbers.length > 1) {
    finalCaseStatus = '多案號';
  }

  const result = {
    sendReceiveCode: llmResult.sendReceiveCode || preprocessed.sendReceiveCode,
    emlFilename: llmResult.emlFilename || '',
    caseCategory: finalCategory,
    caseStatus: finalCaseStatus,
    filingCaseNumbers: filingCaseNumbers,
    problemLabel: null,
    sourceStatus: 'na',
  };

  if (preprocessed.role === 'X') {
    if (llmResult.inferredRole && llmResult.confidence >= CONFIG.CONFIDENCE_INFER) {
      result.problemLabel = '自動辨識來源';
      result.sourceStatus = 'pending';
    } else if (llmResult.inferredRole && llmResult.confidence < CONFIG.CONFIDENCE_INFER) {
      result.sendReceiveCode = preprocessed.sendReceiveCode;
      result.problemLabel = '未知來源';
    } else {
      result.problemLabel = llmResult.confidence < 0.3 ? '已跳過' : '未知來源';
    }
  }

  if (!result.problemLabel) {
    if (llmResult.confidence < CONFIG.CONFIDENCE_AUTO) {
      result.problemLabel = '待確認';
    }
  }

  return result;
}

function _applyLabels(message, result) {
  const thread = message.getThread();
  const tryAdd = (name) => { const l = _getLabel(name); if (l) thread.addLabel(l); };

  if (result.sendReceiveCode) tryAdd(result.sendReceiveCode);
  if (result.caseCategory) tryAdd(result.caseCategory);
  if (result.caseStatus) tryAdd(result.caseStatus);
  if (result.problemLabel) tryAdd(result.problemLabel);
}

function _processEmailBatch(query, limit, shouldDownload) {
  const stats = { processed: 0, auto: 0, needConfirm: 0, autoIdentify: 0, unknown: 0, errors: 0 };

  const senderMap = _loadSenderMap();
  const processedIds = _getProcessedIds();
  const corrections = _getRecentCorrections(20);
  const templates = _loadTemplates();

  _ensureLabels();

  // 搜尋策略：不用 -label 排除（因為同 thread 新回覆會被連帶排除）
  // 改用 Message ID 在 Sheet 裡判斷是否已處理（V1 教訓）
  const searchQuery = query || '';
  const messages = [];
  let searchStart = 0;
  const searchBatch = 100;

  while (messages.length < limit) {
    const threads = GmailApp.search(searchQuery, searchStart, searchBatch);
    if (threads.length === 0) break;

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        if (messages.length >= limit) break;
        if (!processedIds.has(msg.getId())) messages.push(msg);
      }
      if (messages.length >= limit) break;
    }

    searchStart += searchBatch;
    // 安全上限：最多搜尋 500 個 threads
    if (searchStart >= 500) break;
  }

  Logger.log('取得 ' + messages.length + ' 封待處理信件');

  // 分批並行處理：每批 PARALLEL_BATCH 封同時呼叫 Gemini
  const PARALLEL_BATCH = 10;

  for (let batchStart = 0; batchStart < messages.length; batchStart += PARALLEL_BATCH) {
    const batchMessages = messages.slice(batchStart, batchStart + PARALLEL_BATCH);
    Logger.log('並行處理第 ' + (batchStart + 1) + '-' + (batchStart + batchMessages.length) + ' 封...');

    // Step 1: 預處理（規則引擎，不花 API）
    const preprocessedList = [];
    for (let i = 0; i < batchMessages.length; i++) {
      try {
        preprocessedList.push(_preprocessMessage(batchMessages[i], senderMap));
      } catch (e) {
        stats.errors++;
        const msgId = batchMessages[i].getId();
        Logger.log('❌ #' + (batchStart + i + 1) + ' 預處理失敗 (' + msgId + '): ' + e.message);
        _appendLogRecords([{
          messageId: msgId, aiCode: '', date: Utilities.formatDate(batchMessages[i].getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          originalSubject: batchMessages[i].getSubject() || '', sender: '',
          inferredRole: '', filingCaseNumbers: '', allCaseNumbers: '',
          aiSemanticName: '[失敗] ' + e.message.substring(0, 50),
          confidence: 0, caseCategory: '', sourceStatus: 'na',
          inputTokens: 0, outputTokens: 0,
        }]);
      }
    }

    if (preprocessedList.length === 0) continue;

    // Step 2: 並行呼叫 Gemini（一次送出所有請求）
    const apiStart = Date.now();
    const llmResults = _callGeminiBatch(preprocessedList, corrections, templates);
    const apiMs = Date.now() - apiStart;
    Logger.log('⚡ Gemini API 並行呼叫完成：' + preprocessedList.length + ' 封，耗時 ' + (apiMs / 1000).toFixed(1) + ' 秒');

    // Step 3: 逐封處理結果（掛標籤、存 Drive、寫 log）— 每封獨立，一封失敗不影響其他封
    let driveMs = 0;
    for (let i = 0; i < preprocessedList.length; i++) {
      const preprocessed = preprocessedList[i];
      const llmResult = llmResults[i];
      const messageId = preprocessed.messageId;
      const emailIdx = batchStart + i + 1;

      try {
        const finalResult = _determineFinalResult(preprocessed, llmResult);

        // Log 每封信的分類結果
        const confPct = Math.round((llmResult.confidence || 0) * 100);
        const filingInfo = (finalResult.filingCaseNumbers || []).length > 0
          ? ' 歸檔:' + finalResult.filingCaseNumbers.length + '案'
          : '';
        Logger.log('📧 #' + emailIdx + ' [' + (finalResult.sendReceiveCode || '??') + '] '
          + (preprocessed.sender || '').substring(0, 30) + ' → '
          + (llmResult.emlFilename || '(無檔名)').substring(0, 60)
          + ' (信心: ' + confPct + '%' + filingInfo + ')');

        _applyLabels(preprocessed._message, finalResult);

        if (shouldDownload) {
          const driveStart = Date.now();
          const driveResult = _saveEmailToDrive(preprocessed, finalResult);
          driveMs += Date.now() - driveStart;
          if (driveResult.errors.length > 0) {
            Logger.log('  ⚠️ #' + emailIdx + ' 下載問題: ' + driveResult.errors.join('; '));
          }
        }

        // 每封處理完立即寫入 Sheet（失敗不影響其他封）
        _appendLogRecords([{
          messageId, aiCode: finalResult.sendReceiveCode,
          date: Utilities.formatDate(preprocessed.date, 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          originalSubject: preprocessed.originalSubject,
          sender: preprocessed.sender,
          inferredRole: llmResult.inferredRole || '',
          filingCaseNumbers: (finalResult.filingCaseNumbers || []).join(', '),
          allCaseNumbers: preprocessed.caseNumbers.join(', '),
          aiSemanticName: llmResult.emlFilename || '',
          confidence: llmResult.confidence,
          caseCategory: finalResult.caseCategory,
          sourceStatus: finalResult.sourceStatus,
          inputTokens: llmResult.tokenInfo ? llmResult.tokenInfo.inputTokens : 0,
          outputTokens: llmResult.tokenInfo ? llmResult.tokenInfo.outputTokens : 0,
        }]);

        stats.processed++;
        if (finalResult.problemLabel === '自動辨識來源') stats.autoIdentify++;
        else if (finalResult.problemLabel === '未知來源') stats.unknown++;
        else if (finalResult.problemLabel === '待確認') stats.needConfirm++;
        else stats.auto++;

      } catch (e) {
        stats.errors++;
        Logger.log('❌ #' + emailIdx + ' 處理失敗 (' + messageId + '): ' + e.message);
        // 失敗也立即寫入 Sheet，不影響其他封
        _appendLogRecords([{
          messageId, aiCode: '', date: Utilities.formatDate(preprocessed.date, 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          originalSubject: preprocessed.originalSubject || '', sender: preprocessed.sender || '',
          inferredRole: '', filingCaseNumbers: '', allCaseNumbers: '',
          aiSemanticName: '[失敗] ' + e.message.substring(0, 50),
          confidence: 0, caseCategory: '', sourceStatus: 'na',
          inputTokens: 0, outputTokens: 0,
        }]);
      }
    }

    Logger.log('💾 Drive 存檔總耗時：' + (driveMs / 1000).toFixed(1) + ' 秒（' + preprocessedList.length + ' 封）');
  }

  _setSetting('lastProcessedTime', new Date().toISOString());
  _setSetting('累計處理數量', (_getProcessedIds().size));

  return stats;
}


// ===================== 進入點：試跑 =====================

/** 試跑 50 封（完整 Phase 1：分類＋掛標籤＋下載 EML） */
function trialRun() {
  const ui = SpreadsheetApp.getUi();
  const q = ui.prompt('🧪 試跑 50 封',
    '請輸入 Gmail 搜尋條件：\n\n' +
    '範例：\n' +
    '  newer_than:7d          （最近 7 天）\n' +
    '  from:bskb.com          （來自 BSKB）\n' +
    '  subject:filing report  （標題含 filing report）\n\n' +
    '留空 = 所有未處理信件',
    ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;

  try {
    const result = _processEmailBatch(q.getResponseText().trim(), 50, true);
    ui.alert('✅ 試跑完成',
      '處理 ' + result.processed + ' 封（含下載 EML＋掛標籤）\n\n' +
      '✅ 自動處理: ' + result.auto + '\n' +
      '⚠️ 待確認: ' + result.needConfirm + '\n' +
      '🔍 自動辨識來源: ' + result.autoIdentify + '\n' +
      '❓ 未知來源: ' + result.unknown + '\n' +
      '❌ 失敗: ' + result.errors + '\n\n' +
      '→ 到「處理紀錄」Sheet 查看結果\n' +
      '→ 到 Gmail 搜尋 label:AI/自動辨識來源 確認 sender',
      ui.ButtonSet.OK);
  } catch (e) { ui.alert('❌ 失敗', e.message, ui.ButtonSet.OK); }
}

/** 試跑 10 封（快速驗證） */
function trialRunSmall() {
  const ui = SpreadsheetApp.getUi();
  const q = ui.prompt('🧪 快速試跑 10 封', '請輸入 Gmail 搜尋條件（留空 = 所有未處理）：', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;

  try {
    const result = _processEmailBatch(q.getResponseText().trim(), 10, true);
    ui.alert('✅ 完成', '處理 ' + result.processed + ' 封\n自動: ' + result.auto +
      ' | 待確認: ' + result.needConfirm + ' | 辨識來源: ' + result.autoIdentify +
      ' | 未知: ' + result.unknown + ' | 失敗: ' + result.errors, ui.ButtonSet.OK);
  } catch (e) { ui.alert('❌ 失敗', e.message, ui.ButtonSet.OK); }
}

/**
 * 從編輯器直接試跑（不需 UI）
 * 適合第一次測試時使用
 */
function trialRunFromEditor() {
  Logger.log('=== 從編輯器試跑 50 封 ===');
  try {
    const result = _processEmailBatch('newer_than:7d', 50, true);
    Logger.log('✅ 試跑完成！');
    Logger.log('處理: ' + result.processed + ' 封');
    Logger.log('自動: ' + result.auto + ' | 待確認: ' + result.needConfirm);
    Logger.log('辨識來源: ' + result.autoIdentify + ' | 未知: ' + result.unknown);
    Logger.log('失敗: ' + result.errors);
    Logger.log('→ 到 Google Sheet 的「處理紀錄」Tab 查看結果');
  } catch (e) {
    Logger.log('❌ 失敗: ' + e.message);
  }
}

/** 測試單封信：搜尋一封，顯示完整分類結果（不寫入 Sheet、不掛標籤、不下載） */
function testSingleEmail() {
  const ui = SpreadsheetApp.getUi();
  const q = ui.prompt('🔬 測試單封信',
    '請輸入搜尋條件（找到的第一封會被分析）：\n\n' +
    '範例：subject:"Filing Report" from:bskb.com',
    ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;

  try {
    const threads = GmailApp.search(q.getResponseText().trim(), 0, 1);
    if (threads.length === 0) { ui.alert('找不到符合的信件'); return; }

    const message = threads[0].getMessages()[0];
    const senderMap = _loadSenderMap();
    const preprocessed = _preprocessMessage(message, senderMap);

    const corrections = _getRecentCorrections(20);
    const templates = _loadTemplates();
    const llmResult = _callGemini(preprocessed, corrections, templates);
    const finalResult = _determineFinalResult(preprocessed, llmResult);

    ui.alert('🔬 單封信分析結果',
      '標題: ' + preprocessed.originalSubject + '\n' +
      'Sender: ' + preprocessed.sender + '\n' +
      '方向: ' + preprocessed.direction + ' | Role: ' + preprocessed.role + '\n' +
      '規則碼: ' + preprocessed.sendReceiveCode + '\n\n' +
      '── LLM 結果 ──\n' +
      '收發碼: ' + llmResult.sendReceiveCode + '\n' +
      '語義名: ' + llmResult.emlFilename + '\n' +
      '案件類別: ' + llmResult.caseCategory + '\n' +
      '推斷角色: ' + (llmResult.inferredRole || '無') + '\n' +
      '信心: ' + llmResult.confidence + '\n' +
      '理由: ' + llmResult.reasoning + '\n\n' +
      '── 最終結果 ──\n' +
      '最終碼: ' + finalResult.sendReceiveCode + '\n' +
      '問題標籤: ' + (finalResult.problemLabel || '無') + '\n' +
      '案號: ' + preprocessed.caseNumbers.join(', ') + '\n' +
      '附件: ' + preprocessed.attachmentNames.join(', '),
      ui.ButtonSet.OK);
  } catch (e) { ui.alert('❌ 失敗', e.message, ui.ButtonSet.OK); }
}


// ===================== 正式處理 =====================

function processEmails() {
  const startTime = Date.now();
  Logger.log('=== 正式處理 ===');

  try {
    const result = _processEmailBatch('', CONFIG.BATCH_SIZE, true);
    Logger.log('完成: ' + JSON.stringify(result));

    if (Date.now() - startTime < CONFIG.TIMEOUT_SAFETY_MS) {
      runFeedback();
    }
  } catch (e) {
    Logger.log('正式處理失敗: ' + e.message);
  }
}


// ===================== 回授偵測 =====================

/**
 * 決定回授學習的對象
 *
 * 規則：
 * 1. F 方向（收到的信）→ 學習寄件人
 * 2. T 方向（寄出的信）→ 學習收件人（寄件人是自己）
 * 3. OWN_DOMAINS → 永遠跳過（自己不可能是客戶/代理人）
 * 4. 公共 domain（gmail/hotmail 等）→ 用完整 email（因為同 domain 有不同人）
 * 5. 專屬 domain（公司信箱）→ 用 @domain（同公司的人角色一致）
 *
 * @return {string|null} 學習目標（email 或 @domain），null 表示不該學習
 */
function _getFeedbackLearnTarget(aiCode, sender, message) {
  let targetEmail;

  if (aiCode.startsWith('T')) {
    // 寄出的信 → 找收件人
    const toEmails = (message.getTo() || '').split(',').map(r => _extractEmail(r)).filter(Boolean);
    targetEmail = toEmails.find(e => !CONFIG.OWN_DOMAINS.includes(_extractDomain(e)));
    if (!targetEmail) return null;  // 沒有外部收件人
  } else {
    // 收到的信 → 用寄件人
    targetEmail = sender;
  }

  const domain = _extractDomain(targetEmail);

  // OWN_DOMAINS 永遠跳過
  if (CONFIG.OWN_DOMAINS.includes(domain)) return null;

  // 公共 domain → 用完整 email
  if (CONFIG.PUBLIC_DOMAINS.includes(domain)) return targetEmail;

  // 專屬 domain → 用 @domain
  return '@' + domain;
}

function runFeedback() {
  Logger.log('=== 回授偵測 ===');

  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let confirmed = 0, corrected = 0, checked = 0;
  let renamed = 0, renameErrors = 0;

  // ── Part 1: Sender 角色回授（pending 狀態 + Gmail label 變化）──
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][11]).trim() !== 'pending') continue;  // col 11: 來源確認狀態
    checked++;

    try {
      const messageId = String(data[i][0]).trim();
      const message = GmailApp.getMessageById(messageId);
      if (!message) continue;

      const threadLabels = message.getThread().getLabels().map(l => l.getName());
      const hasAutoId = threadLabels.includes(CONFIG.LABEL_PREFIX + '/自動辨識來源');

      if (!hasAutoId) {
        let currentCode = null;
        for (const code of CONFIG.SEND_RECEIVE_CODES) {
          if (threadLabels.includes(CONFIG.LABEL_PREFIX + '/' + code)) { currentCode = code; break; }
        }

        const aiCode = String(data[i][4]).trim();
        const inferredRole = String(data[i][5]).trim();
        const sender = String(data[i][3]).trim();

        const learnTarget = _getFeedbackLearnTarget(aiCode, sender, message);
        const finalRole = currentCode ? currentCode.charAt(1) : null;

        if (currentCode && currentCode !== aiCode) {
          sheet.getRange(i + 1, 12).setValue('corrected');   // col 12: 來源確認狀態 (1-based)
          sheet.getRange(i + 1, 13).setValue(currentCode);   // col 13: 最終收發碼
          sheet.getRange(i + 1, 16).setValue(new Date());    // col 16: 修正時間
          sheet.getRange(i + 1, 17).setValue('tag_change');   // col 17: 修正來源
          if (learnTarget && finalRole && finalRole !== 'X') {
            _addSender(learnTarget, finalRole, 'AI推斷' + inferredRole + '→人修正' + finalRole);
          }
          corrected++;
        } else {
          sheet.getRange(i + 1, 12).setValue('confirmed');
          sheet.getRange(i + 1, 16).setValue(new Date());
          sheet.getRange(i + 1, 17).setValue('tag_change');
          if (learnTarget && inferredRole) {
            _addSender(learnTarget, inferredRole, 'AI推斷確認');
          }
          confirmed++;
        }
      }
    } catch (e) {
      Logger.log('回授(sender)失敗: ' + e.message);
    }
  }

  // ── Part 2: 語義名修正 → Drive 檔案改名 ──
  // 偵測條件：「修正後名稱」有值 + 「修正來源」不含 'name'（代表改名尚未執行）
  // 策略：用 Sheet 裡的日期/收發碼/案號/AI語義名 組出原始檔名，精準比對 Drive 裡的檔案
  // 需要重新讀取 sheet data，因為 Part 1 可能已經更新了修正來源欄
  const freshData = sheet.getDataRange().getValues();
  for (let i = 1; i < freshData.length; i++) {
    const correctedName = String(freshData[i][13] || '').trim();   // col 13: 修正後名稱
    const correctionSource = String(freshData[i][16] || '').trim(); // col 16: 修正來源
    const aiName = String(freshData[i][8] || '').trim();           // col 8: AI語義名

    // 跳過：沒填修正名稱、改名已執行過、跟 AI 名稱相同
    if (!correctedName || correctionSource.indexOf('name') !== -1 || correctedName === aiName) continue;

    try {
      // 從 Sheet 裡組出原始檔名，精準比對 Drive 裡的檔案
      const rowDate = freshData[i][1];                                // col 1: 日期
      const rowCode = String(freshData[i][4] || '').trim();           // col 4: AI收發碼
      const rowFilingCases = String(freshData[i][6] || '').trim();    // col 6: 歸檔案號
      const rowCategory = String(freshData[i][10] || '').trim();      // col 10: AI案件類別

      // 組出日期字串 yyyyMMdd
      // ⚠️ 必須用 Spreadsheet 的時區讀回日期，才能還原當初寫入的值
      //    原因：日期寫入時用 Asia/Taipei 格式化成 "2026-03-13 23:12"，
      //    但 Sheets 可能用不同時區（如 America/LA）自動轉成 Date 物件，
      //    再用 Asia/Taipei 格式化回來會飄移日期（如 3/13 → 3/14）
      let dateStr = '';
      if (rowDate instanceof Date) {
        const ssTz = ss.getSpreadsheetTimeZone();
        dateStr = Utilities.formatDate(rowDate, ssTz, 'yyyyMMdd');
      } else {
        const m = String(rowDate).match(/(\d{4})-(\d{2})-(\d{2})/);
        if (m) dateStr = m[1] + m[2] + m[3];
      }

      // 用歸檔案號組出每個案號資料夾的 baseName
      // 多案號時，每個資料夾的 EML 開頭不同（主案號 vs 副案號）
      // 例：BRIT25710PUS1等2案-... 和 BRIT25711PUS2等2案-...
      const filingArr = rowFilingCases ? rowFilingCases.split(/,\s*/) : [];
      const oldBaseNames = [];
      if (dateStr && rowCode && filingArr.length > 0) {
        const suffix = filingArr.length > 1 ? '等' + filingArr.length + '案' : '';
        for (const caseNum of filingArr) {
          const caseLabel = caseNum + suffix;
          oldBaseNames.push(dateStr + '-' + rowCode + '-' + caseLabel + '-' + aiName);
        }
      }

      Logger.log('📝 改名回授: 「' + aiName + '」→「' + correctedName + '」' +
        (oldBaseNames.length > 0 ? ' (' + oldBaseNames.length + ' 個 baseName)' : ''));

      // 找到 Drive 裡的 EML 和附件並改名（多案號會傳多個 baseName）
      const renameCount = _renameDriveFiles(oldBaseNames, aiName, correctedName, rowCategory);

      // 修正來源：追加而非覆蓋（可能已有 tag_change）
      const newSource = correctionSource ? correctionSource + '+name_change' : 'name_change';
      const notFoundSource = correctionSource ? correctionSource + '+name_not_found' : 'name_not_found';

      if (renameCount > 0) {
        sheet.getRange(i + 1, 16).setValue(new Date());    // col 16: 修正時間 (1-based)
        sheet.getRange(i + 1, 17).setValue(newSource);      // col 17: 修正來源
        renamed += renameCount;
        Logger.log('  ✅ 改名成功: ' + renameCount + ' 個檔案');
      } else {
        sheet.getRange(i + 1, 16).setValue(new Date());
        sheet.getRange(i + 1, 17).setValue(notFoundSource);
        Logger.log('  ⚠️ 未找到符合的 Drive 檔案');
      }
    } catch (e) {
      Logger.log('回授(改名)失敗: ' + e.message);
      renameErrors++;
    }
  }

  // ── Part 3: Sheet 直填最終收發碼 → 學習 Sender ──
  // 偵測條件：最終收發碼有值 + 修正來源不含 'sheet_code'（代表尚未處理）
  // 適用情況：人員直接在 Sheet 填正確的收發碼（不透過 Gmail label）
  let sheetCodeLearned = 0;
  const freshData2 = sheet.getDataRange().getValues();
  for (let i = 1; i < freshData2.length; i++) {
    const finalCode = String(freshData2[i][12] || '').trim();       // col 12: 最終收發碼
    const correctionSource = String(freshData2[i][16] || '').trim(); // col 16: 修正來源
    const aiCode = String(freshData2[i][4] || '').trim();

    // 跳過：沒填、已處理過、跟 AI 碼一樣
    if (!finalCode || correctionSource.indexOf('sheet_code') !== -1 || finalCode === aiCode) continue;
    // 跳過：不是合法的收發碼
    if (!CONFIG.SEND_RECEIVE_CODES.includes(finalCode)) continue;

    try {
      const messageId = String(freshData2[i][0]).trim();
      const sender = String(freshData2[i][3]).trim();
      const message = GmailApp.getMessageById(messageId);
      if (!message) continue;

      const finalRole = finalCode.charAt(1);
      const learnTarget = _getFeedbackLearnTarget(aiCode, sender, message);

      if (learnTarget && finalRole && finalRole !== 'X') {
        _addSender(learnTarget, finalRole, 'Sheet直填' + aiCode + '→' + finalCode);
        sheetCodeLearned++;
      }

      // 追加修正來源
      const newSource = correctionSource ? correctionSource + '+sheet_code' : 'sheet_code';
      sheet.getRange(i + 1, 16).setValue(new Date());    // col 16: 修正時間
      sheet.getRange(i + 1, 17).setValue(newSource);      // col 17: 修正來源
    } catch (e) {
      Logger.log('回授(Sheet收發碼)失敗: ' + e.message);
    }
  }

  const msg = '回授偵測完成\n\n' +
    '【Sender 角色 - Gmail tag】\n' +
    '檢查 pending 紀錄: ' + checked + ' 筆\n' +
    '✅ 確認: ' + confirmed + '\n' +
    '🔄 修正: ' + corrected + '\n' +
    '⏳ 尚未處理: ' + (checked - confirmed - corrected) + '\n\n' +
    '【Sender 角色 - Sheet 直填】\n' +
    '📝 學習: ' + sheetCodeLearned + ' 筆\n\n' +
    '【檔名修正】\n' +
    '📝 改名檔案: ' + renamed + ' 個\n' +
    (renameErrors > 0 ? '❌ 改名失敗: ' + renameErrors + ' 筆\n' : '');

  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('🔄 回授偵測', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* 從 trigger 或編輯器呼叫時沒有 UI */ }
}

/**
 * 在 Drive 裡精準找到原始檔案並改名
 *
 * 比對策略：
 *   1. 如果有 oldBaseNames → 檔名必須以其中一個 baseName 開頭（支援多案號，每個案號資料夾的 baseName 不同）
 *   2. 如果 oldBaseNames 為空 → fallback 到只比對 oldSemanticName（較不精準）
 *
 * 搜尋範圍：
 *   如果有 category → 只搜該分類資料夾底下的案號資料夾
 *   否則 → 搜所有分類/案號資料夾
 *
 * @param {string[]} oldBaseNames    完整的原始 baseName 陣列（多案號會有多個，如 ["20260313-FA-BRIT25710PUS1等2案-委託...", "20260313-FA-BRIT25711PUS2等2案-委託..."]）
 * @param {string} oldSemanticName   原始 AI 語義名（baseName 中的語義部分）
 * @param {string} newSemanticName   修正後名稱
 * @param {string} category          案件類別（專利/商標/未分類），縮小搜尋範圍
 * @return {number} 改名的檔案數量
 */
function _renameDriveFiles(oldBaseNames, oldSemanticName, newSemanticName, category) {
  const rootFolder = _getDriveRootFolder();
  let count = 0;
  const baseNames = oldBaseNames || [];

  Logger.log('  🔍 搜尋參數: oldSemanticName=「' + oldSemanticName + '」');
  if (baseNames.length > 0) {
    baseNames.forEach((bn, idx) => Logger.log('  🔍 baseName[' + idx + ']=「' + bn + '」'));
  }

  /**
   * 掃描一個資料夾內的檔案，符合條件就改名
   */
  function _scanAndRename(folder) {
    const folderName = folder.getName();
    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      const oldName = file.getName();
      fileCount++;

      // 比對條件：檔名包含舊語義名
      if (oldName.indexOf(oldSemanticName) === -1) continue;

      // 如果有 baseNames，確認檔名以其中一個開頭（支援多案號）
      if (baseNames.length > 0) {
        const matchesAny = baseNames.some(bn => oldName.indexOf(bn) === 0);
        if (!matchesAny) {
          Logger.log('  ⚠️ 語義名吻合但 baseName 不符: 「' + oldName + '」');
          continue;
        }
      }

      const newName = oldName.replace(oldSemanticName, newSemanticName);
      if (newName !== oldName) {
        file.setName(newName);
        Logger.log('  📄 ' + oldName + ' → ' + newName);
        count++;
      }
    }
    if (fileCount > 0) {
      Logger.log('  📂 ' + folderName + ': 掃描 ' + fileCount + ' 個檔案');
    }
  }

  // 決定搜尋範圍
  const categoriesToSearch = category ? [category] : ['專利', '商標', '未分類'];
  Logger.log('  🔍 搜尋分類: ' + categoriesToSearch.join(', ') + ' (rootFolder=「' + rootFolder.getName() + '」)');

  for (const cat of categoriesToSearch) {
    try {
      const catFolder = _getOrCreateFolder(rootFolder, cat);
      // 搜尋該分類下所有案號資料夾
      const caseFolders = catFolder.getFolders();
      let folderCount = 0;
      while (caseFolders.hasNext()) {
        _scanAndRename(caseFolders.next());
        folderCount++;
      }
      // 也搜分類資料夾本身（以防檔案直接放在分類層）
      _scanAndRename(catFolder);
      Logger.log('  📁 分類「' + cat + '」: 共 ' + folderCount + ' 個案號資料夾');
    } catch (e) {
      Logger.log('  ⚠️ 搜尋 ' + cat + ' 資料夾失敗: ' + e.message);
    }
  }

  if (count === 0) {
    Logger.log('  ⚠️ 未找到符合的 Drive 檔案');
  }

  return count;
}


// ===================== 查看學習紀錄 =====================

function showLearningLog() {
  const corrections = _getRecentCorrections(50);
  const ui = SpreadsheetApp.getUi();

  if (corrections.length === 0) {
    ui.alert('📝 學習紀錄', '目前沒有任何修正紀錄。\n\n' +
      '修正紀錄的來源：\n' +
      '1. 在 Gmail 移除「自動辨識來源」標籤 → 執行回授偵測（學習 Sender 角色）\n' +
      '2. 在「處理紀錄」Sheet 填寫「修正後名稱」欄 → 執行回授偵測（改名 Drive 檔案 + LLM 學習）\n' +
      '3. 在「修正原因」欄填寫原因，讓 LLM 理解為什麼修改', ui.ButtonSet.OK);
    return;
  }

  let text = '目前有 ' + corrections.length + ' 筆修正紀錄（最新在前）：\n\n';
  corrections.slice(0, 15).forEach((c, i) => {
    text += (i + 1) + '. 「' + c.subject + '」\n';
    if (c.aiCode !== c.finalCode) text += '   碼: ' + c.aiCode + ' → ' + c.finalCode + '\n';
    if (c.aiName !== c.finalName) text += '   名: 「' + c.aiName + '」→「' + c.finalName + '」\n';
    if (c.reason) text += '   原因: ' + c.reason + '\n';
    text += '\n';
  });

  if (corrections.length > 15) text += '...還有 ' + (corrections.length - 15) + ' 筆\n';
  text += '\n這些紀錄會在每次 LLM 分類時自動注入 prompt。';

  ui.alert('📝 學習紀錄', text, ui.ButtonSet.OK);
}


// ===================== LLM Prompt 文件管理 =====================

/**
 * 取得或建立 LLM Prompt Doc（在專案資料夾內，與 Sheet 同層級）
 */
function _getOrCreatePromptDoc() {
  const projectFolder = _getProjectFolder();
  const docName = CONFIG.PROMPT_DOC_NAME;

  // 在專案資料夾內搜尋同名 Doc
  const files = projectFolder.getFilesByName(docName);
  if (files.hasNext()) {
    const file = files.next();
    return DocumentApp.openById(file.getId());
  }

  // 不存在 → 建立新的，然後移到專案資料夾
  const doc = DocumentApp.create(docName);
  const file = DriveApp.getFileById(doc.getId());
  projectFolder.addFile(file);
  // 從根目錄移除（create 預設放在根目錄）
  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== projectFolder.getId()) {
      parent.removeFile(file);
    }
  }

  return doc;
}

/**
 * 匯出 LLM Prompt 到 Google Doc（選單功能）
 * Doc 結構：
 *   1. 完整 SYSTEM_PROMPT（含模板 + 修正紀錄注入後的版本）
 *   2. 待整理學習紀錄（累積區）
 */
function exportPromptDoc() {
  const corrections = _getRecentCorrections(20);
  const templates = _loadTemplates();
  const prompt = _buildPrompt(corrections, templates);

  const doc = _getOrCreatePromptDoc();
  const body = doc.getBody();

  // 清除舊內容
  body.clear();

  // ── 標題 ──
  body.appendParagraph('LLM Prompt 文件')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('最後更新: ' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }))
    .setItalic(true);
  body.appendParagraph('');

  // ── 注入摘要 ──
  body.appendParagraph('注入摘要')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('模板數量: ' + templates.length + ' 條');
  body.appendParagraph('修正紀錄: ' + corrections.length + ' 筆（最近 20 筆 few-shot）');
  body.appendParagraph('Prompt 總長: 約 ' + prompt.length + ' 字元');
  body.appendParagraph('');

  // ── 完整 Prompt ──
  body.appendParagraph('完整 SYSTEM_PROMPT')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(prompt)
    .setFontFamily('Noto Sans Mono')
    .setFontSize(9);
  body.appendParagraph('');

  // ── 待整理學習紀錄（保留區：不會被清除，只會在合併後清空）──
  body.appendParagraph('').setAttributes({});  // 分隔線
  body.appendHorizontalRule();
  body.appendParagraph('待整理學習紀錄')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('以下紀錄會在「整理學習紀錄」時由 LLM 自動摘要，歸納成規則後合併進 SYSTEM_PROMPT。')
    .setItalic(true);
  body.appendParagraph('');

  // 寫入目前所有修正紀錄（含原因）
  if (corrections.length > 0) {
    corrections.forEach((c, i) => {
      let line = (i + 1) + '. 「' + c.subject + '」';
      if (c.aiCode !== c.finalCode) line += ' [碼: ' + c.aiCode + '→' + c.finalCode + ']';
      if (c.aiName !== c.finalName) line += ' [名: 「' + c.aiName + '」→「' + c.finalName + '」]';
      if (c.reason) line += ' [原因: ' + c.reason + ']';
      body.appendParagraph(line).setFontSize(9);
    });
  } else {
    body.appendParagraph('（目前沒有待整理的修正紀錄）');
  }

  doc.saveAndClose();

  const url = doc.getUrl();
  const ui = SpreadsheetApp.getUi();
  ui.alert('📋 LLM Prompt 文件',
    '已匯出到專案資料夾：\n\n' + url + '\n\n' +
    '── 注入摘要 ──\n' +
    '模板: ' + templates.length + ' 條\n' +
    '修正紀錄: ' + corrections.length + ' 筆\n' +
    'Prompt: 約 ' + prompt.length + ' 字元',
    ui.ButtonSet.OK);
}


/**
 * 整理學習紀錄：用 LLM 把累積的修正紀錄歸納成規則，合併進 SYSTEM_PROMPT
 *
 * 流程：
 *   1. 讀取所有修正紀錄
 *   2. 呼叫 LLM 分析修正紀錄 → 產生歸納規則建議
 *   3. 把歸納結果寫入 Prompt Doc 的「已合併規則」區
 *   4. 清空 Sheet 裡已處理的修正紀錄（標記為已合併）
 *   5. 更新 Prompt Doc
 */
function consolidateLearning() {
  const corrections = _getRecentCorrections(100);  // 讀取所有修正紀錄

  if (corrections.length < 3) {
    try {
      SpreadsheetApp.getUi().alert('📝 學習紀錄整理',
        '目前只有 ' + corrections.length + ' 筆修正紀錄，建議累積至少 3 筆再整理。',
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) { Logger.log('修正紀錄不足，跳過整理'); }
    return;
  }

  // 組裝修正紀錄文字
  const correctionText = corrections.map((c, i) => {
    let line = (i + 1) + '. 信件主旨:「' + c.subject + '」';
    if (c.aiCode !== c.finalCode) line += '\n   收發碼修正: ' + c.aiCode + ' → ' + c.finalCode;
    if (c.aiName !== c.finalName) line += '\n   語義名修正: 「' + c.aiName + '」→「' + c.finalName + '」';
    if (c.reason) line += '\n   原因: ' + c.reason;
    return line;
  }).join('\n\n');

  // 呼叫 LLM 做歸納
  const apiKey = getApiKey();
  const endpoint = CONFIG.GEMINI_ENDPOINT + CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const consolidationPrompt = `你是 IP Winner 智財事務所的 email 分類系統開發者。
以下是使用者對 LLM 分類結果的修正紀錄。請分析這些修正，歸納出通用規則。

## 修正紀錄
${correctionText}

## 任務
1. 分析上述修正紀錄，找出重複出現的模式
2. 歸納出 LLM 在分類/命名時應該遵守的規則（每條規則用一行描述）
3. 只產出「新規則」，不要重複已知的基本規則
4. 每條規則要具體、可操作（例如：「TA 委託信中 by [日期] 是我方期限，必須加入語義名」）

## 輸出格式（只回 JSON）
{
  "rules": [
    {
      "category": "期限選擇 | 語義名命名 | 收發碼判定 | 歸檔 | 其他",
      "rule": "具體規則描述",
      "examples": "來源修正紀錄的簡短引用"
    }
  ],
  "summary": "整體修正趨勢的一句話摘要"
}`;

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: consolidationPrompt }] }],
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 2048,
        responseMimeType: 'application/json',
      },
    }),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('❌ 學習整理 LLM 呼叫失敗: ' + response.getResponseCode());
    try {
      SpreadsheetApp.getUi().alert('❌ LLM 呼叫失敗', '錯誤碼: ' + response.getResponseCode(), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
    return;
  }

  let result;
  try {
    const respJson = JSON.parse(response.getContentText());
    const text = respJson.candidates[0].content.parts[0].text;
    result = JSON.parse(text);
  } catch (e) {
    Logger.log('❌ 學習整理解析失敗: ' + e.message);
    try {
      SpreadsheetApp.getUi().alert('❌ 解析失敗', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e2) {}
    return;
  }

  // 寫入 Prompt Doc
  const doc = _getOrCreatePromptDoc();
  const body = doc.getBody();

  // 找到「待整理學習紀錄」標題，在它前面插入「已合併規則」
  const searchResult = body.findText('待整理學習紀錄');
  if (searchResult) {
    const element = searchResult.getElement().getParent();
    const index = body.getChildIndex(element);

    // 在「待整理學習紀錄」前插入合併結果
    const dateStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    body.insertParagraph(index, '').setAttributes({});
    body.insertParagraph(index + 1, '已合併規則 (' + dateStr + ')')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(index + 2, '摘要: ' + (result.summary || ''))
      .setItalic(true);

    let insertIdx = index + 3;
    if (result.rules && result.rules.length > 0) {
      result.rules.forEach(r => {
        const ruleText = '• [' + r.category + '] ' + r.rule + (r.examples ? '（例: ' + r.examples + '）' : '');
        body.insertParagraph(insertIdx, ruleText).setFontSize(10);
        insertIdx++;
      });
    }
  }

  // 清空「待整理學習紀錄」底下的舊紀錄，標記為已整理
  const pendingSearch = body.findText('待整理學習紀錄');
  if (pendingSearch) {
    const pendingElement = pendingSearch.getElement().getParent();
    const pendingIndex = body.getChildIndex(pendingElement);
    // 刪除標題之後的所有內容（除了標題本身和說明文字）
    const totalChildren = body.getNumChildren();
    // 保留: 標題(pendingIndex) + 說明(pendingIndex+1) + 空行(pendingIndex+2)
    const keepUntil = pendingIndex + 2;
    for (let i = totalChildren - 1; i > keepUntil; i--) {
      body.removeChild(body.getChild(i));
    }
    body.appendParagraph('（上次整理: ' + new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }) + '，共處理 ' + corrections.length + ' 筆修正紀錄）')
      .setItalic(true).setFontSize(9);
  }

  doc.saveAndClose();

  // Log 結果
  const ruleCount = result.rules ? result.rules.length : 0;
  Logger.log('✅ 學習整理完成: 從 ' + corrections.length + ' 筆修正紀錄歸納出 ' + ruleCount + ' 條規則');
  result.rules.forEach(r => Logger.log('  📌 [' + r.category + '] ' + r.rule));

  try {
    const url = doc.getUrl();
    SpreadsheetApp.getUi().alert('✅ 學習紀錄整理完成',
      '從 ' + corrections.length + ' 筆修正紀錄歸納出 ' + ruleCount + ' 條新規則。\n\n' +
      '摘要: ' + (result.summary || '') + '\n\n' +
      '已寫入 Prompt 文件: ' + url + '\n\n' +
      '⚠️ 注意: 歸納出的規則目前僅記錄在 Doc 中。\n' +
      '如需正式生效，請將規則手動加入 Code.gs 的 SYSTEM_PROMPT。',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

/**
 * 安裝每週學習整理排程（每週一早上 9 點）
 */
function installConsolidationTrigger() {
  // 先移除舊的
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'weeklyConsolidate')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('weeklyConsolidate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('✅ 已安裝每週學習整理排程（每週一 9:00）');
}

/**
 * 每週自動執行：匯出 Prompt Doc + 整理學習紀錄
 */
function weeklyConsolidate() {
  Logger.log('=== 每週學習整理（自動） ===');

  const corrections = _getRecentCorrections(100);
  if (corrections.length < 3) {
    Logger.log('修正紀錄不足 3 筆，跳過本週整理');
    return;
  }

  // 先匯出最新 Prompt Doc
  exportPromptDoc();
  // 再做整理合併
  consolidateLearning();
}


// ===================== 統計 =====================

function showStats() {
  const ss = _getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
  if (!sheet || sheet.getLastRow() <= 1) {
    try { SpreadsheetApp.getUi().alert('尚無處理紀錄'); } catch (e) { Logger.log('尚無處理紀錄'); }
    return;
  }

  const data = sheet.getDataRange().getValues();
  const codeCounts = {};
  let total = 0, pending = 0, withCustomerName = 0;

  for (let i = 1; i < data.length; i++) {
    total++;
    const code = String(data[i][4]).trim();
    codeCounts[code] = (codeCounts[code] || 0) + 1;
    if (String(data[i][11]).trim() === 'pending') pending++;    // col 11: 來源確認狀態
    if (String(data[i][13] || '').trim()) withCustomerName++;  // col 13: 修正後名稱
  }

  const codeStats = Object.entries(codeCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([c, n]) => c + ': ' + n + ' 封')
    .join('\n');

  const msg = '📊 處理統計\n\n' +
    '總計: ' + total + ' 封\n' +
    '待確認來源: ' + pending + ' 封\n' +
    '客戶已改名: ' + withCustomerName + ' 筆\n\n' +
    '收發碼分布:\n' + codeStats;

  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('📊 處理統計', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* 從編輯器呼叫時沒有 UI，已有 Logger */ }
}


// ===================== 排程 =====================

/** 安裝每日排程：早上 7-8 點執行 */
function installTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Asia/Taipei')
    .create();

  Logger.log('✅ 已安裝每日排程（每天早上 7:00-8:00 台北時間）');
  try {
    SpreadsheetApp.getUi().alert('✅ 已安裝每日排程\n\n每天早上 7:00-8:00（台北時間）自動執行');
  } catch (e) { /* 從編輯器呼叫 */ }
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
