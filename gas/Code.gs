// =============================================
// LINE 備份工具 - Google Apps Script 後端
// 資料夾結構（可在管理介面切換）：
//   A：LINE備份 / 聯絡人名稱 / YYYY-MM / 對話記錄.txt
//   B：LINE備份 / YYYY-MM / 聯絡人名稱.txt
// 自動流程：LINE待處理/ → 每小時自動掃描處理
// =============================================

const ROOT_FOLDER_NAME  = 'LINE備份';
const INBOX_FOLDER_NAME = 'LINE待處理'; // 使用者上傳新檔案到這裡

// ── 設定：讀取 / 儲存 ───────────────────────
// folderStructure: 'A' = 聯絡人/月份/對話記錄.txt
//                  'B' = 月份/聯絡人.txt（舊版）
function getSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    folderStructure: props.getProperty('folderStructure') || 'A'
  };
}

function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (settings.folderStructure) {
      props.setProperty('folderStructure', settings.folderStructure);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── 入口：提供 Web App 管理介面 ──────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LINE 備份管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 核心處理邏輯 ─────────────────────────────
function processLineExport(content) {
  try {
    const parsed = parseLineExport(content);

    if (!parsed.contactName) {
      return {
        success: false,
        error: '無法識別聯絡人名稱。\n請確認上傳的是 LINE「傳送聊天記錄」匯出的 .txt 檔案。'
      };
    }

    if (parsed.messages.length === 0) {
      return {
        success: false,
        error: '未找到任何訊息內容，請確認檔案格式是否正確。'
      };
    }

    // 取得或建立根資料夾 LINE備份
    const rootFolder = getOrCreateFolder(ROOT_FOLDER_NAME, DriveApp.getRootFolder());
    const results = [];
    const monthGroups = groupByMonth(parsed.messages);
    const structure = getSettings().folderStructure; // 'A' or 'B'

    for (const yearMonth of Object.keys(monthGroups).sort()) {
      const messages = monthGroups[yearMonth];
      let targetFolder, fileName;

      if (structure === 'A') {
        // 結構 A：LINE備份 / 聯絡人名稱 / 2024-01 / 對話記錄.txt
        const contactFolder = getOrCreateFolder(sanitizeFileName(parsed.contactName), rootFolder);
        targetFolder = getOrCreateFolder(yearMonth, contactFolder);
        fileName = '對話記錄.txt';
      } else {
        // 結構 B：LINE備份 / 2024-01 / 聯絡人名稱.txt
        targetFolder = getOrCreateFolder(yearMonth, rootFolder);
        fileName = sanitizeFileName(parsed.contactName) + '.txt';
      }

      const existingFiles = targetFolder.getFilesByName(fileName);

      if (existingFiles.hasNext()) {
        const existingFile = existingFiles.next();
        const existingContent = existingFile.getBlob().getDataAsString('UTF-8');

        // 取得舊檔案中所有已存在的訊息特徵（指紋）
        const existingSigs = getExistingSignatures(existingContent);
        
        // 過濾出「未曾備份過」的新訊息
        const newMessages = messages.filter(msg => {
          const sig = `${msg.date}|${msg.time}|${msg.sender}|${msg.message}`;
          return !existingSigs.has(sig);
        });

        if (newMessages.length === 0) {
          results.push({ month: yearMonth, count: 0, action: 'skipped', url: existingFile.getUrl() });
        } else {
          const newContent = formatMessages(parsed.contactName, yearMonth, newMessages);
          existingFile.setContent(existingContent + '\n\n' + '─'.repeat(50) + '\n\n' + newContent);
          results.push({ month: yearMonth, count: newMessages.length, action: 'updated', url: existingFile.getUrl() });
        }
      } else {
        const newContent = formatMessages(parsed.contactName, yearMonth, messages);
        const newFile = targetFolder.createFile(fileName, newContent, MimeType.PLAIN_TEXT);
        results.push({ month: yearMonth, count: messages.length, action: 'created', url: newFile.getUrl() });
      }
    }

    return {
      success: true,
      contactName: parsed.contactName,
      totalMessages: parsed.messages.length,
      monthCount: Object.keys(monthGroups).length,
      results: results,
      rootUrl: rootFolder.getUrl()
    };

  } catch (err) {
    return { success: false, error: '處理失敗：' + err.toString() };
  }
}

// ── LINE 匯出格式解析器 ──────────────────────
// 支援格式（Android 正體中文）：
//   [LINE] 和 王小明 的聊天記錄
//   儲存時間：YYYY/MM/DD HH:MM
//
//   YYYY/MM/DD(週X)
//   HH:MM\t發送者名稱
//   訊息內容
function parseLineExport(content) {
  const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');

  let contactName = '';
  const messages = [];
  let currentDate = null;
  let currentTime = null;
  let currentSender = null;
  let currentMessageLines = [];

  // ── 解析聯絡人名稱（第一行）──
  const headerLine = lines[0] || '';
  const namePatterns = [
    /[和與]\s*(.+?)\s*的聊天記錄/,           // 正體中文（和 / 與）
    /与\s*(.+?)\s*的聊天记录/,               // 簡體中文
    /Chat with\s+(.+)/i,                    // 英文
    /(.+?)とのトーク/,                       // 日文
    /\[LINE\]\s+(.+)/,                      // Fallback
  ];

  for (const pattern of namePatterns) {
    const match = headerLine.match(pattern);
    if (match) {
      contactName = match[1].replace(/^\[LINE\]\s*/, '').trim();
      break;
    }
  }

  // ── 解析訊息 ──
  // 日期行：2024/01/10(週三) 或 2024/01/10（週三）
  const datePattern = /^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/;
  // 時間 + 發送者行（Tab 分隔）：14:30\t王小明\t訊息內容
  // LINE Android 格式為三欄：時間\t發送者\t訊息（同一行）
  const timeSenderPattern = /^(?:上午|下午|午前|午後)?(\d{1,2}:\d{2})\t([^\t]+)(?:\t(.*))?$/;

  const saveMessage = () => {
    if (currentDate && currentTime && currentSender !== null) {
      messages.push({
        date: currentDate,
        time: currentTime,
        sender: currentSender,
        message: currentMessageLines.join('\n').trim()
      });
    }
    currentMessageLines = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // 跳過標頭區域（前 3 行：標題、儲存日期、空行），從第 4 行開始解析
    if (i < 3) continue;

    // 日期行
    const dateMatch = trimmed.match(datePattern);
    if (dateMatch) {
      saveMessage();
      currentTime = null;
      currentSender = null;
      currentDate = `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
      continue;
    }

    // 時間 + 發送者行
    const timeSenderMatch = line.match(timeSenderPattern);
    if (timeSenderMatch && currentDate) {
      saveMessage();
      currentTime = timeSenderMatch[1];
      currentSender = timeSenderMatch[2].trim();
      // 同行的訊息內容（第三欄）直接放入 currentMessageLines
      currentMessageLines = timeSenderMatch[3] !== undefined ? [timeSenderMatch[3]] : [];
      continue;
    }

    // 訊息內容（多行）
    if (currentSender !== null && currentDate) {
      currentMessageLines.push(line);
    }
  }

  saveMessage(); // 儲存最後一則訊息

  return { contactName, messages };
}

// ── 工具函式 ─────────────────────────────────

function groupByMonth(messages) {
  const groups = {};
  for (const msg of messages) {
    if (!msg.date) continue;
    const yearMonth = msg.date.substring(0, 7); // "2024-01"
    if (!groups[yearMonth]) groups[yearMonth] = [];
    groups[yearMonth].push(msg);
  }
  return groups;
}

function formatMessages(contactName, yearMonth, messages) {
  const now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  const lines = [
    `LINE 對話備份`,
    `═`.repeat(50),
    `聯絡人：${contactName}`,
    `月　份：${yearMonth}`,
    `訊息數：${messages.length} 則`,
    `備份時間：${now}`,
    `═`.repeat(50),
    '',
  ];

  let lastDate = '';
  for (const msg of messages) {
    if (msg.date !== lastDate) {
      if (lastDate !== '') lines.push('');
      lines.push(`【 ${msg.date} 】`);
      lines.push('');
      lastDate = msg.date;
    }
    lines.push(`${msg.time}  ${msg.sender}`);
    if (msg.message) {
      msg.message.split('\n').forEach(l => lines.push(`  ${l}`));
    }
    lines.push('');
  }

  return lines.join('\n');
}

function getOrCreateFolder(name, parentFolder) {
  const existing = parentFolder.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(name);
}

function sanitizeFileName(name) {
  return (name || '未知聯絡人').replace(/[\/\\:*?"<>|]/g, '_').trim();
}

// ── 列出所有備份 ─────────────────────────────
function listBackups() {
  try {
    const rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) {
      return { success: true, backups: [], totalSize: 0, rootExists: false };
    }

    const rootFolder = rootFolders.next();
    const backups = [];
    const errorFiles = [];
    let totalSize = 0;
    const structure = getSettings().folderStructure;

    // 取得待處理資料夾中的錯誤檔案
    const inboxFolders = DriveApp.getRootFolder().getFoldersByName(INBOX_FOLDER_NAME);
    if (inboxFolders.hasNext()) {
      const inboxFolder = inboxFolders.next();
      const files = inboxFolder.getFilesByType(MimeType.PLAIN_TEXT);
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName().startsWith('ERROR_')) {
          errorFiles.push({
            id: file.getId(),
            name: file.getName(),
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated().toISOString()
          });
        }
      }
    }

    const firstLevel = rootFolder.getFolders();
    while (firstLevel.hasNext()) {
      const firstFolder = firstLevel.next();
      const firstName = firstFolder.getName();

      if (structure === 'A') {
        // 結構 A：聯絡人 / 月份 / 對話記錄.txt → firstName = 聯絡人名稱
        const monthFolders = firstFolder.getFolders();
        while (monthFolders.hasNext()) {
          const monthFolder = monthFolders.next();
          const monthName = monthFolder.getName();
          const files = monthFolder.getFiles();
          while (files.hasNext()) {
            const file = files.next();
            const size = file.getSize();
            totalSize += size;
            backups.push({
              id: file.getId(),
              contactName: firstName,
              month: monthName,
              size,
              url: file.getUrl(),
              lastUpdated: file.getLastUpdated().toISOString()
            });
          }
        }
      } else {
        // 結構 B：月份 / 聯絡人.txt → firstName = 月份
        const files = firstFolder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          const size = file.getSize();
          totalSize += size;
          backups.push({
            id: file.getId(),
            contactName: file.getName().replace(/\.txt$/, ''),
            month: firstName,
            size,
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated().toISOString()
          });
        }
      }
    }

    // 依月份降冪排序（最新在上），同月份依聯絡人排序
    backups.sort((a, b) =>
      b.month.localeCompare(a.month) || a.contactName.localeCompare(b.contactName)
    );

    return {
      success: true,
      backups,
      errorFiles,
      totalSize,
      rootExists: true,
      rootUrl: rootFolder.getUrl(),
      folderStructure: structure
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 取得備份統計數據 (快速掃描) - 改由背景快取提供
 */
function getBackupStats() {
  try {
    const props = PropertiesService.getScriptProperties();
    const cached = props.getProperty('backupStats');
    if (cached) {
      return { success: true, ...JSON.parse(cached) };
    }
    // 如果沒有快取，即時算一次（只會發生在第一次使用時）
    updateBackupStatsCache();
    const newCached = props.getProperty('backupStats');
    if (newCached) {
       return { success: true, ...JSON.parse(newCached) };
    }
    return { success: true, fileCount: 0, monthCount: 0, totalSize: 0, errorCount: 0 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 實際更新快取的邏輯（由觸發器或修改動作呼叫）
 */
function updateBackupStatsCache() {
  try {
    const settings = getSettings();
    const rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) {
      PropertiesService.getScriptProperties().setProperty('backupStats', JSON.stringify({ fileCount: 0, monthCount: 0, totalSize: 0, errorCount: 0 }));
      return;
    }
    const root = rootFolders.next();

    let fileCount = 0;
    let totalSize = 0;
    const months = new Set();
    let errorCount = 0;

    const inboxFolders = DriveApp.getRootFolder().getFoldersByName(INBOX_FOLDER_NAME);
    if (inboxFolders.hasNext()) {
      const inboxFolder = inboxFolders.next();
      const pFiles = inboxFolder.getFilesByType(MimeType.PLAIN_TEXT);
      while (pFiles.hasNext()) {
        if (pFiles.next().getName().startsWith('ERROR_')) errorCount++;
      }
    }

    const firstLevel = root.getFolders();
    while (firstLevel.hasNext()) {
      const firstFolder = firstLevel.next();
      
      if (settings.folderStructure === 'A') {
        const monthFolders = firstFolder.getFolders();
        while (monthFolders.hasNext()) {
          const monthFolder = monthFolders.next();
          months.add(monthFolder.getName());
          const files = monthFolder.getFiles();
          while (files.hasNext()) {
            const f = files.next();
            fileCount++;
            totalSize += f.getSize();
          }
        }
      } else {
        months.add(firstFolder.getName());
        const files = firstFolder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          fileCount++;
          totalSize += f.getSize();
        }
      }
    }

    PropertiesService.getScriptProperties().setProperty('backupStats', JSON.stringify({
      fileCount, monthCount: months.size, totalSize, errorCount
    }));
  } catch (e) {
    Logger.log('Cache update failed: ' + e.message);
  }
}

// ── 刪除單一備份檔案（移至垃圾桶）──────────────
// 若月份資料夾因此變為空，也一併移至垃圾桶
function deleteBackupFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    const parentFolder = parents.hasNext() ? parents.next() : null;

    file.setTrashed(true);

    if (parentFolder) {
      const hasFiles   = parentFolder.getFiles().hasNext();
      const hasFolders = parentFolder.getFolders().hasNext();
      if (!hasFiles && !hasFolders) parentFolder.setTrashed(true);
    }

    updateBackupStatsCache();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── 刪除全部備份（整個 LINE備份 資料夾移至垃圾桶）
function deleteAllBackups() {
  try {
    const folders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (folders.hasNext()) folders.next().setTrashed(true);
    
    updateBackupStatsCache();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// =============================================
// 自動監控流程（優化版）
// =============================================

// ── 觸發函式：每小時掃描「LINE待處理」並自動處理 ──
// 由時間觸發器自動呼叫，不需手動執行
function watchAndProcess() {
  const rootFolders = DriveApp.getRootFolder().getFoldersByName(INBOX_FOLDER_NAME);
  if (!rootFolders.hasNext()) return; // 資料夾不存在，略過

  const inboxFolder = rootFolders.next();
  const files = inboxFolder.getFilesByType(MimeType.PLAIN_TEXT);
  const log = [];

  while (files.hasNext()) {
    const file = files.next();

    // 跳過之前失敗的檔案（已加 ERROR_ 前綴）
    if (file.getName().startsWith('ERROR_')) continue;

    try {
      const content = file.getBlob().getDataAsString('UTF-8');
      const result  = processLineExport(content);

      if (result.success) {
        file.setTrashed(true); // 處理成功 → 從待處理資料夾移除
        log.push('✅ ' + result.contactName + '（' + result.totalMessages + ' 則）');
      } else {
        file.setName('ERROR_' + file.getName()); // 標記失敗，保留供人工檢查
        log.push('❌ ' + file.getName() + '：' + result.error);
      }
    } catch (e) {
      file.setName('ERROR_' + file.getName());
      log.push('❌ ' + file.getName() + '：' + e.toString());
    }
  }

  if (log.length > 0) Logger.log('LINE備份自動處理結果：\n' + log.join('\n'));
  
  // 更新統計快取
  updateBackupStatsCache();
}

// ── 安裝觸發器（工程師執行一次即可）──────────────
// 執行後：建立「LINE待處理」資料夾 + 設定每小時自動觸發
function installTrigger() {
  // 先移除同名的舊觸發器（避免重複）
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'watchAndProcess')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 建立每小時觸發器
  ScriptApp.newTrigger('watchAndProcess')
    .timeBased()
    .everyHours(1)
    .create();

  // 建立「LINE待處理」資料夾（若不存在）
  const root = DriveApp.getRootFolder();
  let inboxFolder;
  const existing = root.getFoldersByName(INBOX_FOLDER_NAME);
  if (existing.hasNext()) {
    inboxFolder = existing.next();
  } else {
    inboxFolder = root.createFolder(INBOX_FOLDER_NAME);
  }

  const msg = '✅ 安裝完成！\n' +
    '「LINE待處理」資料夾：' + inboxFolder.getUrl() + '\n' +
    '每小時自動掃描一次，新上傳的 .txt 檔案會自動整理到「LINE備份」。';
  Logger.log(msg);
  return msg;
}

// ── 移除觸發器 ──────────────────────────────────
function removeTrigger() {
  const removed = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'watchAndProcess');
  removed.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('已移除 ' + removed.length + ' 個自動觸發器');
}

// ── 新增工具函式：解析舊備份檔的訊息特徵 ──────────────────
function getExistingSignatures(content) {
  const lines = content.split('\n');
  const sigs = new Set();
  let currentDate = '';
  let currentTime = '';
  let currentSender = '';
  let currentMessage = [];

  const saveSig = () => {
    if (currentDate && currentTime && currentSender) {
      sigs.add(`${currentDate}|${currentTime}|${currentSender}|${currentMessage.join('\n').trim()}`);
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // 判斷是否為日期標頭，如：【 2024-01-10 】
    const dateMatch = line.match(/^【 (\d{4}-\d{2}-\d{2}) 】$/);
    if (dateMatch) {
      saveSig();
      currentDate = dateMatch[1];
      currentTime = '';
      currentSender = '';
      currentMessage = [];
      continue;
    }
    
    // 判斷是否為訊息開頭，如：14:30  王小明
    const msgMatch = line.match(/^(\d{1,2}:\d{2})  (.*)$/);
    if (msgMatch) {
      saveSig();
      currentTime = msgMatch[1];
      currentSender = msgMatch[2];
      currentMessage = [];
      continue;
    }

    // 收集訊息內容（縮排兩格的行）
    if (currentTime) {
      if (line.startsWith('  ')) {
        currentMessage.push(line.substring(2));
      } else if (line === '') {
        // 空行忽略
      } else {
        // 遇到其他分隔線或檔頭，代表這則訊息結束
        saveSig();
        currentTime = '';
        currentSender = '';
      }
    }
  }
  saveSig();
  return sigs;
}
