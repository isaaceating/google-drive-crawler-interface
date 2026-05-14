/**
 * Google Drive Crawler interface
 * Version: v2.2.2
 *
 * Fix:
 * 1. 修正 ?playbook=product / ?playbook=presales 仍顯示 Solution 的問題
 * 2. doGet(e) 改用 HtmlService.createTemplateFromFile()
 * 3. 由後端直接注入 INITIAL_PLAYBOOK / INITIAL_LANG 到 index.html
 *
 * Features:
 * 1. Multi-Playbook Snapshot：支援 ?playbook=product / solution / presales
 * 2. Google Sheet Snapshot：每個 playbook 對應一個 Snapshot tab
 * 3. Full Snapshot Preload：開頁時一次載入目前 playbook 的完整 Snapshot
 * 4. Frontend Parent Map：點資料夾不再呼叫後端
 * 5. Manual Refresh：只 refresh 目前 playbook 的 Snapshot tab
 * 6. Daily Auto Refresh：可每日 refresh 所有 playbook Snapshot tabs
 * 7. Title Sync：標題依 playbook 自動切換
 * 8. Language：?lang=en / ?lang=ch，預設英文
 */

const SNAPSHOT_SPREADSHEET_ID = '1JjQ2NAWGLUbIZsc9XjaONGJYyermwE4tigNRTOoHfrY';

const DEFAULT_PLAYBOOK = 'solution';

const PLAYBOOK_CONFIG = {
  product: {
    key: 'product',
    titleEn: 'Product Playbook',
    titleCh: 'Product 彈藥庫',
    rootFolderId: '1zAFat5y1UL-vMqg5yQVy0SAgRD7WG0uY',
    sheetName: 'Snapshot_Product'
  },
  solution: {
    key: 'solution',
    titleEn: 'Solution Playbook',
    titleCh: 'Solution 彈藥庫',
    rootFolderId: '1XtR0qP5DJIQ6jIL6Vh8hhAhJa-wdLVDL',
    sheetName: 'Snapshot_Solution'
  },
  presales: {
    key: 'presales',
    titleEn: 'Pre-sales Playbook',
    titleCh: 'Pre-sales 彈藥庫',
    rootFolderId: '1ltBQJqMoey5jHkRaqjWvRJgL18rXDqLT',
    sheetName: 'Snapshot_PreSales'
  }
};

const SNAPSHOT_HEADERS = [
  'id',
  'parentId',
  'name',
  'itemType',
  'fileType',
  'previewLink',
  'openLink',
  'path',
  'level',
  'sortName',
  'updatedAt'
];

function doGet(e) {
  const lang = e && e.parameter && e.parameter.lang === 'ch' ? 'ch' : 'en';
  const playbookKey = normalizePlaybookKey_(e && e.parameter && e.parameter.playbook);
  const config = getPlaybookConfig_(playbookKey);
  const title = lang === 'ch' ? config.titleCh : config.titleEn;

  const template = HtmlService.createTemplateFromFile('index');

  template.INITIAL_PLAYBOOK = config.key;
  template.INITIAL_LANG = lang;
  template.INITIAL_TITLE_EN = config.titleEn;
  template.INITIAL_TITLE_CH = config.titleCh;

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Frontend API：
 * 開頁時一次回傳目前 playbook 的完整 Snapshot payload
 */
function getFullSnapshotPayload(playbookKey) {
  const config = getPlaybookConfig_(playbookKey);
  const sheet = getSnapshotSheet_(config.sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      playbookKey: config.key,
      titleEn: config.titleEn,
      titleCh: config.titleCh,
      rootFolderId: config.rootFolderId,
      rootItems: [],
      items: [],
      updatedAt: '',
      rowCount: 0
    };
  }

  const headers = values[0];
  const rows = values.slice(1);
  const idx = buildHeaderIndex_(headers);

  const items = rows
    .filter(row => row[idx.id])
    .map(row => ({
      id: String(row[idx.id] || ''),
      parentId: String(row[idx.parentId] || ''),
      name: String(row[idx.name] || ''),
      itemType: String(row[idx.itemType] || ''),
      fileType: String(row[idx.fileType] || ''),
      previewLink: String(row[idx.previewLink] || ''),
      openLink: String(row[idx.openLink] || ''),
      path: String(row[idx.path] || ''),
      level: Number(row[idx.level] || 0),
      hasChildren: String(row[idx.itemType] || '') === 'folder'
    }));

  items.sort(sortItems_);

  let rootItems = items.filter(item => item.parentId === config.rootFolderId);

  // Fallback：如果 parentId 沒對上，就用 level 1 當第一層
  if (!rootItems.length) {
    rootItems = items.filter(item => item.level === 1);
  }

  rootItems.sort(sortItems_);

  const updatedAt = getLatestUpdatedAtString_(rows, idx);

  return {
    playbookKey: config.key,
    titleEn: config.titleEn,
    titleCh: config.titleCh,
    rootFolderId: config.rootFolderId,
    rootItems: rootItems,
    items: items,
    updatedAt: updatedAt,
    rowCount: items.length
  };
}

/**
 * Frontend API：
 * 手動刷新目前 playbook 的 Snapshot tab
 */
function refreshSnapshotManual(playbookKey) {
  const result = rebuildSnapshotByPlaybook(playbookKey);

  return {
    success: true,
    playbookKey: result.playbookKey,
    message: 'Database refreshed successfully.',
    rowCount: result.rowCount,
    updatedAt: result.updatedAt
  };
}

/**
 * 手動重建指定 playbook Snapshot
 * 測試用：
 * rebuildSnapshotByPlaybook('product')
 * rebuildSnapshotByPlaybook('solution')
 * rebuildSnapshotByPlaybook('presales')
 */
function rebuildSnapshotByPlaybook(playbookKey) {
  const config = getPlaybookConfig_(playbookKey);
  return rebuildSnapshot_(config);
}

/**
 * Daily Auto Refresh：
 * 每日刷新全部 playbook Snapshot tabs
 */
function dailyRefreshAllSnapshots() {
  return rebuildAllSnapshots();
}

/**
 * 手動執行：一次刷新全部 playbook Snapshot tabs
 */
function rebuildAllSnapshots() {
  const results = [];

  Object.keys(PLAYBOOK_CONFIG).forEach(key => {
    const config = PLAYBOOK_CONFIG[key];
    const result = rebuildSnapshot_(config);
    results.push(result);
  });

  return results;
}

/**
 * 手動執行一次這個 function，可建立每日自動刷新 trigger
 * 預設每天早上 6 點重建全部 Snapshot tabs
 */
function createDailySnapshotTrigger() {
  deleteSnapshotTriggers_();

  ScriptApp.newTrigger('dailyRefreshAllSnapshots')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  return 'Daily Snapshot trigger created successfully.';
}

/**
 * 手動刪除每日自動刷新 trigger
 */
function deleteDailySnapshotTrigger() {
  deleteSnapshotTriggers_();
  return 'Daily Snapshot trigger deleted successfully.';
}

/**
 * 相容 API：
 * 預設讀 solution
 */
function getRootItems() {
  const config = getPlaybookConfig_(DEFAULT_PLAYBOOK);
  return getSnapshotChildren_(config, config.rootFolderId);
}

function getFolderChildren(folderId) {
  const config = getPlaybookConfig_(DEFAULT_PLAYBOOK);
  return getSnapshotChildren_(config, folderId);
}

function getSearchIndex() {
  const payload = getFullSnapshotPayload(DEFAULT_PLAYBOOK);
  return payload.items
    .filter(item => item.itemType === 'file')
    .sort(sortItems_);
}

/**
 * 核心：依 playbook config 重建 Snapshot
 */
function rebuildSnapshot_(config) {
  const root = DriveApp.getFolderById(config.rootFolderId);
  const rows = [];
  const now = new Date();

  crawlFolderToSnapshot_(
    root,
    config.rootFolderId,
    root.getName(),
    1,
    rows,
    now
  );

  rows.sort((a, b) => {
    const pathA = String(a[7] || '');
    const pathB = String(b[7] || '');
    const parentCompare = pathA.localeCompare(pathB, 'zh-Hant');

    if (parentCompare !== 0) return parentCompare;

    const typeA = a[3];
    const typeB = b[3];

    if (typeA !== typeB) {
      return typeA === 'folder' ? -1 : 1;
    }

    return String(a[2] || '').localeCompare(String(b[2] || ''), 'zh-Hant');
  });

  const sheet = getSnapshotSheet_(config.sheetName);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SNAPSHOT_HEADERS.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, SNAPSHOT_HEADERS.length);

  return {
    success: true,
    playbookKey: config.key,
    sheetName: config.sheetName,
    rowCount: rows.length,
    updatedAt: Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    )
  };
}

/**
 * 從 Snapshot Sheet 讀取指定 parentId 的子項目
 * v2.2.2 前端不再主要使用，但保留相容與測試
 */
function getSnapshotChildren_(config, parentId) {
  const sheet = getSnapshotSheet_(config.sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const headers = values[0];
  const rows = values.slice(1);
  const idx = buildHeaderIndex_(headers);

  const items = rows
    .filter(row => String(row[idx.parentId] || '') === String(parentId || ''))
    .map(row => ({
      id: String(row[idx.id] || ''),
      name: String(row[idx.name] || ''),
      itemType: String(row[idx.itemType] || ''),
      fileType: String(row[idx.fileType] || ''),
      previewLink: String(row[idx.previewLink] || ''),
      openLink: String(row[idx.openLink] || ''),
      path: String(row[idx.path] || ''),
      hasChildren: String(row[idx.itemType] || '') === 'folder'
    }));

  items.sort(sortItems_);
  return items;
}

/**
 * 遞迴掃描 Drive folder，寫成 Snapshot rows
 */
function crawlFolderToSnapshot_(folder, parentId, path, level, rows, now) {
  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const sub = folders.next();

    rows.push([
      sub.getId(),
      parentId,
      sub.getName(),
      'folder',
      '',
      '',
      sub.getUrl(),
      path,
      level,
      sub.getName().toLowerCase(),
      now
    ]);

    crawlFolderToSnapshot_(
      sub,
      sub.getId(),
      path + ' / ' + sub.getName(),
      level + 1,
      rows,
      now
    );
  }

  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const type = mapMimeType_(file.getMimeType());

    rows.push([
      file.getId(),
      parentId,
      file.getName(),
      'file',
      type,
      buildPreviewLink_(file, type),
      file.getUrl(),
      path,
      level,
      file.getName().toLowerCase(),
      now
    ]);
  }
}

/**
 * 取得 Snapshot Sheet，沒有就自動建立
 */
function getSnapshotSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);
  }

  return sheet;
}

/**
 * 取得 playbook config
 */
function getPlaybookConfig_(playbookKey) {
  const key = normalizePlaybookKey_(playbookKey);
  return PLAYBOOK_CONFIG[key] || PLAYBOOK_CONFIG[DEFAULT_PLAYBOOK];
}

/**
 * 正規化 playbook key
 */
function normalizePlaybookKey_(playbookKey) {
  const key = String(playbookKey || DEFAULT_PLAYBOOK)
    .toLowerCase()
    .trim()
    .replace(/_/g, '-');

  if (key === 'pre-sales') return 'presales';
  if (key === 'pre_sales') return 'presales';
  if (key === 'pre') return 'presales';

  return PLAYBOOK_CONFIG[key] ? key : DEFAULT_PLAYBOOK;
}

/**
 * 建立 header index map
 */
function buildHeaderIndex_(headers) {
  const idx = {};

  headers.forEach((h, i) => {
    idx[String(h).trim()] = i;
  });

  return idx;
}

/**
 * 共用排序：folder 在前，file 在後，名稱依 zh-Hant 排序
 */
function sortItems_(a, b) {
  if (a.itemType !== b.itemType) {
    return a.itemType === 'folder' ? -1 : 1;
  }

  return String(a.name || '').localeCompare(String(b.name || ''), 'zh-Hant');
}

/**
 * 取得 Snapshot 最新更新時間，回傳字串
 */
function getLatestUpdatedAtString_(rows, idx) {
  if (idx.updatedAt === undefined) return '';

  let latest = null;

  rows.forEach(row => {
    const value = row[idx.updatedAt];

    if (value instanceof Date) {
      if (!latest || value > latest) latest = value;
    }
  });

  if (!latest) return '';

  return Utilities.formatDate(
    latest,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm'
  );
}

/**
 * 刪除 Daily Snapshot triggers
 */
function deleteSnapshotTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (fn === 'dailyRefreshAllSnapshots') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 檔案類型轉換
 */
function mapMimeType_(mime) {
  if (mime === MimeType.GOOGLE_DOCS) return 'doc';
  if (mime === MimeType.GOOGLE_SLIDES) return 'slides';
  if (mime === MimeType.GOOGLE_SHEETS) return 'sheet';
  if (mime === MimeType.PDF) return 'pdf';

  if (String(mime).startsWith('image/')) return 'image';
  if (String(mime).startsWith('video/')) return 'video';

  return 'file';
}

/**
 * 產生 preview link
 */
function buildPreviewLink_(file, type) {
  const id = file.getId();

  if (type === 'doc') return `https://docs.google.com/document/d/${id}/preview`;
  if (type === 'slides') return `https://docs.google.com/presentation/d/${id}/preview`;
  if (type === 'sheet') return `https://docs.google.com/spreadsheets/d/${id}/preview`;
  if (type === 'pdf') return `https://drive.google.com/file/d/${id}/preview`;
  if (type === 'image') return `https://drive.google.com/thumbnail?id=${id}&sz=w2000`;
  if (type === 'video') return `https://drive.google.com/file/d/${id}/preview`;

  return file.getUrl();
}

/**
 * 測試用：確認指定 playbook Snapshot payload 是否有資料
 */
function testSnapshotPayload() {
  const payload = getFullSnapshotPayload(DEFAULT_PLAYBOOK);

  Logger.log('Playbook: ' + payload.playbookKey);
  Logger.log('Root Folder ID: ' + payload.rootFolderId);
  Logger.log('Total Items: ' + payload.items.length);
  Logger.log('Root Items: ' + payload.rootItems.length);
  Logger.log('Updated At: ' + payload.updatedAt);

  if (payload.rootItems.length > 0) {
    Logger.log('First Root Item: ' + payload.rootItems[0].name);
  }
}

/**
 * 測試用：確認全部 playbook Snapshot payload
 */
function testAllSnapshotPayloads() {
  Object.keys(PLAYBOOK_CONFIG).forEach(key => {
    const payload = getFullSnapshotPayload(key);

    Logger.log('======================');
    Logger.log('Playbook: ' + payload.playbookKey);
    Logger.log('Root Folder ID: ' + payload.rootFolderId);
    Logger.log('Total Items: ' + payload.items.length);
    Logger.log('Root Items: ' + payload.rootItems.length);
    Logger.log('Updated At: ' + payload.updatedAt);

    if (payload.rootItems.length > 0) {
      Logger.log('First Root Item: ' + payload.rootItems[0].name);
    }
  });
}