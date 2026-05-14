/**
 * Google Drive Crawler interface
 * Version: v2.1.1
 *
 * Features:
 * 1. Google Sheet Snapshot：前台瀏覽與搜尋改讀 Snapshot Sheet，不再即時掃 Drive
 * 2. Full Snapshot Preload：開頁時一次載入完整 Snapshot，點資料夾不再呼叫後端
 * 3. Frontend Parent Map：前端建立 parentId map，資料夾切換接近秒開
 * 4. Root Items Fix：後端直接回傳 rootItems，避免第一層顯示空白
 * 5. Manual Refresh：前台按 Refresh Database 可手動重建 Snapshot
 * 6. Daily Auto Refresh：可建立每日自動重建 Snapshot trigger
 * 7. Language：?lang=en / ?lang=ch，預設英文
 */

const ROOT_FOLDER_ID = '1XtR0qP5DJIQ6jIL6Vh8hhAhJa-wdLVDL';

const SNAPSHOT_SPREADSHEET_ID = '1JjQ2NAWGLUbIZsc9XjaONGJYyermwE4tigNRTOoHfrY';
const SNAPSHOT_SHEET_NAME = 'Snapshot';

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
  const title = lang === 'ch' ? 'Solution 彈藥庫' : 'Solution Playbook';

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * v2.1.1 Frontend API：
 * 開頁時一次回傳完整 Snapshot payload
 * 直接包含 rootItems，避免前端第一層 mapping 失敗
 */
function getFullSnapshotPayload() {
  const sheet = getSnapshotSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      rootFolderId: ROOT_FOLDER_ID,
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

  let rootItems = items.filter(item => item.parentId === ROOT_FOLDER_ID);

  // Fallback：如果 parentId 沒對上，就用 level 1 當第一層
  if (!rootItems.length) {
    rootItems = items.filter(item => item.level === 1);
  }

  rootItems.sort(sortItems_);

  const updatedAt = getLatestUpdatedAtString_(rows, idx);

  return {
    rootFolderId: ROOT_FOLDER_ID,
    rootItems: rootItems,
    items: items,
    updatedAt: updatedAt,
    rowCount: items.length
  };
}

/**
 * v2.0 相容 API：
 * 保留給測試或舊版前端使用
 */
function getRootItems() {
  return getSnapshotChildren_(ROOT_FOLDER_ID);
}

function getFolderChildren(folderId) {
  return getSnapshotChildren_(folderId);
}

/**
 * v2.0 相容 API：
 * v2.1.1 前端已在開頁時載入 searchIndex，正常不會再呼叫這個 function
 */
function getSearchIndex() {
  const payload = getFullSnapshotPayload();

  return payload.items
    .filter(item => item.itemType === 'file')
    .sort(sortItems_);
}

/**
 * Frontend API：手動刷新 Snapshot
 */
function refreshSnapshotManual() {
  rebuildSnapshot();

  return {
    success: true,
    message: 'Database refreshed successfully.'
  };
}

/**
 * 手動執行一次這個 function，可建立每日自動刷新 trigger
 * 預設每天早上 6 點重建 Snapshot
 */
function createDailySnapshotTrigger() {
  deleteSnapshotTriggers_();

  ScriptApp.newTrigger('dailyRefreshSnapshot')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  return 'Daily Snapshot trigger created successfully.';
}

/**
 * 每日自動刷新 trigger 呼叫這個 function
 */
function dailyRefreshSnapshot() {
  rebuildSnapshot();
}

/**
 * 手動刪除每日自動刷新 trigger
 */
function deleteDailySnapshotTrigger() {
  deleteSnapshotTriggers_();
  return 'Daily Snapshot trigger deleted successfully.';
}

/**
 * 核心：重建 Snapshot
 */
function rebuildSnapshot() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const rows = [];
  const now = new Date();

  crawlFolderToSnapshot_(
    root,
    ROOT_FOLDER_ID,
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

  const sheet = getSnapshotSheet_();

  sheet.clearContents();
  sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SNAPSHOT_HEADERS.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, SNAPSHOT_HEADERS.length);

  return {
    success: true,
    rowCount: rows.length,
    updatedAt: now
  };
}

/**
 * 從 Snapshot Sheet 讀取指定 parentId 的子項目
 * v2.1.1 前端不再主要使用，但保留相容
 */
function getSnapshotChildren_(parentId) {
  const sheet = getSnapshotSheet_();
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
function getSnapshotSheet_() {
  const ss = SpreadsheetApp.openById(SNAPSHOT_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SNAPSHOT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, SNAPSHOT_HEADERS.length).setValues([SNAPSHOT_HEADERS]);
  }

  return sheet;
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

    if (fn === 'dailyRefreshSnapshot') {
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
 * 測試用：確認 Snapshot payload 是否有資料
 */
function testSnapshotPayload() {
  const payload = getFullSnapshotPayload();

  Logger.log('Root Folder ID: ' + payload.rootFolderId);
  Logger.log('Total Items: ' + payload.items.length);
  Logger.log('Root Items: ' + payload.rootItems.length);
  Logger.log('Updated At: ' + payload.updatedAt);

  if (payload.rootItems.length > 0) {
    Logger.log('First Root Item: ' + payload.rootItems[0].name);
  }
}