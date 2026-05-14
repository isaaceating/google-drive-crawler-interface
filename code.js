/**
 * Google Drive Crawler interface
 * Version: v2.0.0
 *
 * Features:
 * 1. Google Sheet Snapshot：前台瀏覽與搜尋改讀 Snapshot Sheet，不再即時掃 Drive
 * 2. Manual Refresh：前台按 Refresh Database 可手動重建 Snapshot
 * 3. Daily Auto Refresh：可建立每日自動重建 Snapshot trigger
 * 4. Frontend Cache：保留前端 childrenCache，提升同一次使用體驗
 * 5. Language：?lang=en / ?lang=ch，預設英文
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
 * Frontend API：取得 Root folder 第一層內容
 * v2.0 改為讀 Google Sheet Snapshot
 */
function getRootItems() {
  return getSnapshotChildren_(ROOT_FOLDER_ID);
}

/**
 * Frontend API：取得指定 folder 的子項目
 * v2.0 改為讀 Google Sheet Snapshot
 */
function getFolderChildren(folderId) {
  return getSnapshotChildren_(folderId);
}

/**
 * Frontend API：取得搜尋索引
 * v2.0 改為直接讀 Snapshot Sheet，不再使用 CacheService
 */
function getSearchIndex() {
  const sheet = getSnapshotSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const headers = values[0];
  const rows = values.slice(1);
  const idx = buildHeaderIndex_(headers);

  const files = rows
    .filter(row => row[idx.itemType] === 'file')
    .map(row => ({
      id: row[idx.id],
      name: row[idx.name],
      itemType: row[idx.itemType],
      fileType: row[idx.fileType],
      previewLink: row[idx.previewLink],
      openLink: row[idx.openLink],
      path: row[idx.path]
    }));

  files.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));
  return files;
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
 * 會掃描 ROOT_FOLDER_ID 底下所有 folder / file，並寫入 Snapshot Sheet
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
 */
function getSnapshotChildren_(parentId) {
  const sheet = getSnapshotSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  const headers = values[0];
  const rows = values.slice(1);
  const idx = buildHeaderIndex_(headers);

  const items = rows
    .filter(row => row[idx.parentId] === parentId)
    .map(row => ({
      id: row[idx.id],
      name: row[idx.name],
      itemType: row[idx.itemType],
      fileType: row[idx.fileType],
      previewLink: row[idx.previewLink],
      openLink: row[idx.openLink],
      path: row[idx.path],
      hasChildren: row[idx.itemType] === 'folder'
    }));

  items.sort((a, b) => {
    if (a.itemType !== b.itemType) {
      return a.itemType === 'folder' ? -1 : 1;
    }
    return a.name.localeCompare(b.name, 'zh-Hant');
  });

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