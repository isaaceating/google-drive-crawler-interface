/**
 * Google Drive Crawler interface
 * Version: v1.7.1
 *
 * Features:
 * 1. Lazy Load：Folder 展開時才讀取內容
 * 2. Search Index：全域搜尋只搜尋「文件檔名」
 * 3. Panel 1：顯示第 1 層，可展開第 2 層
 * 4. Panel 2：顯示選中第 2 層資料夾內容，可展開第 4 層
 * 5. 第 4 層資料夾不再繼續展開，改顯示不支援並提供 Google Drive 連結
 */

const ROOT_FOLDER_ID = '1zAFat5y1UL-vMqg5yQVy0SAgRD7WG0uY';

const SEARCH_INDEX_CACHE_KEY = 'KNOWLEDGE_SEARCH_INDEX_V171';
const CACHE_TIME = 21600;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lumens 產品彈藥庫')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 初始只抓 Root 第一層
 */
function getRootItems() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  return getFolderItemsLazy_(root);
}

/**
 * Lazy Load：展開或選取 Folder 時才抓子層
 */
function getFolderChildren(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  return getFolderItemsLazy_(folder);
}

/**
 * 只抓指定資料夾的下一層，不遞迴
 */
function getFolderItemsLazy_(folder) {
  const items = [];

  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();

    items.push({
      id: sub.getId(),
      name: sub.getName(),
      itemType: 'folder',
      hasChildren: true,
      openLink: sub.getUrl()
    });
  }

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    items.push(buildFileItem_(file));
  }

  items.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));
  return items;
}

/**
 * 全域搜尋 Index
 * 注意：前端只比對文件檔名，不比對資料夾 path
 */
function getSearchIndex() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SEARCH_INDEX_CACHE_KEY);

  if (cached) return JSON.parse(cached);

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const files = getAllFilesRecursive_(root, root.getName());

  files.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));

  cache.put(SEARCH_INDEX_CACHE_KEY, JSON.stringify(files), CACHE_TIME);
  return files;
}

function refreshSearchIndexCache() {
  CacheService.getScriptCache().remove(SEARCH_INDEX_CACHE_KEY);
}

/**
 * 建立搜尋用文件 Index
 * 只收 file，不收 folder
 */
function getAllFilesRecursive_(folder, path) {
  const results = [];

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const item = buildFileItem_(file);
    item.path = path;
    results.push(item);
  }

  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();
    results.push(...getAllFilesRecursive_(sub, path + ' / ' + sub.getName()));
  }

  return results;
}

function buildFileItem_(file) {
  const type = mapMimeType_(file.getMimeType());

  return {
    id: file.getId(),
    name: file.getName(),
    itemType: 'file',
    fileType: type,
    previewLink: buildPreviewLink_(file, type),
    openLink: file.getUrl()
  };
}

function mapMimeType_(mime) {
  if (mime === MimeType.GOOGLE_DOCS) return 'doc';
  if (mime === MimeType.GOOGLE_SLIDES) return 'slides';
  if (mime === MimeType.GOOGLE_SHEETS) return 'sheet';
  if (mime === MimeType.PDF) return 'pdf';

  if (String(mime).startsWith('image/')) return 'image';
  if (String(mime).startsWith('video/')) return 'video';

  return 'file';
}

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