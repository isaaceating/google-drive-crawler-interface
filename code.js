/**
 * Google Drive Crawler interface
 * 範例: Lumens 產品彈藥庫
 * Version: v1.5 (FAST LOAD + LOADING DIALOG)
 *
 * v1.5：
 * 1. 保留 v1.4 CacheService 秒開架構
 * 2. 前端新增初始讀取視窗
 * 3. 其他功能與資料結構完全不變
 * 結構規則：
 * 1. Root 內可有 folder / file
 * 2. 第一階面板顯示 Root 內所有項目
 * 3. 點第一階 file -> 右側預覽
 * 4. 點第一階 folder -> 第二階面板顯示該 folder 內的 file / subfolder
 * 5. 點第二階 file -> 右側預覽
 * 6. 點第二階 subfolder -> 在第二階面板內展開/收合該 subfolder 內的 files
 * 7. subfolder 內只支援 files，不再往下抓更深層
 * 
 */

const ROOT_FOLDER_ID = '1zAFat5y1UL-vMqg5yQVy0SAgRD7WG0uY';
const CACHE_KEY = 'KNOWLEDGE_TREE_CACHE';
const CACHE_TIME = 21600; // 6 小時

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lumens 產品彈藥庫')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getKnowledgeTree() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached) {
    return JSON.parse(cached);
  }

  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);

  const data = {
    rootName: root.getName(),
    items: getFolderItems_(root, 0)
  };

  cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TIME);

  return data;
}

/**
 * 手動刷新 cache
 */
function refreshCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY);
}

/**
 * 原本邏輯完全保留
 */

function getFolderItems_(folder, level) {
  const items = [];

  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();

    const folderItem = {
      id: sub.getId(),
      name: sub.getName(),
      itemType: 'folder',
      level: level,
      children: []
    };

    if (level === 0) {
      folderItem.children = getFolderItems_(sub, 1);
    } else if (level === 1) {
      folderItem.children = getFilesOnly_(sub, 2);
    }

    items.push(folderItem);
  }

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    items.push(buildFileItem_(file, level));
  }

  items.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));

  return items;
}

function getFilesOnly_(folder, level) {
  const items = [];
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    items.push(buildFileItem_(file, level));
  }

  items.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));
  return items;
}

function buildFileItem_(file, level) {
  const type = mapMimeType_(file.getMimeType());
  return {
    id: file.getId(),
    name: file.getName(),
    itemType: 'file',
    fileType: type,
    level: level,
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