/**
 * Lumens 產品彈藥庫
 * Version: 介面 v1.2
 *
 * 結構規則：
 * 1. Root 內可有 folder / file
 * 2. 第一階面板顯示 Root 內所有項目
 * 3. 點第一階 file -> 右側預覽
 * 4. 點第一階 folder -> 第二階面板顯示該 folder 內的 file / subfolder
 * 5. 點第二階 file -> 右側預覽
 * 6. 點第二階 subfolder -> 在第二階面板內展開/收合該 subfolder 內的 files
 * 7. subfolder 內只支援 files，不再往下抓更深層
 */

const ROOT_FOLDER_ID = '1zAFat5y1UL-vMqg5yQVy0SAgRD7WG0uY';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lumens 產品彈藥庫')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getKnowledgeTree() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);

  return {
    rootName: root.getName(),
    items: getFolderItems_(root, 0)
  };
}

/**
 * level 0 = root
 * level 1 = 主資料夾
 * level 2 = 子資料夾（只抓 files，不再抓 folder）
 */
function getFolderItems_(folder, level) {
  const items = [];

  // 先抓 folders
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
      // Root 下的主資料夾：抓其內 file + subfolder
      folderItem.children = getFolderItems_(sub, 1);
    } else if (level === 1) {
      // 主資料夾下的子資料夾：只抓文件，不再抓更深層
      folderItem.children = getFilesOnly_(sub, 2);
    } else {
      // 不再往下
      folderItem.children = [];
    }

    items.push(folderItem);
  }

  // 再抓 files
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    items.push(buildFileItem_(file, level));
  }

  // 名稱排序
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

  if (type === 'doc') {
    return `https://docs.google.com/document/d/${id}/preview`;
  }

  if (type === 'slides') {
    return `https://docs.google.com/presentation/d/${id}/preview`;
  }

  if (type === 'sheet') {
    return `https://docs.google.com/spreadsheets/d/${id}/preview`;
  }

  if (type === 'pdf') {
    return `https://drive.google.com/file/d/${id}/preview`;
  }

  if (type === 'image') {
    return `https://drive.google.com/thumbnail?id=${id}&sz=w2000`;
  }

  if (type === 'video') {
    return `https://drive.google.com/file/d/${id}/preview`;
  }

  return file.getUrl();
}