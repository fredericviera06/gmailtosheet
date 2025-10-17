function getLabelFolder(labelName) {
  var cache = CacheService.getUserCache();
  var cacheKey = 'folder_' + labelName;
  var cached = cache.get(cacheKey);

  if (cached) {
    try {
      return DriveApp.getFolderById(cached);
    } catch(e) {
      cache.remove(cacheKey);
    }
  }

  var parentFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var emailFolder = getOrCreateSubFolder(parentFolder, "Email");
  var folders = emailFolder.getFoldersByName(labelName);
  var folder = folders.hasNext() ? folders.next() : emailFolder.createFolder(labelName);

  cache.put(cacheKey, folder.getId(), 3600);
  return folder;
}

function getOrCreateSubFolder(parentFolder, subFolderName) {
  var folders = parentFolder.getFoldersByName(subFolderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(subFolderName);
}

function createEmailPdf(msg, pdfFolder, subject, includeHtml) {
  var safeName = (subject || "email").replace(/[^a-zA-Z0-9 \-_.]/g, '').substring(0, 100);
  var fileName = safeName + "_" + msg.getId() + ".pdf";
  var content = includeHtml ? msg.getBody() : msg.getPlainBody();
  var html = '<html><head><meta charset="utf-8"></head><body>' + content + '</body></html>';
  var blob = Utilities.newBlob(html, MimeType.HTML, fileName).getAs(MimeType.PDF);
  var pdfFile = pdfFolder.createFile(blob);
  pdfFile.setName(fileName);
  return pdfFile.getUrl();
}