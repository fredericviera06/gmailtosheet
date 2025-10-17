function getOrCreateSheet(sheetId, newSheetName, labelName) {
  log("getOrCreateSheet - sheetId: " + sheetId + ", newSheetName: " + newSheetName, 'INFO');

  if (newSheetName && newSheetName.trim() !== "") {
    try {
      var parentFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      var emailFolder = getOrCreateSubFolder(parentFolder, "Email");
      var labelFolder = getOrCreateSubFolder(emailFolder, labelName);
      
      var existingFiles = labelFolder.getFilesByName(newSheetName.trim());
      if (existingFiles.hasNext()) {
        var existingFile = existingFiles.next();
        log("Sheet existant trouvé avec ce nom: " + newSheetName + " (ID: " + existingFile.getId() + ")", 'INFO');
        var ss = SpreadsheetApp.openById(existingFile.getId());
        return ss.getSheets()[0];
      }
      
      // Créer nouveau sheet
      var ss = SpreadsheetApp.create(newSheetName.trim());
      var sheet = ss.getSheets()[0];
      
      try {
        var file = DriveApp.getFileById(ss.getId());
        file.moveTo(labelFolder);
        log("Nouveau Sheet créé et déplacé dans le dossier: " + newSheetName, 'INFO');
      } catch(moveError) {
        log("Impossible de déplacer le Sheet: " + moveError.toString(), 'WARN');
      }
      
      return sheet;
    } catch(error) {
      log("Erreur création nouveau Sheet: " + error.toString(), 'ERROR');
      throw new Error("Impossible de créer le Sheet: " + error.message);
    }
  }

  if (sheetId && sheetId.trim() !== "") {
    try {
      var sheet = SpreadsheetApp.openById(sheetId.trim()).getSheets()[0];
      log("Sheet existant ouvert: " + sheetId, 'INFO');
      return sheet;
    } catch(error) {
      log("Erreur ouverture Sheet existant: " + error.toString(), 'ERROR');
      throw new Error("Impossible d'ouvrir le Sheet: " + error.message);
    }
  }

  throw new Error("Veuillez créer un nouveau Sheet (entrez un nom) ou sélectionnez un Sheet existant dans la liste");
}

function setupSheetIfNeeded(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      "Message ID", "From", "From Name", "To", "To Name",
      "Cc", "Cc Name", "Bcc", "Bcc Name",
      "Subject", "Date Sent", "Date Received",
      "PDF Link", "Attachments Count", "Body Snippet", "Labels"
    ];
    for (var i = 1; i <= CONFIG.MAX_ATTACHMENT_COLUMNS; i++) {
      headers.push("Attachment " + i);
    }

    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
}

function getExistingMessageIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  try {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var cleanIds = ids.map(function(row) { return row[0] ? row[0].toString().trim() : ""; })
      .filter(function(id) { return id !== ""; });
    log("IDs existants récupérés: " + cleanIds.length, 'INFO');
    return cleanIds;
  } catch(error) {
    log("Erreur lecture IDs existants: " + error.toString(), 'ERROR');
    return [];
  }
}

function getExportedCountFromSheet(sheet, labelName) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    // Compter toutes les lignes (sauf l'en-tête)
    var exportedCount = lastRow - 1;
    log("Emails exportés comptés dans le Sheet: " + exportedCount, 'INFO');
    
    return exportedCount;
    
  } catch(error) {
    log("Erreur comptage emails exportés: " + error.toString(), 'WARN');
    return 0;
  }
}

function getUserSheets() {
  var startTime = new Date().getTime();
  var files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet" and trashed=false');
  var results = [];
  var count = 0;

  while (files.hasNext() && count < CONFIG.MAX_SHEETS_TO_SHOW) {
    if (new Date().getTime() - startTime > 3000) {
      log("Timeout lors du chargement des Sheets, " + count + " sheets chargés", 'WARN');
      break;
    }
    results.push(files.next());
    count++;
  }

  return results.sort(function(a, b) {
    return b.getLastUpdated() - a.getLastUpdated();
  });
}