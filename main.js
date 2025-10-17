// ========== POINTS D'ENTRÉE PRINCIPAUX ==========
function getContextualAddOn(e) {
  // Appeler directement la fonction globale (pas de UiManager.)
  try {
    validateDriveFolder();
    return buildEnhancedCard(e);
  } catch(error) {
    return buildErrorCard(error.message);
  }
}

function onGmailMessage(e) {
  return getContextualAddOn(e);
}

function onGmailCompose(e) {
  return getContextualAddOn(e);
}

// ========== FONCTIONS UTILITAIRES GLOBALES ==========
function removeDuplicatesFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Supprimer les doublons',
    'Cette action va supprimer toutes les lignes avec des Message ID en double.\nContinuer ?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Opération annulée');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('Aucune donnée à traiter');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var seenIds = {};
  var rowsToDelete = [];
  var duplicatesCount = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    var msgId = data[i][0] ? data[i][0].toString().trim() : "";
    if (!msgId) {
      rowsToDelete.push(i + 2);
      continue;
    }

    if (seenIds[msgId]) {
      rowsToDelete.push(i + 2);
      duplicatesCount++;
    } else {
      seenIds[msgId] = true;
    }
  }

  if (rowsToDelete.length === 0) {
    ui.alert('✅ Aucun doublon trouvé !');
    return;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowIndex) { sheet.deleteRow(rowIndex); });

  ui.alert('✅ Nettoyage terminé !\n\n' + duplicatesCount + ' doublon(s) supprimé(s)\n' + (rowsToDelete.length - duplicatesCount) + ' ligne(s) vide(s) supprimée(s)');
  Logger.log('Doublons supprimés: ' + duplicatesCount + ', lignes vides: ' + (rowsToDelete.length - duplicatesCount));
}

// ========== FONCTION DE TEST ==========
function testExport() {
  log("=== TEST DÉMARRÉ ===", 'INFO');
  
  try {
    var labels = getGmailLabels();
    log("Labels trouvés: " + labels.length, 'INFO');
    
    if (labels.length > 0) {
      var testLabel = labels[0].name;
      log("Test avec label: " + testLabel, 'INFO');
      
      var progress = getProgressInfo(testLabel);
      log("Progression test: " + JSON.stringify(progress), 'INFO');
    }
    
  } catch(e) {
    log("Erreur test: " + e.toString(), 'ERROR');
  }
  
  log("=== TEST TERMINÉ ===", 'INFO');
}