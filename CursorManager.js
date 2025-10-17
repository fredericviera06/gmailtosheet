function getEnhancedCursorForLabel(labelName) {
  try {
    var props = PropertiesService.getUserProperties();
    var cursorData = props.getProperty('cursor_' + labelName);
    var retryData = props.getProperty('retry_' + labelName);
    
    var cursor = cursorData ? JSON.parse(cursorData) : null;
    var retryInfo = retryData ? JSON.parse(retryData) : { count: 0, lastError: null };
    
    // Si trop de tentatives, réinitialiser
    if (retryInfo.count > 3) {
      log("Trop de tentatives échouées, réinitialisation du cursor", 'WARN');
      clearCursorForLabel(labelName);
      return null;
    }
    
    return cursor;
  } catch(e) {
    log("Erreur lecture cursor amélioré: " + e.toString(), 'WARN');
    return null;
  }
}

function setEnhancedCursorForLabel(labelName, cursorObj, success) {
  try {
    var props = PropertiesService.getUserProperties();
    
    if (success) {
      // Réussite : sauvegarder cursor et réinitialiser compteur d'erreurs
      props.setProperty('cursor_' + labelName, JSON.stringify(cursorObj));
      props.deleteProperty('retry_' + labelName);
    } else {
      // Échec : incrémenter compteur d'erreurs
      var retryData = props.getProperty('retry_' + labelName);
      var retryInfo = retryData ? JSON.parse(retryData) : { count: 0, lastError: null };
      retryInfo.count++;
      retryInfo.lastError = new Date().toISOString();
      props.setProperty('retry_' + labelName, JSON.stringify(retryInfo));
    }
  } catch(e) {
    log("Erreur sauvegarde cursor amélioré: " + e.toString(), 'WARN');
  }
}

function getCursorForLabel(labelName) {
  try {
    var prop = PropertiesService.getUserProperties().getProperty('cursor_' + labelName);
    return prop ? JSON.parse(prop) : null;
  } catch(e) {
    log("Erreur lecture cursor: " + e.toString(), 'WARN');
    return null;
  }
}

function setCursorForLabel(labelName, cursorObj) {
  try {
    PropertiesService.getUserProperties().setProperty('cursor_' + labelName, JSON.stringify(cursorObj));
  } catch(e) {
    log("Erreur sauvegarde cursor: " + e.toString(), 'WARN');
  }
}

function clearCursorForLabel(labelName) {
  try {
    PropertiesService.getUserProperties().deleteProperty('cursor_' + labelName);
    PropertiesService.getUserProperties().deleteProperty('retry_' + labelName);
  } catch(e) {
    log("Erreur suppression cursor: " + e.toString(), 'WARN');
  }
}