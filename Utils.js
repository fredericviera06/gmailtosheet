// ========== LOGGING AVANCÉ ==========
var ADVANCED_LOGGER = {
  logSession: function(action, details) {
    var sessionLog = {
      timestamp: new Date().toISOString(),
      action: action,
      details: details
    };
    
    try {
      var props = PropertiesService.getUserProperties();
      var sessionLogs = props.getProperty('session_logs');
      var logs = sessionLogs ? JSON.parse(sessionLogs) : [];
      logs.push(sessionLog);
      
      if (logs.length > 50) {
        logs = logs.slice(-50);
      }
      
      props.setProperty('session_logs', JSON.stringify(logs));
    } catch(e) {
      // Ignorer les erreurs de logging
    }
  }
};

function log(message, level) {
  if (!CONFIG.ENABLE_LOGGING) return;
  
  var prefix = level === 'ERROR' ? '❌' : level === 'WARN' ? '⚠️' : 'ℹ️';
  var logMessage = prefix + ' ' + message;
  
  Logger.log(logMessage);
  
  if (level === 'ERROR') {
    console.error(logMessage);
    ADVANCED_LOGGER.logSession('ERROR', {
      message: message,
      time: new Date().toISOString()
    });
  }
}

function handleError(error, context) {
  log('Erreur dans ' + context + ': ' + error.toString(), 'ERROR');
  try {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('❌ Erreur: ' + error.message)
        .setType(CardService.NotificationType.ERROR))
      .build();
  } catch(e) {
    return;
  }
}

function validateDriveFolder() {
  if (!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID === "ID_DU_DOSSIER") {
    throw new Error("Veuillez configurer DRIVE_FOLDER_ID dans le script");
  }
  try {
    DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    return true;
  } catch(e) {
    throw new Error("ID du dossier Drive invalide: " + CONFIG.DRIVE_FOLDER_ID);
  }
}

function extractName(field) {
  if (!field) return "";
  var matches = field.match(/^([^<]+)<.*>$/);
  return matches && matches[1] ? matches[1].trim() : field.trim();
}

function generateProgressBar(percent) {
  var totalBars = 20;
  var filledBars = Math.round((percent / 100) * totalBars);
  var emptyBars = totalBars - filledBars;
  
  var bar = "▓".repeat(filledBars) + "░".repeat(emptyBars);
  return bar;
}

function calculateOptimalBatchSize(remainingEmails) {
  if (remainingEmails <= 10) return remainingEmails;
  if (remainingEmails <= 50) return 10;
  if (remainingEmails <= 100) return 20;
  return CONFIG.BATCH_SIZE;
}