// ========== CONFIGURATION GLOBALE ==========
var CONFIG = {
  DRIVE_FOLDER_ID: "1SwZIXzRXxVpVCja8N5PVYe2o2nadl8uv",
  BATCH_SIZE: 20,
  MAX_SHEETS_TO_SHOW: 15,
  SNIPPET_LENGTH: 200,
  PDF_QUALITY: false,
  ENABLE_LOGGING: true,
  MAX_THREADS_TO_SCAN: 500,
  MAX_ATTACHMENT_COLUMNS: 10,
  EXECUTION_MARGIN_MS: 30000,
  MIN_BATCH_SIZE: 1,
  MAX_BATCH_SIZE: 50,
  MAX_TOTAL_EMAILS_TO_SCAN: 2000
};

// ========== GESTION DES QUOTAS ==========
var QUOTA_MANAGER = {
  lastExecutionTime: 0,
  dailyEmailCount: 0,
  resetTime: 0,
  
  checkQuotas: function(batchSize) {
    var now = new Date().getTime();
    
    // Reset quotidien
    if (now > this.resetTime) {
      this.dailyEmailCount = 0;
      this.resetTime = now + (24 * 60 * 60 * 1000);
    }
    
    // Vérifier la limite quotidienne (20,000 emails/jour)
    if (this.dailyEmailCount + batchSize > 19000) {
      throw new Error("Quota quotidien presque atteint. " + 
                     this.dailyEmailCount + " emails traités aujourd'hui.");
    }
    
    // Respecter le délai entre les exécutions
    if (now - this.lastExecutionTime < 1000) {
      Utilities.sleep(1000 - (now - this.lastExecutionTime));
    }
    
    this.lastExecutionTime = new Date().getTime();
    this.dailyEmailCount += batchSize;
    
    log("Quotas: " + this.dailyEmailCount + " emails traités aujourd'hui", 'INFO');
  }
};