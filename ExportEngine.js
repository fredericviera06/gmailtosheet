function exportLabelEmailsBatch(e) {
  var startTime = new Date().getTime();
  var maxExecutionTime = CONFIG.EXECUTION_MARGIN_MS;

  try {
    log("=== DÉBUT EXPORT ===", 'INFO');

    validateDriveFolder();

    // Inputs
    var formInput = e && e.formInput ? e.formInput : {};
    var labelName = formInput.labelName;
    if (!labelName) {
      throw new Error("Veuillez sélectionner un label Gmail");
    }
    
    // Validation de la taille du batch
    var batchSize = parseInt(formInput.batchSize) || CONFIG.BATCH_SIZE;
    if (batchSize < CONFIG.MIN_BATCH_SIZE || batchSize > CONFIG.MAX_BATCH_SIZE) {
      throw new Error("La taille du batch doit être entre " + CONFIG.MIN_BATCH_SIZE + " et " + CONFIG.MAX_BATCH_SIZE);
    }
    
    // Vérification des quotas
    QUOTA_MANAGER.checkQuotas(batchSize);
    
    var sheetId = formInput.sheetId;
    var newSheetName = formInput.newSheetName;
    var createPdf = formInput.createPdf && formInput.createPdf.indexOf("yes") !== -1;
    var includeHtml = formInput.includeHtml && formInput.includeHtml.indexOf("yes") !== -1;
    var performanceMode = formInput.enablePerformanceMode && formInput.enablePerformanceMode.indexOf("yes") !== -1;
    var continueOnError = formInput.continueOnError && formInput.continueOnError.indexOf("yes") !== -1;

    // Obtenir ou créer le sheet
    var sheet = getOrCreateSheet(sheetId, newSheetName, labelName);
    log("Sheet obtenu: " + sheet.getName(), 'INFO');

    setupSheetIfNeeded(sheet);
    log("En-têtes configurés", 'INFO');

    var existingIds = getExistingMessageIds(sheet);
    log("IDs existants: " + existingIds.length, 'INFO');

    var cursor = getEnhancedCursorForLabel(labelName);
    log("Cursor pour label '" + labelName + "': " + JSON.stringify(cursor), 'INFO');

    var emailsToProcess = getEmailsToProcess(labelName, existingIds, cursor);
    log("Emails à traiter: " + emailsToProcess.length, 'INFO');

    if (emailsToProcess.length === 0) {
      clearCursorForLabel(labelName);
      clearProgressCache(labelName);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText("✅ Aucun email à exporter (ou déjà exportés)."))
        .setStateChanged(true)
        .build();
    }

    var batch = emailsToProcess.slice(0, batchSize);
    var rowsToAdd = [];
    var processedCount = 0;
    var skippedDuplicates = 0;
    var existingIdsSet = {};
    existingIds.forEach(function(id){ if(id) existingIdsSet[id.toString()] = true; });

    log("Traitement de " + batch.length + " emails...", 'INFO');

    var lastProcessedCursor = null;

    // Traitement par chunks pour éviter les timeouts
    rowsToAdd = processEmailBatchInChunks(batch, labelName, createPdf, includeHtml, batchSize, continueOnError);
    processedCount = rowsToAdd.length;

    // Mettre à jour le cursor
    if (processedCount > 0) {
      var lastMsg = batch[Math.min(processedCount - 1, batch.length - 1)];
      var t = lastMsg.getThread();
      var messagesInThread = t.getMessages();
      var messageIndex = -1;
      for (var mi = 0; mi < messagesInThread.length; mi++) {
        if (messagesInThread[mi].getId() === lastMsg.getId()) { messageIndex = mi; break; }
      }
      
      var previousExported = cursor ? (cursor.exportedCount || 0) : 0;
      lastProcessedCursor = {
        threadId: t.getId(),
        messageIndex: messageIndex,
        exportedCount: previousExported + processedCount
      };
    }

    if (skippedDuplicates > 0) {
      log("⚠️ " + skippedDuplicates + " doublon(s) ignoré(s) pendant le traitement", 'WARN');
    }

    log("Rows à ajouter: " + rowsToAdd.length, 'INFO');

    if (rowsToAdd.length > 0) {
      try {
        var lastRow = sheet.getLastRow();
        if (lastRow < 1) lastRow = 1;
        log("Dernière ligne du Sheet: " + lastRow, 'INFO');

        var targetRange = sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length);
        log("Range cible: " + targetRange.getA1Notation(), 'INFO');

        targetRange.setValues(rowsToAdd);
        log("✅ Données écrites dans le Sheet!", 'INFO');

      } catch(error) {
        log("❌ Erreur écriture dans Sheet: " + error.toString(), 'ERROR');
        throw new Error("Échec de l'écriture dans le Sheet: " + error.message);
      }
    } else {
      log("⚠️ Aucune donnée à écrire", 'WARN');
    }

    if (lastProcessedCursor) {
      setEnhancedCursorForLabel(labelName, lastProcessedCursor, true);
      log("Cursor mis à jour pour reprise: " + JSON.stringify(lastProcessedCursor), 'INFO');
    }

    if (emailsToProcess.length <= batch.length) {
      clearCursorForLabel(labelName);
      clearProgressCache(labelName);
      log("Aucun email restant - cursor effacé", 'INFO');
    }

    var remaining = Math.max(0, emailsToProcess.length - processedCount);
    var message = "✅ " + processedCount + " email(s) exporté(s).\n";
    if (remaining > 0) {
      message += "📊 " + remaining + " email(s) restant(s).\n";
      message += "Cliquez à nouveau pour continuer.";
    } else {
      message += "🎉 Export terminé!";
    }

    var executionTime = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    log("Batch exporté: " + processedCount + " emails en " + executionTime + "s, " + remaining + " restants", 'INFO');
    log("=== FIN EXPORT ===", 'INFO');

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(message))
      .setStateChanged(true)
      .build();

  } catch(error) {
    log("❌ ERREUR FATALE: " + error.toString(), 'ERROR');
    return handleError(error, 'exportLabelEmailsBatch');
  }
}

function processEmailBatchInChunks(messages, labelName, createPdf, includeHtml, batchSize, continueOnError) {
  var rowsToAdd = [];
  var chunkSize = 5;
  
  for (var i = 0; i < messages.length; i += chunkSize) {
    var chunk = messages.slice(i, i + chunkSize);
    var chunkResults = processChunk(chunk, labelName, createPdf, includeHtml, continueOnError);
    rowsToAdd = rowsToAdd.concat(chunkResults);
    
    if (i + chunkSize < messages.length) {
      Utilities.sleep(500);
    }
  }
  
  return rowsToAdd;
}

function processChunk(messages, labelName, createPdf, includeHtml, continueOnError) {
  var chunkResults = [];
  
  for (var i = 0; i < messages.length; i++) {
    try {
      var rowData = processEmailMessage(messages[i], labelName, createPdf, includeHtml);
      chunkResults.push(rowData);
    } catch(error) {
      log("Erreur dans chunk traitement email " + i + ": " + error.toString(), 'ERROR');
      if (!continueOnError) {
        throw error;
      }
    }
  }
  
  return chunkResults;
}