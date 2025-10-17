function getGmailLabels() {
  var cache = CacheService.getUserCache();
  var cached = cache.get('gmail_labels');
  if (cached) return JSON.parse(cached);

  var labels = GmailApp.getUserLabels()
    .map(function(label) {
      return {name: label.getName()};
    })
    .sort(function(a, b) { return a.name.localeCompare(b.name); });

  cache.put('gmail_labels', JSON.stringify(labels), 300);
  return labels;
}

function getPaginatedThreads(labelName, maxThreads) {
  var allThreads = [];
  var start = 0;
  var batchSize = 100;
  
  try {
    while (start < maxThreads) {
      var remaining = maxThreads - start;
      var currentBatchSize = Math.min(batchSize, remaining);
      
      if (currentBatchSize <= 0) break;
      
      var threads = GmailApp.search("label:" + labelName, start, currentBatchSize);
      
      if (threads.length === 0) {
        break;
      }
      
      allThreads = allThreads.concat(threads);
      start += threads.length;
      
      log("Lot de threads récupéré: " + threads.length + " (total: " + allThreads.length + ")", 'INFO');
      
      if (threads.length === currentBatchSize) {
        Utilities.sleep(500);
      }
    }
    
    log("Total threads récupérés pour label '" + labelName + "': " + allThreads.length, 'INFO');
    return allThreads;
    
  } catch(error) {
    log("Erreur lors de la recherche paginée: " + error.toString(), 'ERROR');
    return allThreads;
  }
}

function getEmailsToProcess(labelName, existingIds, cursor) {
  var existingIdsSet = {};
  existingIds.forEach(function(id) { if (id) existingIdsSet[id.toString()] = true; });

  log("Set d'IDs existants créé: " + Object.keys(existingIdsSet).length + " IDs", 'INFO');

  var threads = getPaginatedThreads(labelName, CONFIG.MAX_THREADS_TO_SCAN);
  var emailsToProcess = [];
  var processedCount = 0;
  var duplicatesFound = 0;
  var startAdding = true;

  if (cursor && cursor.threadId) {
    startAdding = false;
  }

  var MAX_EMAILS_TO_SCAN = CONFIG.MAX_TOTAL_EMAILS_TO_SCAN;

  for (var i = 0; i < threads.length && processedCount < MAX_EMAILS_TO_SCAN; i++) {
    var t = threads[i];
    if (!startAdding) {
      if (t.getId() === cursor.threadId) {
        var msgs = t.getMessages();
        for (var mi = cursor.messageIndex + 1; mi < msgs.length; mi++) {
          var m = msgs[mi];
          var id = m.getId();
          if (existingIdsSet[id]) { duplicatesFound++; continue; }
          emailsToProcess.push(m);
          processedCount++;
          if (processedCount >= MAX_EMAILS_TO_SCAN) break;
        }
        startAdding = true;
        if (processedCount >= MAX_EMAILS_TO_SCAN) break;
        continue;
      } else {
        continue;
      }
    }

    var messages = t.getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var msgId = msg.getId();

      if (existingIdsSet[msgId]) {
        duplicatesFound++;
        continue;
      }

      emailsToProcess.push(msg);
      processedCount++;

      if (processedCount >= MAX_EMAILS_TO_SCAN) break;
    }
  }

  log("Emails trouvés à traiter: " + emailsToProcess.length + ", doublons ignorés: " + duplicatesFound, 'INFO');
  return emailsToProcess;
}

function processEmailMessage(msg, labelName, createPdf, includeHtml) {
  var from = msg.getFrom();
  var fromName = extractName(from);
  var to = msg.getTo();
  var toName = extractName(to);
  var cc = msg.getCc() || "";
  var ccName = extractName(cc);
  var bcc = msg.getBcc() || "";
  var bccName = extractName(bcc);
  var subject = msg.getSubject() || "(Sans objet)";
  var dateSent = msg.getDate();

  var snippet = "";
  try {
    snippet = msg.getPlainBody().substring(0, CONFIG.SNIPPET_LENGTH);
  } catch(e) {
    snippet = "(Erreur lecture corps)";
  }

  var labels = "";
  try {
    var thread = msg.getThread();
    labels = thread.getLabels().map(function(l) { return l.getName(); }).join(", ");
  } catch(e) {
    labels = labelName;
  }

  var labelFolder = getLabelFolder(labelName);

  var pdfLink = "";
  if (createPdf) {
    try {
      var pdfFolder = getOrCreateSubFolder(labelFolder, "PDF email");
      pdfLink = createEmailPdf(msg, pdfFolder, subject, includeHtml);
    } catch(error) {
      log("Erreur création PDF: " + error.toString(), 'WARN');
    }
  }

  var attachmentsFolder = getOrCreateSubFolder(labelFolder, "Pièces jointes");
  var attachments = msg.getAttachments();
  var attCount = attachments.length;
  var attColumns = [];

  for (var i = 0; i < attachments.length; i++) {
    if (i >= CONFIG.MAX_ATTACHMENT_COLUMNS) {
      log("Plus de " + CONFIG.MAX_ATTACHMENT_COLUMNS + " PJ, certaines ne seront pas sauvegardées", 'WARN');
      break;
    }

    try {
      var att = attachments[i];
      var saved = attachmentsFolder.createFile(att);
      var fileUrl = saved.getUrl();
      attColumns.push(fileUrl);
    } catch(error) {
      log("Erreur sauvegarde PJ " + i + ": " + error.toString(), 'WARN');
      attColumns.push("(Erreur sauvegarde)");
    }
  }

  while (attColumns.length < CONFIG.MAX_ATTACHMENT_COLUMNS) {
    attColumns.push("");
  }

  var row = [
    msg.getId(), from, fromName, to, toName,
    cc, ccName, bcc, bccName,
    subject, dateSent, dateSent,
    pdfLink, attCount, snippet, labels
  ];

  row = row.concat(attColumns);
  return row;
}