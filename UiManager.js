function getContextualAddOn(e) {
  try {
    validateDriveFolder();
    return buildEnhancedCard(e);
  } catch(error) {
    return buildErrorCard(error.message);
  }
}

function buildEnhancedCard(e) {
  var card = CardService.newCardBuilder();
  var headerSection = CardService.newCardSection()
    .setHeader("🚀 Export Gmail vers Sheets Pro")
    .addWidget(CardService.newTextParagraph()
      .setText("Exportez vos emails intelligemment avec gestion des quotas et reprise sur erreur."));

  // Section configuration principale
  var mainSection = buildMainConfigurationSection(e);
  card.addSection(mainSection);

  // Section options avancées
  var advancedSection = buildAdvancedOptionsSection(e);
  card.addSection(advancedSection);

  // Section progression détaillée
  var progressSection = buildEnhancedProgressSection(e);
  if (progressSection) {
    card.addSection(progressSection);
  }

  // Section actions
  var actionSection = buildActionSection(e);
  card.addSection(actionSection);

  return card.build();
}

function buildMainConfigurationSection(e) {
  var formInput = e && e.formInput ? e.formInput : {};
  var selectedLabel = formInput.labelName || "";
  var selectedBatchSize = formInput.batchSize || CONFIG.BATCH_SIZE.toString();

  var section = CardService.newCardSection()
    .setHeader("⚙️ Configuration Principale");

  // Labels Gmail
  try {
    var labels = getGmailLabels();
    var selectLabel = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle("🏷️ Label Gmail à exporter")
      .setFieldName("labelName")
      .setOnChangeAction(CardService.newAction().setFunctionName("onLabelChange"));

    labels.forEach(function(label) {
      selectLabel.addItem(label.name, label.name, label.name === selectedLabel);
    });
    section.addWidget(selectLabel);
  } catch(e) {
    log("Erreur chargement labels: " + e.toString(), 'ERROR');
  }

  // Taille du batch
  var batchSizeInput = CardService.newTextInput()
    .setFieldName("batchSize")
    .setTitle("📊 Taille du batch")
    .setValue(selectedBatchSize)
    .setHint("Entre " + CONFIG.MIN_BATCH_SIZE + " et " + CONFIG.MAX_BATCH_SIZE + " emails par batch");
  section.addWidget(batchSizeInput);

  // Nouveau sheet
  var newSheetInput = CardService.newTextInput()
    .setFieldName("newSheetName")
    .setTitle("📄 Créer un nouveau Google Sheet")
    .setHint("Ex: Export Emails 2025");
  section.addWidget(newSheetInput);

  section.addWidget(CardService.newTextParagraph()
    .setText("<b>--- OU ---</b>"));

  // Sheets existants
  try {
    var sheets = getUserSheets();
    if (sheets.length > 0) {
      var select = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("📊 Choisir un Sheet existant")
        .setFieldName("sheetId");
      select.addItem("-- Sélectionner un Sheet --", "", true);
      sheets.forEach(function(file) {
        select.addItem(file.getName(), file.getId(), false);
      });
      section.addWidget(select);
    } else {
      section.addWidget(CardService.newTextParagraph()
        .setText("⚠️ Aucun Sheet trouvé. Créez-en un nouveau ci-dessus."));
    }
  } catch(e) {
    log("Erreur chargement Sheets: " + e.toString(), 'ERROR');
    section.addWidget(CardService.newTextParagraph()
      .setText("⚠️ Erreur de chargement des Sheets"));
  }

  return section;
}

function buildAdvancedOptionsSection(e) {
  var formInput = e && e.formInput ? e.formInput : {};
  var section = CardService.newCardSection()
    .setHeader("🔧 Options Avancées")
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(1);

  // Options PDF
  var pdfCheckbox = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("createPdf")
    .addItem("📄 Créer un PDF de chaque email", "yes", false);
  section.addWidget(pdfCheckbox);

  var htmlCheckbox = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("includeHtml")
    .addItem("🌐 Inclure le HTML de l'email (PDF plus riche)", "yes", true);
  section.addWidget(htmlCheckbox);

  // Optimisation performance
  var perfInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("enablePerformanceMode")
    .addItem("🎯 Mode performance", "yes", false);
  section.addWidget(perfInput);

  // Gestion des erreurs
  var errorInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("continueOnError")
    .addItem("🛡️ Continuer en cas d'erreur", "yes", true);
  section.addWidget(errorInput);

  return section;
}

function buildEnhancedProgressSection(e) {
  try {
    var formInput = e && e.formInput ? e.formInput : {};
    var labelName = formInput.labelName;
    
    if (!labelName) {
      return buildProgressHelpSection();
    }
    
    var progressInfo = getProgressInfo(labelName);
    
    if (!progressInfo || progressInfo.total === 0) {
      return buildProgressHelpSection();
    }

    var section = CardService.newCardSection()
      .setHeader("📈 Progression");

    var percent = progressInfo.total > 0 ? Math.round((progressInfo.exported / progressInfo.total) * 100) : 0;
    
    var progressText = "📧 Total estimé: " + progressInfo.total + " emails\n" +
                      "✅ Exportés: " + progressInfo.exported + "\n" +
                      "⏳ Restants: " + progressInfo.remaining + "\n\n" +
                      "Progression: " + percent + "%\n" +
                      generateProgressBar(percent);
    
    if (progressInfo.error) {
      progressText += "\n\n⚠️ Estimation approximative";
    }

    section.addWidget(CardService.newTextParagraph().setText(progressText));
    
    // Bouton de rafraîchissement
    var refreshAction = CardService.newAction()
      .setFunctionName("onLabelChange")
      .setParameters({forceRefresh: 'true'});
    var refreshButton = CardService.newTextButton()
      .setText("🔄 Actualiser la progression")
      .setOnClickAction(refreshAction)
      .setTextButtonStyle(CardService.TextButtonStyle.TEXT);
    section.addWidget(refreshButton);
    
    return section;
    
  } catch(e) {
    log("Erreur construction section progression: " + e.toString(), 'WARN');
    return buildProgressHelpSection();
  }
}

function buildProgressHelpSection() {
  var section = CardService.newCardSection()
    .setHeader("📈 Progression")
    .addWidget(CardService.newTextParagraph()
      .setText("La progression s'affichera après avoir sélectionné un label."));
  return section;
}

function buildActionSection(e) {
  var section = CardService.newCardSection();
  var formInput = e && e.formInput ? e.formInput : {};
  var labelName = formInput.labelName;
  
  // Bouton export principal
  var exportAction = CardService.newAction().setFunctionName("exportLabelEmailsBatch");
  var exportButton = CardService.newTextButton()
    .setText("🚀 Exporter le batch suivant")
    .setOnClickAction(exportAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  
  if (!labelName) {
    exportButton.setDisabled(true);
  }
  
  section.addWidget(exportButton);

  return section;
}

function buildErrorCard(message) {
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText("❌ " + message));
  card.addSection(section);
  return card.build();
}

function onLabelChange(e) {
  try {
    log("onLabelChange appelé", 'INFO');
    
    var formInput = e.formInput || {};
    var labelName = formInput.labelName;
    
    // Forcer le rafraîchissement du cache
    if (labelName) {
      refreshProgressCache(labelName);
    }
    
    var nav = CardService.newNavigation().updateCard(buildEnhancedCard(e));
    return CardService.newActionResponseBuilder()
      .setNavigation(nav)
      .build();
      
  } catch(error) {
    log("Erreur dans onLabelChange: " + error.toString(), 'ERROR');
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("⚠️ Erreur lors du chargement")
        .setType(CardService.NotificationType.WARNING))
      .build();
  }
}