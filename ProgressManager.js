function getProgressInfo(labelName) {
  try {
    var cache = CacheService.getUserCache();
    var cacheKey = 'real_progress_' + labelName;
    var cached = cache.get(cacheKey);
    
    if (cached) {
      var cachedData = JSON.parse(cached);
      if (new Date().getTime() - cachedData.timestamp < 120000) {
        return cachedData.progress;
      }
    }
    
    var totalEstimate = getTotalEmailEstimate(labelName);
    
    var progress = {
      total: totalEstimate,
      exported: 0,
      remaining: totalEstimate,
      basedOnSheet: false,
      timestamp: new Date().getTime()
    };
    
    // Mettre en cache
    cache.put(cacheKey, JSON.stringify({
      progress: progress,
      timestamp: new Date().getTime()
    }), 300);
    
    log("Progression calculée - Total estimé: " + totalEstimate, 'INFO');
    return progress;
    
  } catch(e) {
    log("Erreur calcul progression: " + e.toString(), 'WARN');
    return {
      total: 0,
      exported: 0,
      remaining: 0,
      basedOnSheet: false,
      error: true
    };
  }
}

function getTotalEmailEstimate(labelName) {
  try {
    var cache = CacheService.getUserCache();
    var cacheKey = 'total_estimate_' + labelName;
    var cached = cache.get(cacheKey);
    
    if (cached) {
      return parseInt(cached);
    }
    
    // Méthode d'estimation plus précise
    var totalEstimate = estimateLabelSize(labelName);
    
    cache.put(cacheKey, totalEstimate.toString(), 600);
    return totalEstimate;
    
  } catch(e) {
    log("Erreur estimation totale: " + e.toString(), 'WARN');
    return 0;
  }
}

function estimateLabelSize(labelName) {
  try {
    // Échantillonnage sur plusieurs pages pour meilleure estimation
    var sampleSizes = [];
    
    for (var page = 0; page < 3; page++) {
      var start = page * 100;
      var threads = GmailApp.search("label:" + labelName, start, 50);
      
      if (threads.length === 0) break;
      
      var pageTotal = 0;
      for (var i = 0; i < threads.length; i++) {
        pageTotal += threads[i].getMessageCount();
      }
      sampleSizes.push(pageTotal);
      
      if (threads.length < 50) break;
    }
    
    if (sampleSizes.length === 0) return 0;
    
    // Estimation basée sur la moyenne des échantillons
    var avgPerPage = sampleSizes.reduce((a, b) => a + b, 0) / sampleSizes.length;
    var estimatedTotal = Math.round(avgPerPage * (sampleSizes.length * 2));
    
    log("Estimation taille label '" + labelName + "': " + estimatedTotal + " emails", 'INFO');
    return estimatedTotal;
    
  } catch(e) {
    log("Erreur estimation taille label: " + e.toString(), 'WARN');
    return 0;
  }
}

function refreshProgressCache(labelName) {
  try {
    var cache = CacheService.getUserCache();
    cache.remove('real_progress_' + labelName);
    cache.remove('total_estimate_' + labelName);
    
    log("Cache progression nettoyé pour: " + labelName, 'INFO');
    return true;
    
  } catch(e) {
    log("Erreur rafraîchissement cache: " + e.toString(), 'WARN');
    return false;
  }
}

function clearProgressCache(labelName) {
  try {
    var cache = CacheService.getUserCache();
    cache.remove('progress_' + labelName);
  } catch(e) {
    log("Erreur suppression cache progression: " + e.toString(), 'WARN');
  }
}