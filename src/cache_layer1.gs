// ============================================================
// LAYER 1: Dataset Cache (CacheService) — Dr. Media v2.1
// ============================================================
var DATASET_CACHE_KEY = 'dr_media_dataset_v1';
var DATASET_CACHE_TTL = 3600; // 1 hour in seconds

function _getDatasetCached_() {
  var fromCache = _getDatasetCachedSafe_();
  if (fromCache) return fromCache;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName('dataset');
  if (!sh) throw new Error("No se encontró la hoja 'dataset'");
  var rows = readDataset(sh);
  var raw = sh.getDataRange().getValues();
  var headers = raw[0].map(function(h){ return String(h).trim(); });
  var cache = CacheService.getScriptCache();
  var payload = JSON.stringify({ headers: headers, rows: rows });
  if (payload.length > 90000) {
    var chunkSize = 90000;
    var chunks = Math.ceil(payload.length / chunkSize);
    for (var i = 0; i < chunks; i++) {
      var key = DATASET_CACHE_KEY + '_chunk_' + i;
      cache.put(key, payload.slice(i * chunkSize, (i+1) * chunkSize), DATASET_CACHE_TTL);
    }
    cache.put(DATASET_CACHE_KEY + '_meta', JSON.stringify({ chunked: true, count: chunks }), DATASET_CACHE_TTL);
  } else {
    cache.put(DATASET_CACHE_KEY, payload, DATASET_CACHE_TTL);
  }
  return { headers: headers, rows: rows };
}

function _getDatasetCachedSafe_() {
  var cache = CacheService.getScriptCache();
  var meta = cache.get(DATASET_CACHE_KEY + '_meta');
  if (meta) {
    try {
      var m = JSON.parse(meta);
      if (m.chunked) {
        var parts = [];
        for (var i = 0; i < m.count; i++) {
          var chunk = cache.get(DATASET_CACHE_KEY + '_chunk_' + i);
          if (!chunk) return null;
          parts.push(chunk);
        }
        return JSON.parse(parts.join(''));
      }
    } catch(e) { return null; }
  }
  var single = cache.get(DATASET_CACHE_KEY);
  if (!single) return null;
  try { return JSON.parse(single); } catch(e) { return null; }
}

function invalidateDatasetCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(DATASET_CACHE_KEY);
  cache.remove(DATASET_CACHE_KEY + '_meta');
  for (var i = 0; i < 20; i++) {
    cache.remove(DATASET_CACHE_KEY + '_chunk_' + i);
  }
}

function onEdit(e) {
  if (e && e.source) {
    var sheet = e.range.getSheet();
    if (sheet.getName() === 'dataset') {
      invalidateDatasetCache();
    }
  }
}

function forceRefreshCache() {
  invalidateDatasetCache();
  try { _getDatasetCached_(); } catch(ex) {}
  return { ok: true, message: 'Caché actualizado correctamente.' };
}
