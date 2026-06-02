// ============================================================
// LAYER 2: Result Cache Wrapper for getDashboardData
// getDashboardData_BASE is the original getDashboardData in codigo_base.gs
// This wrapper adds CacheService result-level caching (Layer 2).
// ============================================================

function getDashboardData(targetMonth, targetYear, groupBy, filterPartners, filterClusters, filterPos, filterLobs, filterProductos) {
  // Build cache key from all params
  var cacheKey = 'drm_result_' + Utilities.base64Encode(
    JSON.stringify([targetMonth, targetYear, groupBy, filterPartners, filterClusters, filterPos, filterLobs, filterProductos])
  ).replace(/[^a-zA-Z0-9]/g, '').substring(0, 240);

  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  // Call the original (renamed) function
  var result = getDashboardData_BASE(targetMonth, targetYear, groupBy, filterPartners, filterClusters, filterPos, filterLobs, filterProductos);

  // Cache the result if small enough and no error
  try {
    var resultStr = JSON.stringify(result);
    if (resultStr.length < 90000 && !result.error) {
      cache.put(cacheKey, resultStr, 1800); // 30 min TTL
    }
  } catch(ex) {}

  return result;
}
