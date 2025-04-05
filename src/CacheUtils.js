// src/CacheUtils.gs
// Provides simple caching utility wrappers for optimizing performance

const CACHE_DURATION = 300; // 5 minutes

function getCachedOrFetch(key, fetchFn) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      Logger.log(`Cache parse error for ${key}: ${err.message}`);
    }
  }

  const freshData = fetchFn();
  try {
    cache.put(key, JSON.stringify(freshData), CACHE_DURATION);
  } catch (err) {
    Logger.log(`Failed to cache ${key}: ${err.message}`);
  }

  return freshData;
}

function clearCache(keys = []) {
  const cache = CacheService.getScriptCache();
  if (keys.length > 0) {
    cache.removeAll(keys);
  } else {
    cache.removeAll([
      'all_users',
      'all_questions',
      'all_audits',
      'pending_audits',
      'all_evaluations',
      'all_disputes'
    ]);
  }
}
