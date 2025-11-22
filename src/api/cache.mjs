import { Cache } from '../utils/common.mjs';

// Global cache instance
export const apiCache = new Cache();

/**
 * Cache key builder for consistent cache key construction.
 * Provides centralized cache key generation to avoid manual string construction.
 */
export const CacheKeys = {
  /**
   * Generates cache key for notebooks list.
   * @returns {string} Cache key for notebooks list.
   */
  notebooks: () => 'notebooks:list',

  /**
   * Generates cache key for sections list.
   * @param {string} [notebookId] - Optional notebook ID to scope sections.
   * @param {string} [sectionGroupId] - Optional section group ID to scope sections.
   * @returns {string} Cache key for sections list.
   */
  sections: (notebookId, sectionGroupId) => {
    if (notebookId) return `sections:list:notebook:${notebookId}`;
    if (sectionGroupId) return `sections:list:group:${sectionGroupId}`;
    return 'sections:list';
  },
};

/**
 * Wraps an API call with caching logic.
 * @param {string} cacheKey - The key to use for caching.
 * @param {Function} apiCall - The async function that makes the API call.
 * @param {number} [ttl] - Optional TTL override.
 * @returns {Promise<any>} The API response (cached or fresh).
 */
export async function cachedApiCall(cacheKey, apiCall, ttl) {
  // Check cache first
  const cached = apiCache.get(cacheKey);
  if (cached !== undefined) {
    console.error(`✨ Cache hit: ${cacheKey}`);
    return cached;
  }

  // Cache miss - make the API call
  console.error(`🔄 Cache miss: ${cacheKey}`);
  const result = await apiCall();

  // Store in cache
  apiCache.set(cacheKey, result, ttl);

  return result;
}
