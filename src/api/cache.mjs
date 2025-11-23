import { Cache } from '../utils/common.mjs';
import { logger } from '../utils/logger.mjs';

// Global cache instance
export const apiCache = new Cache();

/**
 * Caching Strategy for OneNote MCP
 * --------------------------------
 * This module implements selective caching for Microsoft Graph API calls to
 * optimize performance while maintaining data consistency.
 *
 * WHAT IS CACHED:
 * - Notebooks list: Rarely changes, safe to cache
 * - Sections list: Changes infrequently, benefits from caching
 *
 * WHAT IS NOT CACHED:
 * - Pages: Dynamic content with frequent mutations (create, update, append, etc.)
 *   Caching pages would risk serving stale content and require complex invalidation
 *   logic across many mutation operations.
 * - Section groups: Rarely used, not worth the complexity
 * - Individual resources: Single-item fetches are fast enough without caching
 * - Search results: Dynamic and query-specific, unsuitable for caching
 *
 * CACHE INVALIDATION:
 * - createNotebook: Invalidates CacheKeys.notebooks()
 * - createSection: Invalidates CacheKeys.sections() and CacheKeys.sections(notebookId)
 * - Page operations: No cache to invalidate (intentional)
 * - Section group operations: No cache to invalidate (intentional)
 *
 * This intentional design keeps the caching layer simple, predictable, and maintainable
 * while providing performance benefits where they matter most (list operations).
 */

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
    logger.debug(`✨ Cache hit: ${cacheKey}`);
    return cached;
  }

  // Cache miss - make the API call
  logger.debug(`🔄 Cache miss: ${cacheKey}`);
  const result = await apiCall();

  // Store in cache
  apiCache.set(cacheKey, result, ttl);

  return result;
}
