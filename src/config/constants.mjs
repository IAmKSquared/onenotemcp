/**
 * Configuration constants for OneNote MCP Server
 * Centralized location for all magic numbers and configuration values
 */

// ===== HTTP Status Codes =====

/**
 * HTTP status codes that should not trigger automatic retry attempts.
 * These typically indicate client errors that won't be resolved by retrying.
 */
export const HTTP_STATUS = {
  // Client errors (4xx) - non-retryable
  BAD_REQUEST: 400,
  UNAUTHORIZED: 401,
  FORBIDDEN: 403,
  NOT_FOUND: 404,
  CONFLICT: 409,
  RATE_LIMIT: 429,

  // Server errors (5xx) - retryable
  SERVER_ERROR_MIN: 500,
  SERVER_ERROR_MAX: 600,
};

/**
 * HTTP status codes that should never trigger retry attempts.
 */
export const NON_RETRYABLE_STATUS_CODES = [
  HTTP_STATUS.BAD_REQUEST,
  HTTP_STATUS.UNAUTHORIZED,
  HTTP_STATUS.FORBIDDEN,
  HTTP_STATUS.NOT_FOUND,
  HTTP_STATUS.CONFLICT,
];

// ===== Retry Configuration =====

/**
 * Retry and backoff configuration for transient API failures.
 */
export const RETRY_CONFIG = {
  /** Maximum number of retry attempts for failed operations */
  MAX_RETRIES: 3,

  /** Base delay in milliseconds for exponential backoff */
  BASE_DELAY_MS: 1000,

  /** Exponential backoff multiplier (delay = base * multiplier^attempt) */
  BACKOFF_MULTIPLIER: 2,
};

// ===== Timeout Configuration =====

/**
 * Timeout values for various operations.
 */
export const TIMEOUTS = {
  /** Delay to allow device code callback to complete (milliseconds) */
  DEVICE_CODE_CALLBACK_MS: 2000,
};

// ===== Cache Configuration =====

/**
 * Cache time-to-live (TTL) settings.
 */
export const CACHE_CONFIG = {
  /** Default cache TTL in milliseconds (5 minutes) */
  DEFAULT_TTL_MS: 5 * 60 * 1000,
};

// ===== Display and Pagination =====

/**
 * Display and pagination limits for result sets.
 */
export const DISPLAY_LIMITS = {
  /** Maximum number of items to display in lists before truncating */
  MAX_DISPLAY_ITEMS: 10,

  /** Microsoft Graph API result limit (used for warning messages) */
  API_RESULT_LIMIT: 50,

  /** Maximum length for text summaries (characters) */
  SUMMARY_MAX_LENGTH: 300,
};

// ===== Validation Rules =====

/**
 * Validation constraints for user inputs and IDs.
 */
export const VALIDATION = {
  /** Minimum allowed length for Microsoft Graph IDs (characters) */
  ID_MIN_LENGTH: 10,

  /** Maximum allowed length for Microsoft Graph IDs (characters) */
  ID_MAX_LENGTH: 200,

  /** Minimum number of rows required in CSV table data (header + data) */
  CSV_MIN_ROWS: 2,
};

// ===== Time Conversion =====

/**
 * Time conversion constants.
 */
export const TIME_CONVERSION = {
  /** Milliseconds per hour (1000 * 60 * 60) */
  MS_PER_HOUR: 3600000,

  /** Milliseconds per minute (1000 * 60) */
  MS_PER_MINUTE: 60000,

  /** Milliseconds per second */
  MS_PER_SECOND: 1000,
};

// ===== Encryption Configuration =====

/**
 * Encryption settings for token storage.
 */
export const ENCRYPTION = {
  /** Algorithm used for token encryption */
  ALGORITHM: 'aes-256-cbc',

  /** Key length in bytes (32 bytes = 256 bits) */
  KEY_LENGTH_BYTES: 32,

  /** Initialization vector length in bytes */
  IV_LENGTH_BYTES: 16,
};
