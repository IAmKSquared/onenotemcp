import pino from 'pino';

/**
 * Centralized logging utility using Pino.
 * Provides structured, high-performance logging with minimal overhead.
 *
 * Log levels (in order of severity):
 * - trace: Very detailed debugging information
 * - debug: Debugging information
 * - info: General informational messages
 * - warn: Warning messages for potentially problematic situations
 * - error: Error messages for failures and exceptions
 * - fatal: Critical errors that may cause application termination
 *
 * Usage:
 *   import { logger } from './utils/logger.mjs';
 *   logger.info('Server started');
 *   logger.warn('Token expiring soon');
 *   logger.error({ err: error }, 'Authentication failed');
 */

// Configure pino with human-readable output during development
// In production, use JSON output for better parsing/analysis
const logger = pino({
  level: process.env.LOG_LEVEL || 'info',
  transport:
    process.env.NODE_ENV !== 'production'
      ? {
          target: 'pino-pretty',
          options: {
            colorize: true,
            translateTime: 'HH:MM:ss',
            ignore: 'pid,hostname',
          },
        }
      : undefined,
});

export { logger };
