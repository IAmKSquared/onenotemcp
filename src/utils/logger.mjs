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

// Configure pino to output to stderr for MCP compatibility.
// MCP servers must keep stdout clean for JSON-RPC protocol messages.
// In development, lazily import pino-pretty (a devDependency) so that
// production installs (npm install --omit=dev) don't crash on a missing
// module; in production we log directly to stderr.
const stream =
  process.env.NODE_ENV !== 'production'
    ? (await import('pino-pretty')).default({
        colorize: true,
        translateTime: 'HH:MM:ss',
        ignore: 'pid,hostname',
        destination: 2, // stderr file descriptor
      })
    : process.stderr;

const logger = pino({ level: process.env.LOG_LEVEL || 'info' }, stream);

export { logger };
