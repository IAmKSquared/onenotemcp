import pino from 'pino';
import pinoPretty from 'pino-pretty';

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

// Configure pino to output to stderr for MCP compatibility
// MCP servers must keep stdout clean for JSON-RPC protocol messages
// Use pino-pretty stream that writes directly to stderr
const prettyStream = pinoPretty({
  colorize: true,
  translateTime: 'HH:MM:ss',
  ignore: 'pid,hostname',
  destination: 2, // stderr file descriptor
});

const logger = pino(
  {
    level: process.env.LOG_LEVEL || 'info',
  },
  process.env.NODE_ENV !== 'production' ? prettyStream : process.stderr
);

export { logger };
