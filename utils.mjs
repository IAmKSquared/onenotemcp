/**
 * Utility functions for OneNote MCP Server
 * Extracted for testing and reusability
 */

import { JSDOM } from 'jsdom';

/**
 * Escapes single quotes in OData filter strings to prevent injection
 * @param {string} str - The string to escape
 * @returns {string} The escaped string
 */
export function escapeODataString(str) {
  if (!str) return '';
  return str.replace(/'/g, "''");
}

/**
 * Validates URL safety and prevents javascript: protocol injection
 * @param {string} url - The URL to sanitize
 * @returns {string} Safe URL or '#' if unsafe
 */
export function sanitizeUrl(url) {
  if (!url) return '#';
  const safeProtocols = /^(https?:\/\/|mailto:)/i;
  const hasProtocol = /^[a-z][a-z0-9+.-]*:/i.test(url.trim());
  if (hasProtocol && !safeProtocols.test(url.trim())) {
    return '#';
  }
  return url.trim();
}

/**
 * Validates resource IDs to prevent injection attacks
 * @param {string} id - The ID to validate
 * @param {string} type - The resource type (for error messages)
 * @throws {Error} If ID is invalid
 */
export function validateId(id, type = 'resource') {
  if (!id || typeof id !== 'string') {
    throw new Error(`Invalid ${type} ID: ID must be a non-empty string.`);
  }

  const trimmedId = id.trim();

  if (trimmedId.length === 0) {
    throw new Error(`Invalid ${type} ID: ID cannot be empty or whitespace.`);
  }

  if (trimmedId.length > 200) {
    throw new Error(`Invalid ${type} ID: ID is too long (max 200 characters).`);
  }

  const dangerousPatterns = [
    /[<>'"]/,           // HTML/JS injection characters
    /javascript:/i,     // JavaScript protocol
    /data:/i,          // Data URI
    /vbscript:/i,      // VBScript protocol
    /\.\./,            // Path traversal
    /[;|&$`\\]/        // Command injection characters
  ];

  for (const pattern of dangerousPatterns) {
    if (pattern.test(trimmedId)) {
      throw new Error(`Invalid ${type} ID: Contains potentially dangerous characters or patterns.`);
    }
  }

  return trimmedId;
}

/**
 * Converts HTML content to readable plain text
 * @param {string} html - The HTML string to convert
 * @returns {string} Plain text content
 */
export function htmlToText(html) {
  if (!html) return '';

  const dom = new JSDOM(html);
  const doc = dom.window.document;

  // Remove script and style elements
  doc.querySelectorAll('script, style').forEach(el => el.remove());

  let text = doc.body.textContent || '';

  // Clean up whitespace
  text = text.replace(/\n\s*\n\s*\n/g, '\n\n'); // Multiple blank lines to double
  text = text.replace(/[ \t]+/g, ' ');           // Multiple spaces to single
  text = text.trim();

  return text;
}

/**
 * Converts plain text or markdown to HTML for OneNote
 * @param {string} text - The text to convert
 * @returns {string} HTML string
 */
export function textToHtml(text) {
  if (!text) return '<p></p>';

  // Escape HTML special characters for security
  let escaped = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');

  // Convert newlines to <br> and wrap in paragraphs
  const paragraphs = escaped.split('\n\n').map(para => {
    const withBreaks = para.split('\n').join('<br>');
    return `<p>${withBreaks}</p>`;
  });

  return paragraphs.join('');
}

/**
 * Cache implementation with TTL support
 */
export class Cache {
  constructor() {
    this.cache = new Map();
    this.defaultTTL = 5 * 60 * 1000; // 5 minutes default
  }

  get(key) {
    const entry = this.cache.get(key);
    if (!entry) return undefined;

    if (Date.now() > entry.expiry) {
      this.cache.delete(key);
      return undefined;
    }

    return entry.value;
  }

  set(key, value, ttl) {
    const expiry = Date.now() + (ttl || this.defaultTTL);
    this.cache.set(key, { value, expiry });
  }

  invalidate(keyOrPattern) {
    if (typeof keyOrPattern === 'string') {
      if (keyOrPattern.includes('*')) {
        const pattern = new RegExp('^' + keyOrPattern.replace(/\*/g, '.*') + '$');
        for (const key of this.cache.keys()) {
          if (pattern.test(key)) {
            this.cache.delete(key);
          }
        }
      } else {
        this.cache.delete(keyOrPattern);
      }
    }
  }

  clear() {
    this.cache.clear();
  }
}
