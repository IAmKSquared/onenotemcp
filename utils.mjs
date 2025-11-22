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
 * Validates resource IDs to prevent injection attacks.
 * Microsoft Graph IDs are typically 20-100 characters, but we allow 10-200 for flexibility.
 * @param {string} id - The ID to validate
 * @param {string} type - The resource type (for error messages)
 * @throws {Error} If ID is invalid
 * @returns {string} The trimmed, validated ID
 */
export function validateId(id, type = 'resource') {
  if (!id || typeof id !== 'string') {
    throw new Error(`Invalid ${type} ID: ID must be a non-empty string.`);
  }

  const trimmedId = id.trim();

  if (trimmedId.length === 0) {
    throw new Error(`Invalid ${type} ID: ID cannot be empty or whitespace.`);
  }

  // Microsoft Graph IDs are typically 20-100 chars, allow 10-200 for flexibility
  if (trimmedId.length < 10 || trimmedId.length > 200) {
    throw new Error(`Invalid ${type} ID: ID length out of expected range (10-200 characters).`);
  }

  // Comprehensive security checks combining both implementations
  const dangerousPatterns = [
    /<script/i,         // Script tags
    /javascript:/i,     // JavaScript protocol
    /data:/i,          // Data URI
    /vbscript:/i,      // VBScript protocol
    /\.\./,            // Path traversal
    /[<>"'`]/,         // HTML/quote characters
    /[\r\n]/,          // Newlines
    /\s{2,}/,          // Multiple consecutive spaces
    /[;|&$\\]/         // Command injection characters
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
 * Converts plain text (with simple markdown) to HTML.
 * Supports markdown features: code blocks, inline code, headers, bold, italic,
 * links, horizontal rules, blockquotes, and lists.
 * @param {string} text - The plain text to convert
 * @returns {string} The HTML representation
 */
export function textToHtml(text) {
  if (!text) return '';

  // Security: Never bypass escaping, even if input looks like HTML
  // All user input must be escaped to prevent XSS

  let html = String(text) // Ensure text is a string
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') // Basic HTML escaping first
    .replace(/```([\s\S]*?)```/g, (match, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/__(.*?)__/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>').replace(/_(.*?)_/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (match, linkText, url) => {
      const safeUrl = sanitizeUrl(url);
      return `<a href="${safeUrl}">${linkText}</a>`;
    })
    .replace(/^---+$/gm, '<hr>')
    .replace(/^&gt; (.+)$/gm, '<blockquote>$1</blockquote>')
    .replace(/^[\*\-\+] (.+)$/gm, '<li>$1</li>')
    .replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');

  html = html.split('\n').map(line => {
    const trimmed = line.trim();
    if (!trimmed) return '';
    if (/^<(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)/.test(trimmed) || /^<\/(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)>/.test(trimmed)) {
      return trimmed; // Already an HTML element we processed or a closing tag
    }
    return `<p>${trimmed}</p>`;
  }).filter(line => line).join('\n');

  html = html.replace(/(<li>.*?<\/li>(?:\s*<li>.*?<\/li>)*)/gs, '<ul>$1</ul>');
  html = html.replace(/(<blockquote>.*?<\/blockquote>(?:\s*<blockquote>.*?<\/blockquote>)*)/gs, '<blockquote>$1</blockquote>');

  return html;
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
