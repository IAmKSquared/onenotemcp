/**
 * Utility functions for OneNote MCP Server
 * Extracted for testing and reusability
 */

import { JSDOM } from 'jsdom';
import { VALIDATION, CACHE_CONFIG, DISPLAY_LIMITS } from '../config/constants.mjs';
import { logger } from './logger.mjs';

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
 * Microsoft Graph IDs are typically 20-100 characters, but we allow flexibility.
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

  // Microsoft Graph IDs are typically 20-100 chars, allow flexibility
  if (trimmedId.length < VALIDATION.ID_MIN_LENGTH || trimmedId.length > VALIDATION.ID_MAX_LENGTH) {
    throw new Error(
      `Invalid ${type} ID: ID length out of expected range (${VALIDATION.ID_MIN_LENGTH}-${VALIDATION.ID_MAX_LENGTH} characters).`
    );
  }

  // Comprehensive security checks combining both implementations
  const dangerousPatterns = [
    /<script/i, // Script tags
    /javascript:/i, // JavaScript protocol
    /data:/i, // Data URI
    /vbscript:/i, // VBScript protocol
    /\.\./, // Path traversal
    /[<>"'`]/, // HTML/quote characters
    /[\r\n]/, // Newlines
    /\s{2,}/, // Multiple consecutive spaces
    /[;|&$\\]/, // Command injection characters
  ];

  for (const pattern of dangerousPatterns) {
    if (pattern.test(trimmedId)) {
      throw new Error(`Invalid ${type} ID: Contains potentially dangerous characters or patterns.`);
    }
  }

  return trimmedId;
}

/**
 * Extracts a short summary from HTML content
 * @param {string} html - The HTML content string
 * @param {number} [maxLength] - The maximum length of the summary
 * @returns {string} A text summary
 */
export function extractTextSummary(html, maxLength = DISPLAY_LIMITS.SUMMARY_MAX_LENGTH) {
  try {
    if (!html) return 'No content to summarize.';
    const dom = new JSDOM(html);
    const document = dom.window.document;
    const bodyText = document.body?.textContent?.trim().replace(/\s+/g, ' ') || '';
    if (!bodyText) return 'No text content found in HTML body.';
    const summary = bodyText.substring(0, maxLength);
    return summary.length < bodyText.length ? `${summary}...` : summary;
  } catch (error) {
    logger.error({ err: error }, 'Error extracting text summary');
    return 'Could not extract text summary.';
  }
}

/**
 * Extracts readable plain text from HTML content with formatted structure.
 * Preserves headings, paragraphs, lists, and tables with visual formatting.
 * Removes scripts and styles, formats headings with underlines, lists with bullets/numbers.
 * @param {string} html - The HTML content string
 * @returns {string} The extracted readable text with formatting
 */
export function extractReadableText(html) {
  try {
    if (!html) return '';
    const dom = new JSDOM(html);
    const document = dom.window.document;

    document.querySelectorAll('script, style').forEach((element) => element.remove());

    let text = '';
    document.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach((heading) => {
      const headingText = heading.textContent?.trim();
      if (headingText) text += `\n${headingText}\n${'-'.repeat(headingText.length)}\n`;
    });
    document.querySelectorAll('p').forEach((paragraph) => {
      const content = paragraph.textContent?.trim();
      if (content) text += `${content}\n\n`;
    });
    document.querySelectorAll('ul, ol').forEach((list) => {
      text += '\n';
      list.querySelectorAll('li').forEach((item, index) => {
        const content = item.textContent?.trim();
        if (content) text += `${list.tagName === 'OL' ? index + 1 + '.' : '-'} ${content}\n`;
      });
      text += '\n';
    });
    document.querySelectorAll('table').forEach((table) => {
      text += '\n📊 Table content:\n';
      table.querySelectorAll('tr').forEach((row) => {
        const cells = Array.from(row.querySelectorAll('td, th'))
          .map((cell) => cell.textContent?.trim())
          .join(' | ');
        if (cells.trim()) text += `${cells}\n`;
      });
      text += '\n';
    });

    if (!text.trim() && document.body) {
      text = document.body.textContent?.trim().replace(/\s+/g, ' ') || '';
    }
    return text.trim();
  } catch (error) {
    logger.error({ err: error }, 'Error extracting readable text');
    return 'Error: Could not extract readable text from HTML content.';
  }
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
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;') // Basic HTML escaping first
    .replace(/```([\s\S]*?)```/g, (match, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/__(.*?)__/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/_(.*?)_/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (match, linkText, url) => {
      const safeUrl = sanitizeUrl(url);
      return `<a href="${safeUrl}">${linkText}</a>`;
    })
    .replace(/^---+$/gm, '<hr>')
    .replace(/^&gt; (.+)$/gm, '<blockquote>$1</blockquote>')
    .replace(/^[*\-+] (.+)$/gm, '<li>$1</li>')
    .replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');

  html = html
    .split('\n')
    .map((line) => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      if (
        /^<(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)/.test(trimmed) ||
        /^<\/(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)>/.test(trimmed)
      ) {
        return trimmed; // Already an HTML element we processed or a closing tag
      }
      return `<p>${trimmed}</p>`;
    })
    .filter((line) => line)
    .join('\n');

  html = html.replace(/(<li>.*?<\/li>(?:\s*<li>.*?<\/li>)*)/gs, '<ul>$1</ul>');
  html = html.replace(
    /(<blockquote>.*?<\/blockquote>(?:\s*<blockquote>.*?<\/blockquote>)*)/gs,
    '<blockquote>$1</blockquote>'
  );

  return html;
}

/**
 * Cache implementation with TTL support
 */
export class Cache {
  /**
   * Creates a new Cache instance with default TTL
   */
  constructor() {
    this.cache = new Map();
    this.defaultTTL = CACHE_CONFIG.DEFAULT_TTL_MS;
  }

  /**
   * Retrieves a value from the cache if it exists and hasn't expired
   * @param {string} key - The cache key to retrieve
   * @returns {*} The cached value, or undefined if not found or expired
   */
  get(key) {
    const entry = this.cache.get(key);
    if (!entry) return undefined;

    if (Date.now() > entry.expiry) {
      this.cache.delete(key);
      return undefined;
    }

    return entry.value;
  }

  /**
   * Stores a value in the cache with an optional TTL
   * @param {string} key - The cache key
   * @param {*} value - The value to cache
   * @param {number} [ttl] - Time to live in milliseconds (defaults to configured TTL)
   */
  set(key, value, ttl) {
    const expiry = Date.now() + (ttl || this.defaultTTL);
    this.cache.set(key, { value, expiry });
  }

  /**
   * Invalidates cache entries by key, wildcard pattern, or RegExp
   * @param {string|RegExp} keyOrPattern - Exact key, wildcard pattern (e.g., "user:*"), or RegExp
   */
  invalidate(keyOrPattern) {
    if (keyOrPattern instanceof RegExp) {
      // Support RegExp patterns
      for (const key of this.cache.keys()) {
        if (keyOrPattern.test(key)) {
          this.cache.delete(key);
        }
      }
    } else if (typeof keyOrPattern === 'string') {
      if (keyOrPattern.includes('*')) {
        // Support wildcard strings like "user:*"
        const pattern = new RegExp('^' + keyOrPattern.replace(/\*/g, '.*') + '$');
        for (const key of this.cache.keys()) {
          if (pattern.test(key)) {
            this.cache.delete(key);
          }
        }
      } else {
        // Exact key match
        this.cache.delete(keyOrPattern);
      }
    }
  }

  /**
   * Clears all entries from the cache
   */
  clear() {
    this.cache.clear();
  }
}

/**
 * Validates and sanitizes CSV data to prevent injection attacks and ensure data integrity.
 * @param {string} csvString - The CSV data as a string (comma-separated values, newline-separated rows).
 * @returns {Array<Array<string>>} Array of rows, each row is an array of sanitized cell values.
 * @throws {Error} If validation fails (too few rows, inconsistent columns, etc.)
 */
export function validateCsvData(csvString) {
  if (!csvString || typeof csvString !== 'string') {
    throw new Error('CSV data must be a non-empty string.');
  }

  // Split into rows and filter out empty lines
  const rows = csvString
    .trim()
    .split('\n')
    .map((row) => row.trim())
    .filter((row) => row.length > 0);

  // Validate minimum rows (header + at least one data row)
  if (rows.length < VALIDATION.CSV_MIN_ROWS) {
    throw new Error('Table data must have at least a header row and one data row.');
  }

  // Parse each row into cells (simple comma-split for now)
  const parsedRows = rows.map((row) => row.split(',').map((cell) => cell.trim()));

  // Validate consistent column count
  const columnCount = parsedRows[0].length;
  if (columnCount === 0) {
    throw new Error('Header row cannot be empty.');
  }

  for (let i = 0; i < parsedRows.length; i++) {
    if (parsedRows[i].length !== columnCount) {
      throw new Error(
        `Row ${i + 1} has ${parsedRows[i].length} columns, but header has ${columnCount} columns. All rows must have the same number of columns.`
      );
    }
  }

  // Sanitize cells to prevent CSV injection attacks
  // CSV injection occurs when cells start with =, +, -, @, which Excel/Sheets interpret as formulas
  const sanitizedRows = parsedRows.map((row) => row.map((cell) => sanitizeCsvCell(cell)));

  return sanitizedRows;
}

/**
 * Sanitizes a single CSV cell to prevent injection attacks.
 * @param {string} cell - The cell value to sanitize.
 * @returns {string} The sanitized cell value.
 */
function sanitizeCsvCell(cell) {
  if (!cell || typeof cell !== 'string') {
    return '';
  }

  const trimmed = cell.trim();

  // Check for dangerous formula prefixes
  const dangerousPrefixes = ['=', '+', '-', '@', '\t', '\r'];
  if (dangerousPrefixes.some((prefix) => trimmed.startsWith(prefix))) {
    // Prepend single quote to neutralize formula interpretation
    return `'${trimmed}`;
  }

  return trimmed;
}
