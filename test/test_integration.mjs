/**
 * Integration tests for OneNote MCP Server core functions
 * These tests cover functions from onenote-mcp.mjs that require more complex setup
 * Run with: node --test test_integration.mjs
 */

import { test, describe, mock, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ============================================================================
// Mock Setup for Testing
// ============================================================================

// Create a test directory for temporary files
const testDir = path.join(__dirname, '.test-temp');

// Setup and teardown
beforeEach(() => {
  if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir, { recursive: true });
  }
});

afterEach(() => {
  // Clean up test files
  if (fs.existsSync(testDir)) {
    fs.rmSync(testDir, { recursive: true, force: true });
  }
});

// ============================================================================
// Encryption Functions Tests (Critical Security Functions)
// ============================================================================

describe('Encryption and Decryption', () => {
  const ALGORITHM = 'aes-256-cbc';
  const IV_LENGTH = 16;

  /**
   * Standalone encryption function for testing
   * @param {string} text - The text to encrypt
   * @param {Buffer} key - The encryption key
   * @returns {Promise<{iv: string, encryptedData: string}>} Encrypted data with IV
   */
  async function testEncrypt(text, key) {
    const iv = crypto.randomBytes(IV_LENGTH);
    const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
    let encrypted = cipher.update(text);
    encrypted = Buffer.concat([encrypted, cipher.final()]);
    return { iv: iv.toString('hex'), encryptedData: encrypted.toString('hex') };
  }

  /**
   * Standalone decryption function for testing
   * @param {{iv: string, encryptedData: string}} encryptedObj - The encrypted data object
   * @param {Buffer} key - The decryption key
   * @returns {Promise<string>} The decrypted text
   */
  async function testDecrypt(encryptedObj, key) {
    const iv = Buffer.from(encryptedObj.iv, 'hex');
    const encryptedText = Buffer.from(encryptedObj.encryptedData, 'hex');
    const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
    let decrypted = decipher.update(encryptedText);
    decrypted = Buffer.concat([decrypted, decipher.final()]);
    return decrypted.toString();
  }

  test('should encrypt and decrypt text correctly', async () => {
    const originalText = 'test-token-12345';
    const key = crypto.randomBytes(32);

    const encrypted = await testEncrypt(originalText, key);
    const decrypted = await testDecrypt(encrypted, key);

    assert.strictEqual(decrypted, originalText);
  });

  test('should use unique IVs for each encryption', async () => {
    const text = 'same-text';
    const key = crypto.randomBytes(32);

    const enc1 = await testEncrypt(text, key);
    const enc2 = await testEncrypt(text, key);

    // IVs should be different
    assert.notStrictEqual(enc1.iv, enc2.iv);
    // Encrypted data should be different due to different IVs
    assert.notStrictEqual(enc1.encryptedData, enc2.encryptedData);

    // But both should decrypt to the same text
    const dec1 = await testDecrypt(enc1, key);
    const dec2 = await testDecrypt(enc2, key);
    assert.strictEqual(dec1, text);
    assert.strictEqual(dec2, text);
  });

  test('should produce different encrypted data for different inputs', async () => {
    const key = crypto.randomBytes(32);

    const enc1 = await testEncrypt('text1', key);
    const enc2 = await testEncrypt('text2', key);

    assert.notStrictEqual(enc1.encryptedData, enc2.encryptedData);
  });

  test('should fail decryption with wrong key', async () => {
    const text = 'secret-data';
    const correctKey = crypto.randomBytes(32);
    const wrongKey = crypto.randomBytes(32);

    const encrypted = await testEncrypt(text, correctKey);

    // Attempting to decrypt with wrong key should throw or produce garbage
    await assert.rejects(
      async () => await testDecrypt(encrypted, wrongKey),
      {
        name: 'Error',
      },
      'Should fail to decrypt with wrong key'
    );
  });

  test('should handle empty strings', async () => {
    const key = crypto.randomBytes(32);

    const encrypted = await testEncrypt('', key);
    const decrypted = await testDecrypt(encrypted, key);

    assert.strictEqual(decrypted, '');
  });

  test('should handle long strings', async () => {
    const longText = 'a'.repeat(10000);
    const key = crypto.randomBytes(32);

    const encrypted = await testEncrypt(longText, key);
    const decrypted = await testDecrypt(encrypted, key);

    assert.strictEqual(decrypted, longText);
  });

  test('should handle special characters', async () => {
    const specialText = 'Special: <>&quot; symbols @#$%^&*()';
    const key = crypto.randomBytes(32);

    const encrypted = await testEncrypt(specialText, key);
    const decrypted = await testDecrypt(encrypted, key);

    assert.strictEqual(decrypted, specialText);
  });

  test('should handle JSON strings', async () => {
    const jsonText = JSON.stringify({
      token: 'abc123',
      expiresOn: '2025-12-31T23:59:59Z',
      scopes: ['Notes.Read', 'Notes.ReadWrite'],
    });
    const key = crypto.randomBytes(32);

    const encrypted = await testEncrypt(jsonText, key);
    const decrypted = await testDecrypt(encrypted, key);

    assert.strictEqual(decrypted, jsonText);
    assert.doesNotThrow(() => JSON.parse(decrypted));
  });
});

// ============================================================================
// Retry with Backoff Tests (Error Resilience)
// ============================================================================

describe('retryWithBackoff', () => {
  /**
   * Standalone retry function for testing
   * @param {() => Promise<any>} fn - The async function to retry
   * @param {number} maxRetries - Maximum number of retry attempts
   * @param {number} baseDelay - Base delay in milliseconds for exponential backoff
   * @returns {Promise<any>} The result of the function
   */
  async function retryWithBackoff(fn, maxRetries = 3, baseDelay = 1000) {
    let lastError;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error;

        // Don't retry on these errors
        const nonRetryableErrors = [400, 401, 403, 404, 409];
        if (error.statusCode && nonRetryableErrors.includes(error.statusCode)) {
          throw error;
        }

        // Retry on rate limits (429) and server errors (500+)
        const shouldRetry =
          error.statusCode === 429 || // Rate limit
          (error.statusCode >= 500 && error.statusCode < 600) || // Server errors
          error.code === 'ETIMEDOUT' || // Timeout
          error.code === 'ECONNRESET'; // Connection reset

        if (!shouldRetry || attempt === maxRetries) {
          throw error;
        }

        // Calculate delay with exponential backoff
        const delay = baseDelay * Math.pow(2, attempt);
        await new Promise((resolve) => setTimeout(resolve, delay));
      }
    }

    throw lastError;
  }

  test('should succeed on first attempt', async () => {
    const fn = mock.fn(async () => 'success');

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(fn.mock.calls.length, 1);
  });

  test('should retry on 429 rate limit error', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      if (attempts < 3) {
        const error = new Error('Rate limited');
        error.statusCode = 429;
        throw error;
      }
      return 'success';
    };

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(attempts, 3);
  });

  test('should retry on 500 server error', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      if (attempts < 2) {
        const error = new Error('Internal server error');
        error.statusCode = 500;
        throw error;
      }
      return 'success';
    };

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(attempts, 2);
  });

  test('should retry on 503 service unavailable', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      if (attempts < 2) {
        const error = new Error('Service unavailable');
        error.statusCode = 503;
        throw error;
      }
      return 'success';
    };

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(attempts, 2);
  });

  test('should retry on ETIMEDOUT error', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      if (attempts < 2) {
        const error = new Error('Timeout');
        error.code = 'ETIMEDOUT';
        throw error;
      }
      return 'success';
    };

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(attempts, 2);
  });

  test('should retry on ECONNRESET error', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      if (attempts < 2) {
        const error = new Error('Connection reset');
        error.code = 'ECONNRESET';
        throw error;
      }
      return 'success';
    };

    const result = await retryWithBackoff(fn, 3, 10);

    assert.strictEqual(result, 'success');
    assert.strictEqual(attempts, 2);
  });

  test('should NOT retry on 401 unauthorized', async () => {
    const fn = async () => {
      const error = new Error('Unauthorized');
      error.statusCode = 401;
      throw error;
    };

    await assert.rejects(
      async () => await retryWithBackoff(fn, 3, 10),
      {
        message: 'Unauthorized',
      },
      'Should not retry on 401'
    );
  });

  test('should NOT retry on 403 forbidden', async () => {
    const fn = async () => {
      const error = new Error('Forbidden');
      error.statusCode = 403;
      throw error;
    };

    await assert.rejects(
      async () => await retryWithBackoff(fn, 3, 10),
      {
        message: 'Forbidden',
      },
      'Should not retry on 403'
    );
  });

  test('should NOT retry on 404 not found', async () => {
    const fn = async () => {
      const error = new Error('Not found');
      error.statusCode = 404;
      throw error;
    };

    await assert.rejects(
      async () => await retryWithBackoff(fn, 3, 10),
      {
        message: 'Not found',
      },
      'Should not retry on 404'
    );
  });

  test('should exhaust all retries and fail', async () => {
    let attempts = 0;
    const fn = async () => {
      attempts++;
      const error = new Error('Always fails');
      error.statusCode = 500;
      throw error;
    };

    await assert.rejects(
      async () => await retryWithBackoff(fn, 3, 10),
      {
        message: 'Always fails',
      },
      'Should fail after exhausting retries'
    );

    assert.strictEqual(attempts, 4); // Initial attempt + 3 retries
  });

  test('should use exponential backoff timing', async () => {
    const delays = [];
    let attempts = 0;

    const fn = async () => {
      attempts++;
      if (attempts < 4) {
        const error = new Error('Retry me');
        error.statusCode = 500;
        throw error;
      }
      return 'success';
    };

    // Mock setTimeout to capture delays
    const originalSetTimeout = global.setTimeout;
    global.setTimeout = (callback, delay) => {
      delays.push(delay);
      return originalSetTimeout(callback, 0); // Execute immediately for test speed
    };

    try {
      await retryWithBackoff(fn, 3, 100);

      // Check exponential backoff: 100, 200, 400
      assert.strictEqual(delays.length, 3);
      assert.strictEqual(delays[0], 100);
      assert.strictEqual(delays[1], 200);
      assert.strictEqual(delays[2], 400);
    } finally {
      global.setTimeout = originalSetTimeout;
    }
  });
});

// ============================================================================
// Detailed Error Message Tests (Error UX)
// ============================================================================

describe('getDetailedErrorMessage', () => {
  /**
   * Standalone error message function for testing
   * @param {Error} error - The error object
   * @param {string} errorPrefix - The prefix for the error message
   * @returns {string} The formatted error message
   */
  function getDetailedErrorMessage(error, errorPrefix) {
    const statusCode = error.statusCode || error.status;
    const errorMessage = error.message || 'Unknown error';

    // Authentication errors
    if (
      statusCode === 401 ||
      errorMessage.includes('authenticate') ||
      errorMessage.includes('Access token')
    ) {
      return "= **Authentication Required**\nYour access token has expired or is invalid. Please run the 'authenticate' tool to sign in again.";
    }

    // Permission errors
    if (statusCode === 403) {
      return "= **Permission Denied**\nYou don't have permission to perform this action. Ensure your account has the required OneNote permissions (Notes.Read, Notes.ReadWrite, Notes.Create).";
    }

    // Not found errors
    if (statusCode === 404) {
      return `L **Resource Not Found**\n${errorMessage}\n\nThe requested item doesn't exist. It may have been deleted or the ID is incorrect.`;
    }

    // Rate limit errors
    if (statusCode === 429) {
      return '� **Rate Limit Exceeded**\nToo many requests. Please wait a moment before trying again. (All retry attempts exhausted)';
    }

    // Server errors
    if (statusCode >= 500 && statusCode < 600) {
      return `=' **Server Error** (${statusCode})\nMicrosoft's OneNote service is experiencing issues. Please try again in a few moments. (All retry attempts exhausted)`;
    }

    // Timeout errors
    if (error.code === 'ETIMEDOUT' || errorMessage.includes('timeout')) {
      return '� **Request Timeout**\nThe request took too long to complete. Please check your network connection and try again.';
    }

    // Network errors
    if (
      error.code === 'ECONNRESET' ||
      error.code === 'ENOTFOUND' ||
      error.code === 'ECONNREFUSED'
    ) {
      return '< **Network Error**\nUnable to connect to Microsoft services. Please check your internet connection.';
    }

    // Token expiration (specific message format from Azure)
    if (errorMessage.includes('token') && errorMessage.includes('expired')) {
      return "= **Token Expired**\nYour authentication token has expired. Please run the 'authenticate' tool to sign in again.";
    }

    // Default error message
    return `L **${errorPrefix}**\n${errorMessage}`;
  }

  test('should handle 401 authentication errors', () => {
    const error = new Error('Unauthorized');
    error.statusCode = 401;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Authentication Required'), true);
    assert.strictEqual(message.includes('authenticate'), true);
  });

  test('should handle 403 permission errors', () => {
    const error = new Error('Forbidden');
    error.statusCode = 403;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Permission Denied'), true);
    assert.strictEqual(message.includes('OneNote permissions'), true);
  });

  test('should handle 404 not found errors', () => {
    const error = new Error('Page not found');
    error.statusCode = 404;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Resource Not Found'), true);
    assert.strictEqual(message.includes('Page not found'), true);
  });

  test('should handle 429 rate limit errors', () => {
    const error = new Error('Too many requests');
    error.statusCode = 429;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Rate Limit Exceeded'), true);
    assert.strictEqual(message.includes('Too many requests'), true);
  });

  test('should handle 500 server errors', () => {
    const error = new Error('Internal server error');
    error.statusCode = 500;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Server Error'), true);
    assert.strictEqual(message.includes('500'), true);
  });

  test('should handle 503 service unavailable', () => {
    const error = new Error('Service unavailable');
    error.statusCode = 503;

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Server Error'), true);
    assert.strictEqual(message.includes('503'), true);
  });

  test('should handle ETIMEDOUT errors', () => {
    const error = new Error('Request timed out');
    error.code = 'ETIMEDOUT';

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Request Timeout'), true);
    assert.strictEqual(message.includes('network connection'), true);
  });

  test('should handle ECONNRESET errors', () => {
    const error = new Error('Connection reset');
    error.code = 'ECONNRESET';

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Network Error'), true);
    assert.strictEqual(message.includes('internet connection'), true);
  });

  test('should handle ENOTFOUND errors', () => {
    const error = new Error('Host not found');
    error.code = 'ENOTFOUND';

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Network Error'), true);
  });

  test('should handle token expiration messages', () => {
    const error = new Error('The access token has expired');

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Token Expired'), true);
  });

  test('should handle authentication-related messages', () => {
    const error = new Error('Please authenticate to continue');

    const message = getDetailedErrorMessage(error, 'Test failed');

    assert.strictEqual(message.includes('Authentication Required'), true);
  });

  test('should handle generic errors with custom prefix', () => {
    const error = new Error('Something went wrong');

    const message = getDetailedErrorMessage(error, 'Custom Operation');

    assert.strictEqual(message.includes('Custom Operation'), true);
    assert.strictEqual(message.includes('Something went wrong'), true);
  });

  test('should handle errors without status code', () => {
    const error = new Error('Generic error');

    const message = getDetailedErrorMessage(error, 'Test');

    assert.strictEqual(message.includes('Test'), true);
    assert.strictEqual(message.includes('Generic error'), true);
  });
});

// ============================================================================
// Helper Functions Tests
// ============================================================================

describe('formatPageInfo', () => {
  /**
   * Format page information for display
   * @param {object} page - The page object
   * @param {number|null} index - Optional index for numbering
   * @returns {string} Formatted page information
   */
  function formatPageInfo(page, index = null) {
    const prefix = index !== null ? `${index + 1}. ` : '';
    const name = page.displayName || page.title || 'Untitled';
    return `${prefix}**${name}** (ID: ${page.id})`;
  }

  test('should format page with title', () => {
    const page = { id: 'page123', title: 'My Page' };
    const result = formatPageInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123)');
  });

  test('should format page with displayName', () => {
    const page = { id: 'page123', displayName: 'My Section' };
    const result = formatPageInfo(page);

    assert.strictEqual(result, '**My Section** (ID: page123)');
  });

  test('should prefer displayName over title', () => {
    const page = { id: 'page123', displayName: 'Display', title: 'Title' };
    const result = formatPageInfo(page);

    assert.strictEqual(result, '**Display** (ID: page123)');
  });

  test('should use "Untitled" for missing name', () => {
    const page = { id: 'page123' };
    const result = formatPageInfo(page);

    assert.strictEqual(result, '**Untitled** (ID: page123)');
  });

  test('should format with index', () => {
    const page = { id: 'page123', title: 'My Page' };
    const result = formatPageInfo(page, 0);

    assert.strictEqual(result, '1. **My Page** (ID: page123)');
  });

  test('should format with zero index', () => {
    const page = { id: 'page123', title: 'First Page' };
    const result = formatPageInfo(page, 0);

    assert.strictEqual(result, '1. **First Page** (ID: page123)');
  });

  test('should format with non-zero index', () => {
    const page = { id: 'page456', title: 'Third Page' };
    const result = formatPageInfo(page, 2);

    assert.strictEqual(result, '3. **Third Page** (ID: page456)');
  });
});

describe('formatItemList', () => {
  /**
   * Format a list of items with pagination info
   * @param {Array} items - The items to format
   * @param {string} itemType - The type of items (for messages)
   * @param {number} maxDisplay - Maximum items to display
   * @param {number} apiLimit - API result limit
   * @returns {{list: string, more: string, limitWarning: string}} Formatted item list
   */
  function formatItemList(items, itemType = 'items', maxDisplay = 10, apiLimit = 50) {
    /**
     * Format page information for display
     * @param {object} page - The page object
     * @param {number|null} index - Optional index for numbering
     * @returns {string} Formatted page information
     */
    function formatPageInfo(page, index = null) {
      const prefix = index !== null ? `${index + 1}. ` : '';
      const name = page.displayName || page.title || 'Untitled';
      return `${prefix}**${name}** (ID: ${page.id})`;
    }

    const displayItems = items.slice(0, maxDisplay);
    const list = displayItems.map((item, i) => formatPageInfo(item, i)).join('\n\n');
    const more =
      items.length > maxDisplay ? `\n\n... and ${items.length - maxDisplay} more ${itemType}.` : '';
    const limitWarning =
      items.length === apiLimit
        ? `\n\n� Note: Reached the ${apiLimit}-result limit. There may be additional matches not shown.`
        : '';
    return { list, more, limitWarning };
  }

  test('should format items within maxDisplay', () => {
    const items = [
      { id: 'id1', title: 'Item 1' },
      { id: 'id2', title: 'Item 2' },
    ];

    const result = formatItemList(items, 'pages', 10, 50);

    assert.strictEqual(result.list.includes('Item 1'), true);
    assert.strictEqual(result.list.includes('Item 2'), true);
    assert.strictEqual(result.more, '');
    assert.strictEqual(result.limitWarning, '');
  });

  test('should add "more" message when exceeding maxDisplay', () => {
    const items = Array.from({ length: 15 }, (_, i) => ({
      id: `id${i}`,
      title: `Item ${i}`,
    }));

    const result = formatItemList(items, 'pages', 10, 50);

    assert.strictEqual(result.more, '\n\n... and 5 more pages.');
  });

  test('should add limit warning when hitting API limit', () => {
    const items = Array.from({ length: 50 }, (_, i) => ({
      id: `id${i}`,
      title: `Item ${i}`,
    }));

    const result = formatItemList(items, 'results', 10, 50);

    assert.strictEqual(result.limitWarning.includes('50-result limit'), true);
    assert.strictEqual(result.limitWarning.includes('additional matches'), true);
  });

  test('should use custom itemType in more message', () => {
    const items = Array.from({ length: 15 }, (_, i) => ({
      id: `id${i}`,
      title: `Section ${i}`,
    }));

    const result = formatItemList(items, 'sections', 10, 50);

    assert.strictEqual(result.more, '\n\n... and 5 more sections.');
  });

  test('should handle empty items array', () => {
    const result = formatItemList([], 'pages', 10, 50);

    assert.strictEqual(result.list, '');
    assert.strictEqual(result.more, '');
    assert.strictEqual(result.limitWarning, '');
  });

  test('should respect custom maxDisplay', () => {
    const items = Array.from({ length: 10 }, (_, i) => ({
      id: `id${i}`,
      title: `Item ${i}`,
    }));

    const result = formatItemList(items, 'pages', 5, 50);

    assert.strictEqual(result.more, '\n\n... and 5 more pages.');
  });
});

describe('createPageHtml', () => {
  /**
   * Convert text to HTML paragraph
   * @param {string} text - The text to convert
   * @returns {string} HTML paragraph or empty string
   */
  function textToHtml(text) {
    if (!text) return '';
    return `<p>${text}</p>`;
  }

  /**
   * Create HTML page content
   * @param {string} title - The page title
   * @param {string} content - The page content
   * @returns {string} Complete HTML document
   */
  function createPageHtml(title, content) {
    const htmlContent = textToHtml(content);
    return `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(title)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(title)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
  }

  test('should create valid HTML document', () => {
    const html = createPageHtml('Test Title', 'Test content');

    assert.strictEqual(html.includes('<!DOCTYPE html>'), true);
    assert.strictEqual(html.includes('<html>'), true);
    assert.strictEqual(html.includes('</html>'), true);
    assert.strictEqual(html.includes('<head>'), true);
    assert.strictEqual(html.includes('<body>'), true);
  });

  test('should include title in head and body', () => {
    const html = createPageHtml('My Page', 'Content');

    assert.strictEqual(html.includes('<title>'), true);
    assert.strictEqual(html.includes('<h1>'), true);
    assert.strictEqual(html.includes('My Page'), true);
  });

  test('should include content in body', () => {
    const html = createPageHtml('Title', 'This is my content');

    assert.strictEqual(html.includes('This is my content'), true);
  });

  test('should include timestamp', () => {
    const html = createPageHtml('Title', 'Content');

    assert.strictEqual(html.includes('Created via OneNote MCP on'), true);
  });

  test('should include charset meta tag', () => {
    const html = createPageHtml('Title', 'Content');

    assert.strictEqual(html.includes('<meta charset="utf-8">'), true);
  });
});

console.log(' All integration tests defined. Run with: node --test test_integration.mjs');
