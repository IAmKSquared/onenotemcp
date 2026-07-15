/**
 * Integration tests for OneNote MCP Server core functions
 * These tests cover functions from onenote-mcp.mjs that require more complex setup
 * Run with: node --test test_integration.mjs
 */

import { test, describe, mock, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { formatItemInfo, formatItemList, createPageHtml } from '../src/utils/html.mjs';
import { retryWithBackoff, getDetailedErrorMessage } from '../src/api/retry.mjs';
import { encrypt, decrypt } from '../src/auth/encryption.mjs';

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
  // These tests exercise the REAL encryption module (AES-256-GCM), which manages
  // its own key via KeyStorage. The basic round-trip, structural DecryptionError
  // cases, and encrypt() input validation are already covered in test_tools.mjs;
  // the tests here add value by verifying IV uniqueness, payload variety, and
  // authentication-tag tamper detection.

  /**
   * Flips the first hex character to produce different-but-valid hex of the
   * same length, simulating tampering with an encrypted field.
   * @param {string} hex - The hex string to tamper with.
   * @returns {string} The tampered hex string.
   */
  function tamperHex(hex) {
    const replacement = hex[0] === 'a' ? 'b' : 'a';
    return replacement + hex.slice(1);
  }

  test('should use unique IVs for each encryption', async () => {
    const text = 'same-text';

    const enc1 = await encrypt(text);
    const enc2 = await encrypt(text);

    // Random IVs mean the same plaintext yields different ciphertext each time.
    assert.notStrictEqual(enc1.iv, enc2.iv);
    assert.notStrictEqual(enc1.encryptedData, enc2.encryptedData);

    // But both still decrypt back to the same text.
    assert.strictEqual(await decrypt(enc1), text);
    assert.strictEqual(await decrypt(enc2), text);
  });

  test('should handle long strings', async () => {
    const longText = 'a'.repeat(10000);

    const encrypted = await encrypt(longText);
    const decrypted = await decrypt(encrypted);

    assert.strictEqual(decrypted, longText);
  });

  test('should handle special characters', async () => {
    const specialText = 'Special: <>&quot; symbols @#$%^&*()';

    const encrypted = await encrypt(specialText);
    const decrypted = await decrypt(encrypted);

    assert.strictEqual(decrypted, specialText);
  });

  test('should handle JSON strings', async () => {
    const jsonText = JSON.stringify({
      token: 'abc123',
      expiresOn: '2025-12-31T23:59:59Z',
      scopes: ['Notes.Read', 'Notes.ReadWrite'],
    });

    const encrypted = await encrypt(jsonText);
    const decrypted = await decrypt(encrypted);

    assert.strictEqual(decrypted, jsonText);
    assert.doesNotThrow(() => JSON.parse(decrypted));
  });

  test('should reject decryption when the authTag is tampered with', async () => {
    const encrypted = await encrypt('secret-data');
    const tampered = { ...encrypted, authTag: tamperHex(encrypted.authTag) };

    // GCM authentication fails inside decipher.final(), which throws a native
    // Error (not a DecryptionError, which only guards structural validation).
    await assert.rejects(
      async () => decrypt(tampered),
      (error) => {
        assert.ok(error instanceof Error);
        assert.notStrictEqual(error.name, 'DecryptionError');
        return true;
      }
    );
  });

  test('should reject decryption when the ciphertext is tampered with', async () => {
    const encrypted = await encrypt('secret-data');
    const tampered = { ...encrypted, encryptedData: tamperHex(encrypted.encryptedData) };

    await assert.rejects(
      async () => decrypt(tampered),
      (error) => {
        assert.ok(error instanceof Error);
        assert.notStrictEqual(error.name, 'DecryptionError');
        return true;
      }
    );
  });
});

// ============================================================================
// Retry with Backoff Tests (Error Resilience)
// ============================================================================

describe('retryWithBackoff', () => {
  // Exercises the REAL retryWithBackoff from src/api/retry.mjs. Signature is
  // (fn, maxRetries, baseDelay); baseDelay is kept small so tests stay fast.
  // Non-retryable status codes come from NON_RETRYABLE_STATUS_CODES in
  // constants (400, 401, 403, 404, 409); backoff multiplier is 2.

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
  // Exercises the REAL getDetailedErrorMessage from src/api/retry.mjs.
  // Assertions match the stable text portions of each message rather than the
  // exact emoji prefixes, so they stay robust against emoji encoding.

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

describe('formatItemInfo', () => {
  test('should format page with title', () => {
    const page = { id: 'page123', title: 'My Page' };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123)');
  });

  test('should format page with displayName', () => {
    const page = { id: 'page123', displayName: 'My Section' };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Section** (ID: page123)');
  });

  test('should prefer displayName over title', () => {
    const page = { id: 'page123', displayName: 'Display', title: 'Title' };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**Display** (ID: page123)');
  });

  test('should use "Untitled" for missing name', () => {
    const page = { id: 'page123' };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**Untitled** (ID: page123)');
  });

  test('should format with index', () => {
    const page = { id: 'page123', title: 'My Page' };
    const result = formatItemInfo(page, 0);

    assert.strictEqual(result, '1. **My Page** (ID: page123)');
  });

  test('should format with zero index', () => {
    const page = { id: 'page123', title: 'First Page' };
    const result = formatItemInfo(page, 0);

    assert.strictEqual(result, '1. **First Page** (ID: page123)');
  });

  test('should format with non-zero index', () => {
    const page = { id: 'page456', title: 'Third Page' };
    const result = formatItemInfo(page, 2);

    assert.strictEqual(result, '3. **Third Page** (ID: page456)');
  });

  test('should include web link from href object', () => {
    const page = {
      id: 'page123',
      title: 'My Page',
      links: { oneNoteWebUrl: { href: 'https://onenote.com/page123' } },
    };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123) - [Web](https://onenote.com/page123)');
  });

  test('should include app link from href object', () => {
    const page = {
      id: 'page123',
      title: 'My Page',
      links: { oneNoteClientUrl: { href: 'onenote://page123' } },
    };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123) - [App](onenote://page123)');
  });

  test('should include both web and app links from href objects', () => {
    const page = {
      id: 'page456',
      title: 'Linked Page',
      links: {
        oneNoteWebUrl: { href: 'https://onenote.com/page456' },
        oneNoteClientUrl: { href: 'onenote://page456' },
      },
    };
    const result = formatItemInfo(page, 1);

    assert.strictEqual(
      result,
      '2. **Linked Page** (ID: page456) - [Web](https://onenote.com/page456) | [App](onenote://page456)'
    );
  });

  test('should handle direct string URLs as fallback', () => {
    const page = {
      id: 'page123',
      title: 'My Page',
      links: { oneNoteWebUrl: 'https://onenote.com/page123' },
    };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123) - [Web](https://onenote.com/page123)');
  });

  test('should not include links when links object is empty', () => {
    const page = {
      id: 'page123',
      title: 'My Page',
      links: {},
    };
    const result = formatItemInfo(page);

    assert.strictEqual(result, '**My Page** (ID: page123)');
  });
});

describe('formatItemList', () => {
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
  // Exercises the REAL createPageHtml from src/utils/html.mjs. The title is
  // plain-HTML-escaped (markdown is NOT applied to it), while the body content
  // still goes through textToHtml. See GitHub issue #4.

  test('should create valid HTML document skeleton', () => {
    const html = createPageHtml('Test Title', 'Test content');

    assert.strictEqual(html.includes('<!DOCTYPE html>'), true);
    assert.strictEqual(html.includes('<html>'), true);
    assert.strictEqual(html.includes('</html>'), true);
    assert.strictEqual(html.includes('<head>'), true);
    assert.strictEqual(html.includes('<body>'), true);
    assert.strictEqual(html.includes('<meta charset="utf-8">'), true);
  });

  test('should include title in head and body', () => {
    const html = createPageHtml('My Page', 'Content');

    assert.strictEqual(html.includes('<title>My Page</title>'), true);
    assert.strictEqual(html.includes('<h1>My Page</h1>'), true);
  });

  test('should apply markdown to body content via textToHtml', () => {
    const html = createPageHtml('Title', '**bold** content');

    assert.strictEqual(html.includes('<strong>bold</strong>'), true);
  });

  test('should include timestamp', () => {
    const html = createPageHtml('Title', 'Content');

    assert.strictEqual(html.includes('Created via OneNote MCP on'), true);
  });

  test('should NOT run the title through the markdown converter', () => {
    // A title starting with a markdown token must stay literal, not become an <h1>.
    const html = createPageHtml('# Plan', 'Body');

    assert.strictEqual(html.includes('<title># Plan</title>'), true);
    assert.strictEqual(html.includes('<h1># Plan</h1>'), true);
    // The buggy behavior would nest a <p> inside <title>; it must not appear.
    assert.strictEqual(html.includes('<title><p>'), false);
  });

  test('should HTML-escape special characters in the title', () => {
    const html = createPageHtml('A < B & C > D', 'Body');

    assert.strictEqual(html.includes('<title>A &lt; B &amp; C &gt; D</title>'), true);
    assert.strictEqual(html.includes('<h1>A &lt; B &amp; C &gt; D</h1>'), true);
  });
});

console.log('All integration tests defined. Run with: node --test test_integration.mjs');
