/**
 * Tests for tool handlers and new functionality
 * Run with: node --test test/test_tools.mjs
 */

import { test, describe } from 'node:test';
import assert from 'node:assert';
import { validateId, collectSettledResults } from '../src/utils/common.mjs';

// ============================================================================
// ID Validation Tests for Tool Parameters
// ============================================================================

describe('validateId for tool parameters', () => {
  test('should validate notebook ID', () => {
    const validId = '0-abc123def456789';
    assert.doesNotThrow(() => validateId(validId, 'notebook'));
  });

  test('should validate section ID', () => {
    const validId = '0-abc123def456789';
    assert.doesNotThrow(() => validateId(validId, 'section'));
  });

  test('should validate sectionGroup ID', () => {
    const validId = '0-abc123def456789';
    assert.doesNotThrow(() => validateId(validId, 'sectionGroup'));
  });

  test('should reject ID that is too short', () => {
    assert.throws(() => validateId('abc', 'notebook'), {
      message: /Invalid notebook ID/,
    });
  });

  test('should reject ID that is too long', () => {
    const longId = 'a'.repeat(250);
    assert.throws(() => validateId(longId, 'notebook'), {
      message: /Invalid notebook ID/,
    });
  });

  test('should reject ID with script injection', () => {
    assert.throws(() => validateId('<script>alert(1)</script>', 'notebook'), {
      message: /Invalid notebook ID/,
    });
  });

  test('should reject ID with path traversal', () => {
    assert.throws(() => validateId('../../../etc/passwd', 'notebook'), {
      message: /Invalid notebook ID/,
    });
  });

  test('should reject ID with command injection characters', () => {
    assert.throws(() => validateId('id;rm -rf /', 'notebook'), {
      message: /Invalid notebook ID/,
    });
  });
});

// ============================================================================
// Encryption Tests
// ============================================================================

describe('encryption', async () => {
  const { encrypt, decrypt } = await import('../src/auth/encryption.mjs');

  test('should encrypt and decrypt round-trip successfully', async () => {
    const original = '{"token":"secret-access-token","expiresOn":"2025-01-01"}';
    const encrypted = await encrypt(original);

    // Verify encrypted structure has required fields
    assert.ok(encrypted.iv, 'encrypted should have iv');
    assert.ok(encrypted.encryptedData, 'encrypted should have encryptedData');
    assert.ok(encrypted.authTag, 'encrypted should have authTag');

    // Verify decryption returns original
    const decrypted = await decrypt(encrypted);
    assert.strictEqual(decrypted, original);
  });

  test('should reject null input with DecryptionError', async () => {
    await assert.rejects(async () => decrypt(null), {
      name: 'DecryptionError',
      message: /Invalid encrypted data: expected object/,
    });
  });

  test('should reject object missing required fields with DecryptionError', async () => {
    await assert.rejects(async () => decrypt({ encryptedData: 'abc', authTag: 'def' }), {
      name: 'DecryptionError',
      message: /Invalid encrypted data: missing required fields/,
    });
  });

  test('should reject empty object with DecryptionError', async () => {
    await assert.rejects(async () => decrypt({}), {
      name: 'DecryptionError',
      message: /Invalid encrypted data: missing required fields/,
    });
  });

  test('should reject non-string input to encrypt()', async () => {
    await assert.rejects(async () => encrypt(null), {
      message: /encrypt\(\) requires a string input/,
    });
    await assert.rejects(async () => encrypt(123), {
      message: /encrypt\(\) requires a string input/,
    });
    await assert.rejects(async () => encrypt({ token: 'abc' }), {
      message: /encrypt\(\) requires a string input/,
    });
  });
});

// ============================================================================
// collectSettledResults Tests
// ============================================================================

describe('collectSettledResults', () => {
  test('should return flattened results when all promises succeed', async () => {
    const promises = [Promise.resolve([1, 2]), Promise.resolve([3, 4]), Promise.resolve([5])];
    const results = await collectSettledResults(promises, 'items');
    assert.deepStrictEqual(results, [1, 2, 3, 4, 5]);
  });

  test('should return only successful results when some promises fail', async () => {
    const promises = [
      Promise.resolve([1, 2]),
      Promise.reject(new Error('fetch failed')),
      Promise.resolve([3]),
    ];
    const results = await collectSettledResults(promises, 'section(s)');
    assert.deepStrictEqual(results, [1, 2, 3]);
  });

  test('should return empty array when all promises fail', async () => {
    const promises = [Promise.reject(new Error('fail 1')), Promise.reject(new Error('fail 2'))];
    const results = await collectSettledResults(promises, 'items');
    assert.deepStrictEqual(results, []);
  });

  test('should return empty array for empty input', async () => {
    const results = await collectSettledResults([], 'items');
    assert.deepStrictEqual(results, []);
  });

  test('should handle non-array values by flattening single items', async () => {
    const promises = [Promise.resolve('single'), Promise.resolve(['array', 'items'])];
    const results = await collectSettledResults(promises, 'items');
    assert.deepStrictEqual(results, ['single', 'array', 'items']);
  });
});
