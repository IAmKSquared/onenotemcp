/**
 * Unit tests for OneNote MCP Server utility functions
 * Run with: node --test test_unit.mjs
 */

import { test, describe } from 'node:test';
import assert from 'node:assert';
import {
  escapeODataString,
  sanitizeUrl,
  validateId,
  htmlToText,
  textToHtml,
  Cache
} from './utils.mjs';

// ============================================================================
// Security Functions Tests
// ============================================================================

describe('escapeODataString', () => {
  test('should escape single quotes by doubling them', () => {
    assert.strictEqual(escapeODataString("O'Reilly"), "O''Reilly");
  });

  test('should handle multiple single quotes', () => {
    assert.strictEqual(escapeODataString("'test'"), "''test''");
  });

  test('should handle empty string', () => {
    assert.strictEqual(escapeODataString(''), '');
  });

  test('should handle null/undefined', () => {
    assert.strictEqual(escapeODataString(null), '');
    assert.strictEqual(escapeODataString(undefined), '');
  });

  test('should not modify strings without quotes', () => {
    assert.strictEqual(escapeODataString('normal text'), 'normal text');
  });

  test('should handle injection attempts', () => {
    const injection = "'; DROP TABLE users; --";
    const escaped = escapeODataString(injection);
    assert.strictEqual(escaped, "''; DROP TABLE users; --");
  });
});

describe('sanitizeUrl', () => {
  test('should allow https URLs', () => {
    assert.strictEqual(sanitizeUrl('https://example.com'), 'https://example.com');
  });

  test('should allow http URLs', () => {
    assert.strictEqual(sanitizeUrl('http://example.com'), 'http://example.com');
  });

  test('should allow mailto URLs', () => {
    assert.strictEqual(sanitizeUrl('mailto:test@example.com'), 'mailto:test@example.com');
  });

  test('should block javascript: protocol', () => {
    assert.strictEqual(sanitizeUrl('javascript:alert(1)'), '#');
  });

  test('should block data: protocol', () => {
    assert.strictEqual(sanitizeUrl('data:text/html,<script>alert(1)</script>'), '#');
  });

  test('should block vbscript: protocol', () => {
    assert.strictEqual(sanitizeUrl('vbscript:msgbox'), '#');
  });

  test('should handle empty/null URLs', () => {
    assert.strictEqual(sanitizeUrl(''), '#');
    assert.strictEqual(sanitizeUrl(null), '#');
    assert.strictEqual(sanitizeUrl(undefined), '#');
  });

  test('should trim whitespace', () => {
    assert.strictEqual(sanitizeUrl('  https://example.com  '), 'https://example.com');
  });

  test('should allow relative URLs without protocol', () => {
    assert.strictEqual(sanitizeUrl('/path/to/page'), '/path/to/page');
  });
});

describe('validateId', () => {
  test('should accept valid alphanumeric IDs', () => {
    assert.doesNotThrow(() => validateId('abc123'));
    assert.strictEqual(validateId('abc123'), 'abc123');
  });

  test('should accept IDs with hyphens and underscores', () => {
    assert.doesNotThrow(() => validateId('abc-123_xyz'));
  });

  test('should trim whitespace', () => {
    assert.strictEqual(validateId('  abc123  '), 'abc123');
  });

  test('should reject empty strings', () => {
    assert.throws(() => validateId(''), /must be a non-empty string/);
  });

  test('should reject null/undefined', () => {
    assert.throws(() => validateId(null), /must be a non-empty string/);
    assert.throws(() => validateId(undefined), /must be a non-empty string/);
  });

  test('should reject IDs with HTML injection characters', () => {
    assert.throws(() => validateId('<script>'), /dangerous characters/);
    assert.throws(() => validateId('test>123'), /dangerous characters/);
    assert.throws(() => validateId('test"123'), /dangerous characters/);
  });

  test('should reject javascript: protocol', () => {
    assert.throws(() => validateId('javascript:alert(1)'), /dangerous characters/);
  });

  test('should reject path traversal attempts', () => {
    assert.throws(() => validateId('../../../etc/passwd'), /dangerous characters/);
  });

  test('should reject command injection characters', () => {
    assert.throws(() => validateId('test;ls'), /dangerous characters/);
    assert.throws(() => validateId('test|cat'), /dangerous characters/);
    assert.throws(() => validateId('test&whoami'), /dangerous characters/);
  });

  test('should reject IDs that are too long', () => {
    const longId = 'a'.repeat(201);
    assert.throws(() => validateId(longId), /too long/);
  });

  test('should accept IDs up to 200 characters', () => {
    const validLongId = 'a'.repeat(200);
    assert.doesNotThrow(() => validateId(validLongId));
  });

  test('should use custom type in error messages', () => {
    assert.throws(() => validateId('', 'page'), /Invalid page ID/);
    assert.throws(() => validateId('', 'notebook'), /Invalid notebook ID/);
  });
});

// ============================================================================
// Content Processing Tests
// ============================================================================

describe('htmlToText', () => {
  test('should extract plain text from simple HTML', () => {
    const html = '<p>Hello World</p>';
    assert.strictEqual(htmlToText(html), 'Hello World');
  });

  test('should remove script tags', () => {
    const html = '<p>Hello</p><script>alert(1)</script><p>World</p>';
    const text = htmlToText(html);
    assert.strictEqual(text.includes('alert'), false);
    assert.strictEqual(text.includes('Hello'), true);
  });

  test('should remove style tags', () => {
    const html = '<p>Hello</p><style>body{color:red}</style>';
    const text = htmlToText(html);
    assert.strictEqual(text.includes('color'), false);
  });

  test('should handle empty HTML', () => {
    assert.strictEqual(htmlToText(''), '');
    assert.strictEqual(htmlToText(null), '');
  });

  test('should normalize whitespace', () => {
    const html = '<p>Hello    World</p>';
    assert.strictEqual(htmlToText(html), 'Hello World');
  });

  test('should collapse multiple blank lines', () => {
    const html = '<p>Line 1</p><br><br><br><p>Line 2</p>';
    const text = htmlToText(html);
    // Should not have more than double newlines
    assert.strictEqual(text.includes('\n\n\n'), false);
  });
});

describe('textToHtml', () => {
  test('should wrap text in paragraph tags', () => {
    const result = textToHtml('Hello World');
    assert.strictEqual(result, '<p>Hello World</p>');
  });

  test('should escape HTML special characters', () => {
    const result = textToHtml('<script>alert("xss")</script>');
    assert.strictEqual(result.includes('<script>'), false);
    assert.strictEqual(result.includes('&lt;script&gt;'), true);
  });

  test('should convert newlines to <br> tags', () => {
    const result = textToHtml('Line 1\nLine 2');
    assert.strictEqual(result, '<p>Line 1<br>Line 2</p>');
  });

  test('should create separate paragraphs for double newlines', () => {
    const result = textToHtml('Para 1\n\nPara 2');
    assert.strictEqual(result, '<p>Para 1</p><p>Para 2</p>');
  });

  test('should handle empty text', () => {
    assert.strictEqual(textToHtml(''), '<p></p>');
    assert.strictEqual(textToHtml(null), '<p></p>');
  });

  test('should escape all dangerous characters', () => {
    const dangerous = '& < > " \'';
    const result = textToHtml(dangerous);
    assert.strictEqual(result.includes('&amp;'), true);
    assert.strictEqual(result.includes('&lt;'), true);
    assert.strictEqual(result.includes('&gt;'), true);
    assert.strictEqual(result.includes('<script'), false); // Should not contain unescaped tags
    assert.strictEqual(result.includes('">'), false); // Should not contain unescaped quote combinations
  });
});

// ============================================================================
// Cache Tests
// ============================================================================

describe('Cache', () => {
  test('should store and retrieve values', () => {
    const cache = new Cache();
    cache.set('key1', 'value1');
    assert.strictEqual(cache.get('key1'), 'value1');
  });

  test('should return undefined for missing keys', () => {
    const cache = new Cache();
    assert.strictEqual(cache.get('nonexistent'), undefined);
  });

  test('should expire entries after TTL', async () => {
    const cache = new Cache();
    cache.set('key1', 'value1', 100); // 100ms TTL

    assert.strictEqual(cache.get('key1'), 'value1');

    // Wait for expiration
    await new Promise(resolve => setTimeout(resolve, 150));

    assert.strictEqual(cache.get('key1'), undefined);
  });

  test('should use default TTL when not specified', () => {
    const cache = new Cache();
    cache.set('key1', 'value1');

    // Should still be there immediately
    assert.strictEqual(cache.get('key1'), 'value1');
  });

  test('should invalidate specific keys', () => {
    const cache = new Cache();
    cache.set('key1', 'value1');
    cache.set('key2', 'value2');

    cache.invalidate('key1');

    assert.strictEqual(cache.get('key1'), undefined);
    assert.strictEqual(cache.get('key2'), 'value2');
  });

  test('should invalidate keys by pattern', () => {
    const cache = new Cache();
    cache.set('user:1', 'data1');
    cache.set('user:2', 'data2');
    cache.set('post:1', 'post1');

    cache.invalidate('user:*');

    assert.strictEqual(cache.get('user:1'), undefined);
    assert.strictEqual(cache.get('user:2'), undefined);
    assert.strictEqual(cache.get('post:1'), 'post1');
  });

  test('should clear all entries', () => {
    const cache = new Cache();
    cache.set('key1', 'value1');
    cache.set('key2', 'value2');

    cache.clear();

    assert.strictEqual(cache.get('key1'), undefined);
    assert.strictEqual(cache.get('key2'), undefined);
  });

  test('should handle complex values', () => {
    const cache = new Cache();
    const obj = { foo: 'bar', nested: { value: 123 } };

    cache.set('key1', obj);

    const retrieved = cache.get('key1');
    assert.deepStrictEqual(retrieved, obj);
  });
});

console.log('✅ All unit tests defined. Run with: node --test test_unit.mjs');
