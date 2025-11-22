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

  test('should create separate paragraphs for each line', () => {
    const result = textToHtml('Para 1\nPara 2');
    assert.strictEqual(result, '<p>Para 1</p>\n<p>Para 2</p>');
  });

  test('should handle empty text', () => {
    assert.strictEqual(textToHtml(''), '');
    assert.strictEqual(textToHtml(null), '');
  });

  test('should escape all dangerous characters', () => {
    const dangerous = '& < > " \'';
    const result = textToHtml(dangerous);
    assert.strictEqual(result.includes('&amp;'), true);
    assert.strictEqual(result.includes('&lt;'), true);
    assert.strictEqual(result.includes('&gt;'), true);
    assert.strictEqual(result.includes('<script'), false);
  });

  test('should convert markdown headers', () => {
    const result = textToHtml('# H1\n## H2\n### H3');
    assert.strictEqual(result.includes('<h1>H1</h1>'), true);
    assert.strictEqual(result.includes('<h2>H2</h2>'), true);
    assert.strictEqual(result.includes('<h3>H3</h3>'), true);
  });

  test('should convert bold text', () => {
    const result = textToHtml('**bold** and __also bold__');
    assert.strictEqual(result.includes('<strong>bold</strong>'), true);
    assert.strictEqual(result.includes('<strong>also bold</strong>'), true);
  });

  test('should convert italic text', () => {
    const result = textToHtml('*italic* and _also italic_');
    assert.strictEqual(result.includes('<em>italic</em>'), true);
    assert.strictEqual(result.includes('<em>also italic</em>'), true);
  });

  test('should convert inline code', () => {
    const result = textToHtml('Use `code` here');
    assert.strictEqual(result.includes('<code>code</code>'), true);
  });

  test('should convert code blocks', () => {
    const result = textToHtml('```\nfunction test() {}\n```');
    assert.strictEqual(result.includes('<pre><code>'), true);
    assert.strictEqual(result.includes('function test() {}'), true);
  });

  test('should convert markdown links with safe URLs', () => {
    const result = textToHtml('[Google](https://google.com)');
    assert.strictEqual(result.includes('<a href="https://google.com">Google</a>'), true);
  });

  test('should sanitize dangerous URLs in links', () => {
    const result = textToHtml('[Click](javascript:alert(1))');
    assert.strictEqual(result.includes('javascript:'), false);
    assert.strictEqual(result.includes('href="#"'), true);
  });

  test('should convert horizontal rules', () => {
    const result = textToHtml('Above\n---\nBelow');
    assert.strictEqual(result.includes('<hr>'), true);
  });

  test('should convert blockquotes', () => {
    const result = textToHtml('> Quote line 1\n> Quote line 2');
    assert.strictEqual(result.includes('<blockquote>Quote line 1</blockquote>'), true);
  });

  test('should convert unordered lists', () => {
    const result = textToHtml('- Item 1\n- Item 2\n* Item 3');
    assert.strictEqual(result.includes('<ul>'), true);
    assert.strictEqual(result.includes('<li>Item 1</li>'), true);
    assert.strictEqual(result.includes('<li>Item 2</li>'), true);
    assert.strictEqual(result.includes('<li>Item 3</li>'), true);
  });

  test('should convert ordered lists', () => {
    const result = textToHtml('1. First\n2. Second\n3. Third');
    assert.strictEqual(result.includes('<ul>'), true);
    assert.strictEqual(result.includes('<li>First</li>'), true);
    assert.strictEqual(result.includes('<li>Second</li>'), true);
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
