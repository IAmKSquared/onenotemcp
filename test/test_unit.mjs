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
  validateIsoDate,
  extractTextSummary,
  extractReadableText,
  textToHtml,
  Cache,
  validateCsvData,
} from '../src/utils/common.mjs';
import {
  formatTimestamp,
  formatModifiedBy,
  formatMetadata,
  formatItemInfo,
} from '../src/utils/html.mjs';

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

  test('should allow onenote: URLs for app deep links', () => {
    assert.strictEqual(sanitizeUrl('onenote://example.com/app'), 'onenote://example.com/app');
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
    assert.doesNotThrow(() => validateId('abc1234567'));
    assert.strictEqual(validateId('abc1234567'), 'abc1234567');
  });

  test('should accept IDs with hyphens and underscores', () => {
    assert.doesNotThrow(() => validateId('abc-123_xyz'));
  });

  test('should trim whitespace', () => {
    assert.strictEqual(validateId('  abc1234567  '), 'abc1234567');
  });

  test('should reject empty strings', () => {
    assert.throws(() => validateId(''), /must be a non-empty string/);
  });

  test('should reject null/undefined', () => {
    assert.throws(() => validateId(null), /must be a non-empty string/);
    assert.throws(() => validateId(undefined), /must be a non-empty string/);
  });

  test('should reject IDs that are too short', () => {
    assert.throws(() => validateId('abc123'), /out of expected range/);
    assert.throws(() => validateId('short'), /out of expected range/);
  });

  test('should accept IDs with minimum length (10 chars)', () => {
    assert.doesNotThrow(() => validateId('abcd123456'));
  });

  test('should reject IDs with HTML/script injection', () => {
    assert.throws(() => validateId('<script>alert(1)</script>'), /dangerous characters/);
    assert.throws(() => validateId('test<script'), /dangerous characters/);
    assert.throws(() => validateId('test>123456'), /dangerous characters/);
    assert.throws(() => validateId('test"1234567'), /dangerous characters/);
  });

  test('should reject javascript: protocol', () => {
    assert.throws(() => validateId('javascript:alert(1)'), /dangerous characters/);
  });

  test('should reject data: protocol', () => {
    assert.throws(() => validateId('data:text/html,<script>'), /dangerous characters/);
  });

  test('should reject vbscript: protocol', () => {
    assert.throws(() => validateId('vbscript:msgbox'), /dangerous characters/);
  });

  test('should reject path traversal attempts', () => {
    assert.throws(() => validateId('../../../etc/passwd'), /dangerous characters/);
  });

  test('should reject newlines', () => {
    assert.throws(() => validateId('test\ninjection'), /dangerous characters/);
    assert.throws(() => validateId('test\rinjection'), /dangerous characters/);
  });

  test('should reject multiple consecutive spaces', () => {
    assert.throws(() => validateId('test  injection'), /dangerous characters/);
  });

  test('should reject command injection characters', () => {
    assert.throws(() => validateId('test;ls1234'), /dangerous characters/);
    assert.throws(() => validateId('test|cat123'), /dangerous characters/);
    assert.throws(() => validateId('test&whoami1'), /dangerous characters/);
    assert.throws(() => validateId('test$var123'), /dangerous characters/);
  });

  test('should reject IDs that are too long', () => {
    const longId = 'a'.repeat(201);
    assert.throws(() => validateId(longId), /out of expected range/);
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

describe('validateIsoDate', () => {
  test('should accept valid YYYY-MM-DD format', () => {
    assert.strictEqual(validateIsoDate('2024-01-15', 'testDate'), '2024-01-15');
    assert.strictEqual(validateIsoDate('2025-12-31', 'testDate'), '2025-12-31');
  });

  test('should accept valid YYYY-MM-DDTHH:MM:SSZ format', () => {
    assert.strictEqual(validateIsoDate('2024-01-15T14:30:00Z', 'testDate'), '2024-01-15T14:30:00Z');
  });

  test('should accept datetime without Z suffix', () => {
    assert.strictEqual(validateIsoDate('2024-01-15T14:30:00', 'testDate'), '2024-01-15T14:30:00');
  });

  test('should reject invalid format', () => {
    assert.throws(() => validateIsoDate('01-15-2024', 'testDate'), /Invalid testDate.*ISO 8601/);
    assert.throws(() => validateIsoDate('2024/01/15', 'testDate'), /Invalid testDate.*ISO 8601/);
    assert.throws(() => validateIsoDate('Jan 15, 2024', 'testDate'), /Invalid testDate.*ISO 8601/);
  });

  test('should reject invalid dates that match format', () => {
    assert.throws(() => validateIsoDate('2024-13-01', 'testDate'), /not a valid date/);
    assert.throws(() => validateIsoDate('2024-01-45', 'testDate'), /not a valid date/);
  });

  test('should reject null/undefined', () => {
    assert.throws(() => validateIsoDate(null, 'testDate'), /expected a date string/);
    assert.throws(() => validateIsoDate(undefined, 'testDate'), /expected a date string/);
  });

  test('should use param name in error messages', () => {
    assert.throws(() => validateIsoDate('bad', 'modifiedAfter'), /Invalid modifiedAfter/);
    assert.throws(() => validateIsoDate('bad', 'modifiedBefore'), /Invalid modifiedBefore/);
  });
});

// ============================================================================
// Content Processing Tests
// ============================================================================

describe('extractTextSummary', () => {
  test('should extract summary from HTML', () => {
    const html = '<p>This is a test paragraph with some content.</p>';
    const summary = extractTextSummary(html, 20);
    assert.strictEqual(summary, 'This is a test parag...');
  });

  test('should use default max length of 300', () => {
    const longText = 'a'.repeat(400);
    const html = `<p>${longText}</p>`;
    const summary = extractTextSummary(html);
    assert.strictEqual(summary.length, 303); // 300 chars + '...'
    assert.strictEqual(summary.endsWith('...'), true);
  });

  test('should not add ellipsis if text is shorter than max length', () => {
    const html = '<p>Short text</p>';
    const summary = extractTextSummary(html, 100);
    assert.strictEqual(summary, 'Short text');
    assert.strictEqual(summary.endsWith('...'), false);
  });

  test('should handle empty HTML', () => {
    assert.strictEqual(extractTextSummary(''), 'No content to summarize.');
    assert.strictEqual(extractTextSummary(null), 'No content to summarize.');
  });

  test('should handle HTML with no body text', () => {
    const html = '<html><body></body></html>';
    assert.strictEqual(extractTextSummary(html), 'No text content found in HTML body.');
  });

  test('should normalize whitespace', () => {
    const html = '<p>Text   with    multiple     spaces</p>';
    const summary = extractTextSummary(html, 50);
    assert.strictEqual(summary, 'Text with multiple spaces');
  });

  test('should extract text from complex HTML', () => {
    const html = '<div><h1>Title</h1><p>First paragraph.</p><p>Second paragraph.</p></div>';
    const summary = extractTextSummary(html, 30);
    assert.strictEqual(summary.startsWith('TitleFirst paragraph.'), true);
    assert.strictEqual(summary.endsWith('...'), true);
  });

  test('should handle HTML with scripts and styles', () => {
    const html =
      '<html><head><style>body{color:red}</style></head><body><p>Content</p><script>alert(1)</script></body></html>';
    const summary = extractTextSummary(html, 50);
    assert.strictEqual(summary.includes('alert'), false); // script content should be stripped
    assert.strictEqual(summary.includes('Content'), true);
  });
});

describe('extractReadableText', () => {
  test('should extract text from simple HTML', () => {
    const html = '<p>Hello World</p>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('Hello World'), true);
  });

  test('should format headings with underlines', () => {
    const html = '<h1>Title</h1><p>Content</p>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('Title'), true);
    assert.strictEqual(text.includes('-----'), true); // Underline
  });

  test('should format unordered lists with bullets', () => {
    const html = '<ul><li>Item 1</li><li>Item 2</li></ul>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('- Item 1'), true);
    assert.strictEqual(text.includes('- Item 2'), true);
  });

  test('should format ordered lists with numbers', () => {
    const html = '<ol><li>First</li><li>Second</li></ol>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('1. First'), true);
    assert.strictEqual(text.includes('2. Second'), true);
  });

  test('should format tables with pipe separators', () => {
    const html = '<table><tr><th>Header</th></tr><tr><td>Data</td></tr></table>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('📊 Table content:'), true);
    assert.strictEqual(text.includes('Header'), true);
    assert.strictEqual(text.includes('Data'), true);
  });

  test('should remove script and style tags', () => {
    const html = '<p>Content</p><script>alert(1)</script><style>body{color:red}</style>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('alert'), false);
    assert.strictEqual(text.includes('color'), false);
    assert.strictEqual(text.includes('Content'), true);
  });

  test('should handle empty HTML', () => {
    assert.strictEqual(extractReadableText(''), '');
    assert.strictEqual(extractReadableText(null), '');
  });

  test('should handle complex HTML with multiple elements', () => {
    const html = '<h1>Title</h1><p>Paragraph 1</p><ul><li>Item 1</li></ul><p>Paragraph 2</p>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('Title'), true);
    assert.strictEqual(text.includes('Paragraph 1'), true);
    assert.strictEqual(text.includes('- Item 1'), true);
    assert.strictEqual(text.includes('Paragraph 2'), true);
  });

  test('should fallback to body text if no structured elements', () => {
    const html = '<div>Plain text in div</div>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('Plain text in div'), true);
  });

  test('should handle nested lists', () => {
    const html = '<ul><li>Parent<ul><li>Child</li></ul></li></ul>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('Parent'), true);
    assert.strictEqual(text.includes('Child'), true);
  });

  test('should preserve paragraph spacing', () => {
    const html = '<p>First paragraph</p><p>Second paragraph</p>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('First paragraph'), true);
    assert.strictEqual(text.includes('Second paragraph'), true);
  });

  test('should handle all heading levels', () => {
    const html = '<h1>H1</h1><h2>H2</h2><h3>H3</h3><h4>H4</h4><h5>H5</h5><h6>H6</h6>';
    const text = extractReadableText(html);
    assert.strictEqual(text.includes('H1'), true);
    assert.strictEqual(text.includes('H2'), true);
    assert.strictEqual(text.includes('H3'), true);
    assert.strictEqual(text.includes('H4'), true);
    assert.strictEqual(text.includes('H5'), true);
    assert.strictEqual(text.includes('H6'), true);
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
    await new Promise((resolve) => setTimeout(resolve, 150));

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

  test('should invalidate keys by wildcard pattern', () => {
    const cache = new Cache();
    cache.set('user:1', 'data1');
    cache.set('user:2', 'data2');
    cache.set('post:1', 'post1');

    cache.invalidate('user:*');

    assert.strictEqual(cache.get('user:1'), undefined);
    assert.strictEqual(cache.get('user:2'), undefined);
    assert.strictEqual(cache.get('post:1'), 'post1');
  });

  test('should invalidate keys by RegExp pattern', () => {
    const cache = new Cache();
    cache.set('user:1', 'data1');
    cache.set('user:2', 'data2');
    cache.set('admin:1', 'admin1');
    cache.set('post:1', 'post1');

    cache.invalidate(/^(user|admin):/);

    assert.strictEqual(cache.get('user:1'), undefined);
    assert.strictEqual(cache.get('user:2'), undefined);
    assert.strictEqual(cache.get('admin:1'), undefined);
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

// ============================================================================
// CSV Validation Tests
// ============================================================================

describe('validateCsvData', () => {
  test('should parse valid CSV data', () => {
    const csv = 'Name,Age,City\nJohn,30,NYC\nJane,25,LA';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 3);
    assert.deepStrictEqual(result[0], ['Name', 'Age', 'City']);
    assert.deepStrictEqual(result[1], ['John', '30', 'NYC']);
    assert.deepStrictEqual(result[2], ['Jane', '25', 'LA']);
  });

  test('should reject CSV with only header row', () => {
    const csv = 'Name,Age,City';
    assert.throws(() => validateCsvData(csv), {
      message: 'Table data must have at least a header row and one data row.',
    });
  });

  test('should reject empty CSV', () => {
    assert.throws(() => validateCsvData(''), {
      message: 'CSV data must be a non-empty string.',
    });
  });

  test('should reject null/undefined', () => {
    assert.throws(() => validateCsvData(null), {
      message: 'CSV data must be a non-empty string.',
    });
    assert.throws(() => validateCsvData(undefined), {
      message: 'CSV data must be a non-empty string.',
    });
  });

  test('should reject inconsistent column counts', () => {
    const csv = 'Name,Age,City\nJohn,30\nJane,25,LA';
    assert.throws(() => validateCsvData(csv), {
      message:
        'Row 2 has 2 columns, but header has 3 columns. All rows must have the same number of columns.',
    });
  });

  test('should reject CSV with empty lines filtered leaving only header', () => {
    const csv = '\nName,Age,City\n\n';
    assert.throws(() => validateCsvData(csv), {
      message: 'Table data must have at least a header row and one data row.',
    });
  });

  test('should handle CSV with extra whitespace', () => {
    const csv = '  Name , Age , City  \n  John , 30 , NYC  \n  Jane , 25 , LA  ';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 3);
    assert.deepStrictEqual(result[0], ['Name', 'Age', 'City']);
    assert.deepStrictEqual(result[1], ['John', '30', 'NYC']);
  });

  test('should sanitize cells with formula prefixes (=)', () => {
    const csv = 'Name,Formula\nJohn,=SUM(A1:A10)';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 2);
    assert.strictEqual(result[1][1], "'=SUM(A1:A10)");
  });

  test('should allow positive numbers with + prefix', () => {
    const csv = 'Name,Value\nJohn,+123';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][1], '+123');
  });

  test('should allow negative numbers with - prefix', () => {
    const csv = 'Name,Value\nJohn,-123';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][1], '-123');
  });

  test('should sanitize formula-like strings starting with + or -', () => {
    const csv = 'Name,Formula1,Formula2\nJohn,-1+1,+1-1';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][1], "'-1+1");
    assert.strictEqual(result[1][2], "'+1-1");
  });

  test('should sanitize cells with @ prefix', () => {
    const csv = 'Name,Handle\nJohn,@user';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][1], "'@user");
  });

  test('should not sanitize normal cells', () => {
    const csv = 'Name,Description\nJohn,This is a normal cell';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][1], 'This is a normal cell');
  });

  test('should filter out completely empty rows', () => {
    const csv = 'Name,Age\n\n\nJohn,30\n\nJane,25\n\n';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 3); // Header + 2 data rows
    assert.deepStrictEqual(result[0], ['Name', 'Age']);
    assert.deepStrictEqual(result[1], ['John', '30']);
    assert.deepStrictEqual(result[2], ['Jane', '25']);
  });

  test('should handle CSV with single column', () => {
    const csv = 'Name\nJohn\nJane';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 3);
    assert.deepStrictEqual(result[0], ['Name']);
    assert.deepStrictEqual(result[1], ['John']);
    assert.deepStrictEqual(result[2], ['Jane']);
  });

  test('should handle CSV with many columns', () => {
    const csv = 'A,B,C,D,E,F\n1,2,3,4,5,6\n7,8,9,10,11,12';
    const result = validateCsvData(csv);

    assert.strictEqual(result.length, 3);
    assert.strictEqual(result[0].length, 6);
    assert.strictEqual(result[1].length, 6);
    assert.strictEqual(result[2].length, 6);
  });

  test('should sanitize multiple dangerous cells but allow numbers', () => {
    const csv = 'Col1,Col2,Col3\n=FORMULA(),+123,-456\n@handle,normal,=ANOTHER()';
    const result = validateCsvData(csv);

    assert.strictEqual(result[1][0], "'=FORMULA()");
    assert.strictEqual(result[1][1], '+123'); // Valid number, not sanitized
    assert.strictEqual(result[1][2], '-456'); // Valid number, not sanitized
    assert.strictEqual(result[2][0], "'@handle");
    assert.strictEqual(result[2][1], 'normal');
    assert.strictEqual(result[2][2], "'=ANOTHER()");
  });
});

// ============================================================================
// Timestamp and Metadata Formatter Tests
// ============================================================================

describe('formatTimestamp', () => {
  test('should format valid ISO timestamp', () => {
    const timestamp = '2025-01-15T14:30:00Z';
    const result = formatTimestamp(timestamp);
    // Result varies by locale, but should contain date components
    assert.strictEqual(typeof result, 'string');
    assert.strictEqual(result.length > 0, true);
  });

  test('should return empty string for null', () => {
    assert.strictEqual(formatTimestamp(null), '');
  });

  test('should return empty string for undefined', () => {
    assert.strictEqual(formatTimestamp(undefined), '');
  });

  test('should return empty string for empty string', () => {
    assert.strictEqual(formatTimestamp(''), '');
  });

  test('should return empty string for invalid date', () => {
    assert.strictEqual(formatTimestamp('not-a-date'), '');
    assert.strictEqual(formatTimestamp('invalid'), '');
  });
});

describe('formatModifiedBy', () => {
  test('should extract user displayName', () => {
    const identitySet = {
      user: { displayName: 'John Doe', email: 'john@example.com' },
    };
    assert.strictEqual(formatModifiedBy(identitySet), 'John Doe');
  });

  test('should fallback to user email if no displayName', () => {
    const identitySet = {
      user: { email: 'john@example.com' },
    };
    assert.strictEqual(formatModifiedBy(identitySet), 'john@example.com');
  });

  test('should fallback to application displayName', () => {
    const identitySet = {
      application: { displayName: 'Test App' },
    };
    assert.strictEqual(formatModifiedBy(identitySet), 'Test App');
  });

  test('should prefer user over application', () => {
    const identitySet = {
      user: { displayName: 'John Doe' },
      application: { displayName: 'Test App' },
    };
    assert.strictEqual(formatModifiedBy(identitySet), 'John Doe');
  });

  test('should return empty string for null', () => {
    assert.strictEqual(formatModifiedBy(null), '');
  });

  test('should return empty string for undefined', () => {
    assert.strictEqual(formatModifiedBy(undefined), '');
  });

  test('should return empty string for empty object', () => {
    assert.strictEqual(formatModifiedBy({}), '');
  });

  test('should return empty string if user has no name or email', () => {
    const identitySet = { user: {} };
    assert.strictEqual(formatModifiedBy(identitySet), '');
  });
});

describe('formatMetadata', () => {
  test('should format lastModifiedDateTime only', () => {
    const item = {
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
    };
    const result = formatMetadata(item);
    assert.strictEqual(result.includes('Modified:'), true);
  });

  test('should format createdDateTime only', () => {
    const item = {
      createdDateTime: '2025-01-10T10:00:00Z',
    };
    const result = formatMetadata(item);
    assert.strictEqual(result.includes('Created:'), true);
  });

  test('should format both timestamps', () => {
    const item = {
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
      createdDateTime: '2025-01-10T10:00:00Z',
    };
    const result = formatMetadata(item);
    assert.strictEqual(result.includes('Modified:'), true);
    assert.strictEqual(result.includes('Created:'), true);
    assert.strictEqual(result.includes(' | '), true);
  });

  test('should include modifiedBy user name', () => {
    const item = {
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
      lastModifiedBy: {
        user: { displayName: 'John Doe' },
      },
    };
    const result = formatMetadata(item);
    assert.strictEqual(result.includes('by John Doe'), true);
  });

  test('should return empty string for item with no metadata', () => {
    const item = { id: 'test-id', displayName: 'Test' };
    assert.strictEqual(formatMetadata(item), '');
  });

  test('should handle null lastModifiedBy gracefully', () => {
    const item = {
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
      lastModifiedBy: null,
    };
    const result = formatMetadata(item);
    assert.strictEqual(result.includes('Modified:'), true);
    assert.strictEqual(result.includes('by'), false);
  });
});

describe('formatItemInfo', () => {
  test('should format basic item with displayName', () => {
    const item = { id: 'test-id-12345', displayName: 'Test Notebook' };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('**Test Notebook**'), true);
    assert.strictEqual(result.includes('ID: test-id-12345'), true);
  });

  test('should format item with title (for pages)', () => {
    const item = { id: 'page-id-12345', title: 'Test Page' };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('**Test Page**'), true);
  });

  test('should show Untitled for items without name', () => {
    const item = { id: 'test-id-12345' };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('**Untitled**'), true);
  });

  test('should include index prefix when provided', () => {
    const item = { id: 'test-id-12345', displayName: 'Test' };
    const result = formatItemInfo(item, 0);
    assert.strictEqual(result.startsWith('1. '), true);
  });

  test('should format web link from links object', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      links: {
        oneNoteWebUrl: { href: 'https://example.com/web' },
      },
    };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('[Web](https://example.com/web)'), true);
  });

  test('should format app link from links object', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      links: {
        oneNoteClientUrl: { href: 'onenote://example.com/app' },
      },
    };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('[App](onenote://example.com/app)'), true);
  });

  test('should include metadata when available', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
      createdDateTime: '2025-01-10T10:00:00Z',
    };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('Modified:'), true);
    assert.strictEqual(result.includes('Created:'), true);
  });

  test('should format with modifiedBy information', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      lastModifiedDateTime: '2025-01-15T14:30:00Z',
      lastModifiedBy: {
        user: { displayName: 'John Doe' },
      },
    };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('by John Doe'), true);
  });

  test('should handle direct URL strings (not objects)', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      links: {
        oneNoteWebUrl: 'https://example.com/direct',
      },
    };
    const result = formatItemInfo(item);
    assert.strictEqual(result.includes('[Web](https://example.com/direct)'), true);
  });

  test('should sanitize dangerous URLs', () => {
    const item = {
      id: 'test-id-12345',
      displayName: 'Test',
      links: {
        oneNoteWebUrl: 'javascript:alert(1)',
        oneNoteClientUrl: 'data:text/html,<script>alert(1)</script>',
      },
    };
    const result = formatItemInfo(item);
    // Dangerous URLs should be filtered out entirely (sanitizeUrl returns '#')
    assert.strictEqual(result.includes('javascript:'), false);
    assert.strictEqual(result.includes('data:'), false);
    assert.strictEqual(result.includes('[Web]'), false);
    assert.strictEqual(result.includes('[App]'), false);
  });
});

console.log('✅ All unit tests defined. Run with: node --test test_unit.mjs');
