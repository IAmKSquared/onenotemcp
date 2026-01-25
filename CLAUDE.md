# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with
code in this repository.

## Project Overview

OneNote MCP Server enables AI assistants to interact with Microsoft OneNote via
the Model Context Protocol (MCP). It provides 25+ tools for reading, writing,
creating, and managing OneNote content through the Microsoft Graph API.

## Commands

```bash
# Start the MCP server
npm start

# Testing
npm test                  # All tests (unit + smoke + integration)
npm run test:unit         # Unit tests only (96 tests)
npm run test:smoke        # Server startup verification
npm run test:integration  # Integration tests (55 tests)

# Code quality
npm run lint              # ESLint analysis
npm run lint:fix          # Auto-fix linting issues
npm run format            # Prettier formatting
npm run quality           # Full check (lint + format + all tests)
```

## Architecture

**Technology:** Node.js 18+, MCP SDK, Microsoft Graph API, Azure Identity
(Device Code Flow)

### Source Structure

```
src/
├── auth/           # Device Code Flow auth, token management, AES-256-GCM encryption, OS keychain
├── api/            # Caching layer (TTL-based) and retry logic with exponential backoff
├── tools/          # MCP tool implementations (auth, read, write, create, delete)
├── utils/          # Security (validation, sanitization), content processing, logging
├── config/         # Centralized constants (no magic numbers)
├── session.mjs     # Session state management (OneNoteSession class)
└── server.mjs      # Entry point
```

### Key Patterns

**Tool Registration:** Each tool category has a
`register*Tools(server, session)` function. All tools except auth tools use
`createToolHandler()` wrapper for automatic retry and error handling.

**Session Management:** `OneNoteSession` class encapsulates token storage, Graph
client initialization, and auth state. Tools call `session.ensureGraphClient()`
to verify authentication.

**Security Functions (always use these):**

- `validateId(id, type)` - Validate IDs before API calls (prevents injection)
- `sanitizeUrl(url)` - Sanitize URLs before returning to users (blocks
  javascript:, data:)
- `escapeODataString(str)` - Escape strings for OData queries

**Content Processing:** Three format options for pages - "text" (readable),
"html" (raw), "summary" (snippet). Use `extractReadableText()` for HTML→text,
`textToHtml()` for Markdown-like→HTML.

**Caching:** Notebooks and sections are cached (5-min TTL). Pages, section
groups, and search results are NOT cached. Use `invalidateCache()` with patterns
like `"sections:*"`.

**Error Handling:** `createToolHandler()` provides standardized MCP error format
with user-friendly messages and emoji indicators.

**Configuration:** All constants in `src/config/constants.mjs` - AUTH_CONFIG,
HTTP_STATUS, RETRY_CONFIG, CACHE_CONFIG, VALIDATION limits, ENCRYPTION settings.

## Conventions

- All source files use `.mjs` extension (ES modules)
- JSDoc required for all function declarations and exports
- Logging via pino to stderr (keeps stdout clean for MCP JSON-RPC)
- Format timestamps with `formatTimestamp()` for user display
- Pre-commit hooks auto-lint and format staged files

## Code Quality Philosophy

**Write code for the problem you have, not the problem you might have.**

Quality code is simple, clear, and correct. It solves today's problem well
without over-engineering for hypothetical futures.

### Craftsmanship

- **Clarity over cleverness** — Code is read more than written; optimize for the
  reader, not the writer
- **Names reveal intent** — Variables, functions, and files should describe what
  they represent or do; avoid abbreviations and generic names like `data`,
  `info`, `handle`
- **Comments explain why, not what** — Good code is self-documenting for _what_
  it does; comments explain _why_ it does it that way (business rules,
  workarounds, non-obvious constraints)
- **Fail fast, fail clearly** — Surface errors early with messages that help
  diagnose the problem; errors are user-facing, make them actionable
- **Security at boundaries** — Validate external input (user data, API
  responses) even when you trust the source; APIs change, bugs happen
- **Consistent patterns** — Follow conventions already in the codebase;
  consistency reduces cognitive load

### Pragmatism

- **Inline over abstract** — Don't extract functions until you have 3+ call
  sites with identical logic; two similar snippets are fine
- **Duplication over wrong abstraction** — Similar code serving different
  purposes may diverge; premature DRY creates coupling
- **Validate once** — Check data at system boundaries; internal code can trust
  what other internal code produces. Boundaries are the _only_ place you
  validate — if data crossed a boundary, downstream functions can trust it.
- **No speculative generality** — Don't add parameters, options, or
  configurability until a second use case exists
- **Delete, don't deprecate** — In a single-codebase project, remove unused
  code; don't leave compatibility shims or TODO comments

### Testing

- **Test behavior, not implementation** — Tests verify what a function promises
  (inputs → outputs), not how it works internally
- **Test at the right level** — Prefer integration tests for tools, unit tests
  for utilities; don't test private functions directly
- **Skip trivial tests** — Don't test that constants equal themselves or that
  simple getters return values

### When to Add Complexity

Add abstraction, configuration, or defensive code when:

- A real bug or user issue requires it
- The same pattern appears 3+ times with identical logic
- External/untrusted input needs validation
- The API or protocol explicitly requires it

### When to Refactor

Refactor existing code when:

- The current code actively causes bugs or makes changes risky
- Multiple features are blocked by the same structural problem
- Understanding the code requires archeological effort every time

Don't refactor just because code looks "ugly" or doesn't match your style
preferences. Working code that's locally messy but globally isolated is fine.

### Examples

**Good:** Clear, inline logic

```javascript
const style = NOTE_STYLES[noteType] || NOTE_STYLES.note;
const label = noteType.charAt(0).toUpperCase() + noteType.slice(1);
const noteHtml = `<div style="...">${label}: ${text}</div>`;
```

**Bad:** Premature abstraction with speculative options

```javascript
const noteHtml = createStyledNoteBlock(text, noteType, {
  includeTimestamp: true,
  wrapInDiv: true,
  customStyles: null,
});
```

**Good:** Direct error with context

```javascript
if (!parsedToken?.token) {
  throw new Error('Token file has invalid structure: missing "token" field');
}
```

**Bad:** Generic error that doesn't help

```javascript
if (!parsedToken?.token) {
  throw new Error('Invalid token');
}
```

**Good:** Simple, single-purpose function

```javascript
await writeFile(tokenPath, JSON.stringify(data), { mode: 0o600 });
```

**Bad:** Over-engineered for hypothetical failures

```javascript
await atomicWriteWithRetryAndRollback(tokenPath, data, {
  retries: 3,
  backoff: 'exponential',
  rollbackOnFailure: true,
});
```

**Bad:** Under-engineered, missing boundary validation

```javascript
async function getPage(pageId) {
  return await graphClient.api(`/pages/${pageId}`).get();
}
```

**Good:** Validates untrusted input at the boundary

```javascript
async function getPage(pageId) {
  validateId(pageId, 'page');
  return await graphClient.api(`/pages/${pageId}`).get();
}
```
