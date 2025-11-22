#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Client } from '@microsoft/microsoft-graph-client';
import { DeviceCodeCredential } from '@azure/identity';
import { JSDOM } from 'jsdom';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import fetch from 'node-fetch';
import { z } from "zod";
import crypto from 'crypto';

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '.access-token.txt');
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Default: Microsoft Graph Explorer App ID
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'];

// --- Encryption Configuration ---
const keyFilePath = path.join(__dirname, '.local-secret-key');
const ALGORITHM = 'aes-256-cbc';
const IV_LENGTH = 16;

// --- Global State ---
let accessToken = null;
let graphClient = null;

// --- MCP Server Initialization ---
const server = new McpServer({
  name: 'onenote',
  version: '1.0.0',
  description: 'OneNote MCP Server - Read, Write, and Edit OneNote content.'
});

// ============================================================================
// AUTHENTICATION & MICROSOFT GRAPH CLIENT MANAGEMENT
// ============================================================================

/**
 * Retrieves or generates the encryption key.
 * @returns {Buffer} The 32-byte encryption key.
 */
function getEncryptionKey() {
  if (fs.existsSync(keyFilePath)) {
    const hexKey = fs.readFileSync(keyFilePath, 'utf8').trim();
    return Buffer.from(hexKey, 'hex');
  } else {
    const newKey = crypto.randomBytes(32);
    fs.writeFileSync(keyFilePath, newKey.toString('hex'), { mode: 0o600 }); // Secure permissions
    console.error('🔑 Generated new secure encryption key.');
    return newKey;
  }
}

/**
 * Encrypts text using AES-256-CBC.
 * @param {string} text - The text to encrypt.
 * @returns {object} The encrypted data { iv, encryptedData }.
 */
function encrypt(text) {
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, getEncryptionKey(), iv);
  let encrypted = cipher.update(text);
  encrypted = Buffer.concat([encrypted, cipher.final()]);
  return { iv: iv.toString('hex'), encryptedData: encrypted.toString('hex') };
}

/**
 * Decrypts text using AES-256-CBC.
 * @param {object} text - The encrypted data object { iv, encryptedData }.
 * @returns {string} The decrypted text.
 */
function decrypt(text) {
  const iv = Buffer.from(text.iv, 'hex');
  const encryptedText = Buffer.from(text.encryptedData, 'hex');
  const decipher = crypto.createDecipheriv(ALGORITHM, getEncryptionKey(), iv);
  let decrypted = decipher.update(encryptedText);
  decrypted = Buffer.concat([decrypted, decipher.final()]);
  return decrypted.toString();
}

/**
 * Loads an existing access token from the local file system.
 */
function loadExistingToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const fileContent = fs.readFileSync(tokenFilePath, 'utf8');
      let tokenDataStr;

      try {
        // Try to parse as JSON first to see if it's our encrypted format
        const parsed = JSON.parse(fileContent);
        if (parsed.iv && parsed.encryptedData) {
          tokenDataStr = decrypt(parsed);
          console.error('🔓 Decrypted token successfully.');
        } else {
          // It's JSON but not encrypted (old format)
          tokenDataStr = fileContent;
          console.error('⚠️  Loaded unencrypted token (legacy format).');
        }
      } catch (e) {
        // Not JSON, likely plain text token (very old format)
        tokenDataStr = fileContent;
        console.error('⚠️  Loaded raw text token (legacy format).');
      }

      try {
        const parsedToken = JSON.parse(tokenDataStr);

        // Check if token has expired
        if (parsedToken.expiresOn) {
          const expiryDate = new Date(parsedToken.expiresOn);
          const now = new Date();

          if (expiryDate <= now) {
            console.error('⚠️  Token has expired. Please re-authenticate.');
            console.error(`   Expired on: ${expiryDate.toLocaleString()}`);
            return; // Don't set accessToken
          }

          const hoursUntilExpiry = Math.floor((expiryDate - now) / (1000 * 60 * 60));
          if (hoursUntilExpiry < 24) {
            console.error(`⏰ Token expires in ${hoursUntilExpiry} hours`);
          }
        }

        accessToken = parsedToken.token;
      } catch (parseError) {
        accessToken = tokenDataStr; // Old format: plain token string
      }
    }
  } catch (error) {
    console.error(`Error loading token: ${error.message}`);
  }
}

/**
 * Initializes the Microsoft Graph client if an access token is available.
 * @returns {Client | null} The initialized Graph client or null.
 */
function initializeGraphClient() {
  if (accessToken && !graphClient) {
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
    console.error('Microsoft Graph client initialized.');
  }
  return graphClient;
}

/**
 * Ensures the Graph client is initialized and authenticated.
 * Loads token if not present, then initializes client.
 * @throws {Error} If no access token is available after attempting to load.
 * @returns {Promise<Client>} The initialized and authenticated Graph client.
 */
async function ensureGraphClient() {
  if (!accessToken) {
    loadExistingToken();
  }
  if (!accessToken) {
    throw new Error('No access token available. Please authenticate first using the "authenticate" tool.');
  }
  if (!graphClient) {
    initializeGraphClient();
  }
  return graphClient;
}

// ============================================================================
// HTML CONTENT PROCESSING UTILITIES
// ============================================================================

/**
 * Extracts readable plain text from HTML content.
 * Removes scripts, styles, and formats headings, paragraphs, lists, and tables.
 * @param {string} html - The HTML content string.
 * @returns {string} The extracted readable text.
 */
function extractReadableText(html) {
  try {
    if (!html) return '';
    const dom = new JSDOM(html);
    const document = dom.window.document;

    document.querySelectorAll('script, style').forEach(element => element.remove());

    let text = '';
    document.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(heading => {
      const headingText = heading.textContent?.trim();
      if (headingText) text += `\n${headingText}\n${'-'.repeat(headingText.length)}\n`;
    });
    document.querySelectorAll('p').forEach(paragraph => {
      const content = paragraph.textContent?.trim();
      if (content) text += `${content}\n\n`;
    });
    document.querySelectorAll('ul, ol').forEach(list => {
      text += '\n';
      list.querySelectorAll('li').forEach((item, index) => {
        const content = item.textContent?.trim();
        if (content) text += `${list.tagName === 'OL' ? index + 1 + '.' : '-'} ${content}\n`;
      });
      text += '\n';
    });
    document.querySelectorAll('table').forEach(table => {
      text += '\n📊 Table content:\n';
      table.querySelectorAll('tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'))
          .map(cell => cell.textContent?.trim())
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
    console.error(`Error extracting readable text: ${error.message}`);
    return 'Error: Could not extract readable text from HTML content.';
  }
}

/**
 * Extracts a short summary from HTML content.
 * @param {string} html - The HTML content string.
 * @param {number} [maxLength=300] - The maximum length of the summary.
 * @returns {string} A text summary.
 */
function extractTextSummary(html, maxLength = 300) {
  try {
    if (!html) return 'No content to summarize.';
    const dom = new JSDOM(html);
    const document = dom.window.document;
    const bodyText = document.body?.textContent?.trim().replace(/\s+/g, ' ') || '';
    if (!bodyText) return 'No text content found in HTML body.';
    const summary = bodyText.substring(0, maxLength);
    return summary.length < bodyText.length ? `${summary}...` : summary;
  } catch (error) {
    console.error(`Error extracting text summary: ${error.message}`);
    return 'Could not extract text summary.';
  }
}

/**
 * Converts plain text (with simple markdown) to HTML.
 * @param {string} text - The plain text to convert.
 * @returns {string} The HTML representation.
 */
function textToHtml(text) {
  if (!text) return '';
  if (text.includes('<html>') || text.includes('<!DOCTYPE html>')) return text; // Already HTML

  let html = String(text) // Ensure text is a string
    .replace(/&/g, '&').replace(/</g, '<').replace(/>/g, '>') // Basic HTML escaping first
    .replace(/```([\s\S]*?)```/g, (match, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/__(.*?)__/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>').replace(/_(.*?)_/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
    .replace(/^---+$/gm, '<hr>')
    .replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>')
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

// ============================================================================
// ONENOTE API UTILITIES
// ============================================================================

/**
 * Escapes a string for use in OData queries to prevent injection attacks.
 * Single quotes in OData must be escaped by doubling them.
 * @param {string} str - The string to escape.
 * @returns {string} The escaped string safe for OData queries.
 */
function escapeODataString(str) {
  if (!str) return '';
  return str.replace(/'/g, "''");
}

/**
 * Fetches the content of a OneNote page.
 * @param {string} pageId - The ID of the page.
 * @param {'httpDirect' | 'direct'} [method='httpDirect'] - The method to use for fetching.
 * @returns {Promise<string>} The HTML content of the page.
 */
async function fetchPageContentAdvanced(pageId, method = 'httpDirect') {
  await ensureGraphClient();
  if (method === 'httpDirect') {
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!response.ok) throw new Error(`HTTP error fetching page content! Status: ${response.status} ${response.statusText}`);
    return await response.text();
  } else { // 'direct'
    return await graphClient.api(`/me/onenote/pages/${pageId}/content`).get();
  }
}

/**
 * Formats OneNote page information for display.
 * @param {object} page - The OneNote page object from Graph API.
 * @param {number | null} [index=null] - Optional index for numbered lists.
 * @returns {string} Formatted page information string.
 */
function formatPageInfo(page, index = null) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  const name = page.displayName || page.title || 'Untitled';
  return `${prefix}**${name}** (ID: ${page.id})`;
}

// ============================================================================
// TOOL HANDLER WRAPPER
// ============================================================================

/**
 * Creates a standardized tool handler with error handling and authentication checks.
 * @param {Function} handler - The async tool implementation function.
 * @param {string} errorPrefix - The prefix for error messages.
 * @returns {Function} A wrapped handler function compatible with McpServer.
 */
function createToolHandler(handler, errorPrefix = 'Tool execution failed') {
  return async (args) => {
    try {
      // 'authenticate' tool is a special case that establishes the session,
      // so we don't check for ensureGraphClient() inside it to avoid recursion/errors.
      // However, since we are wrapping it manually or not wrapping it at all, 
      // this check is mostly for safety if we decide to wrap it later.
      if (handler.name !== 'authenticateHandler') {
        await ensureGraphClient();
      }
      return await handler(args);
    } catch (error) {
      if (error.message.includes('authenticate') || error.message.includes('Access token')) {
        return { isError: true, content: [{ type: 'text', text: `🔐 Authentication Required. Please run the 'authenticate' tool first.` }] };
      }
      return { isError: true, content: [{ type: 'text', text: `❌ ${errorPrefix}: ${error.message}` }] };
    }
  };
}

// ============================================================================
// TOOL DEFINITIONS
// ============================================================================

// --- Authentication Tools ---

server.tool(
  'authenticate',
  {},
  async () => {
    try {
      console.error('Starting device code authentication...');
      let deviceCodeInfo = null;
      const credential = new DeviceCodeCredential({
        clientId: clientId,
        userPromptCallback: (info) => {
          deviceCodeInfo = info;
          console.error(`\n=== AUTHENTICATION REQUIRED ===\n${info.message}\n================================\n`);
        }
      });

      const authPromise = credential.getToken(scopes);
      await new Promise(resolve => setTimeout(resolve, 2000)); // Allow time for userPromptCallback

      if (deviceCodeInfo) {
        const authMessage = `🔐 **AUTHENTICATION REQUIRED**

Please complete the following steps:
1. **Open this URL in your browser:** https://microsoft.com/devicelogin
2. **Enter this code:** ${deviceCodeInfo.userCode}
3. **Sign in with your Microsoft account that has OneNote access.**
4. **After completing authentication, use the 'saveAccessToken' tool.**

Token will be saved automatically upon successful browser authentication.`;

        authPromise.then(tokenResponse => {
          accessToken = tokenResponse.token;
          const tokenData = {
            token: accessToken,
            clientId: clientId,
            scopes: scopes,
            createdAt: new Date().toISOString(),
            expiresOn: tokenResponse.expiresOnTimestamp ? new Date(tokenResponse.expiresOnTimestamp).toISOString() : null
          };

          // Encrypt before saving
          const encryptedToken = encrypt(JSON.stringify(tokenData));
          fs.writeFileSync(tokenFilePath, JSON.stringify(encryptedToken, null, 2));
          console.error('🔒 Token saved securely (encrypted).');
          initializeGraphClient();
        }).catch(error => {
          console.error(`Background authentication failed: ${error.message}`);
        });

        return { content: [{ type: 'text', text: authMessage }] };
      } else {
        return { isError: true, content: [{ type: 'text', text: 'Could not retrieve device code information. Please try again or check console logs.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Authentication failed: ${error.message}` }] };
    }
  }
);

server.tool(
  'saveAccessToken',
  {},
  async () => {
    try {
      loadExistingToken();
      if (accessToken) {
        initializeGraphClient();
        const testResponse = await graphClient.api('/me').get();
        return {
          content: [{
            type: 'text',
            text: `✅ **Authentication Successful!**
    Token loaded and verified.
**Account Info:**
    - Name: ${testResponse.displayName || 'Unknown'}
    - Email: ${testResponse.userPrincipalName || 'Unknown'}
🚀 You can now use OneNote tools!`
          }]
        };
      } else {
        return { isError: true, content: [{ type: 'text', text: `❌ **No Token Found.** Please run the 'authenticate' tool first.` }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to load or verify token: ${error.message}` }] };
    }
  }
);

// --- Page Reading Tools ---

server.tool(
  'listNotebooks',
  {},
  createToolHandler(async () => {
    const response = await graphClient.api('/me/onenote/notebooks').get();
    if (response.value && response.value.length > 0) {
      const notebookList = response.value.map((nb, i) => formatPageInfo(nb, i)).join('\n\n');
      return { content: [{ type: 'text', text: `📚 **Your OneNote Notebooks** (${response.value.length} found):\n\n${notebookList}` }] };
    } else {
      return { content: [{ type: 'text', text: '📚 No OneNote notebooks found.' }] };
    }
  }, 'Failed to list notebooks')
);

server.tool(
  'listSections',
  {
    notebookId: z.string().describe('The ID of the parent notebook.').optional(),
    sectionGroupId: z.string().describe('The ID of the parent section group.').optional()
  },
  createToolHandler(async ({ notebookId, sectionGroupId }) => {
    let endpoint = '/me/onenote/sections';
    if (notebookId) {
      endpoint = `/me/onenote/notebooks/${notebookId}/sections`;
    } else if (sectionGroupId) {
      endpoint = `/me/onenote/sectionGroups/${sectionGroupId}/sections`;
    }

    const response = await graphClient.api(endpoint).get();
    if (response.value && response.value.length > 0) {
      const list = response.value.map((item, i) => formatPageInfo(item, i)).join('\n\n');
      return { content: [{ type: 'text', text: `📂 **Sections** (${response.value.length} found):\n\n${list}` }] };
    } else {
      return { content: [{ type: 'text', text: '📂 No sections found.' }] };
    }
  }, 'Failed to list sections')
);

server.tool(
  'listSectionGroups',
  {
    notebookId: z.string().describe('The ID of the parent notebook.').optional(),
    sectionGroupId: z.string().describe('The ID of the parent section group.').optional()
  },
  createToolHandler(async ({ notebookId, sectionGroupId }) => {
    let endpoint = '/me/onenote/sectionGroups';
    if (notebookId) {
      endpoint = `/me/onenote/notebooks/${notebookId}/sectionGroups`;
    } else if (sectionGroupId) {
      endpoint = `/me/onenote/sectionGroups/${sectionGroupId}/sectionGroups`;
    }

    const response = await graphClient.api(endpoint).get();
    if (response.value && response.value.length > 0) {
      const list = response.value.map((item, i) => formatPageInfo(item, i)).join('\n\n');
      return { content: [{ type: 'text', text: `📁 **Section Groups** (${response.value.length} found):\n\n${list}` }] };
    } else {
      return { content: [{ type: 'text', text: '📁 No section groups found.' }] };
    }
  }, 'Failed to list section groups')
);

server.tool(
  'searchSections',
  {
    query: z.string().describe('The search term for section names.')
  },
  createToolHandler(async ({ query }) => {
    // Escape the query to prevent OData injection and convert to lowercase for case-insensitive search
    const escapedQuery = escapeODataString(query).toLowerCase();

    const response = await graphClient.api('/me/onenote/sections')
      .filter(`contains(tolower(displayName), '${escapedQuery}')`)
      .select('id,displayName,parentNotebook,parentSectionGroup')
      .top(50)
      .get();

    const sections = response.value || [];

    if (sections.length > 0) {
      // Display first 10 results, but keep all 50 fetched for accurate count
      const displaySections = sections.slice(0, 10);
      const list = displaySections.map((item, i) => formatPageInfo(item, i)).join('\n\n');
      const more = sections.length > 10 ? `\n\n... and ${sections.length - 10} more.` : '';
      const limitWarning = sections.length === 50 ? `\n\n⚠️ Note: Reached the 50-result limit. There may be additional matches not shown.` : '';
      return { content: [{ type: 'text', text: `🔍 **Section Search Results** for "${query}" (${sections.length} found):\n\n${list}${more}${limitWarning}` }] };
    } else {
      return { content: [{ type: 'text', text: `🔍 No sections found matching "${query}".` }] };
    }
  }, 'Failed to search sections')
);

server.tool(
  'searchPages',
  {
    query: z.string().describe('The search term for page titles.').optional()
  },
  createToolHandler(async ({ query }) => {
    let request = graphClient.api('/me/onenote/pages')
      .select('id,title,lastModifiedDateTime')
      .top(50);

    if (query) {
      // Escape the query to prevent OData injection and convert to lowercase for case-insensitive search
      const escapedQuery = escapeODataString(query).toLowerCase();
      request = request.filter(`contains(tolower(title), '${escapedQuery}')`);
    }

    const apiResponse = await request.get();
    const pages = apiResponse.value || [];

    if (pages.length > 0) {
      // Display first 10 results, but keep all 50 fetched for accurate count
      const displayPages = pages.slice(0, 10);
      const pageList = displayPages.map((page, i) => formatPageInfo(page, i)).join('\n\n');
      const more = pages.length > 10 ? `\n\n... and ${pages.length - 10} more pages.` : '';
      const limitWarning = pages.length === 50 ? `\n\n⚠️ Note: Reached the 50-result limit. There may be additional matches not shown.` : '';
      return { content: [{ type: 'text', text: `🔍 **Search Results** ${query ? `for "${query}"` : ''} (${pages.length} found):\n\n${pageList}${more}${limitWarning}` }] };
    } else {
      return { content: [{ type: 'text', text: query ? `🔍 No pages found matching "${query}".` : '📄 No pages found.' }] };
    }
  }, 'Failed to search pages')
);

server.tool(
  'getPageContent',
  {
    pageId: z.string().describe('The ID of the page to retrieve content from.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text (readable), html (raw), or summary (brief).')
      .optional()
  },
  createToolHandler(async ({ pageId, format }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
    let resultText = '';

    if (format === 'html') {
      resultText = `📄 **${pageInfo.title}** (HTML Format)\n\n${htmlContent}`;
    } else if (format === 'summary') {
      const summary = extractTextSummary(htmlContent, 300);
      resultText = `📄 **${pageInfo.title}** (Summary)\n\n${summary}`;
    } else { // 'text'
      const textContent = extractReadableText(htmlContent);
      resultText = `📄 **${pageInfo.title}**\n📅 Modified: ${new Date(pageInfo.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
    }
    return { content: [{ type: 'text', text: resultText }] };
  }, 'Failed to get page content')
);

server.tool(
  'getPageByTitle',
  {
    title: z.string().describe('The title (or partial title) of the page to find.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text, html, or summary.')
      .optional()
  },
  createToolHandler(async ({ title, format }) => {
    // Use server-side filtering with proper escaping and case-insensitive search
    const escapedTitle = escapeODataString(title).toLowerCase();
    const pagesResponse = await graphClient.api('/me/onenote/pages')
      .filter(`contains(tolower(title), '${escapedTitle}')`)
      .select('id,title,lastModifiedDateTime')
      .top(50)
      .get();

    const matchingPages = pagesResponse.value || [];

    if (matchingPages.length === 0) {
      // Fetch a few recent pages to show as alternatives
      const recentPages = await graphClient.api('/me/onenote/pages')
        .select('title')
        .top(10)
        .orderby('lastModifiedDateTime desc')
        .get();
      const availablePages = (recentPages.value || []).map(p => `- ${p.title}`).join('\n');
      return { isError: true, content: [{ type: 'text', text: `❌ No page found with title containing "${title}".\n\nRecent pages (up to 10):\n${availablePages || 'None'}` }] };
    }

    // Use the first matching page
    const matchingPage = matchingPages[0];
    const htmlContent = await fetchPageContentAdvanced(matchingPage.id, 'httpDirect');
    let resultText = '';
    if (format === 'html') {
      resultText = `📄 **${matchingPage.title}** (HTML Format)\n\n${htmlContent}`;
    } else if (format === 'summary') {
      const summary = extractTextSummary(htmlContent, 300);
      resultText = `📄 **${matchingPage.title}** (Summary)\n\n${summary}`;
    } else { // 'text'
      const textContent = extractReadableText(htmlContent);
      resultText = `📄 **${matchingPage.title}**\n📅 Modified: ${new Date(matchingPage.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
    }

    // If multiple matches, add a note
    if (matchingPages.length > 1) {
      resultText += `\n\n📌 Note: ${matchingPages.length} pages matched "${title}". Showing the first match.`;
      if (matchingPages.length === 50) {
        resultText += `\n⚠️ Reached the 50-result limit. There may be additional matches not shown.`;
      }
    }

    return { content: [{ type: 'text', text: resultText }] };
  }, 'Failed to get page by title')
);

server.tool(
  'updatePageContent',
  {
    pageId: z.string().describe('The ID of the page to update.'),
    content: z.string().describe('New page content (HTML or markdown-style text).'),
    preserveTitle: z.boolean()
      .default(true)
      .describe('Keep the original title (default: true).')
      .optional()
  },
  createToolHandler(async ({ pageId, content: newContent, preserveTitle }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    console.error(`Updating content for page: "${pageInfo.title}" (ID: ${pageId})`);

    const htmlContentForUpdate = textToHtml(newContent);
    const finalHtml = `
      <div>
        ${preserveTitle ? `<h1>${pageInfo.title}</h1>` : ''}
        ${htmlContentForUpdate}
        <hr>
        <p><em>Updated via OneNote MCP on ${new Date().toLocaleString()}</em></p>
      </div>
    `;

    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'body', action: 'replace', content: finalHtml }])
    });

    if (!response.ok) throw new Error(`Update failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **Page Content Updated!**\nPage: ${pageInfo.title}\nUpdated: ${new Date().toLocaleString()}\nContent Length: ${newContent.length} chars.` }] };
  }, 'Failed to update page content')
);

server.tool(
  'appendToPage',
  {
    pageId: z.string().describe('The ID of the page to append content to.'),
    content: z.string().describe('Content to append (HTML or markdown-style).'),
    addTimestamp: z.boolean().default(true).describe('Add a timestamp (default: true).').optional(),
    addSeparator: z.boolean().default(true).describe('Add a visual separator (default: true).').optional()
  },
  createToolHandler(async ({ pageId, content: newContent, addTimestamp, addSeparator }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    console.error(`Appending content to page: "${pageInfo.title}" (ID: ${pageId})`);

    const htmlContentToAppend = textToHtml(newContent);
    let appendHtml = '';
    if (addSeparator) appendHtml += '<hr>';
    if (addTimestamp) appendHtml += `<p><em>Added on ${new Date().toLocaleString()}</em></p>`;
    appendHtml += htmlContentToAppend;

    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'body', action: 'append', content: appendHtml }])
    });

    if (!response.ok) throw new Error(`Append failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **Content Appended!**\nPage: ${pageInfo.title}\nAppended: ${new Date().toLocaleString()}\nLength: ${newContent.length} chars.` }] };
  }, 'Failed to append content')
);

server.tool(
  'updatePageTitle',
  {
    pageId: z.string().describe('The ID of the page whose title is to be updated.'),
    newTitle: z.string().describe('The new title for the page.')
  },
  createToolHandler(async ({ pageId, newTitle }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    const oldTitle = pageInfo.title;
    console.error(`Updating page title from "${oldTitle}" to "${newTitle}" for page ID "${pageId}"`);

    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'title', action: 'replace', content: newTitle }])
    });

    if (!response.ok) throw new Error(`Title update failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **Page Title Updated!**\nOld Title: ${oldTitle}\nNew Title: ${newTitle}\nUpdated: ${new Date().toLocaleString()}` }] };
  }, 'Failed to update page title')
);

server.tool(
  'replaceTextInPage',
  {
    pageId: z.string().describe('The ID of the page to modify.'),
    findText: z.string().describe('The text to find and replace.'),
    replaceText: z.string().describe('The text to replace with.'),
    caseSensitive: z.boolean().default(false).describe('Case-sensitive search (default: false).').optional()
  },
  createToolHandler(async ({ pageId, findText, replaceText, caseSensitive }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
    console.error(`Replacing text in page: "${pageInfo.title}" (ID: ${pageId})`);

    const flags = caseSensitive ? 'g' : 'gi';
    const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    const matches = (htmlContent.match(regex) || []).length;

    if (matches === 0) {
      return { content: [{ type: 'text', text: `ℹ️ **No matches found** for "${findText}" in page: ${pageInfo.title}.` }] };
    }

    const updatedContent = htmlContent.replace(regex, replaceText);
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'body', action: 'replace', content: `<div>${updatedContent}</div>` }])
    });

    if (!response.ok) throw new Error(`Replace failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **Text Replaced!**\nPage: ${pageInfo.title}\nFound: "${findText}" (${matches} occurrences)\nReplaced with: "${replaceText}".` }] };
  }, 'Failed to replace text')
);

server.tool(
  'addNoteToPage',
  {
    pageId: z.string().describe('The ID of the page to add a note to.'),
    note: z.string().describe('The note/comment content.'),
    noteType: z.enum(['note', 'todo', 'important', 'question'])
      .default('note')
      .describe('Type of note (note, todo, important, question).')
      .optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the note (top or bottom).')
      .optional()
  },
  createToolHandler(async ({ pageId, note, noteType, position }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    console.error(`Adding ${noteType} to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);

    const icons = { note: '📝', todo: '✅', important: '🚨', question: '❓' };
    const colors = { note: '#e3f2fd', todo: '#e8f5e8', important: '#ffebee', question: '#fff3e0' };
    const noteHtml = `
      <div style="border-left: 4px solid #2196f3; background-color: ${colors[noteType]}; padding: 10px; margin: 10px 0;">
        <p><strong>${icons[noteType]} ${noteType.charAt(0).toUpperCase() + noteType.slice(1)}</strong> - <em>${new Date().toLocaleString()}</em></p>
        <p>${textToHtml(note)}</p>
      </div>`;

    const action = position === 'top' ? 'prepend' : 'append';
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'body', action: action, content: noteHtml }])
    });

    if (!response.ok) throw new Error(`Add note failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **${noteType.charAt(0).toUpperCase() + noteType.slice(1)} Added!**\nPage: ${pageInfo.title}\nPosition: ${position}.` }] };
  }, 'Failed to add note')
);

server.tool(
  'addTableToPage',
  {
    pageId: z.string().describe('The ID of the page to add a table to.'),
    tableData: z.string().describe('Table data in CSV format (header row, then data rows).'),
    title: z.string().describe('Optional title for the table.').optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the table (top or bottom).')
      .optional()
  },
  createToolHandler(async ({ pageId, tableData, title, position }) => {
    const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
    console.error(`Adding table to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);

    const rows = tableData.trim().split('\n').map(row => row.split(',').map(cell => cell.trim()));
    if (rows.length < 2) throw new Error('Table data must have at least a header row and one data row.');

    const headerRow = rows[0];
    const dataRows = rows.slice(1);
    let tableHtml = title ? `<h3>📊 ${textToHtml(title)}</h3>` : '';
    tableHtml += `<table style="border-collapse: collapse; width: 100%; margin: 10px 0;"><thead><tr style="background-color: #f5f5f5;">${headerRow.map(cell => `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${textToHtml(cell)}</th>`).join('')}</tr></thead><tbody>${dataRows.map(row => `<tr>${row.map(cell => `<td style="border: 1px solid #ddd; padding: 8px;">${textToHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

    const action = position === 'top' ? 'prepend' : 'append';
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      method: 'PATCH',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify([{ target: 'body', action: action, content: tableHtml }])
    });

    if (!response.ok) throw new Error(`Add table failed: ${response.status} ${response.statusText}`);

    return { content: [{ type: 'text', text: `✅ **Table Added!**\nPage: ${pageInfo.title}\nTitle: ${title || 'Untitled'}\nPosition: ${position}.` }] };
  }, 'Failed to add table')
);

// --- Page Creation Tool ---
server.tool(
  'createPage',
  {
    title: z.string().min(1, { message: "Title cannot be empty." }).describe('The title for the new page.'),
    content: z.string().min(1, { message: "Content cannot be empty." }).describe('The content for the new page (HTML or markdown-style).')
  },
  createToolHandler(async ({ title, content }) => {
    console.error(`Attempting to create page with title: "${title}"`);

    const sectionsResponse = await graphClient.api('/me/onenote/sections').get();
    if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
      throw new Error('No sections found in your OneNote. Cannot create a page.');
    }
    const targetSectionId = sectionsResponse.value[0].id;
    const targetSectionName = sectionsResponse.value[0].displayName;

    const htmlContent = textToHtml(content);
    const pageHtml = `<!DOCTYPE html>
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

    const response = await graphClient
      .api(`/me/onenote/sections/${targetSectionId}/pages`)
      .header('Content-Type', 'application/xhtml+xml')
      .post(pageHtml);

    return {
      content: [{
        type: 'text',
        text: `✅ **Page Created Successfully!**
**Title:** ${response.title}
**Page ID:** ${response.id}
**In Section:** ${targetSectionName}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`
      }]
    };
  }, 'Error creating page')
);

// ============================================================================
// SERVER STARTUP
// ============================================================================

/**
 * Main function to initialize and start the MCP server.
 */
async function main() {
  loadExistingToken(); // Attempt to load token at startup
  if (accessToken) {
    initializeGraphClient(); // Initialize client if token was loaded
  }

  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);

    console.error('🚀✨ OneNote Ultimate MCP Server is now LIVE! ✨🚀');
    console.error(`   Client ID: ${clientId.substring(0, 8)}... (Using ${process.env.AZURE_CLIENT_ID ? 'environment variable' : 'default'})`);
    console.error('   Ready to manage your OneNote like never before!');
    console.error('--- Available Tool Categories ---');
    console.error('   🔐 Auth: authenticate, saveAccessToken');
    console.error('   📚 Read: listNotebooks, searchPages, getPageContent, getPageByTitle, listSections, listSectionGroups, searchSections');
    console.error('   ✏️ Edit: updatePageContent, appendToPage, updatePageTitle, replaceTextInPage, addNoteToPage, addTableToPage');
    console.error('   ➕ Create: createPage');
    console.error('---------------------------------');

    process.on('SIGINT', () => {
      console.error('\n🔌 OneNote MCP Server shutting down gracefully...');
      process.exit(0);
    });
    process.on('SIGTERM', () => {
      console.error('\n🔌 OneNote MCP Server terminated...');
      process.exit(0);
    });

  } catch (error) {
    console.error(`💀 Critical error starting server: ${error.message}`, error.stack);
    process.exit(1);
  }
}

main();