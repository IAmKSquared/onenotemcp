#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadExistingToken } from './auth/token-manager.mjs';
import { initializeGraphClient } from './api/graph-client.mjs';
import { getClientId } from './auth/device-code-flow.mjs';
import { registerAuthTools } from './tools/auth-tools.mjs';
import { registerReadTools } from './tools/read-tools.mjs';
import { registerWriteTools } from './tools/write-tools.mjs';
import { registerCreateTools } from './tools/create-tools.mjs';

// --- MCP Server Initialization ---
const server = new McpServer({
  name: 'onenote',
  version: '1.0.0',
  description: 'OneNote MCP Server - Read, Write, and Edit OneNote content.',
});

// Register all tool categories
registerAuthTools(server);
registerReadTools(server);
registerWriteTools(server);
registerCreateTools(server);

/**
 * Main function to initialize and start the MCP server.
 */
async function main() {
  await loadExistingToken(); // Attempt to load token at startup
  initializeGraphClient(); // Initialize client if token was loaded

  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);

    const clientId = getClientId();
    console.error('🚀✨ OneNote Ultimate MCP Server is now LIVE! ✨🚀');
    console.error(
      `   Client ID: ${clientId.substring(0, 8)}... (Using ${process.env.AZURE_CLIENT_ID ? 'environment variable' : 'default'})`
    );
    console.error('   Ready to manage your OneNote like never before!');
    console.error('--- Available Tool Categories ---');
    console.error('   🔐 Auth: authenticate, saveAccessToken');
    console.error(
      '   📚 Read: listNotebooks, searchPages, getPageContent, getPageByTitle, listSections, listSectionGroups, searchSections'
    );
    console.error(
      '   ✏️ Edit: updatePageContent, appendToPage, updatePageTitle, replaceTextInPage, addNoteToPage, addTableToPage'
    );
    console.error(
      '   ➕ Create: createPage, createPageInSection, createNotebook, createSection, createSectionGroup'
    );
    console.error('   📋 Manage: copyPage');
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
