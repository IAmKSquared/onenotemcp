#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { OneNoteSession } from './session.mjs';
import { getClientId } from './auth/device-code-flow.mjs';
import { registerAuthTools } from './tools/auth-tools.mjs';
import { registerReadTools } from './tools/read-tools.mjs';
import { registerWriteTools } from './tools/write-tools.mjs';
import { registerCreateTools } from './tools/create-tools.mjs';
import { registerDeleteTools } from './tools/delete-tools.mjs';
import { logger } from './utils/logger.mjs';

// --- MCP Server Initialization ---
const server = new McpServer({
  name: 'onenote',
  version: '1.0.0',
  description: 'OneNote MCP Server - Read, Write, and Edit OneNote content.',
});

// --- Session Initialization ---
const session = new OneNoteSession();

// Register all tool categories with session
registerAuthTools(server, session);
registerReadTools(server, session);
registerWriteTools(server, session);
registerCreateTools(server, session);
registerDeleteTools(server, session);

/**
 * Main function to initialize and start the MCP server.
 */
async function main() {
  await session.loadExistingToken(); // Attempt to load token at startup
  session.initializeGraphClient(); // Initialize client if token was loaded

  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);

    const clientId = getClientId();
    logger.info('🚀✨ OneNote Ultimate MCP Server is now LIVE! ✨🚀');
    logger.info(
      `   Client ID: ${clientId.substring(0, 8)}... (Using ${process.env.AZURE_CLIENT_ID ? 'environment variable' : 'default'})`
    );
    logger.info('   Ready to manage your OneNote like never before!');
    logger.info('--- Available Tool Categories ---');
    logger.info('   🔐 Auth: authenticate, saveAccessToken');
    logger.info(
      '   📚 Read: listNotebooks, searchPages, getPageContent, getPageByTitle, listSections, listSectionGroups, searchSections'
    );
    logger.info(
      '   ✏️ Edit: updatePageContent, appendToPage, updatePageTitle, replaceTextInPage, addNoteToPage, addTableToPage'
    );
    logger.info(
      '   ➕ Create: createPage, createPageInSection, createNotebook, createSection, createSectionGroup'
    );
    logger.info('   📋 Manage: copyPage');
    logger.info('   🗑️ Delete: deletePage');
    logger.info('---------------------------------');

    process.on('SIGINT', () => {
      logger.info('\n🔌 OneNote MCP Server shutting down gracefully...');
      process.exit(0);
    });
    process.on('SIGTERM', () => {
      logger.info('\n🔌 OneNote MCP Server terminated...');
      process.exit(0);
    });
  } catch (error) {
    logger.fatal({ err: error }, '💀 Critical error starting server');
    process.exit(1);
  }
}

main();
