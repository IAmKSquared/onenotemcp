import path from 'path';
import { fileURLToPath } from 'url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const serverPath = path.join(__dirname, '..', 'src', 'server.mjs');

// Critical tools to verify
const expectedTools = [
  // Authentication
  'authenticate',
  'saveAccessToken',
  // Reading
  'listNotebooks',
  'listSections',
  'listSectionGroups',
  'searchSections',
  'listPagesInSection',
  'searchPages',
  'getRecentPages',
  'getPageContent',
  'getPageByTitle',
  'getPageLink',
  // Creating Structure
  'createNotebook',
  'createSection',
  'createSectionGroup',
  // Page Creation & Editing
  'createPage',
  'createPageInSection',
  'updatePageContent',
  'appendToPage',
  'updatePageTitle',
  'replaceTextInPage',
  'addNoteToPage',
  'addTableToPage',
  // Page Management
  'copyPage',
  // Delete
  'deletePage',
];

// Guard against hangs so CI can never block indefinitely.
const timeout = setTimeout(() => {
  console.error('❌ Verification TIMEOUT: No response received.');
  process.exit(1);
}, 10000);

console.log(`Starting server at: ${serverPath}`);

const transport = new StdioClientTransport({
  command: 'node',
  args: [serverPath],
});
const client = new Client({ name: 'smoke-test', version: '1.0.0' });

try {
  // connect() performs the initialize/initialized handshake automatically.
  await client.connect(transport);

  const { tools } = await client.listTools();
  const toolNames = tools.map((t) => t.name);
  console.log('Tools found:', toolNames);

  const missing = expectedTools.filter((tool) => !toolNames.includes(tool));

  clearTimeout(timeout);
  await client.close();

  if (missing.length === 0) {
    console.log(`✅ Verification PASSED: All ${expectedTools.length} expected tools found.`);
    process.exit(0);
  } else {
    console.error('❌ Verification FAILED: Missing tools:');
    missing.forEach((tool) => console.error(`  - ${tool}`));
    console.error(`\nFound ${toolNames.length} tools, expected ${expectedTools.length}`);
    process.exit(1);
  }
} catch (error) {
  clearTimeout(timeout);
  console.error('❌ Verification FAILED:', error);
  await client.close().catch(() => {});
  process.exit(1);
}
