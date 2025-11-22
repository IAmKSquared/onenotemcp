import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const serverPath = path.join(__dirname, 'onenote-mcp.mjs');
const serverProcess = spawn('node', [serverPath], {
  stdio: ['pipe', 'pipe', 'inherit'], // Pipe stdin/stdout, inherit stderr for logs
});

let requestCount = 0;

/**
 *
 * @param method
 * @param params
 */
function sendRequest(method, params = {}) {
  const id = requestCount++;
  const request = {
    jsonrpc: '2.0',
    id,
    method,
    params,
  };
  const json = JSON.stringify(request);
  // console.error(`Sending: ${json}`);
  serverProcess.stdin.write(json + '\n');
}

serverProcess.stdout.on('data', (data) => {
  const lines = data
    .toString()
    .split('\n')
    .filter((line) => line.trim());
  for (const line of lines) {
    try {
      const response = JSON.parse(line);
      // console.error(`Received: ${JSON.stringify(response, null, 2)}`);
      handleResponse(response);
    } catch (_e) {
      // console.error(`Non-JSON output: ${line}`);
    }
  }
});

/**
 *
 * @param response
 */
function handleResponse(response) {
  if (response.id === 0) {
    // Initialize response
    console.log('✅ Server initialized');
    sendRequest('tools/list');
  } else if (response.id === 1) {
    // tools/list response
    const tools = response.result.tools;
    const toolNames = tools.map((t) => t.name);
    console.log('Available tools:', toolNames);

    const expectedTools = [
      // Section tools
      'listSections',
      'listSectionGroups',
      'searchSections',
      // Creation tools
      'createNotebook',
      'createSection',
      'createSectionGroup',
      // Page management
      'copyPage',
      // Advanced editing
      'replaceTextInPage',
      'addNoteToPage',
      'addTableToPage',
    ];
    const missing = expectedTools.filter((t) => !toolNames.includes(t));

    if (missing.length === 0) {
      console.log(`✅ All ${expectedTools.length} expected tools are present.`);
      // Call listSections
      sendRequest('tools/call', { name: 'listSections', arguments: {} });
    } else {
      console.error('❌ Missing tools:', missing);
      process.exit(1);
    }
  } else if (response.id === 2) {
    // listSections response
    if (response.error) {
      console.error('❌ listSections failed:', response.error);
    } else {
      console.log('✅ listSections response received.');
      // console.log(response.result.content[0].text);

      // Call searchSections
      sendRequest('tools/call', { name: 'searchSections', arguments: { query: 'Test' } });
    }
  } else if (response.id === 3) {
    // searchSections response
    if (response.error) {
      console.error('❌ searchSections failed:', response.error);
    } else {
      console.log('✅ searchSections response received.');
      // Test searchPages with date filtering
      const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
      sendRequest('tools/call', {
        name: 'searchPages',
        arguments: {
          query: '',
          modifiedAfter: oneWeekAgo,
        },
      });
    }
  } else if (response.id === 4) {
    // searchPages with date filter response
    if (response.error) {
      console.error('❌ searchPages with date filter failed:', response.error);
    } else {
      console.log('✅ searchPages with date filter response received.');
      console.log('🎉 Integration Test Complete!');
      process.exit(0);
    }
  }
}

// Start sequence
sendRequest('initialize', {
  protocolVersion: '2024-11-05',
  capabilities: {},
  clientInfo: { name: 'tester', version: '1.0.0' },
});

// Timeout
setTimeout(() => {
  console.error('❌ Timeout waiting for responses.');
  process.exit(1);
}, 10000);
