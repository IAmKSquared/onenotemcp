import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const serverPath = path.join(__dirname, 'onenote-mcp.mjs');

console.log(`Starting server at: ${serverPath}`);

const server = spawn('node', [serverPath], {
    stdio: ['pipe', 'pipe', 'inherit']
});

const request = {
    jsonrpc: "2.0",
    method: "tools/list",
    id: 1
};

server.stdout.on('data', (data) => {
    const output = data.toString();
    console.log('Received output:', output);

    try {
        // MCP uses JSON-RPC, output might be line-delimited JSON
        const lines = output.split('\n').filter(line => line.trim());
        for (const line of lines) {
            try {
                const json = JSON.parse(line);
                if (json.id === 1 && json.result && json.result.tools) {
                    const tools = json.result.tools;
                    const toolNames = tools.map(t => t.name);
                    console.log('Tools found:', toolNames);

                    const hasAuthenticate = toolNames.includes('authenticate');
                    const hasListNotebooks = toolNames.includes('listNotebooks');
                    const hasGetPageContent = toolNames.includes('getPageContent');

                    if (hasAuthenticate && hasListNotebooks && hasGetPageContent) {
                        console.log('✅ Verification PASSED: All critical tools found.');
                        process.exit(0);
                    } else {
                        console.error('❌ Verification FAILED: Missing tools.');
                        console.error(`Expected authenticate: ${hasAuthenticate}`);
                        console.error(`Expected listNotebooks: ${hasListNotebooks}`);
                        process.exit(1);
                    }
                }
            } catch (e) {
                // Ignore non-JSON lines or partial chunks
            }
        }
    } catch (error) {
        console.error('Error parsing output:', error);
        process.exit(1);
    }
});

server.on('error', (err) => {
    console.error('Failed to start server:', err);
    process.exit(1);
});

// Send the request
const requestString = JSON.stringify(request) + '\n';
console.log('Sending request:', requestString);
server.stdin.write(requestString);

// Timeout
setTimeout(() => {
    console.error('❌ Verification TIMEOUT: No response received.');
    server.kill();
    process.exit(1);
}, 5000);
