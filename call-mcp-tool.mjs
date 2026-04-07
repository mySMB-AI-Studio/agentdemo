/**
 * Temporary script: call create_demo via MCP stdio protocol
 */
import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Read tool name and args from command line: node call-mcp-tool.mjs <tool_name> '<json_args>'
const toolName = process.argv[2] || 'create_demo';
const args = process.argv[3] ? JSON.parse(process.argv[3]) : {};

const server = spawn('node', ['src/mcp-server.js'], {
  cwd: __dirname,
  stdio: ['pipe', 'pipe', 'pipe'],
  env: { ...process.env },
});

let lineBuffer = '';
let toolCallSent = false;
let toolResponseReceived = false;

function processLine(line) {
  const trimmed = line.trim();
  if (!trimmed) return;

  // Print all non-JSON lines as console output
  try {
    const msg = JSON.parse(trimmed);

    if (msg.id === 1 && msg.result && !toolCallSent) {
      // Initialize response: now send the tool call exactly once
      toolCallSent = true;
      console.log('[CLIENT] MCP server initialized. Sending create_demo tool call...');
      const callRequest = JSON.stringify({
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/call',
        params: {
          name: toolName,
          arguments: args,
        },
      }) + '\n';
      server.stdin.write(callRequest);

    } else if (msg.id === 2 && !toolResponseReceived) {
      // Tool call response
      toolResponseReceived = true;
      console.log('\n=== MCP Tool Response (JSON-RPC) ===');
      console.log(JSON.stringify(msg, null, 2));
      setTimeout(() => {
        server.kill();
        process.exit(0);
      }, 500);
    }
  } catch {
    // Not JSON — it's console output from the tool, print it
    console.log('[TOOL OUTPUT] ' + line);
  }
}

server.stdout.on('data', (data) => {
  lineBuffer += data.toString();
  const lines = lineBuffer.split('\n');
  // Keep the last potentially incomplete line in buffer
  lineBuffer = lines.pop();
  for (const line of lines) {
    processLine(line);
  }
});

server.stderr.on('data', (data) => {
  process.stderr.write('[MCP stderr] ' + data.toString());
});

server.on('close', (code) => {
  // Process any remaining buffer
  if (lineBuffer.trim()) processLine(lineBuffer);
  console.log(`\n[MCP server] exited with code ${code}`);
});

server.on('error', (err) => {
  console.error('Failed to start MCP server:', err);
  process.exit(1);
});

// Send initialize request
const initRequest = JSON.stringify({
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2024-11-05',
    capabilities: {},
    clientInfo: { name: 'test-client', version: '1.0.0' },
  },
}) + '\n';

server.stdin.write(initRequest);

// Safety timeout: 10 minutes
setTimeout(() => {
  console.log('\n[CLIENT] Timeout reached (10 min). Killing server.');
  server.kill();
  process.exit(1);
}, 600000);
