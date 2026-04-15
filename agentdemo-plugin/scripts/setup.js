#!/usr/bin/env node

/**
 * AgentDemo Setup Script
 * Creates ~/.agentdemo/.env with user credentials.
 * Run this after installing the plugin: node scripts/setup.js
 */

import readline from 'readline';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';
import { dirname, resolve, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

// Check all known .env locations
const envLocations = [
  resolve(__dirname, '.env'),
  resolve(__dirname, '../.env'),
  join(os.homedir(), '.agentdemo', '.env'),
  resolve(process.cwd(), '.env'),
];

function ask(rl, question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => resolve(answer.trim()));
  });
}

async function main() {
  console.log('');
  console.log('AgentDemo Setup');
  console.log('───────────────');
  console.log('');

  // Check if .env already exists somewhere
  for (const loc of envLocations) {
    if (fs.existsSync(loc)) {
      console.log(`Found existing .env at: ${loc}`);
      const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
      const overwrite = await ask(rl, 'Overwrite? (y/N): ');
      rl.close();
      if (overwrite.toLowerCase() !== 'y') {
        console.log('Setup cancelled. Existing .env will be used.');
        return;
      }
      break;
    }
  }

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  console.log('Enter your M365 demo account details:');
  console.log('');

  const email = await ask(rl, '  Email: ');
  const password = await ask(rl, '  Password: ');
  const tenant = await ask(rl, '  Tenant (e.g. mysmb): ');
  const apiKey = await ask(rl, '  Anthropic API key (optional, press Enter to skip): ');

  rl.close();

  // Build .env content
  const lines = [
    `DEMO_EMAIL=${email}`,
    `DEMO_PASSWORD=${password}`,
    `DEMO_TENANT=${tenant}`,
    `FFMPEG_PATH=`,
    `DEMO_ENVIRONMENT=`,
    `ANTHROPIC_API_KEY=${apiKey}`,
  ];

  // Save to ~/.agentdemo/.env
  const agentdemoDir = join(os.homedir(), '.agentdemo');
  const envPath = join(agentdemoDir, '.env');

  if (!fs.existsSync(agentdemoDir)) {
    fs.mkdirSync(agentdemoDir, { recursive: true });
  }

  fs.writeFileSync(envPath, lines.join('\n') + '\n', 'utf8');

  console.log('');
  console.log(`Setup complete. Credentials saved to: ${envPath}`);
  console.log('');
  console.log('Next step — authenticate with M365:');
  console.log('  node scripts/auth-standalone.js');
}

main().catch((err) => {
  console.error('Setup failed:', err.message);
  process.exit(1);
});
