#!/usr/bin/env node

/**
 * AgentDemo Auth Standalone
 * Authenticates with M365 and saves browser session to ~/.agentdemo/.browser-session/
 * Run: node scripts/auth-standalone.js
 */

// Pre-flight: ensure npm dependencies are installed before importing them.
// This keeps auth-standalone runnable even if users skipped `npm install`.
import { resolve as _resolve, dirname as _dirname, join as _join } from 'path';
import { fileURLToPath as _fileURLToPath } from 'url';
import { existsSync as _existsSync } from 'fs';
import { execSync as _execSync } from 'child_process';

const __preflightDir = _dirname(_fileURLToPath(import.meta.url));
const __preflightRoot = _resolve(__preflightDir, '..');
const __preflightNm = _join(__preflightRoot, 'node_modules');

if (!_existsSync(__preflightNm)) {
  console.log('');
  console.log('Installing dependencies (first-time setup)...');
  console.log('This will take 1–2 minutes.');
  console.log('');
  try {
    _execSync('npm install', { cwd: __preflightRoot, stdio: 'inherit' });
    console.log('');
  } catch (err) {
    console.error('');
    console.error('npm install failed. Please run manually:');
    console.error(`  cd "${__preflightRoot}"`);
    console.error('  npm install');
    process.exit(1);
  }
}

const { config } = await import('dotenv');
import { resolve, dirname, join } from 'path';
import { fileURLToPath } from 'url';
import { existsSync, mkdirSync } from 'fs';
import os from 'os';

const __dirname = dirname(fileURLToPath(import.meta.url));

// Load .env from multiple locations (first match wins)
const envLocations = [
  resolve(__dirname, '.env'),
  resolve(__dirname, '../.env'),
  join(os.homedir(), '.agentdemo', '.env'),
  resolve(process.cwd(), '.env'),
];

let envLoaded = false;
for (const envPath of envLocations) {
  if (existsSync(envPath)) {
    config({ path: envPath });
    console.log('Loaded .env from:', envPath);
    envLoaded = true;
    break;
  }
}

if (!envLoaded) {
  console.error('No .env file found. Run setup first: node scripts/setup.js');
  process.exit(1);
}

if (!process.env.DEMO_EMAIL || !process.env.DEMO_PASSWORD) {
  console.error('DEMO_EMAIL and DEMO_PASSWORD must be set in .env');
  process.exit(1);
}

// Ensure ~/.agentdemo/.browser-session/ exists
const homeSessionDir = join(os.homedir(), '.agentdemo', '.browser-session');
if (!existsSync(homeSessionDir)) {
  mkdirSync(homeSessionDir, { recursive: true });
}

// Patch the session dir to use ~/.agentdemo/.browser-session/ before importing auth
// We do this by setting an env var that auth.js can check
process.env.AGENTDEMO_SESSION_DIR = homeSessionDir;

// Import and run auth
const { chromium } = await import('playwright');
import fs from 'fs';
import path from 'path';
import readline from 'readline';

const STATE_FILE = path.join(homeSessionDir, 'state.json');

function waitForKeypress(message) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(message, () => {
      rl.close();
      resolve();
    });
  });
}

async function main() {
  const email = process.env.DEMO_EMAIL;
  const password = process.env.DEMO_PASSWORD;

  console.log('');
  console.log(`Authenticating as: ${email}`);
  console.log('');

  const contextOpts = {
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    viewport: null,
  };

  // Load existing session if available
  if (existsSync(STATE_FILE)) {
    contextOpts.storageState = STATE_FILE;
  }

  const browser = await chromium.launch({
    headless: false,
    channel: 'chromium',
    args: ['--start-maximized'],
  });
  const context = await browser.newContext(contextOpts);

  // Check if already logged in
  const checkPage = await context.newPage();
  try {
    await checkPage.goto('https://m365.cloud.microsoft', {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await checkPage.waitForTimeout(3000);
    const checkUrl = checkPage.url();
    if (!checkUrl.includes('login.microsoftonline.com') && !checkUrl.includes('login.live.com')) {
      await context.storageState({ path: STATE_FILE });
      await checkPage.close();
      await browser.close();
      console.log(`Already logged in as ${email} — session saved to ${homeSessionDir}`);
      return;
    }
  } catch { /* not yet logged in */ }
  await checkPage.close();

  // Fresh login
  const page = await context.newPage();
  await page.goto('https://login.microsoftonline.com', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 60000 });

  // Enter email
  const EMAIL_SELECTORS = [
    'input[type="email"]',
    'input[name="loginfmt"]',
    'input[id="i0116"]',
    '[data-testid="i0116"]',
  ];
  let emailFilled = false;
  for (const sel of EMAIL_SELECTORS) {
    try {
      await page.waitForSelector(sel, { timeout: 60000 });
      await page.fill(sel, email);
      emailFilled = true;
      break;
    } catch { /* try next */ }
  }
  if (!emailFilled) {
    throw new Error('Could not find email input on login page.');
  }

  await page.click('input[type="submit"]');
  await page.waitForTimeout(2000);

  // Enter password
  try {
    await page.waitForSelector('input[type="password"]', { timeout: 60000 });
    await page.fill('input[type="password"]', password);
    await page.click('input[type="submit"]');
  } catch { /* different auth flow */ }

  await page.waitForTimeout(3000);

  // Check for MFA
  const mfaIndicators = [
    '#idDiv_SAOTCAS_Title',
    'text="Approve sign in request"',
    'text="Verify your identity"',
    'text="Enter code"',
    '#idDiv_SAOTCC_Description',
  ];

  let mfaDetected = false;
  for (const selector of mfaIndicators) {
    try {
      const el = await page.$(selector);
      if (el) { mfaDetected = true; break; }
    } catch { /* ignore */ }
  }

  if (mfaDetected) {
    console.log('');
    await waitForKeypress(
      'ACTION REQUIRED: Complete MFA in the browser window, then press Enter to continue...'
    );
    await page.waitForTimeout(3000);
  }

  // Handle "Stay signed in?" prompt
  try {
    const staySignedIn = await page.$('text="Stay signed in?"');
    if (staySignedIn) {
      await page.click('input[type="submit"]');
      await page.waitForTimeout(2000);
    }
  } catch { /* ignore */ }

  // Verify login
  await page.waitForTimeout(3000);
  const currentUrl = page.url();
  if (currentUrl.includes('login.microsoftonline.com')) {
    throw new Error('Login did not complete. Check credentials and try again.');
  }

  await context.storageState({ path: STATE_FILE });
  await page.close();
  await browser.close();

  console.log('');
  console.log(`Authenticated as ${email}`);
  console.log(`Session saved to: ${homeSessionDir}`);
}

main().catch((err) => {
  console.error('Authentication failed:', err.message);
  process.exit(1);
});
