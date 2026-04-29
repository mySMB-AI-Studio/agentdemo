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

// Parse --profile flag (e.g. --profile recipient)
const profileArgIndex = process.argv.indexOf('--profile');
const profile = profileArgIndex !== -1 ? process.argv[profileArgIndex + 1] : null;

// Load .env — profile-specific first, then fallback locations
// Profile .env: ~/.agentdemo/.env-{profile}  (e.g. ~/.agentdemo/.env-recipient)
const envLocations = [
  resolve(__dirname, '.env'),
  resolve(__dirname, '../.env'),
  ...(profile ? [join(os.homedir(), '.agentdemo', `.env-${profile}`)] : []),
  join(os.homedir(), '.agentdemo', '.env'),
  resolve(process.cwd(), '.env'),
];

let envLoaded = false;
for (const envPath of envLocations) {
  if (existsSync(envPath)) {
    config({ path: envPath });
    const label = profile ? ` (profile: ${profile})` : '';
    console.log(`Loaded .env from: ${envPath}${label}`);
    envLoaded = true;
    break;
  }
}

if (!envLoaded) {
  console.error('No .env file found. Run setup first: node scripts/setup.js');
  process.exit(1);
}

// Resolve credentials — named profiles use PROFILE_{NAME}_EMAIL / PROFILE_{NAME}_PASSWORD
// falling back to DEMO_EMAIL / DEMO_PASSWORD if the profile-specific vars are absent.
let resolvedEmail = process.env.DEMO_EMAIL;
let resolvedPassword = process.env.DEMO_PASSWORD;
if (profile && profile !== 'default') {
  const envPrefix = `PROFILE_${profile.toUpperCase()}`;
  if (process.env[`${envPrefix}_EMAIL`])    resolvedEmail    = process.env[`${envPrefix}_EMAIL`];
  if (process.env[`${envPrefix}_PASSWORD`]) resolvedPassword = process.env[`${envPrefix}_PASSWORD`];
}

if (!resolvedEmail || !resolvedPassword) {
  const varHint = profile && profile !== 'default'
    ? `PROFILE_${profile.toUpperCase()}_EMAIL / PROFILE_${profile.toUpperCase()}_PASSWORD (or DEMO_EMAIL / DEMO_PASSWORD as fallback)`
    : 'DEMO_EMAIL and DEMO_PASSWORD';
  console.error(`Credentials not found. Set ${varHint} in ~/.agentdemo/.env`);
  process.exit(1);
}

// Resolve session directory — named profiles get their own folder
const homeSessionDir = (profile && profile !== 'default')
  ? join(os.homedir(), '.agentdemo', `.browser-session-${profile}`)
  : join(os.homedir(), '.agentdemo', '.browser-session');

if (!existsSync(homeSessionDir)) {
  mkdirSync(homeSessionDir, { recursive: true });
}

// Set session dir env var (auth.js reads this for the default profile)
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

async function checkStatus() {
  const email = resolvedEmail || '(unknown)';

  console.log('');
  console.log('AgentDemo Auth Status');
  console.log('─────────────────────');
  if (profile) console.log(`Profile:  ${profile}`);
  console.log(`Account:  ${email}`);
  console.log(`Session:  ${homeSessionDir}`);

  if (!existsSync(STATE_FILE)) {
    console.log('Status: NOT AUTHENTICATED');
    console.log('');
    const runCmd = profile ? `node scripts/auth-standalone.js --profile ${profile}` : 'node scripts/auth-standalone.js';
    console.log(`No saved session found. Run: ${runCmd}`);
    process.exit(1);
  }

  console.log(`State file: ${STATE_FILE}`);
  console.log('Checking session validity...');

  const browser = await chromium.launch({
    headless: true,
    channel: 'chromium',
  });
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    viewport: { width: 1920, height: 1080 },
    storageState: STATE_FILE,
  });

  const page = await context.newPage();
  let isValid = false;
  try {
    await page.goto('https://www.office.com', {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await page.waitForTimeout(3000);
    const url = page.url();
    isValid = !url.includes('login.microsoftonline.com') && !url.includes('login.live.com');
  } catch (err) {
    console.log(`Validation error: ${err.message}`);
  }

  await browser.close();

  if (isValid) {
    console.log('Status: AUTHENTICATED');
    console.log('');
    console.log(`Session is valid for ${email}`);
  } else {
    console.log('Status: EXPIRED');
    console.log('');
    const rerunCmd = profile ? `node scripts/auth-standalone.js --profile ${profile}` : 'node scripts/auth-standalone.js';
    console.log(`Session exists but has expired. Run: ${rerunCmd}`);
    process.exit(1);
  }
}

async function main() {
  // --status flag: check auth status without logging in
  if (process.argv.includes('--status')) {
    await checkStatus();
    return;
  }

  const email = resolvedEmail;
  const password = resolvedPassword;

  console.log('');
  if (profile) console.log(`Profile:       ${profile}`);
  console.log(`Authenticating as: ${email}`);
  console.log(`Session dir:   ${homeSessionDir}`);
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
      const profileLabel = profile ? ` [${profile}]` : '';
      console.log(`✓ Already logged in as ${email}${profileLabel} — session saved to ${homeSessionDir}`);
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
  const profileLabel = profile ? ` [${profile}]` : '';
  console.log(`✓ Authenticated as ${email}${profileLabel}`);
  console.log(`Session saved to: ${homeSessionDir}`);
}

main().catch((err) => {
  console.error('Authentication failed:', err.message);
  process.exit(1);
});
