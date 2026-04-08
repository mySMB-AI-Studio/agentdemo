import { chromium } from 'playwright';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import readline from 'readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.join(__dirname, '..');

const SESSION_DIR = path.join(ROOT, '.browser-session');
const STATE_FILE = path.join(SESSION_DIR, 'state.json');

function ensureSessionDir() {
  if (!fs.existsSync(SESSION_DIR)) {
    fs.mkdirSync(SESSION_DIR, { recursive: true });
  }
}

/**
 * Returns the state file path for a given profile name.
 * - undefined / null / 'default' → .browser-session/state.json
 * - 'sean' → .browser-session/state-sean.json
 */
function sessionFilePath(profile) {
  if (!profile || profile === 'default') return STATE_FILE;
  return path.join(SESSION_DIR, `state-${profile}.json`);
}

/**
 * Returns an array of { profile, file } for all saved sessions.
 */
export function listSessions() {
  if (!fs.existsSync(SESSION_DIR)) return [];
  const entries = [];
  // Default session
  if (fs.existsSync(STATE_FILE)) {
    entries.push({ profile: 'default', file: STATE_FILE });
  }
  // Named profiles: state-{profile}.json
  for (const f of fs.readdirSync(SESSION_DIR)) {
    const m = f.match(/^state-(.+)\.json$/);
    if (m) {
      entries.push({ profile: m[1], file: path.join(SESSION_DIR, f) });
    }
  }
  return entries;
}

function waitForKeypress(message) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(message, () => {
      rl.close();
      resolve();
    });
  });
}

export async function createBrowserContext(options = {}) {
  ensureSessionDir();
  const isHeadless = options.headless ?? false;
  const stateFile = sessionFilePath(options.profile);
  const launchOpts = {
    headless: isHeadless,
    channel: 'chromium',
    args: ['--start-maximized'],
  };
  // Use viewport: null to inherit the maximized window size.
  // Playwright ignores --start-maximized when viewport is set explicitly.
  // We fall back to 1920x1080 for headless mode where there is no window.
  const contextOpts = {
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    viewport: isHeadless ? { width: 1920, height: 1080 } : null,
  };

  if (fs.existsSync(stateFile)) {
    contextOpts.storageState = stateFile;
  }

  const browser = await chromium.launch(launchOpts);
  const context = await browser.newContext(contextOpts);
  return { browser, context };
}

export async function saveSession(context, profile) {
  ensureSessionDir();
  const stateFile = sessionFilePath(profile);
  await context.storageState({ path: stateFile });
}

export async function isSessionValid(context) {
  const page = await context.newPage();
  try {
    await page.goto('https://www.office.com', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(3000);
    const url = page.url();
    const isValid = !url.includes('login.microsoftonline.com') && !url.includes('login.live.com');
    await page.close();
    return isValid;
  } catch {
    await page.close();
    return false;
  }
}

export async function performLogin(context, profile) {
  const email = process.env.DEMO_EMAIL;
  const password = process.env.DEMO_PASSWORD;

  if (!email || !password) {
    throw new Error('DEMO_EMAIL and DEMO_PASSWORD must be set in .env file');
  }

  // ── Already logged in? ──────────────────────────────────────────────────────
  // Navigate to M365 first. If it loads without redirecting to login, save
  // the current session and return early — no re-login needed.
  const checkPage = await context.newPage();
  try {
    await checkPage.goto('https://m365.cloud.microsoft', {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await checkPage.waitForTimeout(3000);
    const checkUrl = checkPage.url();
    if (!checkUrl.includes('login.microsoftonline.com') && !checkUrl.includes('login.live.com')) {
      // Also warm up Studio so its cookies are included in the saved session
      console.log('  Warming up Copilot Studio session...');
      try {
        await checkPage.goto('https://copilotstudio.microsoft.com', {
          waitUntil: 'domcontentloaded',
          timeout: 60000,
        });
        await checkPage.waitForTimeout(5000);
        const studioUrl = checkPage.url();
        if (!studioUrl.includes('login.microsoftonline.com') && !studioUrl.includes('login.live.com')) {
          console.log('  ✓ Copilot Studio session established.');
        } else {
          console.log('  ⚠ Copilot Studio redirected to login — Studio access may not be available for this account.');
        }
      } catch {
        console.log('  ⚠ Could not reach Copilot Studio during auth (non-fatal).');
      }
      await checkPage.close();
      await saveSession(context, profile);
      console.log(`✓ Already logged in as ${email} — session saved.`);
      return;
    }
  } catch { /* not yet logged in, fall through */ }
  await checkPage.close();

  // ── Fresh login ─────────────────────────────────────────────────────────────
  const page = await context.newPage();
  await page.goto('https://login.microsoftonline.com', { waitUntil: 'domcontentloaded', timeout: 60000 });

  // Wait for page to fully settle before searching for inputs
  await page.waitForLoadState('networkidle', { timeout: 60000 });

  // Enter email — try fallback selectors in order
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
    throw new Error('Could not find email input on login page. Check the login URL or increase timeout.');
  }

  await page.click('input[type="submit"]');

  // Wait for password field or redirect
  await page.waitForTimeout(2000);

  // Enter password
  try {
    await page.waitForSelector('input[type="password"]', { timeout: 60000 });
    await page.fill('input[type="password"]', password);
    await page.click('input[type="submit"]');
  } catch {
    // Password field might not appear if using different auth flow
  }

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
      if (el) {
        mfaDetected = true;
        break;
      }
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
      await page.click('input[type="submit"]'); // Click Yes
      await page.waitForTimeout(2000);
    }
  } catch { /* ignore */ }

  // Verify login succeeded
  await page.waitForTimeout(3000);
  const currentUrl = page.url();
  if (currentUrl.includes('login.microsoftonline.com')) {
    throw new Error('Login did not complete successfully. Check credentials and try again.');
  }

  // Warm up the Copilot Studio domain so its session cookies are captured.
  // Without this, Studio may redirect to login even when M365 session is valid.
  console.log('  Warming up Copilot Studio session...');
  try {
    await page.goto('https://copilotstudio.microsoft.com', {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await page.waitForTimeout(5000);
    const studioUrl = page.url();
    if (!studioUrl.includes('login.microsoftonline.com') && !studioUrl.includes('login.live.com')) {
      console.log('  ✓ Copilot Studio session established.');
    } else {
      console.log('  ⚠ Copilot Studio redirected to login — Studio access may not be available for this account.');
    }
  } catch {
    console.log('  ⚠ Could not reach Copilot Studio during auth (non-fatal).');
  }

  await saveSession(context, profile);
  await page.close();
  console.log(`✓ Session saved. You are logged in as ${email}`);
}

export async function verifyCopilot(context) {
  const page = await context.newPage();
  try {
    await page.goto('https://m365.cloud.microsoft', {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await page.waitForTimeout(5000);

    const url = page.url();
    if (url.includes('login.microsoftonline.com')) {
      await page.close();
      console.log('✗ Session expired — run agentdemo auth first');
      return false;
    }

    // Check if Business Chat / Copilot loads
    const pageContent = await page.content();
    const hasCopilot = pageContent.includes('Copilot') ||
                       pageContent.includes('copilot') ||
                       pageContent.includes('Business Chat') ||
                       pageContent.includes('chat');

    await page.close();

    if (hasCopilot) {
      console.log('✓ M365 Copilot accessible');
      return true;
    } else {
      console.log('✗ M365 Copilot page loaded but Copilot UI not detected');
      console.log('  → Verify the Copilot license is assigned to the demo account');
      console.log('  → Try opening https://m365.cloud.microsoft manually in the browser');
      return false;
    }
  } catch (err) {
    await page.close();
    console.log(`✗ Failed to reach M365 Copilot: ${err.message}`);
    console.log('  → Check network connectivity');
    console.log('  → Verify https://m365.cloud.microsoft is not blocked');
    return false;
  }
}

export async function runAuth(opts) {
  const profile = opts.profile || null;
  const stateFile = sessionFilePath(profile);
  const profileLabel = profile && profile !== 'default' ? `profile "${profile}"` : 'default profile';

  if (opts.status) {
    const sessions = listSessions();
    if (sessions.length === 0) {
      console.log('No sessions found. Run agentdemo auth to log in.');
      return;
    }
    // If a specific profile is requested, just check that one
    const toCheck = profile
      ? sessions.filter(s => s.profile === (profile === 'default' ? 'default' : profile))
      : sessions;
    for (const session of toCheck) {
      if (!fs.existsSync(session.file)) {
        console.log(`  ✗ No session file for profile "${session.profile}"`);
        continue;
      }
      const { browser: cb, context: cc } = await createBrowserContext({ headless: true, profile: session.profile === 'default' ? null : session.profile });
      const valid = await isSessionValid(cc);
      await cb.close();
      if (valid) {
        console.log(`  ✓ Session active for profile "${session.profile}"`);
      } else {
        console.log(`  ✗ Session expired for profile "${session.profile}". Run: agentdemo auth${session.profile !== 'default' ? ` --profile ${session.profile}` : ''}`);
      }
    }
    return;
  }

  // Resolve credentials: for named profiles, check PROFILE_{NAME}_EMAIL / PROFILE_{NAME}_PASSWORD first
  let email = process.env.DEMO_EMAIL;
  let password = process.env.DEMO_PASSWORD;
  if (profile && profile !== 'default') {
    const envPrefix = `PROFILE_${profile.toUpperCase()}`;
    if (process.env[`${envPrefix}_EMAIL`]) email = process.env[`${envPrefix}_EMAIL`];
    if (process.env[`${envPrefix}_PASSWORD`]) password = process.env[`${envPrefix}_PASSWORD`];
  }

  // Check if an existing session is still valid before opening the browser
  if (fs.existsSync(stateFile)) {
    const { browser: checkBrowser, context: checkContext } = await createBrowserContext({ headless: true, profile });
    const valid = await isSessionValid(checkContext);
    await checkBrowser.close();

    if (valid && !opts.verifyCopilot) {
      console.log(`✓ Existing session is valid for ${email} (${profileLabel})`);
      return;
    }

    // Session is expired — wipe just this profile's state file
    console.log(`Session expired for ${profileLabel}. Clearing stale session data before re-logging in...`);
    fs.rmSync(stateFile, { force: true });
  }

  const { browser, context } = await createBrowserContext({ profile });

  // Temporarily override env vars so performLogin picks up the right credentials
  const origEmail = process.env.DEMO_EMAIL;
  const origPassword = process.env.DEMO_PASSWORD;
  if (profile && profile !== 'default') {
    if (email) process.env.DEMO_EMAIL = email;
    if (password) process.env.DEMO_PASSWORD = password;
  }

  // Perform fresh login — pass profile so saveSession writes to the correct file
  await performLogin(context, profile);

  // Restore env vars
  if (profile && profile !== 'default') {
    process.env.DEMO_EMAIL = origEmail;
    process.env.DEMO_PASSWORD = origPassword;
    console.log(`✓ Session saved for ${profileLabel}`);
  }

  if (opts.verifyCopilot) {
    await verifyCopilot(context);
  }

  await browser.close();
}
