import fs from 'fs';
import path from 'path';
import readline from 'readline';
import yaml from 'js-yaml';
import slugify from 'slugify';
import { fileURLToPath } from 'url';
import { parseConfig, getDemoDir, PLATFORM_COLORS } from './config-parser.js';
import { createBrowserContext, isSessionValid } from './auth.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.join(__dirname, '..');
const DEMOS_DIR = path.join(ROOT, 'demos');

const VALID_PLATFORMS = ['m365-copilot', 'sharepoint', 'power-automate', 'teams', 'outlook', 'xero', 'custom'];
const MAX_SLIDES = 50;

function makeSlug(name) {
  return slugify(name, { lower: true, strict: true });
}

// ──────────────────────────────────────────────
// INPUT SANITISATION
// ──────────────────────────────────────────────

/** Strip multi-line pastes: keep only the first line. */
function sanitise(raw) {
  if (!raw) return '';
  return raw.split('\n')[0].split('\r')[0].trim();
}

// ──────────────────────────────────────────────
// READLINE HELPERS — replaces enquirer entirely
// ──────────────────────────────────────────────

let _activeRL = null; // track active instance for cleanup

function createRL() {
  // Close any previous instance that was not cleaned up
  if (_activeRL && !_activeRL.closed) {
    try { _activeRL.close(); } catch { /* ignore */ }
  }
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  _activeRL = rl;
  return rl;
}

function closeRL(rl) {
  if (rl && !rl.closed) {
    try { rl.close(); } catch { /* ignore */ }
  }
  if (_activeRL === rl) _activeRL = null;
}

/**
 * Ask a question. Returns sanitised first-line answer.
 * If `validate` is provided it receives the sanitised value and must return
 * true or an error-message string. The question is re-asked on failure.
 */
function ask(rl, question, defaultVal, validate) {
  const suffix = defaultVal ? ` [${defaultVal}]` : '';
  const doAsk = () => new Promise((resolve) => {
    if (rl.closed) { resolve(defaultVal || ''); return; }
    rl.question(`  ${question}${suffix}: `, (raw) => {
      const answer = sanitise(raw) || defaultVal || '';
      if (validate) {
        const result = validate(answer);
        if (result !== true) {
          console.log(`    ✗ ${result}`);
          resolve(doAsk()); // re-ask
          return;
        }
      }
      resolve(answer);
    });
    // If stdin closes (piped input), fall through with default
    rl.once('close', () => resolve(defaultVal || ''));
  });
  return doAsk();
}

/**
 * Strict y/n confirmation. Only accepts 'y' or 'n' (or empty for default).
 * Any other input re-asks the question.
 */
function askYN(rl, question, defaultYes = false) {
  const hint = defaultYes ? '(Y/n)' : '(y/N)';
  const doAsk = () => new Promise((resolve) => {
    if (rl.closed) { resolve(defaultYes); return; }
    rl.question(`  ${question} ${hint}: `, (raw) => {
      const a = sanitise(raw).toLowerCase();
      if (a === '') { resolve(defaultYes); return; }
      if (a === 'y' || a === 'yes') { resolve(true); return; }
      if (a === 'n' || a === 'no') { resolve(false); return; }
      console.log('    Please enter y or n.');
      resolve(doAsk()); // re-ask
    });
    rl.once('close', () => resolve(defaultYes));
  });
  return doAsk();
}

function askSelect(rl, question, choices) {
  const doAsk = () => new Promise((resolve) => {
    if (rl.closed) { resolve(choices[0]); return; }
    console.log(`\n  ${question}`);
    choices.forEach((c, i) => console.log(`    ${i + 1}. ${c}`));
    rl.question(`  Enter number (1-${choices.length}): `, (raw) => {
      const idx = parseInt(sanitise(raw), 10) - 1;
      if (idx >= 0 && idx < choices.length) {
        resolve(choices[idx]);
      } else {
        console.log(`    Please enter a number between 1 and ${choices.length}.`);
        resolve(doAsk()); // re-ask
      }
    });
    rl.once('close', () => resolve(choices[0]));
  });
  return doAsk();
}

// ──────────────────────────────────────────────
// FIELD VALIDATORS
// ──────────────────────────────────────────────

const V = {
  required:   (max) => (v) => !v ? 'This field is required.' : v.length > max ? `Max ${max} characters.` : true,
  hexColor:   (v) => /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v) ? true : 'Must be a hex colour like #00C9A7.',
  url:        (v) => !v || v.startsWith('https://') ? true : 'Must start with https://',
  urlRequired:(v) => !v ? 'URL is required.' : !v.startsWith('https://') ? 'Must start with https://' : true,
  prompt:     (v) => v.length > 150 ? 'Max 150 characters for a prompt.' : true,
};

// ──────────────────────────────────────────────
// INIT — Guided Interview
// ──────────────────────────────────────────────

export async function runInit(opts) {
  if (opts.discover || opts.fromCopilotStudio) {
    return runAutoDiscoveryInit(opts);
  }

  // Pre-filled values from auto-discovery
  const d = opts._discovered || {};

  const rl = createRL();

  // ── BUG 2 FIX: Ctrl+C handler ──
  const sigintHandler = () => {
    closeRL(rl);
    console.log('\n\n  Setup cancelled. No files were written.');
    process.exit(0);
  };
  process.on('SIGINT', sigintHandler);

  try {
    console.log('\n  AgentDemo — New Demo Setup\n');

    // STEP 1: Agent basics (with validation)
    const title = await ask(rl, 'Agent name', opts.name || d.title || '', V.required(60));
    const description = await ask(rl, 'One-line description', d.description || '', V.required(200));
    let brand_color = await ask(rl, 'Brand color (hex)', '#00C9A7', V.hexColor);
    if (brand_color && !brand_color.startsWith('#')) brand_color = '#' + brand_color;
    const agent_icon = await ask(rl, 'Agent icon path (optional, press Enter to skip)', '');
    const m365_copilot_url = await ask(rl, 'M365 Copilot URL', d.m365_copilot_url || '', V.url);

    const basics = { title, description, brand_color, agent_icon, m365_copilot_url };

    // STEP 2: Platform selection — text-based, not multi-select
    console.log('\n  Platforms involved (besides M365 Copilot):');
    console.log('  Type each platform and press Enter. Press Enter alone when done.');
    console.log('  Options: sharepoint | power-automate | teams | outlook | xero | custom');

    const platforms = ['m365-copilot'];

    // Pre-fill from discovery
    if (d.platforms && d.platforms.length > 1) {
      const extras = d.platforms.filter(p => p !== 'm365-copilot');
      if (extras.length > 0) {
        console.log(`  (Detected: ${extras.join(', ')})`);
      }
      for (const p of extras) {
        if (VALID_PLATFORMS.includes(p) && !platforms.includes(p)) {
          platforms.push(p);
        }
      }
    }

    let addingPlatforms = true;
    while (addingPlatforms) {
      const raw = await ask(rl, '>', '');
      const p = sanitise(raw);
      if (!p) {
        addingPlatforms = false;
      } else if (VALID_PLATFORMS.includes(p) && !platforms.includes(p)) {
        platforms.push(p);
      } else if (p === 'm365-copilot') {
        console.log('    m365-copilot is always included.');
      } else if (!VALID_PLATFORMS.includes(p)) {
        console.log(`    Unknown platform "${p}". Valid: ${VALID_PLATFORMS.filter(x => x !== 'm365-copilot').join(', ')}`);
      } else {
        console.log(`    "${p}" already added.`);
      }
    }

    console.log(`\n  Platforms: ${platforms.join(', ')}`);

    // STEP 3: Slide builder loop
    const slides = [];
    let addMore = true;
    let slideId = 1;

    // Pre-fill prompts from discovery
    const discoveredPrompts = d.prompts ? [...d.prompts] : [];

    while (addMore) {
      // ── BUG 1 FIX: slide limit ──
      if (slideId > MAX_SLIDES) {
        console.log(`\n  ⚠ Reached maximum number of slides (${MAX_SLIDES}).`);
        const action = await ask(rl, "Type 'done' to finish or 'remove' to delete the last slide", 'done');
        if (sanitise(action) === 'remove' && slides.length > 0) {
          const removed = slides.pop();
          slideId--;
          console.log(`  Removed slide ${removed.id}.`);
          continue;
        }
        break;
      }

      console.log(`\n  ── Slide ${slideId} ──`);

      const platform = platforms.length === 1
        ? platforms[0]
        : await askSelect(rl, 'Platform for this slide:', platforms);

      console.log(`  Platform: ${platform}`);

      const urlHints = {
        'power-automate': 'Open flow in Power Automate, copy URL',
        'sharepoint': 'Navigate to the list/page, copy URL',
        'm365-copilot': 'Auto-populated from agent URL, override if needed',
        'teams': 'Right-click channel → Get link to channel',
        'outlook': 'Open the email/folder, copy the URL',
      };
      if (urlHints[platform]) {
        console.log(`  Hint: ${urlHints[platform]}`);
      }

      const story_label = await ask(rl, 'Story label (short headline)', '', V.required(80));
      const narrative = await ask(rl, 'Narrative (1-2 sentences)', '', V.required(300));
      const defaultUrl = platform === 'm365-copilot' ? m365_copilot_url : '';
      const url = await ask(rl, 'URL', defaultUrl, V.url);

      const slide = {
        id: slideId,
        platform,
        story_label,
        narrative,
        url,
        sample_prompts: [],
        record_clip: false,
        annotations: [],
      };

      // M365 Copilot-specific: sample prompts
      if (platform === 'm365-copilot') {
        console.log('\n  Sample prompts (press Enter alone when done):');

        // Show discovered prompts as suggestions
        if (discoveredPrompts.length > 0 && slideId <= 2) {
          console.log('  Discovered starter prompts you can use:');
          discoveredPrompts.forEach((p, i) => console.log(`    ${i + 1}. ${p}`));
          console.log('');
        }

        let promptNum = 1;
        while (promptNum <= 5) {
          const defaultPrompt = discoveredPrompts.length > 0 ? discoveredPrompts.shift() : '';
          const prompt = await ask(rl, `Prompt ${promptNum}`, defaultPrompt, V.prompt);
          if (!prompt) break;
          slide.sample_prompts.push(prompt);
          promptNum++;
        }

        slide.record_clip = await askYN(rl, 'Record video clip?', true);
      }

      // Annotations
      const wantAnnotations = await askYN(rl, 'Add annotations to this slide?', false);
      if (wantAnnotations) {
        let addAnn = true;
        while (addAnn) {
          const type = await askSelect(rl, 'Annotation type:', ['box', 'arrow', 'badge', 'spotlight']);
          const label = await ask(rl, 'Label', '', V.required(80));
          const description = await ask(rl, 'Description (optional)', '');
          const x = await ask(rl, 'Position X (0-100%)', '50', (v) => {
            const n = Number(v);
            return (!isNaN(n) && n >= 0 && n <= 100) ? true : 'Must be a number between 0 and 100.';
          });
          const y = await ask(rl, 'Position Y (0-100%)', '50', (v) => {
            const n = Number(v);
            return (!isNaN(n) && n >= 0 && n <= 100) ? true : 'Must be a number between 0 and 100.';
          });

          slide.annotations.push({
            type,
            label,
            description,
            position: { x: Number(x), y: Number(y) },
          });

          addAnn = await askYN(rl, 'Add another annotation?', false);
        }
      }

      slides.push(slide);
      slideId++;

      // ── BUG 1 FIX: strict y/n for "add another slide" ──
      addMore = await askYN(rl, 'Add another slide?', false);
    }

    // STEP 4: Review & confirm
    const slug = makeSlug(basics.title);
    console.log('');
    console.log('  ┌─────────────────────────────────────────────────────────────┐');
    console.log(`  │  Demo Summary — ${basics.title.substring(0, 42).padEnd(42)} │`);
    console.log('  ├───────┬────────────────┬─────────────────────────────────────┤');
    console.log('  │ Slide │ Platform       │ Story Label                         │');
    console.log('  ├───────┼────────────────┼─────────────────────────────────────┤');
    for (const s of slides) {
      const label = s.story_label.substring(0, 35).padEnd(35);
      console.log(`  │ ${String(s.id).padEnd(5)} │ ${s.platform.padEnd(14)} │ ${label}  │`);
    }
    console.log('  └───────┴────────────────┴─────────────────────────────────────┘');
    console.log('');
    console.log('  Write these files?');
    console.log(`  - demos/${slug}/demo.yaml`);
    console.log(`  - demos/${slug}/PREFLIGHT-GUIDE.md`);
    console.log('');

    const confirmed = await askYN(rl, 'Confirm?', true);

    if (!confirmed) {
      console.log('  Cancelled. No files written.');
      return;
    }

    // Write files
    await writeInitFiles(basics, slides, platforms);

  } finally {
    // ── BUG 2 FIX: always clean up ──
    closeRL(rl);
    process.removeListener('SIGINT', sigintHandler);
  }
}

// ──────────────────────────────────────────────
// AUTO-DISCOVERY INIT (Copilot Studio)
// ──────────────────────────────────────────────

// Connector-name → AgentDemo platform slug mapping
const CONNECTOR_PLATFORM_MAP = {
  'sharepoint':       'sharepoint',
  'power automate':   'power-automate',
  'microsoft teams':  'teams',
  'teams':            'teams',
  'outlook':          'outlook',
  'office 365 outlook': 'outlook',
  'xero':             'xero',
  'dataverse':        'custom',
  'http':             'custom',
  'http webhook':     'custom',
};

// Narrative templates per platform
const NARRATIVE_TEMPLATES = {
  'power-automate': (connector) =>
    `This automated flow${connector ? ` connects ${connector} and` : ''} keeps your data current — no manual work needed.`,
  'sharepoint': () =>
    'All records live here in SharePoint, organized and always ready for the agent to query.',
  'teams': () =>
    'The agent is available directly inside Teams, where your team already works every day.',
  'outlook': () =>
    'Email notifications and actions happen automatically through Outlook.',
  'xero': () =>
    'Financial data from Xero flows directly into the agent for real-time answers.',
  'm365-copilot': () =>
    'Your team asks questions in plain English inside Microsoft 365 — no new tools to learn.',
  'custom': (connector) =>
    `Data from ${connector || 'this platform'} is integrated directly into the agent workflow.`,
};

// Story label templates per platform
const STORY_TEMPLATES = {
  'power-automate': 'Data flows in automatically',
  'sharepoint':     'A single source of truth',
  'teams':          'Works where your team already works',
  'outlook':        'Automated email in your inbox',
  'xero':           'Financial data at your fingertips',
  'm365-copilot':   'Ask the agent anything',
  'custom':         'Connected and integrated',
};

/** Helper: try clicking the first matching visible selector */
async function clickFirst(page, selectors, waitMs = 3000) {
  for (const sel of selectors) {
    try {
      const el = await page.$(sel);
      if (el && await el.isVisible()) {
        await el.click();
        await page.waitForTimeout(waitMs);
        return true;
      }
    } catch { /* next */ }
  }
  return false;
}

/** Helper: dismiss common popups/modals (what's new, updates, etc.)
 *  Tries multiple rounds because some pages show cascading popups.
 *  Optimised for speed — short waits, parallel-safe selectors. */
async function dismissPopups(page, maxAttempts = 3) {
  // Ordered by likelihood on Copilot Studio
  const POPUP_SELECTORS = [
    '[role="dialog"] button:has-text("Skip")',
    '[role="dialog"] button:has-text("Got it")',
    '[role="dialog"] button:has-text("Close")',
    '[role="dialog"] button[aria-label="Close"]',
    'button:has-text("Skip")',
    'button:has-text("Got it")',
    'button:has-text("Dismiss")',
    'button:has-text("Maybe later")',
    'button:has-text("Not now")',
    'button:has-text("No thanks")',
    'button[aria-label="Close"]',
    'button[aria-label="Dismiss"]',
  ];

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    let dismissed = false;
    for (const sel of POPUP_SELECTORS) {
      try {
        const el = await page.$(sel);
        if (el && await el.isVisible()) {
          await el.click();
          console.log(`  Dismissed popup: ${sel.substring(0, 50)}`);
          await page.waitForTimeout(800);
          dismissed = true;
          break;
        }
      } catch { /* next */ }
    }
    if (!dismissed) break;
    await page.waitForTimeout(500);
  }
}

/** Save DEMO_ENVIRONMENT back to .env file */
function saveEnvironmentToEnv(envId) {
  const envPath = path.join(ROOT, '.env');
  if (!fs.existsSync(envPath)) return;
  let content = fs.readFileSync(envPath, 'utf8');
  if (content.includes('DEMO_ENVIRONMENT=')) {
    content = content.replace(/DEMO_ENVIRONMENT=.*/, `DEMO_ENVIRONMENT=${envId}`);
  } else {
    content += `\nDEMO_ENVIRONMENT=${envId}\n`;
  }
  fs.writeFileSync(envPath, content);
}

async function runAutoDiscoveryInit(opts) {
  console.log('\n  AgentDemo — Auto-Discovery\n');

  // ═══════════════════════════════════════
  // STEP 0: Get the two URLs (from flags or ask user)
  // ═══════════════════════════════════════
  let studioUrl = opts.studioUrl || '';
  let m365Url = opts.m365Url || '';

  if (!studioUrl || !m365Url) {
    const urlRL = createRL();
    const urlSigint = () => { closeRL(urlRL); console.log('\n\n  Cancelled.'); process.exit(0); };
    process.on('SIGINT', urlSigint);
    try {
      if (!studioUrl) {
        console.log('  Paste the direct link to your agent in Copilot Studio.');
        console.log('  (Open the agent in Copilot Studio → copy URL from browser address bar)\n');
        studioUrl = await ask(urlRL, 'Copilot Studio URL', '', (v) => {
          if (!v) return 'URL is required';
          if (!v.startsWith('https://')) return 'Must start with https://';
          if (!v.includes('copilotstudio')) return 'Must be a Copilot Studio URL';
          return true;
        });
      }
      if (!m365Url) {
        console.log('\n  Paste the direct link to your agent in M365 Copilot.');
        console.log('  (Open M365 Copilot → right-click the agent → Copy link)\n');
        m365Url = await ask(urlRL, 'M365 Copilot URL', '', (v) => {
          if (!v) return 'URL is required';
          if (!v.startsWith('https://')) return 'Must start with https://';
          return true;
        });
      }
    } finally {
      closeRL(urlRL);
      process.removeListener('SIGINT', urlSigint);
    }
  }

  console.log(`\n  Studio URL:  ${studioUrl}`);
  console.log(`  M365 URL:    ${m365Url}\n`);

  const { browser, context } = await createBrowserContext();
  const valid = await isSessionValid(context);
  if (!valid) {
    console.log('  Session expired — run agentdemo auth first.');
    await browser.close();
    return;
  }

  const page = await context.newPage();
  const debugDir = path.join(ROOT, '.browser-session');
  if (!fs.existsSync(debugDir)) fs.mkdirSync(debugDir, { recursive: true });

  // All data collected from the browser
  let discoveredData = null;
  const connectionWarnings = []; // connectors with non-Connected status

  try {
    // ═══════════════════════════════════════
    // PHASE 1: Navigate directly to agent in Copilot Studio
    // ═══════════════════════════════════════
    console.log('  Opening agent in Copilot Studio...');
    await page.goto(studioUrl, {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(5000);

    // Dismiss any popups (what's new, feature tours, etc.)
    await dismissPopups(page);
    console.log(`  Agent page URL: ${page.url()}`);

    // ═══════════════════════════════════════
    // PHASE 2: Read agent Overview
    // ═══════════════════════════════════════
    // The Studio URL takes us directly to the agent detail page.
    // No environment picking or agent list needed.
    console.log('  Reading Overview...');

    // Extract agent name from the page
    const agentName = await page.evaluate(() => {
      // Try page title first
      const title = document.title?.replace(/\s*[-–|].*$/, '').trim();
      if (title && title.length > 1 && title.length < 100) return title;
      // Try prominent heading
      for (const sel of ['h1', 'h2', '[class*="agentName" i]', '[class*="botName" i]']) {
        const el = document.querySelector(sel);
        if (el) {
          const t = el.textContent?.trim();
          if (t && t.length > 1 && t.length < 100) return t;
        }
      }
      return '';
    }) || 'My Agent';

    console.log(`  Agent name: ${agentName}`);

    discoveredData = {
      title: agentName,
      description: '',
      m365_copilot_url: m365Url,
      brand_color: '#00C9A7',
      platforms: ['m365-copilot'],
      topics: [],       // { name, triggerPhrases[], connector? }
      connections: [],   // { name, status, platform }
      flowUrls: [],      // direct URLs to Power Automate flows
      sharepointUrls: [],// SharePoint site URLs
      slides: [],        // pre-built slides
    };

    // Read description from Overview
    const overview = await page.evaluate(() => {
      const desc = document.querySelector(
        '[class*="description" i], [data-testid*="description"], [aria-label*="description" i], ' +
        '[class*="subtitle" i]'
      );
      return {
        description: desc?.textContent?.trim()?.split('\n')[0]?.trim() || '',
      };
    });
    if (overview.description) {
      discoveredData.description = overview.description;
      console.log(`  Description: ${overview.description.substring(0, 60)}...`);
    }

    // Try to save agent icon
    try {
      const iconEl = await page.$('[class*="avatar" i] img, [class*="icon" i] img, [class*="agent" i] img');
      if (iconEl) {
        const iconSrc = await iconEl.getAttribute('src');
        if (iconSrc && iconSrc.startsWith('http')) {
          console.log('  Agent icon found');
          // Will be saved later after demo folder is created
        }
      }
    } catch { /* no icon */ }

    // ═══════════════════════════════════════
    // PHASE 6: Read Topics & Trigger Phrases
    // ═══════════════════════════════════════
    console.log('  Reading Topics...');
    const topicsClicked = await clickFirst(page, [
      'button:has-text("Topics")', 'a:has-text("Topics")',
      '[role="tab"]:has-text("Topics")', '[aria-label*="Topics"]',
      'nav >> text=/topics/i',
    ], 4000);

    if (topicsClicked) {
      await page.waitForLoadState('networkidle').catch(() => {});
      await page.waitForTimeout(4000);

      const topics = await page.evaluate(() => {
        const results = [];
        const rows = document.querySelectorAll(
          '[role="row"], [role="listitem"], tr, [class*="topic" i], [class*="Topic"]'
        );
        for (const row of rows) {
          // Name
          const nameEl = row.querySelector(
            'a, [class*="name" i], [class*="Name"], [class*="title" i], td:first-child, ' +
            '[role="gridcell"]:first-child, h3, h4'
          );
          if (!nameEl) continue;
          const name = nameEl.textContent?.trim()?.split('\n')[0]?.trim();
          if (!name || name.length < 2 || name.length > 150) continue;

          // Status
          const statusEl = row.querySelector('[class*="status" i], [class*="badge" i]');
          const status = statusEl?.textContent?.trim()?.toLowerCase() || '';
          // Skip system/disabled topics
          if (status.includes('off') || status.includes('disabled')) continue;
          const skipNames = new Set([
            'name', 'status', 'trigger', 'modified', 'system', 'fallback',
            'greeting', 'on error', 'conversational boosting', 'topics',
          ]);
          if (skipNames.has(name.toLowerCase())) continue;

          // Trigger phrases (may be in sub-elements or a separate column)
          const triggerEl = row.querySelector('[class*="trigger" i], [class*="phrase" i], td:nth-child(2)');
          const triggerText = triggerEl?.textContent?.trim() || '';
          const phrases = triggerText
            .split(/[;,\n]/)
            .map(p => p.trim())
            .filter(p => p.length >= 5 && p.length <= 150)
            .slice(0, 3);

          results.push({ name, phrases });
        }
        return results.slice(0, 20);
      });

      if (topics.length > 0) {
        discoveredData.topics = topics;
        const totalPhrases = topics.reduce((n, t) => n + t.phrases.length, 0);
        console.log(`  Found ${topics.length} topic(s) with ${totalPhrases} trigger phrase(s)`);
      } else {
        console.log('  No topics found (or page layout not recognised)');
      }
    }

    // ═══════════════════════════════════════
    // PHASE 7: Read Connections / Actions
    // ═══════════════════════════════════════
    console.log('  Reading Connections...');

    // Try Actions tab first, then Settings > Connections
    let connFound = await clickFirst(page, [
      'button:has-text("Actions")', 'a:has-text("Actions")',
      '[role="tab"]:has-text("Actions")', '[aria-label*="Actions"]',
    ], 4000);
    if (!connFound) {
      connFound = await clickFirst(page, [
        'button:has-text("Settings")', 'a:has-text("Settings")',
        '[role="tab"]:has-text("Settings")',
      ], 4000);
      if (connFound) {
        await clickFirst(page, [
          'button:has-text("Connections")', 'a:has-text("Connections")',
          '[role="tab"]:has-text("Connections")', 'text=/connections/i',
        ], 3000);
      }
    }

    if (connFound) {
      await page.waitForLoadState('networkidle').catch(() => {});
      await page.waitForTimeout(4000);

      const connections = await page.evaluate(() => {
        const found = [];
        const seen = new Set();
        const rows = document.querySelectorAll(
          '[role="row"], [role="listitem"], tr, [class*="connection" i], [class*="connector" i], ' +
          '[class*="action" i][class*="item" i], [class*="Action"][class*="Item"]'
        );
        for (const row of rows) {
          const nameEl = row.querySelector(
            'a, [class*="name" i], [class*="Name"], [class*="title" i], td:first-child, ' +
            '[role="gridcell"]:first-child, h3, h4, span'
          );
          if (!nameEl) continue;
          const name = nameEl.textContent?.trim()?.split('\n')[0]?.trim();
          if (!name || name.length < 2 || name.length > 100) continue;

          const statusEl = row.querySelector(
            '[class*="status" i], [class*="badge" i], [class*="state" i]'
          );
          const status = statusEl?.textContent?.trim() || 'Unknown';

          const skipNames = new Set([
            'name', 'status', 'type', 'connections', 'actions', 'connector',
            'modified', 'add', 'create', 'new',
          ]);
          if (skipNames.has(name.toLowerCase())) continue;

          if (!seen.has(name.toLowerCase())) {
            seen.add(name.toLowerCase());
            found.push({ name, status });
          }
        }
        return found;
      });

      if (connections.length > 0) {
        console.log(`  Found ${connections.length} connection(s):`);
        for (const conn of connections) {
          // Map to platform slug
          const slug = Object.entries(CONNECTOR_PLATFORM_MAP).find(
            ([key]) => conn.name.toLowerCase().includes(key)
          )?.[1] || 'custom';

          const connEntry = { name: conn.name, status: conn.status, platform: slug };
          discoveredData.connections.push(connEntry);

          if (!discoveredData.platforms.includes(slug)) {
            discoveredData.platforms.push(slug);
          }

          const statusIcon = conn.status.toLowerCase().includes('connect') ? '✓' : '⚠';
          console.log(`    ${statusIcon} ${conn.name} → ${slug} (${conn.status})`);

          // Track warnings for non-connected connectors
          if (!conn.status.toLowerCase().includes('connect')) {
            connectionWarnings.push(connEntry);
          }
        }
      } else {
        console.log('  No connections found (or page layout not recognised)');
      }
    }

    // ═══════════════════════════════════════
    // PHASE 8: Read flow URLs and SharePoint URLs from page text
    // ═══════════════════════════════════════
    try {
      const pageUrls = await page.evaluate(() => {
        const text = document.body.innerText || '';
        const urls = text.match(/https?:\/\/[^\s"'<>]+/g) || [];
        return urls;
      });
      for (const u of pageUrls) {
        if (u.includes('make.powerautomate.com') || u.includes('flow.microsoft.com')) {
          if (!discoveredData.flowUrls.includes(u)) discoveredData.flowUrls.push(u);
        }
        if (u.includes('.sharepoint.com')) {
          if (!discoveredData.sharepointUrls.includes(u)) discoveredData.sharepointUrls.push(u);
        }
      }
      if (discoveredData.flowUrls.length > 0) {
        console.log(`  Found ${discoveredData.flowUrls.length} Power Automate flow URL(s)`);
      }
      if (discoveredData.sharepointUrls.length > 0) {
        console.log(`  Found ${discoveredData.sharepointUrls.length} SharePoint URL(s)`);
      }
    } catch { /* ignore URL extraction errors */ }

    // M365 Copilot URL already provided by user — no need to discover it
    console.log(`  M365 Copilot URL: ${m365Url}`);

  } catch (err) {
    try {
      const debugScreenshot = path.join(debugDir, 'debug-screenshot.png');
      await page.screenshot({ path: debugScreenshot, fullPage: true });
      console.log(`  Debug screenshot saved: ${debugScreenshot}`);
    } catch { /* ignore */ }
    console.log(`\n  Discovery error: ${err.message}`);
    console.log(`  Page URL at failure: ${page.url()}`);
  }

  // ═══════════════════════════════════════════════
  // CRITICAL: Close browser BEFORE any readline prompts
  // ═══════════════════════════════════════════════
  console.log('  Closing browser...');
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  console.log('  Browser closed.\n');

  if (!discoveredData) {
    console.log('  Discovery failed. Falling back to manual interview.\n');
    opts.discover = false;
    opts.fromCopilotStudio = false;
    opts._discovered = {};
    await runInit(opts);
    return;
  }

  // ═══════════════════════════════════════════════
  // PHASE 10: Auto-build slide order from discovery
  // ═══════════════════════════════════════════════
  const slides = [];
  let slideId = 1;

  // A) Supporting platform slides first (Power Automate → SharePoint → Xero → custom)
  const platformOrder = ['power-automate', 'sharepoint', 'xero', 'custom'];
  for (const plat of platformOrder) {
    if (!discoveredData.platforms.includes(plat)) continue;
    const conn = discoveredData.connections.find(c => c.platform === plat);
    const narrativeFn = NARRATIVE_TEMPLATES[plat] || NARRATIVE_TEMPLATES['custom'];
    let url = '';
    if (plat === 'power-automate' && discoveredData.flowUrls.length > 0) {
      url = discoveredData.flowUrls.shift();
    } else if (plat === 'sharepoint' && discoveredData.sharepointUrls.length > 0) {
      url = discoveredData.sharepointUrls.shift();
    }
    slides.push({
      id: slideId++,
      platform: plat,
      story_label: STORY_TEMPLATES[plat] || STORY_TEMPLATES['custom'],
      narrative: narrativeFn(conn?.name || ''),
      url,
      sample_prompts: [],
      record_clip: false,
      annotations: [],
    });
  }

  // B) M365 Copilot slides — one per topic group (max 3 prompts each)
  const topicGroups = [];
  if (discoveredData.topics.length > 0) {
    for (const topic of discoveredData.topics) {
      const prompts = topic.phrases.length > 0 ? topic.phrases : [topic.name];
      topicGroups.push({ name: topic.name, prompts: prompts.slice(0, 3) });
    }
  }

  if (topicGroups.length === 0) {
    // Fallback: single M365 Copilot slide with generic prompt
    topicGroups.push({ name: 'Ask the agent anything', prompts: [] });
  }

  for (const group of topicGroups) {
    slides.push({
      id: slideId++,
      platform: 'm365-copilot',
      story_label: group.name.length <= 80 ? group.name : STORY_TEMPLATES['m365-copilot'],
      narrative: NARRATIVE_TEMPLATES['m365-copilot'](),
      url: discoveredData.m365_copilot_url || '',
      sample_prompts: group.prompts,
      record_clip: true,
      annotations: [],
    });
  }

  // C) Teams / Outlook at the end
  for (const plat of ['teams', 'outlook']) {
    if (!discoveredData.platforms.includes(plat)) continue;
    const narrativeFn = NARRATIVE_TEMPLATES[plat] || NARRATIVE_TEMPLATES['custom'];
    slides.push({
      id: slideId++,
      platform: plat,
      story_label: STORY_TEMPLATES[plat],
      narrative: narrativeFn(),
      url: '',
      sample_prompts: [],
      record_clip: false,
      annotations: [],
    });
  }

  discoveredData.slides = slides;

  // ═══════════════════════════════════════════════
  // PHASE 11: Print summary + minimal confirmation
  // ═══════════════════════════════════════════════
  const topicCount = discoveredData.topics.length;
  const phraseCount = discoveredData.topics.reduce((n, t) => n + t.phrases.length, 0);
  const flowCount = discoveredData.flowUrls.length;
  const spCount = discoveredData.sharepointUrls.length;

  console.log(`  AgentDemo discovered the following for: ${discoveredData.title}\n`);
  console.log(`   Platforms found:     ${discoveredData.platforms.join(', ')}`);
  console.log(`   Topics found:        ${topicCount} (${phraseCount} trigger phrases)`);
  if (flowCount > 0) console.log(`   Flows found:         ${flowCount} (URLs pre-filled)`);
  if (spCount > 0)   console.log(`   SharePoint sites:    ${spCount} (URLs pre-filled)`);
  console.log('');

  // Connection warnings
  for (const warn of connectionWarnings) {
    console.log(`   ⚠ ${warn.name} connector status: ${warn.status}`);
    console.log(`     This connection needs attention before capture.`);
    console.log(`     See PREFLIGHT-GUIDE.md for fix steps.\n`);
  }

  console.log('   Suggested slides:');
  for (const s of slides) {
    const promptInfo = s.sample_prompts.length > 0 ? ` (${s.sample_prompts.length} prompts)` : '';
    console.log(`   ${s.id}. ${s.platform.padEnd(16)} — ${s.story_label}${promptInfo}`);
  }
  console.log('');
  console.log('   Items needing your input:');
  console.log('   ─────────────────────────');
  console.log('   - Narrative text for each slide (pre-filled, press Enter to accept)');
  console.log('   - URLs for slides without pre-filled values');
  console.log('   - Any custom annotations (optional, can add later in demo.yaml)');
  console.log('');

  const rl = createRL();
  const sigintHandler = () => {
    closeRL(rl);
    console.log('\n\n  Setup cancelled. No files were written.');
    process.exit(0);
  };
  process.on('SIGINT', sigintHandler);

  try {
    const skipReview = await ask(rl, "Press Enter to review slides, or type 'skip' to accept all defaults", '');

    if (sanitise(skipReview).toLowerCase() === 'skip') {
      console.log('\n  Skipping review — writing files with all defaults...');
    } else {
      // Minimal review: show each slide, only ask for narrative + URL if missing
      for (const slide of slides) {
        console.log(`\n  ── Slide ${slide.id} — ${slide.platform} ──`);
        console.log(`  Story label: ${slide.story_label}`);

        const newNarrative = await ask(rl, 'Narrative (press Enter to accept)', slide.narrative, V.required(300));
        slide.narrative = newNarrative;

        if (!slide.url) {
          const urlHints = {
            'power-automate': 'Paste Power Automate flow URL',
            'sharepoint': 'Paste SharePoint list/page URL',
            'teams': 'Paste Teams channel URL',
            'outlook': 'Paste Outlook email/folder URL',
          };
          const hint = urlHints[slide.platform] || 'Paste the URL';
          console.log(`  Hint: ${hint}`);
          slide.url = await ask(rl, 'URL', slide.url || discoveredData.m365_copilot_url || '', V.url);
        } else {
          console.log(`  URL: ${slide.url}`);
        }
      }
    }

    // Final confirmation
    const slug = makeSlug(discoveredData.title);
    console.log('');
    console.log('  ┌─────────────────────────────────────────────────────────────┐');
    console.log(`  │  Demo Summary — ${discoveredData.title.substring(0, 42).padEnd(42)} │`);
    console.log('  ├───────┬────────────────┬─────────────────────────────────────┤');
    console.log('  │ Slide │ Platform       │ Story Label                         │');
    console.log('  ├───────┼────────────────┼─────────────────────────────────────┤');
    for (const s of slides) {
      const label = s.story_label.substring(0, 35).padEnd(35);
      console.log(`  │ ${String(s.id).padEnd(5)} │ ${s.platform.padEnd(14)} │ ${label}  │`);
    }
    console.log('  └───────┴────────────────┴─────────────────────────────────────┘');
    console.log('');
    console.log('  Write these files?');
    console.log(`  - demos/${slug}/demo.yaml`);
    console.log(`  - demos/${slug}/PREFLIGHT-GUIDE.md`);
    console.log('');

    const confirmed = await askYN(rl, 'Confirm?', true);
    if (!confirmed) {
      console.log('  Cancelled. No files written.');
      return;
    }

    // Build basics object for writeInitFiles
    const basics = {
      title: discoveredData.title,
      description: discoveredData.description || `Interactive demo for ${discoveredData.title}`,
      brand_color: discoveredData.brand_color || '#00C9A7',
      agent_icon: '',
      m365_copilot_url: discoveredData.m365_copilot_url || '',
    };
    const platforms = [...new Set(discoveredData.platforms)];

    await writeInitFiles(basics, slides, platforms, connectionWarnings);

  } finally {
    closeRL(rl);
    process.removeListener('SIGINT', sigintHandler);
  }
}

// ──────────────────────────────────────────────
// FILE OUTPUT
// ──────────────────────────────────────────────

async function writeInitFiles(basics, slides, platforms, connectionWarnings = []) {
  const slug = makeSlug(basics.title);
  const demoDir = path.join(DEMOS_DIR, slug);

  // Create directories
  for (const sub of ['screenshots', 'clips', 'output']) {
    fs.mkdirSync(path.join(demoDir, sub), { recursive: true });
  }

  // 1. Write demo.yaml with detailed comments
  const yamlLines = [
    '# ──────────────────────────────────────────────',
    `# AgentDemo configuration for: ${basics.title}`,
    `# Generated: ${new Date().toISOString()}`,
    '#',
    '# Edit this file to customise your demo slides.',
    '# Then run: agentdemo run --config <path-to-this-file>',
    '# ──────────────────────────────────────────────',
    '',
    'demo:',
    '  # Agent display name shown in the demo header.',
    `  title: "${basics.title}"`,
    '',
    '  # One-line agent summary shown below the title.',
    `  description: "${basics.description}"`,
    '',
    '  # Primary hex colour for the demo UI chrome.',
    `  brand_color: "${basics.brand_color}"`,
    '',
  ];

  if (basics.agent_icon) {
    yamlLines.push('  # Path to a logo/icon image for the agent.');
    yamlLines.push(`  agent_icon: "${basics.agent_icon}"`);
    yamlLines.push('');
  }

  yamlLines.push('  # Deep link to the agent inside M365 Copilot.');
  yamlLines.push('  # Get this by opening the agent in M365 Copilot and copying the URL.');
  yamlLines.push(`  m365_copilot_url: "${basics.m365_copilot_url}"`);
  yamlLines.push('');
  yamlLines.push('  slides:');

  for (const slide of slides) {
    yamlLines.push(`    - id: ${slide.id}`);
    yamlLines.push('      # Platform for this slide.');
    yamlLines.push('      # Options: m365-copilot | sharepoint | power-automate |');
    yamlLines.push('      #          teams | outlook | xero | custom');
    yamlLines.push(`      platform: ${slide.platform}`);
    yamlLines.push('');
    yamlLines.push('      # Large headline shown above the screenshot in the demo.');
    yamlLines.push('      # Keep short and business-focused, not technical.');
    yamlLines.push(`      story_label: "${slide.story_label}"`);
    yamlLines.push('');
    yamlLines.push('      # 1-2 sentences for a business audience.');
    yamlLines.push('      # Avoid technical jargon — explain what the user sees.');
    yamlLines.push(`      narrative: "${slide.narrative}"`);
    yamlLines.push('');
    yamlLines.push('      # Direct URL Playwright will navigate to for this slide.');
    yamlLines.push(`      url: "${slide.url}"`);

    if (slide.sample_prompts.length > 0) {
      yamlLines.push('');
      yamlLines.push('      # Prompts to type into the agent (m365-copilot slides only).');
      yamlLines.push('      # Each prompt is typed with realistic speed and the response is recorded.');
      yamlLines.push('      sample_prompts:');
      for (const p of slide.sample_prompts) {
        yamlLines.push(`        - "${p}"`);
      }
    }

    if (slide.record_clip) {
      yamlLines.push('');
      yamlLines.push('      # Set to true to record a video clip of the agent responding.');
      yamlLines.push('      record_clip: true');
    }

    if (slide.annotations.length > 0) {
      yamlLines.push('');
      yamlLines.push('      # Visual annotations overlaid on the screenshot.');
      yamlLines.push('      # Types: box | arrow | badge | spotlight');
      yamlLines.push('      annotations:');
      for (const ann of slide.annotations) {
        yamlLines.push(`        - type: ${ann.type}`);
        yamlLines.push(`          label: "${ann.label}"`);
        if (ann.description) {
          yamlLines.push(`          description: "${ann.description}"`);
        }
        yamlLines.push('          position:');
        yamlLines.push(`            x: ${ann.position.x}`);
        yamlLines.push(`            y: ${ann.position.y}`);
      }
    }
    yamlLines.push('');
  }

  fs.writeFileSync(path.join(demoDir, 'demo.yaml'), yamlLines.join('\n'));

  // 2. Write PREFLIGHT-GUIDE.md
  generatePreflightGuide(basics, slides, platforms, demoDir, slug, connectionWarnings);

  // 3. Write .session-meta.json
  const meta = {
    agent_name: basics.title,
    created_at: new Date().toISOString(),
    last_captured: null,
    last_generated: null,
    slide_count: slides.length,
    platforms: [...new Set(platforms)],
  };
  fs.writeFileSync(path.join(demoDir, '.session-meta.json'), JSON.stringify(meta, null, 2));

  // Print completion
  console.log(`\n  ✓ Demo scaffolded: demos/${slug}/`);
  console.log(`  ✓ PREFLIGHT-GUIDE.md generated`);
  console.log('');
  console.log('  Your next steps:');
  console.log('  ──────────────────────────────────────────────');
  console.log(`  1. Review your demo plan:`);
  console.log(`     cat demos/${slug}/demo.yaml`);
  console.log('');
  console.log(`  2. Read the pre-flight guide:`);
  console.log(`     cat demos/${slug}/PREFLIGHT-GUIDE.md`);
  console.log('');
  console.log(`  3. Run pre-flight check:`);
  console.log(`     agentdemo check --config demos/${slug}/demo.yaml`);
  console.log('');
  console.log(`  4. When ready, start capture:`);
  console.log(`     agentdemo run --config demos/${slug}/demo.yaml`);
  console.log('  ──────────────────────────────────────────────');
}

// ──────────────────────────────────────────────
// PREFLIGHT GUIDE GENERATOR
// ──────────────────────────────────────────────

function generatePreflightGuide(basics, slides, platforms, demoDir, slug, connectionWarnings = []) {
  const usedPlatforms = new Set(slides.map(s => s.platform));
  const lines = [];

  lines.push(`# Pre-Flight Guide — ${basics.title}`);
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');

  // Connection warnings section
  if (connectionWarnings.length > 0) {
    lines.push('## ⚠ Connection Warnings');
    lines.push('');
    lines.push('The following connectors need attention before capture:');
    lines.push('');
    for (const warn of connectionWarnings) {
      lines.push(`- **${warn.name}** — Status: ${warn.status}`);
      lines.push(`  Fix: Open Copilot Studio → Settings → Connections → Re-authenticate ${warn.name}`);
    }
    lines.push('');
  }

  // Section A: One-Time Setup
  lines.push('## Section A — One-Time Demo Account Setup');
  lines.push('');

  lines.push('### A1 — Create the Demo Account');
  lines.push('- [ ] Create a dedicated M365 account: demo@yourtenant.onmicrosoft.com');
  lines.push('      Regular licensed user, not personal or admin account');
  lines.push('- [ ] Assign licenses:');
  lines.push('      - Microsoft 365');
  lines.push('      - Microsoft Copilot');
  if (usedPlatforms.has('power-automate')) {
    lines.push('      - Power Automate');
  }
  lines.push('- [ ] Set non-expiring password:');
  lines.push('      Entra ID → Users → {account} → Properties →');
  lines.push('      Password policies → Disable password expiration');
  lines.push('- [ ] If tenant enforces MFA: set up Microsoft Authenticator app');
  lines.push('');

  lines.push('### A2 — Configure the .env File');
  lines.push('- [ ] Copy .env.example to .env in agentdemo root');
  lines.push('- [ ] Fill in DEMO_EMAIL, DEMO_PASSWORD, DEMO_TENANT');
  lines.push('- [ ] Confirm .env is in .gitignore before proceeding');
  lines.push('');

  lines.push('### A3 — First Login & Session Initialization');
  lines.push('- [ ] Run: `agentdemo auth`');
  lines.push('- [ ] If MFA appears: approve in Authenticator, press Enter in terminal');
  lines.push('- [ ] Confirm: "✓ Session saved. You are logged in as demo@..."');
  lines.push('');

  lines.push('### A4 — Verify M365 Copilot Access');
  lines.push('- [ ] Run: `agentdemo auth --verify-copilot`');
  lines.push('- [ ] Confirm: "✓ M365 Copilot accessible"');
  lines.push('');

  lines.push('### A5 — Verify ffmpeg');
  lines.push('- [ ] Run: `agentdemo check --ffmpeg-only`');
  lines.push('- [ ] Confirm: "✓ ffmpeg available at {path}"');
  lines.push('');

  // Section B: Per-Demo Setup
  lines.push('## Section B — Per-Demo Setup');
  lines.push('');

  lines.push('### B1 — Publish the Agent');
  lines.push('- [ ] Open Copilot Studio: https://copilotstudio.microsoft.com');
  lines.push('- [ ] Confirm agent has no errors in topics or actions');
  lines.push('- [ ] Confirm agent responds to at least one prompt in Test panel');
  lines.push('- [ ] Click Publish → wait for "Your agent is published."');
  lines.push('');

  lines.push('### B2 — Add M365 Copilot Channel');
  lines.push('- [ ] Copilot Studio → Channels → Microsoft 365 Copilot → Add channel');
  lines.push('- [ ] Wait for status: "Connected" (may take 2–5 minutes)');
  lines.push('- [ ] Open https://m365.cloud.microsoft as demo account');
  lines.push('- [ ] Confirm agent appears in the agent list');
  lines.push('- [ ] Copy the deep link and paste into demo.yaml as m365_copilot_url');
  lines.push('');

  if (usedPlatforms.has('sharepoint')) {
    lines.push('### B3 — Approve SharePoint Connection');
    lines.push('- [ ] In M365 Copilot, open the agent and run a prompt that triggers SharePoint lookup');
    lines.push('- [ ] If connection dialog appears: click Connect, sign in as demo account');
    lines.push('- [ ] Confirm agent returns data (not a connection error)');
    lines.push('- [ ] Confirm demo account has Read access to the SharePoint site');
    lines.push('');
  }

  if (usedPlatforms.has('power-automate')) {
    lines.push('### B4 — Verify Power Automate Flow Access');
    lines.push('- [ ] Open https://make.powerautomate.com as demo account');
    lines.push('- [ ] Confirm the flow is visible');
    lines.push('- [ ] Confirm at least one "Succeeded" entry in Run History');
    lines.push('- [ ] Copy direct URL to flow detail page → paste into demo.yaml slide url');
    lines.push('');
  }

  if (usedPlatforms.has('xero')) {
    lines.push('### B5 — Approve Xero Connection');
    lines.push('- [ ] Copilot Studio → Settings → Connections → find Xero connector');
    lines.push('- [ ] Confirm status: "Connected"');
    lines.push('- [ ] In M365 Copilot, run a prompt triggering Xero data');
    lines.push('- [ ] Confirm agent returns Xero data without errors');
    lines.push('');
  }

  if (usedPlatforms.has('teams')) {
    lines.push('### B6 — Verify Teams Access');
    lines.push('- [ ] Open https://teams.microsoft.com as demo account');
    lines.push('- [ ] Confirm demo account is a member of the relevant channel');
    lines.push('- [ ] Copy channel URL and paste into demo.yaml slide url');
    lines.push('');
  }

  if (usedPlatforms.has('outlook')) {
    lines.push('### B7 — Verify Outlook Access');
    lines.push('- [ ] Open https://outlook.office.com as demo account');
    lines.push('- [ ] Confirm the relevant email or folder is visible');
    lines.push('- [ ] Copy URL from browser address bar → paste into demo.yaml slide url');
    lines.push('');
  }

  if (usedPlatforms.has('custom')) {
    lines.push('### B8 — Custom Platform Setup');
    lines.push('- [ ] Confirm demo account has login access to the custom platform');
    lines.push('- [ ] Navigate to exact URL that will be captured — confirm it loads');
    lines.push('- [ ] If platform requires separate login: pre-login manually before capture');
    lines.push('');
  }

  // Per-slide checklist
  lines.push('---');
  lines.push('');
  lines.push('## Per-Slide Checklist');
  lines.push('');
  for (const slide of slides) {
    lines.push(`### Slide ${slide.id} — ${slide.platform}`);
    lines.push(`- **Platform to open:** ${slide.platform}`);
    lines.push(`- **URL to navigate to:** ${slide.url || '(uses M365 Copilot URL)'}`);
    lines.push(`- **Screenshot will be saved as:** screenshots/${slide.id}-${slide.platform}-final.png`);
    if (slide.platform === 'm365-copilot' && slide.sample_prompts.length > 0) {
      lines.push(`- **Prompts to run:** ${slide.sample_prompts.join('; ')}`);
    }
    lines.push('');
  }

  fs.writeFileSync(path.join(demoDir, 'PREFLIGHT-GUIDE.md'), lines.join('\n'));
}

// ──────────────────────────────────────────────
// GUIDE COMMAND
// ──────────────────────────────────────────────

export async function runGuide(opts) {
  const config = parseConfig(opts.config);
  const demoDir = getDemoDir(opts.config);
  const slug = path.basename(demoDir);
  const platforms = [...new Set(config.slides.map(s => s.platform))];

  generatePreflightGuide(
    { title: config.title, description: config.description },
    config.slides,
    platforms,
    demoDir,
    slug,
  );

  console.log(`✓ PREFLIGHT-GUIDE.md generated at ${path.join(demoDir, 'PREFLIGHT-GUIDE.md')}`);
}

// ──────────────────────────────────────────────
// LIST COMMAND
// ──────────────────────────────────────────────

export async function runList() {
  if (!fs.existsSync(DEMOS_DIR)) {
    console.log('No demos found. Run agentdemo init to create one.');
    return;
  }

  const entries = fs.readdirSync(DEMOS_DIR, { withFileTypes: true });
  const demos = entries.filter(e => e.isDirectory());

  if (demos.length === 0) {
    console.log('No demos found. Run agentdemo init to create one.');
    return;
  }

  console.log('\n  AgentDemo — All Demos\n');
  console.log('  Name                   Slides  Status         Last Captured');
  console.log('  ─────────────────────  ──────  ─────────────  ─────────────');

  for (const demo of demos) {
    const demoPath = path.join(DEMOS_DIR, demo.name);
    const metaPath = path.join(demoPath, '.session-meta.json');
    const sessionPath = path.join(demoPath, '.session.json');

    let name = demo.name;
    let slideCount = '?';
    let status = 'not started';
    let lastCaptured = '—';

    if (fs.existsSync(metaPath)) {
      const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
      name = meta.agent_name || demo.name;
      slideCount = String(meta.slide_count || '?');
      lastCaptured = meta.last_captured
        ? new Date(meta.last_captured).toLocaleDateString()
        : '—';
    }

    if (fs.existsSync(sessionPath)) {
      const session = JSON.parse(fs.readFileSync(sessionPath, 'utf8'));
      status = session.status || 'unknown';
    }

    console.log(
      `  ${name.substring(0, 22).padEnd(23)}  ${slideCount.padEnd(6)}  ${status.padEnd(13)}  ${lastCaptured}`
    );
  }
  console.log('');
}

// ──────────────────────────────────────────────
// NEW COMMAND
// ──────────────────────────────────────────────

export async function runNew(opts) {
  const slug = makeSlug(opts.name);
  const demoDir = path.join(DEMOS_DIR, slug);

  if (fs.existsSync(demoDir)) {
    console.log(`Demo folder already exists: demos/${slug}/`);
    return;
  }

  for (const sub of ['screenshots', 'clips', 'output']) {
    fs.mkdirSync(path.join(demoDir, sub), { recursive: true });
  }

  const yamlContent = `demo:
  title: "${opts.name}"
  description: ""
  brand_color: "#00C9A7"
  m365_copilot_url: ""

  slides:
    - id: 1
      # Platform captured for this slide.
      # Options: m365-copilot | sharepoint | power-automate |
      #          teams | outlook | xero | custom
      platform: m365-copilot

      # Large headline shown above the screenshot.
      story_label: "Your first slide"

      # 1-2 sentence plain-English explanation.
      narrative: "Describe what this slide shows."

      url: ""

      sample_prompts:
        - "Ask the agent something"
      record_clip: true
`;

  fs.writeFileSync(path.join(demoDir, 'demo.yaml'), yamlContent);

  const meta = {
    agent_name: opts.name,
    created_at: new Date().toISOString(),
    last_captured: null,
    last_generated: null,
    slide_count: 1,
    platforms: ['m365-copilot'],
  };
  fs.writeFileSync(path.join(demoDir, '.session-meta.json'), JSON.stringify(meta, null, 2));

  console.log(`✓ Demo scaffolded: demos/${slug}/`);
  console.log(`  Edit demos/${slug}/demo.yaml to configure your slides.`);
}
