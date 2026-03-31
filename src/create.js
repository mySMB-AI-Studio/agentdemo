/**
 * create.js — Single command that does everything end to end.
 *
 * Usage:  agentdemo create
 *
 * STEP 1: Ask for two URLs (Copilot Studio + M365 Copilot)
 * STEP 2: Auto-discover agent info from Copilot Studio
 * STEP 3: Auto-capture from M365 Copilot + connected platforms
 * STEP 4: Auto-generate callout text via Anthropic API
 * STEP 5: Auto-generate Storylane-style demo.html
 * STEP 6: Open demo.html in default browser
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import readline from 'readline';
import slugify from 'slugify';
import { createBrowserContext, isSessionValid, saveSession } from './auth.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, '..');
const DEMOS_DIR = path.join(ROOT, 'demos');

// ───────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────

function makeSlug(name) {
  return slugify(name, { lower: true, strict: true }) || 'my-agent';
}

function rl_ask(question) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(question, (answer) => {
      rl.close();
      resolve((answer || '').split('\n')[0].split('\r')[0].trim());
    });
  });
}

/** Fully blocking Enter wait — does NOT return until user presses Enter */
function waitForEnter(message) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
      terminal: false,
    });
    process.stdout.write('\n' + message + '\nPress Enter to continue... ');
    rl.once('line', () => {
      rl.close();
      resolve();
    });
  });
}

/** Ensure the browser window is as large as possible on Windows */
async function maximizeWindow(page) {
  try {
    await page.evaluate(() => {
      window.moveTo(0, 0);
      window.resizeTo(screen.availWidth, screen.availHeight);
    });
  } catch { /* ignore */ }
  // Guaranteed fallback: set viewport explicitly to 1920x1080
  // This works even when window maximization fails
  try {
    const vp = page.viewportSize();
    if (!vp || vp.width < 1600) {
      await page.setViewportSize({ width: 1920, height: 1080 });
    }
  } catch { /* ignore */ }
}

async function clickFirst(page, selectors, waitMs = 2000) {
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

async function dismissPopups(page) {
  // Dismiss ALL known popups: onboarding, welcome, consent, feature tours, etc.
  // Try multiple rounds (popups can cascade)
  for (let attempt = 0; attempt < 5; attempt++) {
    let found = false;

    // Priority 1: Consent confirmation dialog ("Confirm" button)
    for (const sel of [
      '[role="dialog"] button:has-text("Confirm")',
      'button:has-text("Confirm")',
      '[role="dialog"] button:has-text("Accept")',
    ]) {
      try {
        const el = await page.$(sel);
        if (el && await el.isVisible()) {
          // Verify it's actually a consent/confirmation dialog
          const dialogText = await page.evaluate(() => {
            const d = document.querySelector('[role="dialog"], [class*="modal" i], [class*="dialog" i]');
            return d?.textContent?.toLowerCase() || '';
          });
          if (dialogText.includes('consent') || dialogText.includes('shared with you') ||
              dialogText.includes('confirm') || dialogText.includes('permission')) {
            await el.click();
            await page.waitForTimeout(2000);
            found = true;
            break;
          }
        }
      } catch { /* next */ }
    }
    if (found) continue;

    // Priority 2: Welcome / onboarding / "What's new" popups
    for (const sel of [
      'button:has-text("Skip")', 'button:has-text("skip")',
      'button:has-text("Dismiss")', 'button:has-text("Got it")',
      'button:has-text("Maybe later")', 'button:has-text("Not now")',
      'button:has-text("No thanks")', 'button:has-text("Close")',
      'button[aria-label="Close"]', 'button[aria-label="Dismiss"]',
      'button[aria-label="close"]',
      '[role="dialog"] button:has-text("Skip")',
      '[role="dialog"] button:has-text("Close")',
      '[role="dialog"] button:has-text("Got it")',
      '[role="dialog"] button[aria-label="Close"]',
      // Fluent UI specific
      '.ms-Dialog button[class*="close" i]',
      '.ms-Modal button[class*="close" i]',
      '[class*="teaching" i] button:has-text("Got it")',
      '[class*="teaching" i] button:has-text("Skip")',
      'button[class*="dismiss" i]', 'button[class*="close" i]',
    ]) {
      try {
        const el = await page.$(sel);
        if (el && await el.isVisible()) {
          await el.click();
          await page.waitForTimeout(1500);
          found = true;
          break;
        }
      } catch { /* next */ }
    }

    if (!found) break;
  }
}

/** Connector name → AgentDemo platform slug */
const CONNECTOR_MAP = {
  'sharepoint': 'sharepoint', 'power automate': 'power-automate',
  'microsoft teams': 'teams', 'teams': 'teams',
  'outlook': 'outlook', 'office 365 outlook': 'outlook',
  'xero': 'xero', 'dataverse': 'custom', 'http': 'custom',
};

// ───────────────────────────────────────────
// STEP 1: Ask for two URLs
// ───────────────────────────────────────────

async function askForUrls(opts) {
  let studioUrl = opts.studioUrl || '';
  let m365Url = opts.m365Url || '';

  console.log('\n  AgentDemo — Create a new demo');
  console.log('  ─────────────────────────────\n');

  if (!studioUrl) {
    studioUrl = await rl_ask('  Paste your Copilot Studio agent URL:\n  > ');
    while (!studioUrl.startsWith('https://')) {
      console.log('  Must start with https://');
      studioUrl = await rl_ask('  > ');
    }
  }

  if (!m365Url) {
    m365Url = await rl_ask('\n  Paste your M365 Copilot agent URL:\n  > ');
    while (!m365Url.startsWith('https://')) {
      console.log('  Must start with https://');
      m365Url = await rl_ask('  > ');
    }
  }

  console.log("\n  That's all we need. Starting now...\n");
  return { studioUrl, m365Url };
}

// ───────────────────────────────────────────
// STEP 2: Auto-discover from Copilot Studio
// ───────────────────────────────────────────

async function discoverFromStudio(page, studioUrl) {
  console.log('  ● Discovering agent from Copilot Studio...');
  console.log('    (this may take up to 60 seconds on first load)');

  const discovered = {
    name: 'My Agent',
    description: '',
    topics: [],       // { name, phrases[] }
    connections: [],   // { name, platform }
    flowUrls: [],
    sharepointUrls: [],
  };

  /** Navigate to Studio URL with retry */
  async function loadStudioPage(attempt = 1) {
    await maximizeWindow(page);
    await page.goto(studioUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    // Wait for networkidle with generous timeout
    await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
    // Wait for ANY meaningful content element to confirm SPA loaded
    try {
      await page.waitForSelector(
        'h1, h2, [role="tab"], [role="tablist"], [data-testid*="bot"], [class*="agent" i]',
        { timeout: 30000 }
      );
    } catch {
      if (attempt === 1) {
        console.log('    ⚠ Copilot Studio took too long to load. Retrying once...');
        return loadStudioPage(2);
      }
      throw new Error('Copilot Studio page did not load after 2 attempts');
    }
    await dismissPopups(page);
  }

  try {
    await loadStudioPage();

    // Read agent name — multiple strategies with known-label filtering
    const TAB_NAMES = new Set([
      'overview', 'knowledge', 'tools', 'agents', 'topics', 'activity',
      'evaluation', 'analytics', 'channels', 'settings', 'actions',
      'home', 'flows', 'copilot studio', 'microsoft copilot studio',
    ]);

    discovered.name = await page.evaluate((tabNames) => {
      // Strategy 1: Browser tab title
      // Format: "Overview - Humanitix Event Attendance Agent"
      // or: "Humanitix Event Attendance Agent - Microsoft Copilot Studio"
      const pageTitle = document.title || '';
      const titleParts = pageTitle.split(/\s*[-–|]\s*/);
      // Try each part — skip known tab/app names
      for (const part of titleParts) {
        const cleaned = part.trim();
        if (cleaned && cleaned.length > 2 && cleaned.length < 80
            && !tabNames.has(cleaned.toLowerCase())) {
          return cleaned;
        }
      }
      // Strategy 2: Try from the end of title parts (agent name is often second)
      if (titleParts.length >= 2) {
        const last = titleParts[titleParts.length - 1].trim();
        const secondLast = titleParts[titleParts.length - 2].trim();
        if (secondLast && !tabNames.has(secondLast.toLowerCase()) && secondLast.length > 2) {
          return secondLast;
        }
        if (last && !tabNames.has(last.toLowerCase()) && last.length > 2) {
          return last;
        }
      }

      // Strategy 3: Look for the agent name heading near the agent icon
      // It appears as a large text element at the top of the agent detail page
      for (const sel of [
        'h1', 'h2',
        '[class*="agentName" i]', '[class*="botName" i]',
        '[class*="agent-name" i]', '[class*="bot-name" i]',
        '[data-testid*="agent-name"]', '[data-testid*="bot-name"]',
      ]) {
        const el = document.querySelector(sel);
        if (el) {
          const t = el.textContent?.trim();
          if (t && t.length > 2 && t.length < 80 && !tabNames.has(t.toLowerCase())) {
            return t;
          }
        }
      }
      return '';
    }, Array.from(TAB_NAMES)) || 'My Agent';

    console.log(`  ✓ Agent found: ${discovered.name}`);

    // Read description from Overview tab
    const descr = await page.evaluate(() => {
      const el = document.querySelector('[class*="description" i], [data-testid*="description"]');
      return el?.textContent?.trim()?.substring(0, 200) || '';
    });
    discovered.description = descr;

    // Navigate to Topics tab
    const topicsClicked = await clickFirst(page, [
      'button:has-text("Topics")', 'a:has-text("Topics")',
      '[role="tab"]:has-text("Topics")', '[aria-label*="Topics"]',
    ], 5000);

    if (topicsClicked) {
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(3000);

      // Read topics and trigger phrases
      const topics = await page.evaluate(() => {
        const found = [];
        const rows = document.querySelectorAll('[data-testid*="topic"], [role="row"], tr');
        for (const row of rows) {
          const cells = row.querySelectorAll('[role="gridcell"], td, > div');
          if (cells.length === 0) continue;
          const name = cells[0]?.getAttribute('title') || cells[0]?.textContent?.trim()?.split('\n')[0]?.trim();
          if (!name || name.length < 2 || name.length > 100) continue;
          const lower = name.toLowerCase();
          if (lower.includes('greeting') || lower.includes('goodbye') ||
              lower.includes('escalate') || lower.includes('fallback') ||
              lower.includes('start over') || lower === 'name' || lower === 'status') continue;
          found.push({ name, phrases: [] });
        }
        return found;
      });

      // For each topic, try to read trigger phrases
      for (const topic of topics.slice(0, 8)) {
        try {
          await clickFirst(page, [
            `text="${topic.name}"`, `a:has-text("${topic.name}")`,
          ], 3000);
          await page.waitForTimeout(2000);

          const phrases = await page.evaluate(() => {
            const found = [];
            const phraseEls = document.querySelectorAll(
              '[class*="trigger" i] li, [class*="phrase" i], [data-testid*="phrase"], ' +
              '[class*="trigger" i] span, [class*="utterance" i]'
            );
            for (const el of phraseEls) {
              const t = el.textContent?.trim();
              if (t && t.length > 2 && t.length < 200 && !found.includes(t)) found.push(t);
            }
            return found.slice(0, 3);
          });
          topic.phrases = phrases;
          await page.goBack().catch(() => {});
          await page.waitForTimeout(2000);
        } catch { /* skip */ }
      }

      discovered.topics = topics;
      const phraseCount = topics.reduce((n, t) => n + t.phrases.length, 0);
      console.log(`  ✓ Topics found: ${topics.length} (${phraseCount} trigger phrases)`);
    }

    // Navigate to Actions/Connections
    const actionsClicked = await clickFirst(page, [
      'button:has-text("Actions")', 'a:has-text("Actions")',
      '[role="tab"]:has-text("Actions")', 'button:has-text("Connections")',
    ], 5000);

    if (actionsClicked) {
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(3000);

      const connections = await page.evaluate((connMap) => {
        const found = [];
        const rows = document.querySelectorAll('[role="row"], tr, [class*="connector" i], [class*="connection" i]');
        for (const row of rows) {
          const text = row.textContent?.trim()?.toLowerCase() || '';
          for (const [keyword, platform] of Object.entries(connMap)) {
            if (text.includes(keyword) && !found.some(c => c.platform === platform)) {
              found.push({ name: keyword, platform });
            }
          }
        }
        return found;
      }, CONNECTOR_MAP);

      discovered.connections = connections;
      if (connections.length > 0) {
        console.log(`  ✓ Connections found: ${connections.map(c => c.name).join(', ')}`);
      }
    }

    console.log('');
  } catch (err) {
    console.log(`    ⚠ Could not read agent details from Copilot Studio.`);
    console.log(`      Proceeding with M365 Copilot URL only.`);
    console.log(`      Demo will have fewer slides — you can add more later by editing demo.yaml\n`);
  }

  return discovered;
}

// ───────────────────────────────────────────
// STEP 2B: Manual fallback when discovery fails
// ───────────────────────────────────────────

const VALID_PLATFORMS = new Set(['sharepoint', 'power-automate', 'teams', 'outlook', 'xero', 'custom']);

async function askManualFallback(discovered) {
  console.log('');
  console.log('  ⚠ Could not read agent details from Copilot Studio.\n');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const ask = (q) => new Promise(r => rl.question(q, a => r((a || '').split('\n')[0].split('\r')[0].trim())));

  try {
    const doManual = await ask('  Would you like to provide agent details manually? (y/n): ');
    if (doManual.toLowerCase() !== 'y') {
      console.log('  Continuing with M365 Copilot URL only.');
      console.log('  Demo will use generic text — you can improve it later.\n');
      rl.close();
      return;
    }

    // Agent name
    const name = await ask('\n  Agent name:\n  > ');
    if (name) discovered.name = name;

    // Description
    const desc = await ask('\n  One-line description (shown on demo intro slide):\n  > ');
    if (desc) discovered.description = desc;

    // Instructions (multi-line — read until BLANK LINE)
    console.log('\n  Paste the agent\'s instructions (from Copilot Studio Overview page).');
    console.log('  Press Enter on an empty line when done:');
    let instructions = '';
    let instrLineCount = 0;
    while (true) {
      const line = await ask('  > ');
      if (!line) break;
      instructions += (instructions ? ' ' : '') + line;
      instrLineCount++;
    }
    if (instructions) {
      discovered.instructions = instructions;
      console.log(`\n  ✓ Instructions saved (${instrLineCount} lines)`);
    }

    // Clear separator before platform input
    console.log('\n  ─────────────────────────────────────────────');
    console.log('  Now, what platforms does this agent connect to?');
    console.log('  ─────────────────────────────────────────────');
    console.log('  Type each one and press Enter.');
    console.log('  Press Enter alone when done.');
    console.log('  Options: sharepoint | power-automate | teams | outlook | xero | custom\n');
    while (true) {
      const plat = await ask('  > ');
      if (!plat) break;
      const p = plat.toLowerCase().trim();
      if (VALID_PLATFORMS.has(p)) {
        if (!discovered.connections.some(c => c.platform === p)) {
          discovered.connections.push({ name: p, platform: p });
          console.log(`    ✓ Added: ${p}`);
        } else {
          console.log(`    Already added: ${p}`);
        }
      } else {
        console.log(`    Not recognised: '${plat}'.`);
        console.log('    Valid options: sharepoint | power-automate | teams | outlook | xero | custom');
      }
    }

    console.log('\n  ✓ Got it. Continuing with M365 Copilot capture...\n');
  } finally {
    rl.close();
  }
}

// ───────────────────────────────────────────
// STEP 2C: Generate demo prompts via Anthropic API
// ───────────────────────────────────────────

async function generateDemoPrompts(discovered) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null; // caller will use welcome screen prompts instead

  console.log('  ● Generating demo prompts using AI...');

  let Anthropic;
  try {
    const mod = await import('@anthropic-ai/sdk');
    Anthropic = mod.default || mod.Anthropic;
  } catch {
    console.log('    ⚠ @anthropic-ai/sdk not installed. Will read prompts from agent welcome screen.');
    return null;
  }

  const client = new Anthropic({ apiKey });
  const platforms = discovered.connections.map(c => c.platform).join(', ') || 'none detected';

  try {
    const response = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: `You are helping create a product demo for an AI agent.

Agent name: ${discovered.name}
Agent description: ${discovered.description || 'Not provided'}
Agent instructions: ${discovered.instructions || 'Not provided'}
Connected platforms: ${platforms}

Generate 3-5 realistic demo prompts that:
1. Showcase the agent's most impressive capabilities
2. Are written as a real user would type them
3. Progress from simple to more complex
4. Each prompt should produce a visually interesting response (lists, tables, summaries — not just yes/no)
5. Are specific to THIS agent's actual purpose

Also generate a one-sentence demo hook for the intro slide that would make a business decision maker want to watch.

Respond in JSON only, no markdown:
{"hook": "one sentence that sells the demo", "prompts": [{"text": "the prompt text", "purpose": "what this prompt demonstrates", "slide_label": "short business-level headline for this slide"}]}`,
      }],
    });

    const jsonStr = response.content[0]?.text || '';
    const cleaned = jsonStr.replace(/```json?\s*/g, '').replace(/```/g, '').trim();
    const result = JSON.parse(cleaned);

    if (result.prompts && result.prompts.length > 0) {
      console.log('');
      console.log('  ✓ Generated demo prompts:');
      result.prompts.forEach((p, i) => {
        console.log(`    ${i + 1}. "${p.text}"`);
        console.log(`       → ${p.purpose}`);
      });
      console.log('');
      return result;
    }
  } catch (err) {
    console.log(`    ⚠ Prompt generation failed: ${err.message?.substring(0, 60)}`);
  }

  return null;
}

// ───────────────────────────────────────────
// STEP 3: Auto-capture
// ───────────────────────────────────────────

/** Verify the correct agent is loaded after navigation (not generic Copilot) */
async function verifyAgentLoaded(page, agentName, m365Url) {
  if (!agentName || agentName === 'My Agent') return; // can't verify generic name
  await page.waitForTimeout(3000);
  const bodyText = await page.textContent('body').catch(() => '');
  if (bodyText.includes(agentName)) {
    console.log(`    ✓ Agent verified: ${agentName}`);
    return;
  }
  // Retry navigation once
  console.log(`    ⚠ Agent not loaded correctly, retrying navigation...`);
  await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle').catch(() => {});
  await page.waitForTimeout(5000);
  await dismissPopups(page);
  const retryText = await page.textContent('body').catch(() => '');
  if (retryText.includes(agentName)) {
    console.log(`    ✓ Agent verified on retry: ${agentName}`);
  } else {
    console.log(`    ⚠ Could not verify agent loaded correctly`);
  }
}

/** Wait for agent response using stop-button detection (most reliable method).
 *  Returns { complete: bool, lastMsgText: string } */
async function waitForAgentResponse(page, maxSeconds = 120) {
  const start = Date.now();
  const sec = () => Math.round((Date.now() - start) / 1000);

  // Selectors for the stop/cancel button that appears during generation
  const STOP_SELS = 'button[aria-label*="Stop" i], button[title*="Stop" i], [data-testid*="stop" i], button[aria-label*="Cancel" i]';
  // Phrases indicating still loading
  const LOADING_PHRASES = ['lining things up', 'generating', 'working on it', 'just a moment', 'searching for', 'thinking'];

  // ── PHASE 1: Wait for stop button to APPEAR (agent started) ──
  let responseStarted = false;
  try {
    await page.waitForSelector(STOP_SELS, { timeout: 30000 });
    responseStarted = true;
    process.stdout.write(`\r    ● Agent started responding...              `);
  } catch {
    // Fallback: check if text changed (stop button may not appear for fast responses)
    process.stdout.write(`\r    ● Could not detect stop button, waiting...`);
    await page.waitForTimeout(5000);
    responseStarted = true; // assume it started
  }

  // ── PHASE 2: Wait for stop button to DISAPPEAR (response complete) ──
  let responseComplete = false;
  try {
    await page.waitForSelector(STOP_SELS, { state: 'hidden', timeout: maxSeconds * 1000 });
    responseComplete = true;
    process.stdout.write(`\r    ● Stop button gone — checking final state...`);
  } catch {
    process.stdout.write(`\r    ● Timed out waiting (${sec()}s). Taking screenshot anyway.`);
  }

  // ── PHASE 3: Wait for loading text to be gone (up to 15 more seconds) ──
  const guardEnd = Date.now() + 15000;
  while (Date.now() < guardEnd) {
    const stillLoading = await page.evaluate((phrases) => {
      const els = document.querySelectorAll('span, p, div');
      for (const el of els) {
        const t = el.textContent?.trim()?.toLowerCase() || '';
        if (t.length > 0 && t.length < 60) {
          for (const phrase of phrases) {
            if (t.includes(phrase)) return true;
          }
        }
      }
      // Check for visible spinners
      for (const sel of ['[class*="spinner" i]', '[class*="loading" i]', '[class*="streaming" i]']) {
        try { const el = document.querySelector(sel); if (el && el.offsetParent !== null) return true; } catch {}
      }
      return false;
    }, LOADING_PHRASES).catch(() => false);
    if (!stillLoading) break;
    await page.waitForTimeout(1000);
  }

  // Extra settle time
  await page.waitForTimeout(1500);

  process.stdout.write('\r' + ' '.repeat(60) + '\r');

  // Get last message text
  const lastMsgText = await page.evaluate(() => {
    const sels = ['[data-testid*="message"]', '[class*="message" i]', '[class*="response" i]', '[class*="assistant" i]', '[role="listitem"]'];
    for (const sel of sels) {
      const els = document.querySelectorAll(sel);
      if (els.length > 0) return els[els.length - 1]?.innerText?.substring(0, 800) || '';
    }
    return '';
  }).catch(() => '');

  const elapsed = sec();
  if (responseComplete) {
    console.log(`    ✓ Response complete (${elapsed} seconds)`);
  } else {
    console.log(`    ⚠ Response may be incomplete (${elapsed}s timeout)`);
  }

  return { complete: responseComplete, lastMsgText, elapsed };
}

/** Scroll the last agent response to the top of the viewport, then screenshot */
async function scrollAndScreenshot(page, screenshotPath) {
  await page.evaluate(() => {
    const msgSelectors = [
      '[data-testid*="assistant"]', '[class*="assistant" i]',
      '[class*="agent-message" i]', '[class*="bot-message" i]',
      '[class*="copilot" i][class*="message" i]',
      '[data-testid*="message"]', '[class*="message" i]',
      '[class*="response" i]', '[role="listitem"]',
    ];
    for (const sel of msgSelectors) {
      const els = document.querySelectorAll(sel);
      if (els.length > 0) {
        els[els.length - 1].scrollIntoView({ behavior: 'instant', block: 'start' });
        return;
      }
    }
    window.scrollTo(0, document.body.scrollHeight);
  }).catch(() => {});
  await page.waitForTimeout(1500);

  await page.screenshot({ path: screenshotPath, fullPage: false });

  // Verify screenshot not blank
  const size = fs.statSync(screenshotPath).size;
  if (size < 50000) {
    await page.waitForTimeout(2000);
    await page.screenshot({ path: screenshotPath, fullPage: false });
  }
}

/** Detect connection approval requests in the page */
async function detectConnectionRequest(page) {
  return page.evaluate(() => {
    const CONNECTION_KEYWORDS = [
      'connect to', 'sign in to', 'authorize', 'approve',
      'connection required', 'action required', 'please connect',
      'needs access to', 'connect your', 'sign-in required',
      'connection manager',
    ];

    // Words that are UI labels, NOT platform names
    const SKIP_WORDS = new Set([
      'open', 'close', 'cancel', 'confirm', 'ok', 'yes', 'no',
      'connect', 'sign', 'authorize', 'connection', 'manager',
      'settings', 'here', 'click', 'link', 'button', 'more',
      'learn', 'this', 'the', 'your', 'a', 'an', 'to', 'in',
      'required', 'action', 'please', 'now', 'continue',
    ]);

    function cleanPlatformName(raw) {
      if (!raw) return '';
      const name = raw.trim().replace(/[\.\,\!\?]+$/, '').trim();
      if (name.length < 2 || name.length > 50) return '';
      if (SKIP_WORDS.has(name.toLowerCase())) return '';
      // Also skip if ALL words are skip words
      const words = name.split(/\s+/);
      if (words.every(w => SKIP_WORDS.has(w.toLowerCase()))) return '';
      return name;
    }

    // Check for modal/dialog
    const dialog = document.querySelector('[role="dialog"], [class*="modal" i], [class*="dialog" i]');
    if (dialog && dialog.offsetParent !== null) {
      const dt = dialog.textContent?.toLowerCase() || '';
      for (const kw of CONNECTION_KEYWORDS) {
        if (dt.includes(kw)) {
          const match = dt.match(/(?:connect to|sign in to|authorize|access to)\s+([A-Za-z0-9\s]+?)(?:\.|,|$|\s+to\s)/i);
          const name = cleanPlatformName(match?.[1]);
          return { detected: true, platform: name || 'your connected platform', source: 'dialog' };
        }
      }
    }

    // Check the last few messages in chat
    const msgs = document.querySelectorAll('[class*="message" i], [class*="response" i], [role="listitem"]');
    for (let i = Math.max(0, msgs.length - 3); i < msgs.length; i++) {
      const text = msgs[i]?.textContent?.toLowerCase() || '';
      for (const kw of CONNECTION_KEYWORDS) {
        if (text.includes(kw)) {
          const match = text.match(/(?:connect to|sign in to|authorize|access to)\s+([A-Za-z0-9\s]+?)(?:\.|,|$|\s+to\s)/i);
          const name = cleanPlatformName(match?.[1]);
          return { detected: true, platform: name || 'your connected platform', source: 'chat' };
        }
      }
    }

    // Check for "Connect" or "Authorize" buttons in chat area
    const btns = document.querySelectorAll('button, a');
    for (const b of btns) {
      const t = b.textContent?.trim() || '';
      const tl = t.toLowerCase();
      if ((tl === 'connect' || tl === 'authorize' || tl === 'sign in' || tl.startsWith('connect to'))
          && b.offsetParent !== null) {
        const name = cleanPlatformName(t.replace(/^connect\s+to\s*/i, ''));
        return { detected: true, platform: name || 'your connected platform', source: 'button' };
      }
    }

    return { detected: false };
  }).catch(() => ({ detected: false }));
}

/** Pause and show connection approval instructions */
async function handleConnectionRequest(page, platform, slideId, screenshotsDir) {
  // 1. Screenshot the connection request (browser stays OPEN)
  const ssPath = path.join(screenshotsDir, `${slideId}-connection-request.png`);
  await page.screenshot({ path: ssPath }).catch(() => {});

  // 2. Extract better platform name from visible page
  const detectedName = await page.evaluate(() => {
    const SKIP = new Set([
      'open', 'close', 'cancel', 'confirm', 'ok', 'yes', 'no',
      'connect', 'sign', 'authorize', 'connection', 'manager',
      'settings', 'here', 'click', 'link', 'button', 'more',
      'learn', 'this', 'the', 'your', 'a', 'an', 'to', 'in',
      'required', 'action', 'please', 'now', 'continue',
    ]);
    function clean(raw) {
      if (!raw) return '';
      const n = raw.trim().replace(/[\.\,\!\?]+$/, '').trim();
      if (n.length < 2 || n.length > 50) return '';
      if (SKIP.has(n.toLowerCase())) return '';
      if (n.split(/\s+/).every(w => SKIP.has(w.toLowerCase()))) return '';
      return n;
    }
    // Check buttons for "Connect to X" / "Sign in to X"
    const btns = document.querySelectorAll('button, a, [role="button"]');
    for (const b of btns) {
      const t = b.textContent?.trim() || '';
      for (const re of [/connect\s+to\s+(.+)/i, /sign\s+in\s+to\s+(.+)/i, /authorize\s+(.+)/i]) {
        const m = t.match(re);
        const name = clean(m?.[1]);
        if (name) return name;
      }
    }
    // Check chat messages for platform name patterns
    const msgs = document.querySelectorAll('[class*="message" i], [class*="response" i], [role="listitem"]');
    for (const msg of msgs) {
      const t = msg.textContent || '';
      for (const re of [/connect\s+to\s+([A-Z][A-Za-z0-9\s]+?)[\.,]/i, /([A-Z][A-Za-z0-9]+)\s+connection/i]) {
        const m = t.match(re);
        const name = clean(m?.[1]);
        if (name) return name;
      }
    }
    return '';
  }).catch(() => '');

  const displayPlatform = detectedName || platform || 'your connected platform';

  // 3. Print pause message — browser is still open and visible
  console.log('');
  console.log('  ┌─────────────────────────────────────────────┐');
  console.log('  │  ⚠ CONNECTION APPROVAL REQUIRED             │');
  console.log('  │                                             │');
  console.log(`  │  The agent needs you to approve a           │`);
  console.log(`  │  connection to: ${displayPlatform.substring(0, 27).padEnd(27)}│`);
  console.log('  │                                             │');
  console.log('  │  Steps:                                     │');
  console.log('  │  1. Look at the browser window              │');
  console.log('  │  2. Click "Connect" or follow the link      │');
  console.log('  │  3. Complete the authorization flow         │');
  console.log('  │  4. Return here and press Enter to retry    │');
  console.log('  └─────────────────────────────────────────────┘');

  // 4. BLOCKING WAIT — nothing runs until user presses Enter
  await waitForEnter('  Approve the connection in the browser, then:');
  console.log('  RESUMED — user pressed Enter');

  return { ssPath, platform: displayPlatform };
}

async function autoCapture(page, m365Url, discovered, demoDir, generatedPrompts = null) {
  const screenshotsDir = path.join(demoDir, 'screenshots');
  const clipsDir = path.join(demoDir, 'clips');
  fs.mkdirSync(screenshotsDir, { recursive: true });
  fs.mkdirSync(clipsDir, { recursive: true });

  const slides = [];
  const connectionEvents = []; // { platform, screenshotPath } — for connection setup slide
  let slideId = 1;

  // ── Capture connected platform slides first ──
  for (const conn of discovered.connections) {
    const platform = conn.platform;
    if (platform === 'm365-copilot') continue;

    console.log(`  ● Capturing ${platform}...`);

    // These external platforms typically can't be auto-captured:
    // they need a trigger to have fired (email sent, flow run, invoice updated).
    // Create a placeholder slide with metadata for the developer.
    slides.push({
      id: slideId++,
      platform,
      type: 'platform',
      placeholder: true,
      connectorName: conn.name,
      screenshot: null,
      clip: null,
      prompt: null,
      callout: null,
      placeholderInfo: null, // filled by AI later
    });
    console.log(`    → Placeholder (needs manual screenshot)`);
  }

  // ── Capture M365 Copilot agent interaction slides ──
  console.log('  ● Opening M365 Copilot...');
  try {
    await maximizeWindow(page);
    await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(8000);
    await dismissPopups(page);

    // Find chat input — try targeted selectors in order.
    // M365 Copilot uses a Lexical rich text editor, so the actual input
    // is a contenteditable <div> or <p>, NOT a regular <input>/<textarea>.
    // We must verify the matched element is editable and NOT a button.
    const INPUT_SELECTORS = [
      'div[data-testid="chat-input"] [contenteditable]',
      '[role="textbox"][contenteditable="true"]',
      'p[data-lexical-editor="true"]',
      'div[contenteditable="true"][data-lexical-editor]',
      'div[contenteditable="true"].notranslate',
      '[placeholder*="Message Copilot" i]',
      '[placeholder*="Ask Copilot" i]',
      '[placeholder*="Type a message" i]',
      '[placeholder*="Message" i]',
      'div[contenteditable="true"]',
      'textarea[placeholder]',
    ];

    let INPUT_SEL = ''; // the selector that worked
    let chatLoaded = false;
    for (const sel of INPUT_SELECTORS) {
      try {
        const loc = page.locator(sel).last();
        await loc.waitFor({ state: 'visible', timeout: 3000 });

        // Verify this is actually an editable element, not a button
        const elInfo = await loc.evaluate(node => ({
          tag: node.tagName,
          role: node.getAttribute('role'),
          contenteditable: node.getAttribute('contenteditable'),
          ariaLabel: node.getAttribute('aria-label') || '',
        }));

        if (elInfo.tag === 'BUTTON' || elInfo.role === 'button') {
          console.log(`    Skipping button element: ${sel} (aria-label="${elInfo.ariaLabel}")`);
          continue;
        }

        INPUT_SEL = sel;
        chatLoaded = true;
        console.log(`  ✓ Found chat input: ${sel} (tag: ${elInfo.tag}, contenteditable: ${elInfo.contenteditable})`);
        break;
      } catch { continue; }
    }

    if (chatLoaded) {
      // Pre-check: detect connection requests on the welcome screen
      const preCheck = await detectConnectionRequest(page);
      if (preCheck.detected) {
        const connResult = await handleConnectionRequest(page, preCheck.platform, 'pre', screenshotsDir);
        connectionEvents.push({ platform: connResult.platform, screenshotPath: connResult.ssPath });
        await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await page.waitForLoadState('networkidle').catch(() => {});
        await page.waitForTimeout(8000);
        await dismissPopups(page);
      }
    }

    if (!chatLoaded) {
      // Debug: save screenshot and log all input-like elements
      const debugSS = path.join(screenshotsDir, `debug-no-input-${Date.now()}.png`);
      await page.screenshot({ path: debugSS, fullPage: false }).catch(() => {});
      const elements = await page.evaluate(() => {
        return Array.from(document.querySelectorAll(
          'input, textarea, [contenteditable], [role="textbox"]'
        )).map(el => ({
          tag: el.tagName,
          type: el.type || '',
          placeholder: el.placeholder || '',
          ariaLabel: el.getAttribute('aria-label') || '',
          role: el.getAttribute('role') || '',
          contenteditable: el.getAttribute('contenteditable') || '',
          id: el.id || '',
          className: (el.className || '').substring(0, 60),
        }));
      }).catch(() => []);
      console.log('    ⚠ Could not find chat input after trying all selectors.');
      console.log(`    Debug screenshot: ${debugSS}`);
      console.log('    Input elements found on page:');
      if (elements.length === 0) {
        console.log('      (none)');
      } else {
        for (const el of elements) {
          console.log(`      <${el.tag}> placeholder="${el.placeholder}" aria-label="${el.ariaLabel}" role="${el.role}" id="${el.id}"`);
        }
      }
      // Still save a landing screenshot as a slide
      const ssPath = path.join(screenshotsDir, `${slideId}-m365-landing.png`);
      await page.screenshot({ path: ssPath, fullPage: false });
      slides.push({
        id: slideId++, platform: 'm365-copilot', type: 'agent',
        screenshot: ssPath, clip: null, prompt: null, callout: null,
      });
    } else {
      // ── Collect prompts — three strategies in priority order ──
      let prompts = [];
      let promptLabels = {}; // prompt text → slide label

      // Strategy 1: Use AI-generated prompts if available
      if (generatedPrompts && generatedPrompts.prompts && generatedPrompts.prompts.length > 0) {
        for (const p of generatedPrompts.prompts) {
          prompts.push(p.text);
          promptLabels[p.text] = p.slide_label || '';
        }
        console.log(`  Using ${prompts.length} AI-generated prompts`);
      }

      // Strategy 2: Use discovered trigger phrases from Copilot Studio
      if (prompts.length === 0 && discovered.topics.length > 0) {
        for (const topic of discovered.topics) {
          if (topic.phrases.length > 0) {
            prompts.push(...topic.phrases.slice(0, 2));
          } else if (topic.name.length > 5) {
            prompts.push(topic.name);
          }
        }
        if (prompts.length > 0) {
          console.log(`  Using ${prompts.length} prompts from Copilot Studio topics`);
        }
      }

      // Strategy 3: Read starter prompts from M365 Copilot welcome screen
      if (prompts.length === 0) {
        console.log('  ● Reading starter prompts from agent welcome screen...');
        const starterPrompts = await page.evaluate(() => {
          const found = [];
          const sels = [
            '[data-testid*="suggestion"]', '[class*="suggestion" i]',
            '[class*="starter" i]', '[class*="prompt-chip" i]',
            '[class*="prompt" i][class*="card" i]',
            '[class*="recommended" i]', '[class*="sample" i]',
          ];
          for (const sel of sels) {
            const els = document.querySelectorAll(sel);
            for (const el of els) {
              const t = el.textContent?.trim();
              if (t && t.length > 5 && t.length < 150 && !found.includes(t)) {
                found.push(t);
              }
            }
            if (found.length > 0) break;
          }
          // Fallback: look for any clickable elements in the welcome area with short text
          if (found.length === 0) {
            const buttons = document.querySelectorAll('button, [role="button"]');
            for (const b of buttons) {
              const t = b.textContent?.trim();
              if (t && t.length > 10 && t.length < 120 && !t.toLowerCase().includes('send')
                  && !t.toLowerCase().includes('attach') && !t.toLowerCase().includes('mic')) {
                found.push(t);
              }
            }
          }
          return found.slice(0, 5);
        }).catch(() => []);

        if (starterPrompts.length > 0) {
          prompts = starterPrompts;
          console.log(`  ✓ Found ${prompts.length} starter prompts from welcome screen`);
        } else {
          // Absolute last resort: use agent name as prompt context
          prompts = [`Tell me about your capabilities and what you can help with`];
          console.log('  ⚠ No starter prompts found. Using generic prompt.');
        }
      }

      // Limit to 5 prompts max
      prompts = prompts.slice(0, 5);

      const totalPrompts = prompts.length;
      const RESPONSE_TIMEOUT = 120; // 2 minutes max wait
      const ERROR_PHRASES = ['something went wrong', "i'm having trouble", "couldn't complete", 'connection failed'];

      // Selectors for the "New chat" button to start a fresh conversation per prompt
      const NEW_CHAT_SELS = [
        '[aria-label*="New chat" i]',
        '[title*="New chat" i]',
        'button:has-text("New chat")',
        '[data-testid*="new-chat"]',
        '[data-testid*="newChat"]',
      ].join(',');

      for (let i = 0; i < prompts.length; i++) {
        const prompt = prompts[i];
        console.log(`\n  ● Slide ${slideId} — M365 Copilot`);
        console.log(`    Typing prompt: "${prompt.substring(0, 60)}${prompt.length > 60 ? '...' : ''}"`);

        try {
          // Start a fresh chat for each prompt (except the first)
          // so each slide shows one clean exchange only
          if (i > 0) {
            // Navigate fresh to the agent URL for a clean conversation
            await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
            await page.waitForLoadState('networkidle').catch(() => {});
            await page.waitForTimeout(5000);
            await dismissPopups(page);
            await verifyAgentLoaded(page, discovered.name, m365Url);
          }

          // Type prompt — always re-query the input (fresh locator, no stale refs)
          // M365 Copilot uses a Lexical rich text editor which intercepts
          // keyboard events. fill() does not work — must use keyboard.type().
          const inputLocator = page.locator(INPUT_SEL).last();
          await inputLocator.waitFor({ state: 'visible', timeout: 15000 });
          await inputLocator.click();
          await page.waitForTimeout(300);
          // Select all existing text and clear it before typing
          await page.keyboard.press('Control+A');
          await page.keyboard.press('Backspace');
          await page.waitForTimeout(200);
          // Type with realistic keystroke speed (50ms per char)
          await page.keyboard.type(prompt, { delay: 50 });
          await page.waitForTimeout(500);

          // Press Enter to send
          await page.keyboard.press('Enter');

          let isPartial = false;
          let connectionDetected = false;

          // ── Wait for response using stop-button detection ──
          let result = await waitForAgentResponse(page, RESPONSE_TIMEOUT);

          // Check for connection request if response seems stalled
          if (!result.complete && result.lastMsgText.length < 10) {
            const connCheck = await detectConnectionRequest(page);
            if (connCheck.detected) {
              const connResult = await handleConnectionRequest(page, connCheck.platform, slideId, screenshotsDir);
              connectionEvents.push({ platform: connResult.platform, screenshotPath: connResult.ssPath });
              connectionDetected = true;
            }
          }

          // If connection was detected, retry this prompt
          if (connectionDetected) {
            console.log('    Retrying prompt after connection approval...');
            await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
            await page.waitForLoadState('networkidle').catch(() => {});
            await page.waitForTimeout(8000);
            await dismissPopups(page);
            await verifyAgentLoaded(page, discovered.name, m365Url);
            try {
              const retryLocator = page.locator(INPUT_SEL).last();
              await retryLocator.waitFor({ state: 'visible', timeout: 15000 });
              await retryLocator.click();
              await page.waitForTimeout(300);
              await page.keyboard.press('Control+A');
              await page.keyboard.press('Backspace');
              await page.waitForTimeout(200);
              await page.keyboard.type(prompt, { delay: 50 });
              await page.waitForTimeout(500);
              await page.keyboard.press('Enter');
              result = await waitForAgentResponse(page, RESPONSE_TIMEOUT);
            } catch (retryErr) {
              console.log(`    ⚠ Retry failed: ${retryErr.message?.substring(0, 50)}`);
              isPartial = true;
            }
          }

          if (!result.complete) isPartial = true;

          // Check for error response
          const hasError = ERROR_PHRASES.some(phrase => (result.lastMsgText || '').toLowerCase().includes(phrase));
          if (hasError) {
            console.log(`    ⚠ Agent returned an error message — saved as needs-review`);
          }

          // Scroll and screenshot
          const ssPath = path.join(screenshotsDir, `${slideId}-m365-prompt-${i + 1}.png`);
          await scrollAndScreenshot(page, ssPath);
          console.log(`    ✓ Screenshot saved`);

          slides.push({
            id: slideId++,
            platform: 'm365-copilot',
            type: 'agent',
            screenshot: ssPath,
            clip: null,
            prompt,
            callout: null,
            partial: isPartial,
            needsReview: hasError,
          });

          // Wait before next prompt
          await page.waitForTimeout(2000);
        } catch (err) {
          process.stdout.write('\r' + ' '.repeat(60) + '\r');
          console.log(`    ⚠ Prompt failed: ${err.message}`);
          const ssPath = path.join(screenshotsDir, `${slideId}-m365-error-${i + 1}.png`);
          await page.screenshot({ path: ssPath }).catch(() => {});
          slides.push({
            id: slideId++, platform: 'm365-copilot', type: 'agent',
            screenshot: ssPath, clip: null, prompt, callout: null,
          });
        }
      }
    }
  } catch (err) {
    console.log(`    ⚠ M365 Copilot capture failed: ${err.message}`);
  }

  return { slides, connectionEvents };
}

// ───────────────────────────────────────────
// STEP 4: Auto-generate callout text via API
// ───────────────────────────────────────────

async function generateCallouts(slides, agentName, description) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.log('  ● Skipping AI callout generation (no ANTHROPIC_API_KEY in .env)');
    console.log('    Using auto-generated placeholder text instead.\n');
    // Generate simple callouts without API
    for (const slide of slides) {
      if (slide.type === 'connection-setup') {
        const platName = slide.connectionPlatform || 'your accounts';
        slide.callout = {
          text: `Before the agent can access ${platName}, users approve a one-time connection. After approval, the agent works seamlessly every time.`,
          position: 'bottom-left',
          point_to: { x: 50, y: 50 },
        };
      } else if (slide.type === 'agent' && slide.prompt) {
        slide.callout = {
          text: `Here the agent responds to "${slide.prompt}" with accurate, sourced answers from your connected data.`,
          position: 'bottom-right',
          point_to: { x: 50, y: 50 },
        };
      } else if (slide.type === 'platform') {
        slide.callout = {
          text: `This ${slide.platform.replace(/-/g, ' ')} integration keeps your data connected and always up to date.`,
          position: 'top-right',
          point_to: { x: 50, y: 40 },
        };
      }
    }
    return;
  }

  console.log('  ● Generating callout text via Anthropic API...');

  let Anthropic;
  try {
    const mod = await import('@anthropic-ai/sdk');
    Anthropic = mod.default || mod.Anthropic;
  } catch {
    console.log('    ⚠ @anthropic-ai/sdk not installed. Using placeholder text.');
    // Fall back to simple callouts
    for (const slide of slides) {
      if (slide.type === 'connection-setup') {
        const platName = slide.connectionPlatform || 'your accounts';
        slide.callout = { text: `Before the agent can access ${platName}, users approve a one-time connection. After that, the agent works seamlessly.`, position: 'bottom-left', point_to: { x: 50, y: 50 } };
      } else if (slide.type === 'agent' && slide.prompt) {
        slide.callout = { text: `The agent responds to "${slide.prompt}" with accurate, real-time answers.`, position: 'bottom-left', point_to: { x: 50, y: 50 } };
      } else if (slide.type === 'platform') {
        slide.callout = { text: `This ${slide.platform.replace(/-/g, ' ')} integration powers the agent with connected data.`, position: 'top-right', point_to: { x: 50, y: 40 } };
      }
    }
    return;
  }

  const client = new Anthropic({ apiKey });

  for (const slide of slides) {
    if (!slide.screenshot || !fs.existsSync(slide.screenshot)) continue;

    try {
      const imgBuffer = fs.readFileSync(slide.screenshot);
      const base64 = imgBuffer.toString('base64');
      const mediaType = 'image/png';

      const platformLabel = slide.platform === 'm365-copilot' ? 'M365 Copilot (agent chat)' : slide.platform;

      const response = await client.messages.create({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 300,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'image',
              source: { type: 'base64', media_type: mediaType, data: base64 },
            },
            {
              type: 'text',
              text: `You are writing guided demo narration for a business audience. This screenshot is from slide ${slide.id} of a demo for an AI agent called "${agentName}".

Agent description: ${description || 'An AI agent built with Microsoft Copilot Studio.'}
Platform shown: ${platformLabel}
${slide.prompt ? `User prompt: "${slide.prompt}"` : ''}

Write a single callout bubble text (2-3 sentences max) that:
- Explains what the viewer is seeing in plain English
- Highlights the most impressive or useful thing visible
- Sounds like a friendly guided tour, not a manual
- Is written for a business decision maker, not a developer

Also suggest where the callout bubble should be positioned:
top-left | top-right | bottom-left | bottom-right | center
Choose the position that covers the least important part of the screenshot.

Respond in JSON only:
{"text": "...", "position": "...", "point_to": {"x": 0, "y": 0}}
where x and y are percentages (0-100) indicating where the arrow should point.`,
            },
          ],
        }],
      });

      const jsonStr = response.content[0]?.text || '';
      // Parse JSON from response (handle markdown code blocks)
      const cleaned = jsonStr.replace(/```json?\s*/g, '').replace(/```/g, '').trim();
      const parsed = JSON.parse(cleaned);
      slide.callout = {
        text: parsed.text || '',
        position: parsed.position || 'bottom-right',
        point_to: parsed.point_to || { x: 50, y: 50 },
      };
      console.log(`    Slide ${slide.id}: ✓`);
    } catch (err) {
      console.log(`    Slide ${slide.id}: ⚠ ${err.message?.substring(0, 60)}`);
      // Fallback
      slide.callout = {
        text: slide.prompt
          ? `The agent responds to "${slide.prompt}" with accurate, real-time answers.`
          : `This shows the ${slide.platform} integration powering the agent.`,
        position: 'bottom-right',
        point_to: { x: 50, y: 50 },
      };
    }
  }
}

// ───────────────────────────────────────────
// STEP 4B: Generate placeholder instructions
// ───────────────────────────────────────────

async function generatePlaceholderInfo(slides, agentName, description) {
  const placeholders = slides.filter(s => s.placeholder);
  if (placeholders.length === 0) return;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    // Generate simple placeholder info without API
    for (const slide of placeholders) {
      slide.placeholderInfo = {
        what_to_capture: `Open ${slide.platform} and screenshot the screen showing data related to ${agentName}.`,
        suggested_callout: `This ${slide.platform.replace(/-/g, ' ')} integration keeps your data connected and always up to date for the agent.`,
        suggested_position: 'top-right',
        setup_required: `Ensure the ${slide.connectorName || slide.platform} connection is active and has recent data.`,
      };
    }
    return;
  }

  console.log('  ● Generating placeholder instructions...');
  let Anthropic;
  try {
    const mod = await import('@anthropic-ai/sdk');
    Anthropic = mod.default || mod.Anthropic;
  } catch {
    return generatePlaceholderInfo(slides, agentName, description); // fallback
  }

  const client = new Anthropic({ apiKey });

  for (const slide of placeholders) {
    try {
      const response = await client.messages.create({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 400,
        messages: [{
          role: 'user',
          content: `An AI agent called "${agentName}" has an integration with ${slide.platform}.
Agent description: ${description || 'An AI agent built with Microsoft Copilot Studio.'}
Connected action/connector: ${slide.connectorName || slide.platform}

Generate instructions for a developer who needs to manually capture a screenshot of ${slide.platform} to use in a product demo.

Respond in JSON only:
{"what_to_capture": "1-2 sentences describing exactly what screen/state to screenshot. Be specific.", "suggested_callout": "2-3 sentence callout bubble text written for a business audience explaining what this platform does in the context of the agent.", "suggested_position": "top-left|top-right|bottom-left|bottom-right", "setup_required": "1 sentence describing any state that needs to exist before taking the screenshot."}`,
        }],
      });

      const jsonStr = response.content[0]?.text || '';
      const cleaned = jsonStr.replace(/```json?\s*/g, '').replace(/```/g, '').trim();
      slide.placeholderInfo = JSON.parse(cleaned);
      console.log(`    Slide ${slide.id} (${slide.platform}): ✓`);
    } catch (err) {
      console.log(`    Slide ${slide.id}: ⚠ ${err.message?.substring(0, 50)}`);
      slide.placeholderInfo = {
        what_to_capture: `Open ${slide.platform} and capture the relevant screen showing data for ${agentName}.`,
        suggested_callout: `This ${slide.platform.replace(/-/g, ' ')} integration powers the agent with real-time data.`,
        suggested_position: 'top-right',
        setup_required: `Ensure ${slide.connectorName || slide.platform} has recent data.`,
      };
    }
  }
}

// ───────────────────────────────────────────
// STEP 4C: Generate PLACEHOLDER-GUIDE.md
// ───────────────────────────────────────────

function generatePlaceholderGuide(slides, agentName, demoDir, slug, connectionEvents = []) {
  const placeholders = slides.filter(s => s.placeholder);
  if (placeholders.length === 0 && connectionEvents.length === 0) return null;

  const lines = [
    `# Placeholder Guide — ${agentName}`,
    `Generated: ${new Date().toISOString()}`,
    '',
    `${placeholders.length} slide(s) need manual screenshots. Instructions for each:`,
    '',
  ];

  for (const slide of placeholders) {
    const info = slide.placeholderInfo || {};
    lines.push(`## Slide ${slide.id} — ${slide.platform}`);
    lines.push('');
    lines.push(`**What to capture:**`);
    lines.push(info.what_to_capture || `Screenshot of ${slide.platform} showing agent-related data.`);
    lines.push('');
    lines.push(`**Setup required:**`);
    lines.push(info.setup_required || 'Ensure the connection is active.');
    lines.push('');
    lines.push(`**How to complete this slide:**`);
    lines.push(`1. Take the screenshot and save it to: \`screenshots/${slide.id}-${slide.platform}-final.png\``);
    lines.push(`2. In \`demo.yaml\`, find slide ${slide.id} and change \`status: placeholder\` to \`status: ready\``);
    lines.push(`3. Set \`screenshot: screenshots/${slide.id}-${slide.platform}-final.png\``);
    lines.push(`4. Run: \`agentdemo generate --config demos/${slug}/demo.yaml\``);
    lines.push('');
    lines.push('---');
    lines.push('');
  }

  // Connection approvals section
  if (connectionEvents.length > 0) {
    const connPlatforms = [...new Set(connectionEvents.map(c => c.platform))];
    lines.push('## Connection Approvals');
    lines.push('');
    lines.push('If this is the first time running a demo for this agent,');
    lines.push('the agent may ask to approve connections to:');
    for (const p of connPlatforms) {
      lines.push(`- **${p}**`);
    }
    lines.push('');
    lines.push('To pre-approve these BEFORE running capture:');
    lines.push('1. Open M365 Copilot and navigate to the agent');
    lines.push('2. Type any prompt that uses each connection');
    lines.push('3. When the connection dialog appears, approve it');
    lines.push('4. Confirm the agent responds successfully');
    lines.push('5. You only need to do this once per demo account');
    lines.push('');
    lines.push('If you skip this step, AgentDemo will pause during capture');
    lines.push('and prompt you to approve connections at that point instead.');
    lines.push('');
  }

  const guidePath = path.join(demoDir, 'PLACEHOLDER-GUIDE.md');
  fs.writeFileSync(guidePath, lines.join('\n'));
  return guidePath;
}

// ───────────────────────────────────────────
// STEP 5: Generate Storylane-style demo.html
// ───────────────────────────────────────────

function generateDemoHTML(slides, agentName, description, m365Url, brandColor = '#6B2FD9') {
  // Build slides data for the HTML
  const slidesData = [];

  // Intro slide
  slidesData.push({
    type: 'intro',
    agentName,
    description: description || `Interactive demo for ${agentName}`,
    brandColor,
  });

  // Content slides
  for (const slide of slides) {
    // Use pre-computed base64 if available (connection-setup slides)
    let screenshotBase64 = slide.screenshotBase64 || '';
    if (!screenshotBase64 && slide.screenshot && fs.existsSync(slide.screenshot)) {
      const buf = fs.readFileSync(slide.screenshot);
      if (buf.length >= 50000) {
        screenshotBase64 = buf.toString('base64');
      }
    }

    const isPlaceholder = slide.placeholder || (!screenshotBase64 && slide.type !== 'connection-setup');
    const info = slide.placeholderInfo || {};

    slidesData.push({
      type: slide.type,
      platform: slide.platform,
      prompt: slide.prompt || '',
      screenshot: screenshotBase64,
      placeholder: isPlaceholder,
      storyLabel: slide.storyLabel || '',
      callout: isPlaceholder
        ? { text: info.suggested_callout || '', position: info.suggested_position || 'top-right', point_to: { x: 50, y: 40 } }
        : (slide.callout || { text: '', position: 'bottom-right', point_to: { x: 50, y: 50 } }),
      placeholderCapture: info.what_to_capture || '',
    });
  }

  // Outro slide
  slidesData.push({
    type: 'outro',
    agentName,
    m365Url: 'https://www.mysmb.com/',
    brandColor,
  });

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${agentName} — Interactive Demo</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0a0a0a; color: #fff; overflow: hidden; height: 100vh; }

/* Top bar */
.top-bar { position: fixed; top: 0; left: 0; right: 0; height: 48px; background: #1a1a1a; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; z-index: 100; border-bottom: 1px solid #333; }
.top-bar .agent-name { font-size: 14px; font-weight: 600; color: #ccc; }
.top-bar .slide-counter { font-size: 13px; color: #888; }
.progress-bar { position: fixed; top: 48px; left: 0; right: 0; height: 3px; background: #333; z-index: 100; }
.progress-fill { height: 100%; background: ${brandColor}; transition: width 0.35s ease; }

/* Browser frame */
.browser-frame { position: fixed; top: 51px; left: 40px; right: 40px; bottom: 60px; background: #1e1e1e; border-radius: 12px; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.5); }
.browser-toolbar { height: 40px; background: #2d2d2d; display: flex; align-items: center; padding: 0 16px; gap: 8px; }
.browser-dots { display: flex; gap: 6px; }
.browser-dots span { width: 12px; height: 12px; border-radius: 50%; }
.browser-dots span:nth-child(1) { background: #ff5f57; }
.browser-dots span:nth-child(2) { background: #febc2e; }
.browser-dots span:nth-child(3) { background: #28c840; }
.browser-url { flex: 1; margin-left: 12px; height: 26px; background: #1e1e1e; border-radius: 6px; padding: 0 12px; font-size: 12px; color: #888; display: flex; align-items: center; }
.browser-content { position: relative; height: calc(100% - 40px); overflow: hidden; background: #f5f5f5; }
.browser-content img { width: 100%; height: 100%; object-fit: contain; background: #f5f5f5; }

/* Callout bubble — Storylane style */
.callout { position: absolute; max-width: 320px; background: ${brandColor}; color: #fff; border-radius: 12px; padding: 20px; font-size: 14px; line-height: 1.5; box-shadow: 0 8px 24px rgba(0,0,0,0.3); z-index: 50; opacity: 0; transform: translateY(10px); animation: calloutIn 0.35s ease forwards; animation-delay: 0.3s; }
.callout .callout-text { margin-bottom: 16px; }
.callout .callout-btn { display: inline-block; padding: 8px 20px; background: rgba(255,255,255,0.2); color: #fff; border: 1px solid rgba(255,255,255,0.3); border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
.callout .callout-btn:hover { background: rgba(255,255,255,0.35); }
.callout-pointer { position: absolute; width: 12px; height: 12px; background: ${brandColor}; border-radius: 50%; }
@keyframes calloutIn { to { opacity: 1; transform: translateY(0); } }

/* Callout positions */
.callout.top-left { top: 60px; left: 40px; }
.callout.top-right { top: 60px; right: 40px; }
.callout.bottom-left { bottom: 40px; left: 40px; }
.callout.bottom-right { bottom: 40px; right: 40px; }
.callout.center { top: 50%; left: 50%; transform: translate(-50%, -50%); }
.callout.center { animation: calloutInCenter 0.35s ease forwards; animation-delay: 0.3s; }
@keyframes calloutInCenter { to { opacity: 1; transform: translate(-50%, -50%); } }

/* Prompt badge */
.prompt-badge { position: absolute; top: 12px; left: 50%; transform: translateX(-50%); background: rgba(0,0,0,0.7); color: #fff; padding: 8px 20px; border-radius: 20px; font-size: 13px; max-width: 80%; text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; z-index: 40; backdrop-filter: blur(10px); }

/* Intro / Outro slides */
.slide-intro, .slide-outro { position: fixed; top: 51px; left: 40px; right: 40px; bottom: 60px; display: flex; flex-direction: column; align-items: center; justify-content: center; background: ${brandColor}; border-radius: 12px; }
.slide-intro h1 { font-size: 42px; font-weight: 700; margin-bottom: 12px; }
.slide-intro p { font-size: 18px; opacity: 0.85; margin-bottom: 32px; max-width: 500px; text-align: center; }
.slide-intro .start-btn, .slide-outro .cta-btn { padding: 14px 40px; background: #fff; color: ${brandColor}; font-size: 16px; font-weight: 700; border: none; border-radius: 8px; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }
.slide-intro .start-btn:hover, .slide-outro .cta-btn:hover { transform: scale(1.05); box-shadow: 0 8px 20px rgba(0,0,0,0.2); }
.slide-outro h2 { font-size: 32px; margin-bottom: 12px; }
.slide-outro p { font-size: 16px; opacity: 0.85; margin-bottom: 28px; text-align: center; max-width: 450px; }

/* Bottom bar */
.bottom-bar { position: fixed; bottom: 0; left: 0; right: 0; height: 60px; display: flex; align-items: center; justify-content: center; gap: 8px; background: #0a0a0a; z-index: 100; }
.dot { width: 8px; height: 8px; border-radius: 50%; background: #444; cursor: pointer; transition: all 0.2s; }
.dot.active { background: ${brandColor}; width: 24px; border-radius: 4px; }

/* Placeholder slide */
.placeholder-content { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; background: #fafafa; border: 3px dashed #ddd; border-radius: 8px; margin: 20px; }
.placeholder-icon { font-size: 48px; margin-bottom: 12px; opacity: 0.5; }
.placeholder-platform { font-size: 18px; font-weight: 600; color: #666; margin-bottom: 8px; }
.placeholder-label { font-size: 14px; color: #999; }
.incomplete-badge { position: absolute; top: 12px; right: 12px; background: #f59e0b; color: #fff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 12px; z-index: 60; text-transform: uppercase; letter-spacing: 0.5px; }
.callout.placeholder-callout { border: 2px dashed rgba(255,255,255,0.5); background: rgba(107,47,217,0.75); }
.callout .callout-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; opacity: 0.8; }
.callout .callout-instruction { font-size: 12px; opacity: 0.7; margin-top: 12px; padding-top: 12px; border-top: 1px solid rgba(255,255,255,0.2); line-height: 1.4; }

/* Hidden slides */
.slide { display: none; }
.slide.active { display: block; }
</style>
</head>
<body>

<div class="top-bar">
  <span class="agent-name" id="agentName"></span>
  <span class="slide-counter" id="slideCounter"></span>
</div>
<div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>

<div id="slideContainer"></div>

<div class="bottom-bar" id="dots"></div>

<script>
const SLIDES = ${JSON.stringify(slidesData)};
let current = 0;

function render() {
  const container = document.getElementById('slideContainer');
  const s = SLIDES[current];
  container.innerHTML = '';

  document.getElementById('agentName').textContent = SLIDES[0].agentName || '';
  document.getElementById('slideCounter').textContent = (current + 1) + ' of ' + SLIDES.length;
  document.getElementById('progressFill').style.width = ((current + 1) / SLIDES.length * 100) + '%';

  // Dots
  const dots = document.getElementById('dots');
  dots.innerHTML = '';
  for (let i = 0; i < SLIDES.length; i++) {
    const d = document.createElement('div');
    d.className = 'dot' + (i === current ? ' active' : '');
    d.onclick = () => { current = i; render(); };
    dots.appendChild(d);
  }

  if (s.type === 'intro') {
    container.innerHTML = '<div class="slide-intro active">' +
      '<h1>' + esc(s.agentName) + '</h1>' +
      '<p>' + esc(s.description) + '</p>' +
      '<button class="start-btn" onclick="next()">Start Demo &rarr;</button>' +
      '</div>';
    return;
  }

  if (s.type === 'outro') {
    container.innerHTML = '<div class="slide-outro active">' +
      '<h2>That\\u2019s the demo!</h2>' +
      '<p>You just saw ' + esc(s.agentName) + ' in action across Microsoft 365.</p>' +
      '<a href="' + esc(s.m365Url) + '" target="_blank" class="cta-btn">Try it yourself &rarr;</a>' +
      '</div>';
    return;
  }

  // Connection-setup slide — slightly different visual treatment
  if (s.type === 'connection-setup' && s.storyLabel) {
    var connHtml = '<div class="browser-frame" style="border: 2px solid #f59e0b;"><div class="browser-toolbar">' +
      '<div class="browser-dots"><span></span><span></span><span></span></div>' +
      '<div class="browser-url" style="color:#f59e0b;">\\u2699\\ufe0f ' + esc(s.storyLabel) + '</div></div>' +
      '<div class="browser-content">';
    if (s.screenshot) {
      connHtml += '<img src="data:image/png;base64,' + s.screenshot + '" alt="Connection setup" style="opacity:0.9">';
    }
    connHtml += '<div style="position:absolute;top:12px;left:12px;background:#f59e0b;color:#fff;font-size:11px;font-weight:700;padding:4px 10px;border-radius:12px;text-transform:uppercase;letter-spacing:0.5px;">One-time setup</div>';
    if (s.callout && s.callout.text) {
      var cpos = s.callout.position || 'bottom-right';
      connHtml += '<div class="callout ' + cpos + '" style="background:#f59e0b;">' +
        '<div class="callout-text">' + esc(s.callout.text) + '</div>' +
        '<button class="callout-btn" onclick="next()">Next &rarr;</button></div>';
    }
    connHtml += '</div></div>';
    container.innerHTML = connHtml;
    return;
  }

  // Content slide
  let html = '<div class="browser-frame"><div class="browser-toolbar">' +
    '<div class="browser-dots"><span></span><span></span><span></span></div>' +
    '<div class="browser-url">' + esc(s.platform) + '</div></div>' +
    '<div class="browser-content">';

  if (s.placeholder) {
    // Placeholder slide — dashed box with platform icon
    var platformIcons = {'sharepoint':'\\ud83d\\udcc1','power-automate':'\\u26a1','teams':'\\ud83d\\udcac','outlook':'\\ud83d\\udce7','xero':'\\ud83d\\udcb0','custom':'\\ud83d\\udd17'};
    var icon = platformIcons[s.platform] || '\\ud83d\\udcf7';
    html += '<div class="placeholder-content">' +
      '<div class="placeholder-icon">' + icon + '</div>' +
      '<div class="placeholder-platform">' + esc(s.platform) + '</div>' +
      '<div class="placeholder-label">Screenshot to be added</div></div>';
    html += '<div class="incomplete-badge">Incomplete</div>';
  } else if (s.screenshot) {
    html += '<img src="data:image/png;base64,' + s.screenshot + '" alt="' + esc(s.platform) + '">';
  } else {
    html += '<div class="placeholder-content">' +
      '<div class="placeholder-icon">\\ud83d\\udcf7</div>' +
      '<div class="placeholder-label">Screenshot not yet captured</div></div>';
    html += '<div class="incomplete-badge">Incomplete</div>';
  }

  // Prompt badge
  if (s.prompt) {
    html += '<div class="prompt-badge">&ldquo;' + esc(s.prompt) + '&rdquo;</div>';
  }

  // Callout bubble
  if (s.callout && s.callout.text) {
    var pos = s.callout.position || 'bottom-right';
    if (s.placeholder) {
      html += '<div class="callout placeholder-callout ' + pos + '">' +
        '<div class="callout-label">\\ud83d\\udca1 Suggested callout</div>' +
        '<div class="callout-text">' + esc(s.callout.text) + '</div>' +
        (s.placeholderCapture ? '<div class="callout-instruction">To complete this slide: ' + esc(s.placeholderCapture) + '</div>' : '') +
        '<button class="callout-btn" onclick="next()">Next &rarr;</button></div>';
    } else {
      html += '<div class="callout ' + pos + '">' +
        '<div class="callout-text">' + esc(s.callout.text) + '</div>' +
        '<button class="callout-btn" onclick="next()">Next &rarr;</button></div>';
    }
  }

  html += '</div></div>';
  container.innerHTML = html;
}

function next() { if (current < SLIDES.length - 1) { current++; render(); } }
function prev() { if (current > 0) { current--; render(); } }
function esc(s) { const d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }

document.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowRight' || e.key === ' ') next();
  if (e.key === 'ArrowLeft') prev();
});

// Click outside callout advances
document.addEventListener('click', (e) => {
  if (!e.target.closest('.callout') && !e.target.closest('.dot') && !e.target.closest('.start-btn') && !e.target.closest('.cta-btn')) {
    next();
  }
});

// Override browser back button
window.addEventListener('popstate', (e) => { e.preventDefault(); prev(); });
history.pushState(null, '', '');

render();
</script>
</body>
</html>`;

  return html;
}

// ───────────────────────────────────────────
// Main orchestrator
// ───────────────────────────────────────────

export async function runCreate(opts) {
  const headless = opts.headless || false;
  if (headless) {
    console.log('  Running in headless mode. Note: if a connection approval is');
    console.log('  required, the script will pause and ask you to re-run without --headless.\n');
  }

  // STEP 1: Get URLs
  const { studioUrl, m365Url } = await askForUrls(opts);

  // Check auth
  const { browser, context } = await createBrowserContext({ headless });
  const valid = await isSessionValid(context);
  if (!valid) {
    console.log('  ✗ Session expired. Run: agentdemo auth');
    await browser.close();
    return;
  }

  let activePage = await context.newPage();
  let activeBrowser = browser;

  let discovered, slides;
  try {
    // STEP 2: Auto-discover from Copilot Studio
    discovered = await discoverFromStudio(activePage, studioUrl);

    // If discovery returned nothing useful, ask user for manual input
    // Close browser first so readline works on MINGW64
    if (discovered.name === 'My Agent' || (discovered.topics.length === 0 && discovered.connections.length === 0)) {
      await activePage.close().catch(() => {});
      await activeBrowser.close().catch(() => {});
      await askManualFallback(discovered);
      // Re-open browser for capture
      const reopened = await createBrowserContext({ headless });
      const validAgain = await isSessionValid(reopened.context);
      if (!validAgain) {
        console.log('  ✗ Session expired. Run: agentdemo auth');
        await reopened.browser.close();
        return;
      }
      activeBrowser = reopened.browser;
      activePage = await reopened.context.newPage();
    }

    // STEP 2C: Generate demo prompts via AI (before capture)
    const generatedPrompts = await generateDemoPrompts(discovered);
    // Update intro hook if AI provided one
    if (generatedPrompts?.hook) {
      discovered.hook = generatedPrompts.hook;
    }

    // Set up demo directory
    const slug = makeSlug(discovered.name);
    const demoDir = path.join(DEMOS_DIR, slug);
    const outputDir = path.join(demoDir, 'output');
    fs.mkdirSync(outputDir, { recursive: true });

    // STEP 3: Auto-capture from M365 Copilot
    const captureResult = await autoCapture(activePage, m365Url, discovered, demoDir, generatedPrompts);
    slides = captureResult.slides;
    const connectionEvents = captureResult.connectionEvents || [];

    // Close browser before API calls
    await activePage.close().catch(() => {});
    await activeBrowser.close().catch(() => {});
    console.log('');

    // Insert connection setup slide if connections were encountered
    if (connectionEvents.length > 0) {
      const connPlatforms = [...new Set(connectionEvents.map(c => c.platform))];
      const connScreenshot = connectionEvents[0].screenshotPath;
      let connBase64 = '';
      if (connScreenshot && fs.existsSync(connScreenshot)) {
        connBase64 = fs.readFileSync(connScreenshot).toString('base64');
      }
      const connText = connPlatforms.length === 1
        ? `Before the agent can access ${connPlatforms[0]}, users approve a one-time connection. After approval, the agent works seamlessly every time.`
        : `The agent connects to ${connPlatforms.join(' and ')}. Each connection is approved once and remembered — no repeated sign-ins required.`;

      // Insert as the first content slide (after intro)
      slides.unshift({
        id: 0,
        platform: 'm365-copilot',
        type: 'connection-setup',
        screenshot: connScreenshot,
        screenshotBase64: connBase64,
        clip: null,
        prompt: null,
        callout: {
          text: connText,
          position: 'bottom-right',
          point_to: { x: 50, y: 50 },
        },
        storyLabel: 'One-time setup: connect your accounts',
      });
      // Re-number slide IDs
      slides.forEach((s, idx) => { s.id = idx + 1; });
      console.log(`  ✓ Added connection setup slide (${connPlatforms.join(', ')})`);
    }

    // STEP 4: Generate callout text for captured slides
    await generateCallouts(slides, discovered.name, discovered.description);

    // STEP 4B: Generate placeholder instructions
    await generatePlaceholderInfo(slides, discovered.name, discovered.description);

    // STEP 5: Generate HTML
    console.log('  ● Generating demo.html...');
    const introDesc = discovered.hook || discovered.description || `Interactive demo for ${discovered.name}`;
    const html = generateDemoHTML(slides, discovered.name, introDesc, m365Url);
    const htmlPath = path.join(outputDir, 'demo.html');
    fs.writeFileSync(htmlPath, html);
    console.log(`    ✓ Written: ${htmlPath}`);

    // STEP 5B: Generate PLACEHOLDER-GUIDE.md if needed
    const demoSlug = makeSlug(discovered.name);
    const guidePath = generatePlaceholderGuide(slides, discovered.name, demoDir, demoSlug, connectionEvents);

    // STEP 6: Summary
    const capturedSlides = slides.filter(s => s.screenshot && !s.placeholder);
    const placeholderSlides = slides.filter(s => s.placeholder);
    console.log('');
    console.log(`  ✓ Demo created: ${htmlPath}`);
    console.log('');

    console.log('  ────────────────────────────────');
    if (capturedSlides.length > 0) {
      console.log(`  ✓ ${capturedSlides.length} slides captured`);
    }

    // Check for partial/needs-review slides
    const partialSlides = slides.filter(s => s.partial);
    const reviewSlides = slides.filter(s => s.needsReview);
    if (partialSlides.length > 0) {
      console.log(`  ~ ${partialSlides.length} partial (agent response timed out)`);
    }
    if (reviewSlides.length > 0) {
      console.log(`  ⚠ ${reviewSlides.length} needs review (agent returned error)`);
    }

    if (placeholderSlides.length > 0) {
      console.log(`  ⚠ ${placeholderSlides.length} placeholder (manual screenshot needed)`);
    }
    console.log('  ────────────────────────────────');

    if (placeholderSlides.length > 0) {
      console.log('');
      console.log('  ┌─────────────────────────────────────────────────────┐');
      for (const ps of placeholderSlides) {
        const info = ps.placeholderInfo || {};
        console.log(`  │  Slide ${ps.id} — ${ps.platform.padEnd(44)}│`);
        const capture = (info.what_to_capture || '').substring(0, 50);
        if (capture) {
          console.log(`  │  Capture: ${capture.padEnd(42)}│`);
        }
        const setup = (info.setup_required || '').substring(0, 50);
        if (setup) {
          console.log(`  │  Setup: ${setup.padEnd(44)}│`);
        }
        console.log(`  │${''.padEnd(53)}│`);
      }
      console.log('  └─────────────────────────────────────────────────────┘');
      console.log('');
      if (guidePath) {
        console.log(`  Instructions saved to: ${guidePath}`);
      }
      console.log(`\n  When screenshots are ready:`);
      console.log(`  agentdemo generate --config demos/${demoSlug}/demo.yaml`);
    }

    console.log('');

    // Open in default browser
    try {
      const open = (await import('open')).default;
      await open(htmlPath);
      console.log('  ✓ Opened in browser\n');
    } catch {
      console.log(`  Open manually: ${htmlPath}\n`);
    }

  } catch (err) {
    console.log(`\n  ✗ Error: ${err.message}`);
    await activePage.close().catch(() => {});
    await activeBrowser.close().catch(() => {});
  }
}
