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

    // Priority 2: Copilot Studio "What's new" / release notes panel
    // This panel appears on every Studio visit and blocks tab navigation if not dismissed.
    try {
      const studioWhatsNew = await page.evaluate(() => {
        // Look for a panel/dialog whose heading contains "What's new" or "New features"
        const headings = document.querySelectorAll('h1, h2, h3, [role="heading"], [class*="title" i]');
        for (const h of headings) {
          const t = h.textContent?.trim().toLowerCase() || '';
          if (t.includes("what's new") || t.includes("whats new") || t.includes("new features") || t.includes("release notes")) {
            // Find nearest dismiss button in the same panel
            const panel = h.closest('[role="dialog"], [role="complementary"], [class*="panel" i], [class*="modal" i], [class*="callout" i]') || h.parentElement?.parentElement;
            if (panel) {
              const btn = panel.querySelector('button[aria-label="Close"], button[aria-label="Dismiss"], button[aria-label="close"], button[class*="close" i]');
              if (btn) { btn.click(); return true; }
            }
          }
        }
        return false;
      });
      if (studioWhatsNew) {
        await page.waitForTimeout(1500);
        continue;
      }
    } catch { /* non-fatal */ }

    // Priority 3: Welcome / onboarding / generic popups
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

  const DEBUG_SCREENSHOT = path.join(ROOT, '.debug-studio-discovery.png');

  /** Navigate to Studio URL with retry */
  async function loadStudioPage(attempt = 1) {
    await maximizeWindow(page);
    // Skip networkidle — Studio keeps persistent WebSocket connections that
    // prevent it from ever reaching idle. Just wait for DOM content.
    await page.goto(studioUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(3000);

    // Bail early if auth redirected us to login — Studio may need its own
    // session cookie warmup (run: agentdemo auth)
    const landedUrl = page.url();
    if (landedUrl.includes('login.microsoftonline.com') || landedUrl.includes('login.live.com')) {
      throw new Error(
        'Copilot Studio redirected to login. The saved session may not cover the Studio domain. ' +
        'Run: agentdemo auth'
      );
    }

    // Studio sometimes restores the last-visited agent from session state instead
    // of navigating to the requested bot ID. Detect this and force a hard reload.
    const botIdMatch = studioUrl.match(/bots\/([a-f0-9-]+)\//i);
    if (botIdMatch && !landedUrl.includes(botIdMatch[1])) {
      console.log(`    ⚠ Studio loaded a different agent (session restore). Forcing reload...`);
      await page.goto(studioUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForTimeout(3000);
      const reloadedUrl = page.url();
      if (!reloadedUrl.includes(botIdMatch[1])) {
        console.log(`    ⚠ Still on wrong agent after reload — Studio may redirect to home on first visit.`);
      }
    }

    // Studio's React SPA can take 90-120 s to hydrate in a fresh browser.
    // Use Studio-specific selectors and a generous timeout.
    try {
      await page.waitForSelector(
        '[role="tablist"], [role="tab"], [data-testid*="bot"], [class*="botName" i], ' +
        '[class*="agentName" i], nav[aria-label*="agent" i], h1, h2',
        { timeout: 120000 }
      );
    } catch (selErr) {
      await page.screenshot({ path: DEBUG_SCREENSHOT, fullPage: false }).catch(() => {});
      if (attempt === 1) {
        console.log('    ⚠ Copilot Studio content not found yet. Retrying once...');
        return loadStudioPage(2);
      }
      throw new Error(
        `Copilot Studio page did not render after 2 attempts (${selErr.message}). ` +
        `Debug screenshot saved to: ${DEBUG_SCREENSHOT}`
      );
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

    // TAB_NAMES is a Set in Node but page.evaluate serialises it as a plain array.
    // Use Array.prototype.includes inside evaluate (not Set.prototype.has).
    discovered.name = await page.evaluate((tabNames) => {
      const isTabName = (s) => tabNames.includes(s.toLowerCase());

      // Strategy 1: Browser tab title
      // Format: "Overview - Licence Renewal Agent"
      // or: "Licence Renewal Agent - Microsoft Copilot Studio"
      const pageTitle = document.title || '';
      const titleParts = pageTitle.split(/\s*[-–|]\s*/);
      for (const part of titleParts) {
        const cleaned = part.trim();
        if (cleaned && cleaned.length > 2 && cleaned.length < 80 && !isTabName(cleaned)) {
          return cleaned;
        }
      }
      // Strategy 2: Try from the end of title parts (agent name is often second)
      if (titleParts.length >= 2) {
        const last = titleParts[titleParts.length - 1].trim();
        const secondLast = titleParts[titleParts.length - 2].trim();
        if (secondLast && !isTabName(secondLast) && secondLast.length > 2) return secondLast;
        if (last && !isTabName(last) && last.length > 2) return last;
      }

      // Strategy 3: Look for the agent name heading near the agent icon
      for (const sel of [
        'h1', 'h2',
        '[class*="agentName" i]', '[class*="botName" i]',
        '[class*="agent-name" i]', '[class*="bot-name" i]',
        '[data-testid*="agent-name"]', '[data-testid*="bot-name"]',
      ]) {
        const el = document.querySelector(sel);
        if (el) {
          const t = el.textContent?.trim();
          if (t && t.length > 2 && t.length < 80 && !isTabName(t)) return t;
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

    // Read agent instructions from Overview tab
    // Instructions live in a textarea or large text region labelled "Instructions"
    const instrText = await page.evaluate(() => {
      // Try labelled textarea first
      for (const label of document.querySelectorAll('label, [class*="label" i], legend')) {
        const text = label.textContent?.trim()?.toLowerCase() || '';
        if (text.includes('instruction')) {
          // Look for associated input/textarea via for= or sibling
          const forId = label.getAttribute('for');
          if (forId) {
            const el = document.getElementById(forId);
            if (el) return (el.value || el.textContent || '').trim().substring(0, 2000);
          }
          // Try next sibling or parent's textarea
          const parent = label.closest('[class*="field" i], [class*="section" i], div') || label.parentElement;
          if (parent) {
            const ta = parent.querySelector('textarea, [contenteditable="true"], [class*="editor" i]');
            if (ta) return (ta.value || ta.textContent || '').trim().substring(0, 2000);
          }
        }
      }
      // Fallback: any large textarea on the page (likely the instructions field)
      const allTa = document.querySelectorAll('textarea, [contenteditable="true"]');
      for (const ta of allTa) {
        const val = (ta.value || ta.textContent || '').trim();
        if (val.length > 50) return val.substring(0, 2000);
      }
      return '';
    });
    if (instrText) {
      discovered.instructions = instrText;
      console.log(`  ✓ Agent instructions read (${instrText.length} chars)`);
    }

    // Navigate to Topics tab
    await dismissPopups(page); // dismiss any popup that appeared after page load
    const topicsClicked = await clickFirst(page, [
      'button:has-text("Topics")', 'a:has-text("Topics")',
      '[role="tab"]:has-text("Topics")', '[aria-label*="Topics"]',
    ], 5000);

    if (topicsClicked) {
      await page.waitForTimeout(3000);

      // Read topics and trigger phrases
      const topics = await page.evaluate(() => {
        const found = [];
        const rows = document.querySelectorAll('[data-testid*="topic"], [role="row"], tr');
        for (const row of rows) {
          const cells = row.querySelectorAll('[role="gridcell"], td, div');
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
    await dismissPopups(page); // dismiss any popup that reappeared after Topics navigation
    const actionsClicked = await clickFirst(page, [
      'button:has-text("Actions")', 'a:has-text("Actions")',
      '[role="tab"]:has-text("Actions")', 'button:has-text("Connections")',
    ], 5000);

    if (actionsClicked) {
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
    // Save screenshot of whatever the browser is showing so it's easy to diagnose
    await page.screenshot({ path: DEBUG_SCREENSHOT, fullPage: false }).catch(() => {});
    console.log(`    ⚠ Could not read agent details from Copilot Studio.`);
    console.log(`      Reason: ${err.message}`);
    console.log(`      Debug screenshot: ${DEBUG_SCREENSHOT}`);
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
// STEP 2C: Generate demo script via Anthropic API
// ───────────────────────────────────────────

async function generateDemoScript(discovered) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null; // caller will use welcome screen prompts instead

  console.log('  ● Generating demo script using AI...');

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
  const topicsSummary = discovered.topics.length > 0
    ? discovered.topics.map(t => {
        const phrases = t.phrases.length > 0 ? ` (e.g. "${t.phrases[0]}")` : '';
        return `  - ${t.name}${phrases}`;
      }).join('\n')
    : '  (none discovered)';

  try {
    const response = await client.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      messages: [{
        role: 'user',
        content: `You are helping create a product demo for an AI agent. Your job is to design a realistic, flowing CONVERSATION that tells a compelling business story in a single chat thread.

Agent name: ${discovered.name}
Agent description: ${discovered.description || 'Not provided'}
Agent instructions: ${discovered.instructions || 'Not provided'}
Copilot Studio topics (agent capabilities):
${topicsSummary}
Connected platforms: ${platforms}

Design a conversation of 3–6 steps. The steps must flow logically as a real user session — each step builds on the previous response. The conversation should demonstrate a complete end-to-end workflow from discovery through to action.

Rules:
- Use [ENTITY] as a placeholder in a prompt when the actual value (e.g. a record name, licence number, item name) will only be known after reading the previous agent response
- Set extract_entity to true when the response will contain a specific named item that the NEXT step needs to reference
- entity_context describes what kind of thing to extract (e.g. "licence name", "invoice number", "project name")
- Set capture_platform_after to the platform slug (e.g. "sharepoint", "power-automate") when this step causes a visible change in that platform worth screenshotting — otherwise null
- slide_label is a short business-level headline (≤8 words) shown in the demo viewer
- Make prompts sound natural, like a real business user would type them

Respond in JSON only, no markdown:
{
  "hook": "one sentence that sells this demo to a business decision maker",
  "steps": [
    {
      "prompt": "the exact text to type in chat",
      "slide_label": "short business headline",
      "extract_entity": false,
      "entity_context": null,
      "capture_platform_after": null
    }
  ]
}`,
      }],
    });

    const jsonStr = response.content[0]?.text || '';
    const cleaned = jsonStr.replace(/```json?\s*/g, '').replace(/```/g, '').trim();
    const result = JSON.parse(cleaned);

    if (result.steps && result.steps.length > 0) {
      console.log('');
      console.log('  ✓ Generated demo script:');
      result.steps.forEach((s, i) => {
        console.log(`    ${i + 1}. "${s.prompt.substring(0, 70)}${s.prompt.length > 70 ? '...' : ''}"`);
        console.log(`       → ${s.slide_label}`);
        if (s.extract_entity) console.log(`       → extracts: ${s.entity_context}`);
        if (s.capture_platform_after) console.log(`       → platform capture after: ${s.capture_platform_after}`);
      });
      console.log('');
      return result;
    }
  } catch (err) {
    console.log(`    ⚠ Script generation failed: ${err.message}`);
  }

  return null;
}

// ───────────────────────────────────────────
// Entity extraction helper
// ───────────────────────────────────────────

async function extractEntityFromResponse(responseText, entityContext) {
  if (!responseText || !entityContext) return null;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (apiKey) {
    try {
      const mod = await import('@anthropic-ai/sdk');
      const Anthropic = mod.default || mod.Anthropic;
      const client = new Anthropic({ apiKey });
      const resp = await client.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 60,
        messages: [{
          role: 'user',
          content: `Extract the ${entityContext} from this text. Reply with ONLY the extracted value — no explanation, no punctuation, no quotes.\n\nText:\n${responseText.substring(0, 1200)}`,
        }],
      });
      const extracted = (resp.content[0]?.text || '').trim();
      if (extracted && extracted.length > 0 && extracted.length < 120) {
        console.log(`    ✓ Extracted ${entityContext}: "${extracted}"`);
        return extracted;
      }
    } catch (err) {
      console.log(`    ⚠ Entity extraction failed: ${err.message?.substring(0, 50)}`);
    }
  }

  // Regex fallback
  // 1. Bold markdown: **Some Value**
  const boldMatch = responseText.match(/\*\*([^*\n]{2,80})\*\*/);
  if (boldMatch) {
    const val = boldMatch[1].trim();
    console.log(`    ✓ Extracted (bold) ${entityContext}: "${val}"`);
    return val;
  }
  // 2. First list item: "- Something" or "1. Something" or "• Something"
  const listMatch = responseText.match(/^[\-\*•][ \t]+(.+)|^\d+\.\s+(.+)/m);
  if (listMatch) {
    const val = (listMatch[1] || listMatch[2]).trim();
    if (val.length > 1 && val.length < 100) {
      console.log(`    ✓ Extracted (list item) ${entityContext}: "${val}"`);
      return val;
    }
  }
  // 3. Capitalized multi-word phrase
  const capMatch = responseText.match(/\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\b/);
  if (capMatch) {
    console.log(`    ✓ Extracted (phrase) ${entityContext}: "${capMatch[1]}"`);
    return capMatch[1];
  }

  console.log(`    ⚠ Could not extract ${entityContext} from response`);
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

async function autoCapture(page, m365Url, discovered, demoDir, generatedScript = null) {
  const screenshotsDir = path.join(demoDir, 'screenshots');
  const clipsDir = path.join(demoDir, 'clips');
  fs.mkdirSync(screenshotsDir, { recursive: true });
  fs.mkdirSync(clipsDir, { recursive: true });

  const slides = [];
  const connectionEvents = []; // { platform, screenshotPath } — for connection setup slide
  let slideId = 1;

  // ── Platform display names ──
  const PLATFORM_DISPLAY = {
    'sharepoint': 'SharePoint',
    'power-automate': 'Power Automate',
    'teams': 'Microsoft Teams',
    'outlook': 'Outlook',
    'xero': 'Xero',
    'custom': 'Custom Platform',
  };

  // ── Platform capture helper — opens a new tab, screenshots, closes it ──
  // Keeps the main M365 chat page intact throughout the entire capture.
  async function capturePlatformInNewTab(conn, ssPath) {
    const platform = conn.platform;
    const tab = await page.context().newPage();
    try {
      await maximizeWindow(tab);
      await tab.goto(conn.url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await tab.waitForLoadState('networkidle').catch(() => {});

      if (tab.url().includes('login.microsoftonline.com') || tab.url().includes('login.live.com')) {
        throw new Error('Auth expired — redirected to login');
      }

      if (platform === 'sharepoint') {
        await tab.waitForTimeout(3000);
        const spSelectors = ['[role="grid"]', '.ms-List', '.ms-DetailsList', 'table', '[data-automationid="ListCell"]'];
        for (const sel of spSelectors) {
          try { await tab.waitForSelector(sel, { timeout: 10000 }); break; } catch { /* next */ }
        }
        await tab.waitForTimeout(2000);
        await tab.evaluate(() => window.scrollBy(0, 200));
        await tab.waitForTimeout(1000);

      } else if (platform === 'power-automate') {
        await tab.waitForTimeout(3000);
        const paSelectors = ['[role="grid"]', '.ms-DetailsRow', '[data-automationid]', 'table', '.ms-List', '.ms-DetailsList'];
        for (const sel of paSelectors) {
          try { await tab.waitForSelector(sel, { timeout: 10000 }); break; } catch { /* next */ }
        }
        await tab.waitForTimeout(2000);
        for (const sel of ['[data-testid="connection-string"]', '[class*="credential"]', '[class*="secret"]', 'input[type="password"]']) {
          const els = await tab.$$(sel);
          for (const el of els) { await el.evaluate(n => { n.style.visibility = 'hidden'; }); }
        }

      } else if (platform === 'teams') {
        await tab.waitForTimeout(5000);
        const tSelectors = ['[data-tid="messageBodyContent"]', '[role="main"]', '.message-body', '.ts-message-list-container'];
        for (const sel of tSelectors) {
          try { await tab.waitForSelector(sel, { timeout: 15000 }); break; } catch { /* next */ }
        }
        await tab.waitForTimeout(2000);

      } else if (platform === 'outlook') {
        await tab.waitForTimeout(5000);
        const oSelectors = ['[role="listbox"]', '[data-testid="MailList"]', '.customScrollBar', '[aria-label*="Message list"]', '[role="main"]'];
        for (const sel of oSelectors) {
          try { await tab.waitForSelector(sel, { timeout: 15000 }); break; } catch { /* next */ }
        }
        await tab.waitForTimeout(2000);

      } else {
        await tab.waitForTimeout(3000);
      }

      await tab.screenshot({ path: ssPath, fullPage: false });
      console.log('    ✓ Page loaded');
      console.log('    ✓ Screenshot saved');
      return true;
    } catch (err) {
      console.log(`    ✗ Platform tab capture failed: ${err.message}`);
      return false;
    } finally {
      await tab.close().catch(() => {});
    }
  }

  // Platform connections are captured AFTER the welcome screen (below),
  // using new tabs so the M365 chat page is never navigated away.
  const platformConns = discovered.connections.filter(c => c.platform !== 'm365-copilot');


  // ── Capture M365 Copilot agent interaction slides ──
  console.log('  ● Capturing M365 Copilot...\n');
  console.log('  ● Opening M365 Copilot...');
  try {
    await maximizeWindow(page);
    await page.goto(m365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(8000);
    await dismissPopups(page);

    // ── SLIDE 1: Welcome screen (before any prompt) ──
    const welcomeSsPath = path.join(screenshotsDir, `${slideId}-m365-welcome.png`);
    await page.screenshot({ path: welcomeSsPath, fullPage: false });
    console.log(`  ✓ Slide ${slideId} — Welcome screen captured`);
    slides.push({
      id: slideId++,
      platform: 'm365-copilot',
      type: 'agent',
      screenshot: welcomeSsPath,
      clip: null,
      prompt: null,
      callout: null,
      storyLabel: 'Meet your AI agent',
    });

    // ── Platform slides — captured in new tabs so chat page stays open ──
    if (platformConns.length > 0) {
      console.log('\n  ● Capturing supporting platforms...\n');
    }
    for (const conn of platformConns) {
      const platform = conn.platform;
      const displayName = PLATFORM_DISPLAY[platform] || platform;
      console.log(`  ● Slide ${slideId} — ${displayName}`);
      if (conn.url) {
        console.log(`    Navigating to ${displayName} in new tab...`);
        const ssFilename = `${slideId}-${platform}-initial.png`;
        const ssPath = path.join(screenshotsDir, ssFilename);
        const captured = await capturePlatformInNewTab(conn, ssPath);
        slides.push({
          id: slideId++, platform, type: 'platform',
          placeholder: !captured, connectorName: conn.name,
          screenshot: captured ? ssPath : null,
          clip: null, prompt: null, callout: null, placeholderInfo: null,
        });
      } else {
        slides.push({
          id: slideId++, platform, type: 'platform',
          placeholder: true, connectorName: conn.name,
          screenshot: null, clip: null, prompt: null, callout: null, placeholderInfo: null,
        });
        console.log(`    → Placeholder (needs manual screenshot)`);
      }
      console.log('');
    }

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
      // ── Build steps — three strategies in priority order ──
      let steps = [];

      // Strategy 1: AI-generated conversation script
      if (generatedScript && generatedScript.steps && generatedScript.steps.length > 0) {
        steps = generatedScript.steps.slice(0, 6);
        console.log(`  Using ${steps.length} AI-generated script steps`);
      }

      // Strategy 2: Copilot Studio topic phrases
      if (steps.length === 0 && discovered.topics.length > 0) {
        for (const topic of discovered.topics) {
          const text = topic.phrases.length > 0 ? topic.phrases[0] : (topic.name.length > 5 ? topic.name : null);
          if (text) steps.push({ prompt: text, slide_label: topic.name || '', extract_entity: false, entity_context: null, capture_platform_after: null });
        }
        if (steps.length > 0) {
          steps = steps.slice(0, 5);
          console.log(`  Using ${steps.length} steps from Copilot Studio topics`);
        }
      }

      // Strategy 3: Starter prompts from agent welcome screen
      if (steps.length === 0) {
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
              if (t && t.length > 5 && t.length < 150 && !found.includes(t)) found.push(t);
            }
            if (found.length > 0) break;
          }
          if (found.length === 0) {
            const buttons = document.querySelectorAll('button, [role="button"]');
            for (const b of buttons) {
              const t = b.textContent?.trim();
              if (t && t.length > 10 && t.length < 120
                  && !t.toLowerCase().includes('send') && !t.toLowerCase().includes('attach')
                  && !t.toLowerCase().includes('mic')) found.push(t);
            }
          }
          return found.slice(0, 5);
        }).catch(() => []);

        if (starterPrompts.length > 0) {
          steps = starterPrompts.map(t => ({ prompt: t, slide_label: '', extract_entity: false, entity_context: null, capture_platform_after: null }));
          console.log(`  ✓ Found ${steps.length} starter prompts from welcome screen`);
        } else {
          steps = [{ prompt: 'Tell me about your capabilities and what you can help with', slide_label: '', extract_entity: false, entity_context: null, capture_platform_after: null }];
          console.log('  ⚠ No starter prompts found. Using generic prompt.');
        }
      }

      const RESPONSE_TIMEOUT = 120;
      const ERROR_PHRASES = ['something went wrong', "i'm having trouble", "couldn't complete", 'connection failed'];

      // Carries the last extracted entity forward across steps
      let extractedEntity = null;

      for (let i = 0; i < steps.length; i++) {
        const step = steps[i];

        // Substitute [ENTITY] with the extracted value from the previous step
        let promptText = step.prompt;
        if (extractedEntity && promptText.includes('[ENTITY]')) {
          promptText = promptText.replace(/\[ENTITY\]/g, extractedEntity);
        }

        console.log(`\n  ● Slide ${slideId} — M365 Copilot`);
        if (step.slide_label) console.log(`    "${step.slide_label}"`);
        console.log(`    Typing prompt: "${promptText.substring(0, 60)}${promptText.length > 60 ? '...' : ''}"`);

        try {
          // ── All prompts stay in the SAME chat thread ──
          // M365 Copilot uses a Lexical rich text editor — fill() won't work.
          const inputLocator = page.locator(INPUT_SEL).last();
          await inputLocator.waitFor({ state: 'visible', timeout: 15000 });
          await inputLocator.click();
          await page.waitForTimeout(300);
          await page.keyboard.press('Control+A');
          await page.keyboard.press('Backspace');
          await page.waitForTimeout(200);
          await page.keyboard.type(promptText, { delay: 50 });
          await page.waitForTimeout(500);
          await page.keyboard.press('Enter');

          let isPartial = false;
          let connectionDetected = false;

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

          // Connection approval resets the thread — re-navigate and retry
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
              await page.keyboard.type(promptText, { delay: 50 });
              await page.waitForTimeout(500);
              await page.keyboard.press('Enter');
              result = await waitForAgentResponse(page, RESPONSE_TIMEOUT);
            } catch (retryErr) {
              console.log(`    ⚠ Retry failed: ${retryErr.message?.substring(0, 50)}`);
              isPartial = true;
            }
          }

          if (!result.complete) isPartial = true;

          // Extract entity from this response if needed for the next step
          if (step.extract_entity && step.entity_context && result.lastMsgText) {
            extractedEntity = await extractEntityFromResponse(result.lastMsgText, step.entity_context);
          }

          const hasError = ERROR_PHRASES.some(phrase => (result.lastMsgText || '').toLowerCase().includes(phrase));
          if (hasError) console.log(`    ⚠ Agent returned an error message — saved as needs-review`);

          const ssPath = path.join(screenshotsDir, `${slideId}-m365-prompt-${i + 1}.png`);
          await scrollAndScreenshot(page, ssPath);
          console.log(`    ✓ Screenshot saved`);

          slides.push({
            id: slideId++,
            platform: 'm365-copilot',
            type: 'agent',
            screenshot: ssPath,
            clip: null,
            prompt: promptText,
            callout: null,
            storyLabel: step.slide_label || '',
            partial: isPartial,
            needsReview: hasError,
          });

          // ── Mid-conversation platform capture ──
          // Opened in a new tab so the chat thread is never interrupted.
          if (step.capture_platform_after) {
            const capPlatform = step.capture_platform_after;
            const capConn = discovered.connections.find(c => c.platform === capPlatform && c.url);
            if (capConn) {
              const displayName = PLATFORM_DISPLAY[capPlatform] || capPlatform;
              console.log(`\n  ● Slide ${slideId} — ${displayName} (post-action capture)`);
              const platSsPath = path.join(screenshotsDir, `${slideId}-${capPlatform}-action.png`);
              const captured = await capturePlatformInNewTab(capConn, platSsPath);
              slides.push({
                id: slideId++, platform: capPlatform, type: 'platform',
                placeholder: !captured, connectorName: capConn.name,
                screenshot: captured ? platSsPath : null,
                clip: null, prompt: null, callout: null,
                storyLabel: `Result in ${displayName}`, placeholderInfo: null,
              });
            } else {
              console.log(`    ⚠ capture_platform_after="${capPlatform}" — no URL found, skipping`);
            }
          }

          await page.waitForTimeout(2000);
        } catch (err) {
          process.stdout.write('\r' + ' '.repeat(60) + '\r');
          console.log(`    ⚠ Prompt failed: ${err.message}`);
          const ssPath = path.join(screenshotsDir, `${slideId}-m365-error-${i + 1}.png`);
          await page.screenshot({ path: ssPath }).catch(() => {});
          slides.push({
            id: slideId++, platform: 'm365-copilot', type: 'agent',
            screenshot: ssPath, clip: null, prompt: promptText, callout: null,
            storyLabel: step.slide_label || '',
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
        model: 'claude-sonnet-4-6',
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
        model: 'claude-sonnet-4-6',
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

function generateDemoHTML(slides, agentName, description, m365Url, brandColor = '#00C9A7') {
  // Platform display names and accent colors for browser chrome
  const PLAT_NAMES = {
    'm365-copilot': 'Microsoft 365 Copilot',
    'sharepoint': 'SharePoint',
    'power-automate': 'Power Automate',
    'teams': 'Microsoft Teams',
    'outlook': 'Outlook',
    'xero': 'Xero',
    'custom': 'Custom Platform',
  };
  const PLAT_COLORS = {
    'm365-copilot': '#0078D4',
    'sharepoint': '#036C70',
    'power-automate': '#0066FF',
    'teams': '#5558AF',
    'outlook': '#0072C6',
    'xero': '#13B5EA',
    'custom': brandColor,
  };

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
      platformName: PLAT_NAMES[slide.platform] || slide.platform,
      platformColor: PLAT_COLORS[slide.platform] || brandColor,
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
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #F5F5F0; color: #fff; overflow: hidden; height: 100vh; }

/* Top bar */
.top-bar { position: fixed; top: 0; left: 0; right: 0; height: 48px; background: #1A1A1A; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; z-index: 100; border-bottom: 1px solid #333; }
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
.callout .callout-btn { display: inline-block; padding: 8px 20px; background: #1A1A1A; color: #fff; border: none; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
.callout .callout-btn:hover { background: #333; }
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
.slide-intro .start-btn, .slide-outro .cta-btn { padding: 14px 40px; background: #1A1A1A; color: #fff; font-size: 16px; font-weight: 700; border: none; border-radius: 8px; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }
.slide-intro .start-btn:hover, .slide-outro .cta-btn:hover { transform: scale(1.05); box-shadow: 0 8px 20px rgba(0,0,0,0.2); }
.slide-outro h2 { font-size: 32px; margin-bottom: 12px; }
.slide-outro p { font-size: 16px; opacity: 0.85; margin-bottom: 28px; text-align: center; max-width: 450px; }

/* Bottom bar */
.bottom-bar { position: fixed; bottom: 0; left: 0; right: 0; height: 60px; display: flex; align-items: center; justify-content: center; gap: 8px; background: #1A1A1A; z-index: 100; }
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

  // Content slide — platform-colored top border + display name in URL bar
  var platColor = s.platformColor || '${brandColor}';
  var platName = s.platformName || s.platform;
  let html = '<div class="browser-frame" style="border-top: 3px solid ' + platColor + ';">' +
    '<div class="browser-toolbar">' +
    '<div class="browser-dots"><span></span><span></span><span></span></div>' +
    '<div class="browser-url"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + platColor + ';margin-right:8px;"></span>' + esc(platName) + '</div></div>' +
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

    // Always apply explicitly-provided agent name/instructions (takes priority over auto-discovery)
    if (opts.agentName) discovered.name = opts.agentName;
    if (opts.instructions) discovered.instructions = opts.instructions;

    // If discovery returned nothing useful, use opts overrides or ask user for manual input
    // Close browser first so readline works on MINGW64
    if (discovered.name === 'My Agent' || (discovered.topics.length === 0 && discovered.connections.length === 0)) {
      // MCP-supplied overrides already applied above
      if (opts.platforms) {
        const platList = opts.platforms.split(',').map(p => p.trim().toLowerCase()).filter(Boolean);
        for (const p of platList) {
          if (!discovered.connections.some(c => c.platform === p)) {
            discovered.connections.push({ name: p, platform: p });
          }
        }
      }

      // In MCP mode with opts provided, skip interactive fallback
      const hasOverrides = opts.agentName || opts.instructions || opts.platforms;
      if (!opts.mcpMode || !hasOverrides) {
        await activePage.close().catch(() => {});
        await activeBrowser.close().catch(() => {});
        await askManualFallback(discovered);
        // Re-open browser for capture
      } else {
        console.log(`  ✓ Using provided agent details: ${discovered.name}`);
        await activePage.close().catch(() => {});
        await activeBrowser.close().catch(() => {});
      }

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

    // STEP 2B: Attach platform URLs from CLI flags / MCP params
    const PLATFORM_URL_MAP = {
      'sharepoint':     opts.sharepointUrl,
      'power-automate': opts.powerAutomateUrl,
      'teams':          opts.teamsUrl,
      'outlook':        opts.outlookUrl,
      'xero':           opts.xeroUrl,
    };

    // Ensure --platforms flag adds connections (even outside the discovery-fallback block)
    if (opts.platforms) {
      const platList = opts.platforms.split(',').map(p => p.trim().toLowerCase()).filter(Boolean);
      for (const p of platList) {
        if (!discovered.connections.some(c => c.platform === p)) {
          discovered.connections.push({ name: p, platform: p });
        }
      }
    }

    // Attach URLs to matching connections, or create the connection entry
    for (const [platform, url] of Object.entries(PLATFORM_URL_MAP)) {
      if (!url) continue;
      const existing = discovered.connections.find(c => c.platform === platform);
      if (existing) {
        existing.url = url;
      } else {
        discovered.connections.push({ name: platform, platform, url });
      }
    }

    // Handle --custom-url (variadic — can be string or array)
    if (opts.customUrl) {
      const customUrls = Array.isArray(opts.customUrl) ? opts.customUrl : [opts.customUrl];
      for (let i = 0; i < customUrls.length; i++) {
        discovered.connections.push({ name: `custom-${i + 1}`, platform: 'custom', url: customUrls[i] });
      }
    }

    // STEP 2C: Generate demo script via AI (before capture)
    const generatedScript = await generateDemoScript(discovered);
    // Update intro hook if AI provided one
    if (generatedScript?.hook) {
      discovered.hook = generatedScript.hook;
    }

    // Set up demo directory
    const slug = makeSlug(discovered.name);
    const demoDir = path.join(DEMOS_DIR, slug);
    const outputDir = path.join(demoDir, 'output');
    fs.mkdirSync(outputDir, { recursive: true });

    // STEP 3: Auto-capture from M365 Copilot
    const captureResult = await autoCapture(activePage, m365Url, discovered, demoDir, generatedScript);
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
    const platformCaptured = capturedSlides.filter(s => s.type === 'platform');
    const m365Captured = capturedSlides.filter(s => s.platform === 'm365-copilot');
    console.log('');
    console.log(`  ✓ Demo created: ${htmlPath}`);
    console.log('');

    console.log('  ────────────────────────────────');
    if (platformCaptured.length > 0) {
      console.log(`  ✓ ${platformCaptured.length} platform slides captured`);
    }
    if (m365Captured.length > 0) {
      console.log(`  ✓ ${m365Captured.length} M365 Copilot slides captured`);
    }
    console.log(`  ✓ ${slides.length} total slides`);

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
