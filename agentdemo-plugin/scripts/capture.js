import path from 'path';
import fs from 'fs';
import { parseConfig, getDemoDir } from './config-parser.js';
import { createBrowserContext, isSessionValid, performLogin, saveSession } from './auth.js';
import { captureM365Copilot } from './capture-steps/m365-copilot.js';
import { captureSharePoint } from './capture-steps/sharepoint.js';
import { capturePowerAutomate } from './capture-steps/power-automate.js';
import { captureTeams } from './capture-steps/teams.js';
import { captureOutlook } from './capture-steps/outlook.js';
import { captureGeneric } from './capture-steps/generic.js';
import readline from 'readline';
import { execSync } from 'child_process';

const CAPTURE_HANDLERS = {
  'm365-copilot': captureM365Copilot,
  'sharepoint': captureSharePoint,
  'power-automate': capturePowerAutomate,
  'teams': captureTeams,
  'outlook': captureOutlook,
  'xero': captureGeneric,
  'custom': captureGeneric,
};

function waitForKeypress(msg) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(msg, () => { rl.close(); resolve(); });
  });
}

function getSessionPath(demoDir) {
  return path.join(demoDir, '.session.json');
}

function loadSession(demoDir) {
  const p = getSessionPath(demoDir);
  if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, 'utf8'));
  return null;
}

function saveSessionState(demoDir, session) {
  session.last_updated = new Date().toISOString();
  const completedSlides = [];
  const failedSlides = [];
  const pendingSlides = [];
  for (const [id, info] of Object.entries(session.slides)) {
    if (info.status === 'done') completedSlides.push(Number(id));
    else if (info.status === 'failed' || info.status === 'needs-review') failedSlides.push(Number(id));
    else if (info.status === 'pending' || info.status === 'skipped') pendingSlides.push(Number(id));
    else if (info.status === 'partial') failedSlides.push(Number(id));
  }
  session.completed_slides = completedSlides.sort((a, b) => a - b);
  session.failed_slides = failedSlides.sort((a, b) => a - b);
  session.pending_slides = pendingSlides.sort((a, b) => a - b);

  if (pendingSlides.length === 0 && failedSlides.length === 0) session.status = 'completed';
  else if (completedSlides.length > 0) session.status = 'partial';

  fs.writeFileSync(getSessionPath(demoDir), JSON.stringify(session, null, 2));
}

function appendLog(demoDir, message) {
  const logPath = path.join(demoDir, 'capture.log');
  fs.appendFileSync(logPath, message + '\n');
}

function initSession(config, demoDir) {
  const session = {
    demo: path.basename(demoDir),
    started_at: new Date().toISOString(),
    last_updated: new Date().toISOString(),
    status: 'in-progress',
    slides: {},
    completed_slides: [],
    failed_slides: [],
    pending_slides: [],
  };
  for (const slide of config.slides) {
    session.slides[slide.id] = { status: 'pending', screenshot: null, clip: null };
  }
  saveSessionState(demoDir, session);
  return session;
}

function getDiskSpaceMB() {
  try {
    if (process.platform === 'win32') {
      const out = execSync('wmic logicaldisk get freespace,caption', { encoding: 'utf8' });
      const lines = out.trim().split('\n').slice(1);
      for (const line of lines) {
        const parts = line.trim().split(/\s+/);
        if (parts.length >= 2) return Math.floor(Number(parts[1]) / 1024 / 1024);
      }
    } else {
      const out = execSync("df -m . | tail -1 | awk '{print $4}'", { encoding: 'utf8' });
      return parseInt(out.trim(), 10);
    }
  } catch { /* ignore */ }
  return 999999;
}

async function getFfmpegPath() {
  if (process.env.FFMPEG_PATH) return process.env.FFMPEG_PATH;
  try {
    const mod = await import('ffmpeg-static');
    return mod.default || mod;
  } catch { /* ignore */ }
  try {
    execSync('ffmpeg -version', { stdio: 'ignore' });
    return 'ffmpeg';
  } catch { /* ignore */ }
  return null;
}

function checkFfmpegSync() {
  if (process.env.FFMPEG_PATH && fs.existsSync(process.env.FFMPEG_PATH)) {
    return process.env.FFMPEG_PATH;
  }
  try {
    const modPath = path.join(process.cwd(), 'node_modules', 'ffmpeg-static', 'index.js');
    if (fs.existsSync(modPath)) {
      // Read the actual binary path from the module
      const content = fs.readFileSync(modPath, 'utf8');
      const binDir = path.join(process.cwd(), 'node_modules', 'ffmpeg-static');
      const candidates = fs.readdirSync(binDir).filter(f => f.startsWith('ffmpeg'));
      if (candidates.length > 0) return path.join(binDir, candidates[0]);
      return 'ffmpeg-static (bundled)';
    }
  } catch { /* ignore */ }
  try {
    execSync('ffmpeg -version', { stdio: 'ignore' });
    return 'ffmpeg (system)';
  } catch { /* ignore */ }
  return null;
}

export async function runCheck(opts) {
  console.log('');

  // ffmpeg-only shortcut
  if (opts.ffmpegOnly) {
    const ff = checkFfmpegSync();
    if (ff) {
      console.log(`✓ ffmpeg available at ${ff}`);
    } else {
      console.log('✗ ffmpeg not found — install ffmpeg or run npm install');
    }
    return { ready: !!ff };
  }

  const configPath = opts.config;
  const demoDir = getDemoDir(configPath);
  const result = { ready: true, warnings: [], errors: [] };

  let config;
  try {
    config = parseConfig(configPath);
    console.log(`AgentDemo Pre-Flight Check — ${config.title}`);
    console.log('─'.repeat(40));
  } catch (err) {
    console.log(`✗ Config error: ${err.message}`);
    return { ready: false };
  }

  // .env
  const envPath = path.join(path.dirname(path.dirname(demoDir)), '.env');
  if (fs.existsSync(path.join(demoDir, '..', '..', '.env')) || process.env.DEMO_EMAIL) {
    console.log('✓ .env configured');
  } else {
    console.log('✗ .env file not found');
    result.errors.push('.env missing');
    result.ready = false;
  }

  // Auth session
  const stateFile = path.join(path.dirname(demoDir), '..', '.browser-session', 'state.json');
  const rootStateFile = path.join(demoDir, '..', '..', '.browser-session', 'state.json');
  const sessionExists = fs.existsSync(stateFile) || fs.existsSync(rootStateFile);
  if (sessionExists) {
    console.log(`✓ Auth session file exists`);
  } else {
    console.log('✗ No auth session — run agentdemo auth first');
    result.errors.push('No auth session');
    result.ready = false;
  }

  // ffmpeg
  const ff = checkFfmpegSync();
  if (ff) {
    console.log(`✓ ffmpeg available`);
  } else {
    console.log('✗ ffmpeg not found');
    result.warnings.push('ffmpeg missing — clips will not be converted to mp4');
  }

  // Disk space
  const diskMB = getDiskSpaceMB();
  if (diskMB > 2000) {
    console.log(`✓ Disk space OK (${Math.round(diskMB / 1024)}GB free)`);
  } else if (diskMB > 1000) {
    console.log(`⚠ Disk space low (${Math.round(diskMB / 1024)}GB free) — may run into issues with clips`);
    result.warnings.push('Low disk space');
  } else {
    console.log(`✗ Disk space critically low (${diskMB}MB free)`);
    result.errors.push('Insufficient disk space');
    result.ready = false;
  }

  // Folders
  for (const sub of ['screenshots', 'clips', 'output']) {
    const dir = path.join(demoDir, sub);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  }
  console.log('✓ Output directories exist');

  // Slide URLs (just print them, no network check in basic mode)
  for (const slide of config.slides) {
    if (slide.url) {
      console.log(`✓ Slide ${slide.id} URL configured: ${slide.url.substring(0, 60)}...`);
    }
  }

  console.log('─'.repeat(40));
  if (result.ready) {
    console.log('Result: READY — all critical checks passed.');
  } else {
    console.log(`Result: NOT READY — ${result.errors.length} issue(s) must be resolved before capture.`);
  }
  if (result.warnings.length > 0) {
    console.log(`Warnings: ${result.warnings.join(', ')}`);
  }
  console.log('');
  return result;
}

export async function runCapture(opts) {
  const configPath = opts.config;
  const config = parseConfig(configPath);
  const demoDir = getDemoDir(configPath);

  // Ensure output directories
  for (const sub of ['screenshots', 'clips', 'output']) {
    const dir = path.join(demoDir, sub);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  const session = initSession(config, demoDir);
  appendLog(demoDir, `\n[${new Date().toISOString()}] Capture started for ${config.title}`);

  // Determine which slides to capture
  let slidesToCapture = config.slides;
  if (opts.slide != null) {
    slidesToCapture = config.slides.filter(s => s.id === opts.slide);
    if (slidesToCapture.length === 0) {
      console.log(`Slide ${opts.slide} not found in config.`);
      return;
    }
  }

  // Set up signal handler for graceful shutdown
  let interrupted = false;
  const onSignal = () => {
    interrupted = true;
    console.log('\n\nInterrupted — saving progress...');
  };
  process.on('SIGINT', onSignal);
  process.on('SIGTERM', onSignal);

  // Launch browser
  const { browser, context } = await createBrowserContext();

  // Check session validity
  const valid = await isSessionValid(context);
  if (!valid) {
    console.log('Session expired — re-authenticating...');
    await performLogin(context);
  }

  // Capture slides
  for (const slide of slidesToCapture) {
    if (interrupted) break;

    console.log(`\nCapturing slide ${slide.id} (${slide.platform})...`);
    session.slides[slide.id].status = 'in-progress';
    saveSessionState(demoDir, session);

    const handler = CAPTURE_HANDLERS[slide.platform] || captureGeneric;
    const captureCtx = {
      context,
      config,
      slide,
      demoDir,
      screenshotsDir: path.join(demoDir, 'screenshots'),
      clipsDir: path.join(demoDir, 'clips'),
    };

    let retries = 0;
    const maxRetries = 1;
    let success = false;

    while (retries <= maxRetries && !interrupted) {
      try {
        // Check disk space before clip recording
        if (slide.record_clip) {
          const diskMB = getDiskSpaceMB();
          if (diskMB < 200) {
            console.log(`⚠ Low disk space (${diskMB}MB) — skipping clip for slide ${slide.id}`);
            slide.record_clip = false;
          }
        }

        const result = await handler(captureCtx);

        session.slides[slide.id] = {
          status: result.status || 'done',
          screenshot: result.screenshot || null,
          clip: result.clip || null,
        };
        success = true;
        console.log(`✓ Slide ${slide.id} captured (${slide.platform})`);
        appendLog(demoDir,
          `[${new Date().toISOString()}] [SLIDE ${slide.id}] [SUCCESS] ${slide.platform}`
        );
        break;
      } catch (err) {
        const errorType = classifyError(err);
        appendLog(demoDir,
          `[${new Date().toISOString()}] [SLIDE ${slide.id}] [${errorType}]\n` +
          `  URL: ${slide.url}\n` +
          `  Error: ${err.message}\n` +
          `  Action taken: ${retries < maxRetries ? 'Retrying' : 'Marked as failed'}`
        );

        if (errorType === 'AUTH_EXPIRED') {
          console.log('Session expired mid-run — re-authenticating...');
          try {
            await performLogin(context);
            continue; // Don't count as retry
          } catch {
            session.slides[slide.id] = { status: 'failed', error: 'Auth re-login failed', retries };
            break;
          }
        }

        if (errorType === 'MFA_REQUIRED') {
          await waitForKeypress(
            'ACTION REQUIRED: Complete MFA in the browser window, then press Enter to continue...'
          );
          await saveSession(context);
          continue;
        }

        if (retries < maxRetries && isRetryable(errorType)) {
          retries++;
          console.log(`⚠ Slide ${slide.id} failed (${err.message}). Retrying in 5s... (${retries}/${maxRetries})`);
          await new Promise(r => setTimeout(r, 5000));
        } else {
          session.slides[slide.id] = {
            status: errorType === 'AGENT_RESPONSE_TIMEOUT' ? 'partial' : 'failed',
            error: err.message,
            retries,
          };
          console.log(`✗ Slide ${slide.id} failed: ${err.message}`);
          break;
        }
      }
    }

    saveSessionState(demoDir, session);
  }

  process.removeListener('SIGINT', onSignal);
  process.removeListener('SIGTERM', onSignal);

  await browser.close();

  // Update final status
  const hasFailures = session.failed_slides.length > 0;
  const hasPending = session.pending_slides.length > 0;
  if (!hasFailures && !hasPending) {
    session.status = 'completed';
  } else {
    session.status = interrupted ? 'partial' : 'failed';
  }
  saveSessionState(demoDir, session);

  // Generate resume guide if needed
  if (hasFailures || hasPending) {
    generateResumeGuide(config, session, demoDir);
  }

  // Print summary
  printSummary(session);
}

export async function runResume(opts) {
  const configPath = opts.config;
  const config = parseConfig(configPath);
  const demoDir = getDemoDir(configPath);
  const session = loadSession(demoDir);

  if (!session) {
    console.log('No previous session found. Run agentdemo capture first.');
    return;
  }

  const toResume = [];
  for (const slide of config.slides) {
    const s = session.slides[slide.id];
    if (s && (s.status === 'pending' || s.status === 'failed' || s.status === 'partial' || s.status === 'skipped')) {
      toResume.push(slide);
    }
  }

  if (toResume.length === 0) {
    console.log('All slides are captured. Run agentdemo generate to build the demo.');
    return;
  }

  console.log(`\nResuming ${config.title} demo capture`);
  console.log(`✓ Completed: slides ${session.completed_slides.join(', ') || 'none'}`);
  console.log(`✗ Failed:    ${session.failed_slides.map(id => `slide ${id} (${session.slides[id]?.error || 'unknown'})`).join(', ') || 'none'}`);
  console.log(`○ Pending:   slides ${session.pending_slides.join(', ') || 'none'}`);
  console.log(`\nWill attempt: slides ${toResume.map(s => s.id).join(', ')}`);

  await waitForKeypress('Press Enter to continue or Ctrl+C to cancel.');

  // Re-run capture for outstanding slides
  const originalSlides = config.slides;
  config.slides = toResume;

  // Use same capture flow
  await runCapture({ config: configPath, slide: undefined });
}

function classifyError(err) {
  const msg = err.message || '';
  if (msg.includes('login.microsoftonline.com')) return 'AUTH_EXPIRED';
  if (msg.includes('MFA') || msg.includes('Verify your identity')) return 'MFA_REQUIRED';
  if (msg.includes('404') || msg.includes('403') || msg.includes('redirect')) return 'NAVIGATION_FAILED';
  if (msg.includes('waitForSelector') || msg.includes('Timeout')) return 'ELEMENT_NOT_FOUND';
  if (msg.includes('agent') && msg.includes('timeout')) return 'AGENT_RESPONSE_TIMEOUT';
  if (msg.includes('Something went wrong') || msg.includes('having trouble')) return 'AGENT_ERROR_RESPONSE';
  return 'UNKNOWN';
}

function isRetryable(errorType) {
  return ['ELEMENT_NOT_FOUND', 'AGENT_RESPONSE_TIMEOUT', 'CLIP_EMPTY_OR_CORRUPT'].includes(errorType);
}

function generateResumeGuide(config, session, demoDir) {
  const lines = [
    `# Resume Guide — ${config.title}`,
    `Generated: ${new Date().toISOString()}`,
    `Run status: ${session.status} (${session.failed_slides.length} failed, ${session.completed_slides.length} done)`,
    '',
    '## Before you resume, check these things:',
    '',
  ];

  for (const id of [...session.failed_slides, ...session.pending_slides]) {
    const slideSession = session.slides[id];
    const slideConfig = config.slides.find(s => s.id === id);
    if (!slideConfig) continue;

    lines.push(`### Slide ${id} — ${slideConfig.platform} (${(slideSession.status || 'pending').toUpperCase()}: ${slideSession.error || 'not attempted'})`);
    lines.push(`**What happened:** ${slideSession.error || 'Slide was not reached during the run'}`);
    lines.push('');
    lines.push('**Check before retrying:**');
    lines.push(`- [ ] Verify URL is accessible: ${slideConfig.url}`);
    lines.push('- [ ] Confirm demo account has necessary permissions');
    lines.push(`- [ ] Once confirmed working, run: agentdemo resume --config ${path.basename(demoDir)}/demo.yaml`);
    lines.push('');

    if (slideSession.screenshot || slideSession.clip) {
      lines.push('**Assets saved from failed attempt:**');
      if (slideSession.screenshot) lines.push(`- Screenshot: ${slideSession.screenshot}`);
      if (slideSession.clip) lines.push(`- Clip: ${slideSession.clip}`);
      lines.push('');
    }
    lines.push('---');
    lines.push('');
  }

  lines.push('## Ready to resume?');
  lines.push(`Run: agentdemo resume --config ${path.relative(process.cwd(), path.join(demoDir, 'demo.yaml'))}`);
  lines.push('');
  lines.push('## Want to generate with current assets?');
  lines.push(`Run: agentdemo generate --config ${path.relative(process.cwd(), path.join(demoDir, 'demo.yaml'))}`);
  lines.push('(Slides with missing assets will show a placeholder card)');

  fs.writeFileSync(path.join(demoDir, 'RESUME-GUIDE.md'), lines.join('\n'));
  console.log(`\n✓ RESUME-GUIDE.md generated`);
}

function printSummary(session) {
  const done = session.completed_slides.length;
  const partial = Object.values(session.slides).filter(s => s.status === 'partial').length;
  const failed = session.failed_slides.length;

  console.log('');
  console.log('┌─────────────────────────────────────┐');
  console.log('│  AgentDemo Capture Complete          │');
  console.log('│                                      │');
  console.log(`│  ✓ Done     ${String(done).padEnd(2)} slides${' '.repeat(16)}│`);
  if (partial > 0)
    console.log(`│  ~ Partial  ${String(partial).padEnd(2)} slide(s)${' '.repeat(14)}│`);
  console.log(`│  ✗ Failed   ${String(failed).padEnd(2)} slides${' '.repeat(16)}│`);
  console.log('│                                      │');
  if (failed > 0 || partial > 0) {
    console.log('│  Run `agentdemo resume` to retry     │');
    console.log('│  failed slides, or                   │');
  }
  console.log('│  Run `agentdemo generate` to build   │');
  console.log('│  demo with current assets.            │');
  console.log('└─────────────────────────────────────┘');
  console.log('');
}
