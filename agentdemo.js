#!/usr/bin/env node

import { Command } from 'commander';
import dotenv from 'dotenv';
import { createRequire } from 'module';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.join(__dirname, '.env') });

const program = new Command();

program
  .name('agentdemo')
  .description('Automated Storylane-style interactive demo recorder for Microsoft Copilot Studio agents')
  .version('1.0.0');

// ── create (the main command — does everything) ──
program
  .command('create')
  .description('Create a complete demo — paste two URLs and done')
  .option('--studio-url <url>', 'Direct link to agent in Copilot Studio')
  .option('--m365-url <url>', 'Direct link to agent in M365 Copilot')
  .option('--headless', 'Run browser invisibly in the background')
  .option('--platforms <list>', 'Comma-separated platforms, e.g. "sharepoint,power-automate"')
  .option('--sharepoint-url <url>', 'SharePoint list or site URL to capture')
  .option('--power-automate-url <url>', 'Power Automate cloud flows page URL to capture')
  .option('--teams-url <url>', 'Microsoft Teams channel URL to capture')
  .option('--outlook-url <url>', 'Outlook email or folder URL to capture')
  .option('--xero-url <url>', 'Xero page URL to capture')
  .option('--custom-url <urls...>', 'Custom platform URL(s) to capture (can be repeated)')
  .action(async (opts) => {
    const { runCreate } = await import('./src/create.js');
    await runCreate(opts);
  });

// ── auth ──
program
  .command('auth')
  .description('Manage demo account session')
  .option('--status', 'Print current session status')
  .option('--verify-copilot', 'Verify M365 Copilot is accessible')
  .action(async (opts) => {
    const { runAuth } = await import('./src/auth.js');
    await runAuth(opts);
  });

// ── init ──
program
  .command('init')
  .description('Guided interview to create demo.yaml')
  .option('--name <name>', 'Agent name')
  .option('--discover', 'Auto-discover agent from Copilot Studio + M365 Copilot links')
  .option('--from-copilot-studio', '(deprecated alias for --discover)')
  .option('--studio-url <url>', 'Direct link to agent in Copilot Studio')
  .option('--m365-url <url>', 'Direct link to agent in M365 Copilot')
  .action(async (opts) => {
    // Backward compat: treat --from-copilot-studio same as --discover
    if (opts.fromCopilotStudio) opts.discover = true;
    // If either URL is provided, auto-enable discover mode
    if (opts.studioUrl || opts.m365Url) opts.discover = true;
    const { runInit } = await import('./src/init.js');
    await runInit(opts);
  });

// ── check ──
program
  .command('check')
  .description('Pre-flight verification')
  .option('--config <path>', 'Path to demo.yaml')
  .option('--ffmpeg-only', 'Only check ffmpeg availability')
  .action(async (opts) => {
    if (!opts.config && !opts.ffmpegOnly) {
      console.error('error: --config <path> is required unless using --ffmpeg-only');
      process.exit(1);
    }
    const { runCheck } = await import('./src/capture.js');
    await runCheck(opts);
  });

// ── capture ──
program
  .command('capture')
  .description('Run capture mode')
  .requiredOption('--config <path>', 'Path to demo.yaml')
  .option('--slide <id>', 'Capture a specific slide only', parseInt)
  .action(async (opts) => {
    const { runCapture } = await import('./src/capture.js');
    await runCapture(opts);
  });

// ── generate ──
program
  .command('generate')
  .description('Run generate mode')
  .requiredOption('--config <path>', 'Path to demo.yaml')
  .action(async (opts) => {
    const { runGenerate } = await import('./src/generate.js');
    await runGenerate(opts);
  });

// ── run ──
program
  .command('run')
  .description('Capture then generate')
  .requiredOption('--config <path>', 'Path to demo.yaml')
  .action(async (opts) => {
    const { runCapture, runCheck } = await import('./src/capture.js');
    const { runGenerate } = await import('./src/generate.js');
    const checkResult = await runCheck(opts);
    if (!checkResult.ready) {
      process.exit(1);
    }
    await runCapture(opts);
    await runGenerate(opts);
  });

// ── resume ──
program
  .command('resume')
  .description('Resume from last failed/partial run')
  .requiredOption('--config <path>', 'Path to demo.yaml')
  .action(async (opts) => {
    const { runResume } = await import('./src/capture.js');
    await runResume(opts);
  });

// ── guide ──
program
  .command('guide')
  .description('Generate PREFLIGHT-GUIDE.md')
  .requiredOption('--config <path>', 'Path to demo.yaml')
  .action(async (opts) => {
    const { runGuide } = await import('./src/init.js');
    await runGuide(opts);
  });

// ── list ──
program
  .command('list')
  .description('List all demos and capture status')
  .action(async () => {
    const { runList } = await import('./src/init.js');
    await runList();
  });

// ── new ──
program
  .command('new')
  .description('Scaffold blank demo folder + YAML')
  .requiredOption('--name <name>', 'Demo name')
  .action(async (opts) => {
    const { runNew } = await import('./src/init.js');
    await runNew(opts);
  });

program.parse();
