/**
 * AgentDemo MCP Server
 * Exposes AgentDemo functionality as MCP tools for use in Claude Code.
 */

import { createRequire } from 'module';
import { fileURLToPath as _fileURLToPath } from 'url';
const require = createRequire(import.meta.url);
const dotenv = require('dotenv');
const path = require('path');

// Load .env from agentdemo root folder.
// Use fileURLToPath so that URL-encoded spaces (%20) in the path are decoded
// correctly on Windows — new URL(import.meta.url).pathname leaves them encoded.
const __mcpDir = path.dirname(_fileURLToPath(import.meta.url));
dotenv.config({ path: path.resolve(__mcpDir, '../.env') });

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { runCreate } from './create.js';
import { runGenerate } from './generate.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEMOS_DIR = path.join(__dirname, '..', 'demos');

// ── Tool definitions ──────────────────────────────────────────────────────────

const TOOLS = [
  {
    name: 'create_demo',
    description:
      'Create a new interactive demo by automating Copilot Studio discovery and M365 Copilot capture. ' +
      'Provide the Copilot Studio URL and the M365 Copilot URL. ' +
      'Optionally supply a slug (folder name) and whether to run headless.',
    inputSchema: {
      type: 'object',
      properties: {
        studio_url: {
          type: 'string',
          description: 'Copilot Studio agent URL (the "Go to Copilot Studio" link)',
        },
        m365_url: {
          type: 'string',
          description: 'M365 Copilot chat URL where the agent is published',
        },
        agent_name: {
          type: 'string',
          description: 'Agent name to use if Copilot Studio discovery fails',
        },
        instructions: {
          type: 'string',
          description: 'Agent instructions to use if Copilot Studio discovery fails',
        },
        platforms: {
          type: 'string',
          description: 'Comma-separated list of platforms the agent connects to (e.g. "sharepoint,power-automate")',
        },
        sharepoint_url: {
          type: 'string',
          description: 'URL of the SharePoint list or site to capture',
        },
        power_automate_url: {
          type: 'string',
          description: 'URL of the Power Automate cloud flows page to capture',
        },
        teams_url: {
          type: 'string',
          description: 'URL of the Teams channel to capture',
        },
        outlook_url: {
          type: 'string',
          description: 'URL of the Outlook email or folder to capture',
        },
        xero_url: {
          type: 'string',
          description: 'URL of the Xero page to capture',
        },
        custom_urls: {
          type: 'string',
          description: 'Comma-separated URLs for custom platform captures',
        },
        slug: {
          type: 'string',
          description: 'Optional folder name for the demo (auto-derived from agent name if omitted)',
        },
        headless: {
          type: 'boolean',
          description: 'Run browser in headless mode (default: false)',
          default: false,
        },
      },
      required: ['studio_url', 'm365_url'],
    },
  },
  {
    name: 'get_demo_status',
    description:
      'Get the current status of a demo — whether it has screenshots, a generated HTML file, and when it was last updated.',
    inputSchema: {
      type: 'object',
      properties: {
        slug: {
          type: 'string',
          description: 'Demo folder name (as shown by list_demos)',
        },
      },
      required: ['slug'],
    },
  },
  {
    name: 'resume_demo',
    description:
      'Resume a previously started demo capture session that was interrupted. ' +
      'Picks up from the last captured slide.',
    inputSchema: {
      type: 'object',
      properties: {
        slug: {
          type: 'string',
          description: 'Demo folder name to resume',
        },
        headless: {
          type: 'boolean',
          description: 'Run browser in headless mode (default: false)',
          default: false,
        },
      },
      required: ['slug'],
    },
  },
  {
    name: 'list_demos',
    description: 'List all demos that have been created, with their status and output paths.',
    inputSchema: {
      type: 'object',
      properties: {},
    },
  },
  {
    name: 'generate_demo_html',
    description:
      'Re-generate the demo HTML output file from existing captured screenshots without re-running the browser. ' +
      'Useful for refreshing the HTML after manual edits to screenshots or config.',
    inputSchema: {
      type: 'object',
      properties: {
        slug: {
          type: 'string',
          description: 'Demo folder name',
        },
      },
      required: ['slug'],
    },
  },
];

// ── Tool handlers ─────────────────────────────────────────────────────────────

async function handleCreateDemo(args) {
  const {
    studio_url, m365_url, agent_name, instructions, platforms,
    sharepoint_url, power_automate_url, teams_url, outlook_url, xero_url, custom_urls,
    slug, headless = false,
  } = args;

  if (!studio_url || !m365_url) {
    return { error: 'studio_url and m365_url are required' };
  }

  // Convert custom_urls CSV string into array for runCreate
  const customUrlArr = custom_urls ? custom_urls.split(',').map(u => u.trim()).filter(Boolean) : undefined;

  try {
    await runCreate({
      studioUrl: studio_url,
      m365Url: m365_url,
      agentName: agent_name || undefined,
      instructions: instructions || undefined,
      platforms: platforms || undefined,
      sharepointUrl: sharepoint_url || undefined,
      powerAutomateUrl: power_automate_url || undefined,
      teamsUrl: teams_url || undefined,
      outlookUrl: outlook_url || undefined,
      xeroUrl: xero_url || undefined,
      customUrl: customUrlArr,
      slug: slug || undefined,
      headless,
      mcpMode: true,
    });

    // Derive the expected slug the same way create.js does, then find that demo.
    // Fallback order: explicit slug arg → slugify(agent_name) → most-recently-modified demo.
    let expectedSlug = slug;
    if (!expectedSlug && agent_name) {
      expectedSlug = agent_name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
    }
    const demos = getDemoList();
    const created = expectedSlug
      ? demos.find(d => d.slug === expectedSlug)
      : demos.sort((a, b) => (b.last_updated || '').localeCompare(a.last_updated || ''))[0];

    return {
      success: true,
      message: `Demo created successfully.`,
      demo: created || null,
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
    };
  }
}

async function handleGetDemoStatus(args) {
  const { slug } = args;
  const demoDir = path.join(DEMOS_DIR, slug);

  if (!fs.existsSync(demoDir)) {
    return { error: `Demo "${slug}" not found. Run list_demos to see available demos.` };
  }

  return getDemoStatus(slug, demoDir);
}

async function handleResumeDemo(args) {
  const { slug, headless = false } = args;
  const demoDir = path.join(DEMOS_DIR, slug);

  if (!fs.existsSync(demoDir)) {
    return { error: `Demo "${slug}" not found.` };
  }

  // Find session meta for config path
  const metaPath = path.join(demoDir, '.session-meta.json');
  if (!fs.existsSync(metaPath)) {
    return { error: `No session metadata found for "${slug}". Cannot resume.` };
  }

  try {
    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    const configPath = meta.config_path;

    if (!configPath || !fs.existsSync(configPath)) {
      return { error: `Config path not found in session metadata.` };
    }

    const { runResume } = await import('./capture.js');
    await runResume({ config: configPath, headless });

    return {
      success: true,
      message: `Demo "${slug}" resumed successfully.`,
      status: getDemoStatus(slug, demoDir),
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

async function handleListDemos() {
  if (!fs.existsSync(DEMOS_DIR)) {
    return { demos: [], message: 'No demos directory found. Run create_demo to get started.' };
  }

  const demos = getDemoList();
  if (demos.length === 0) {
    return { demos: [], message: 'No demos found. Run create_demo to get started.' };
  }

  return { demos };
}

async function handleGenerateDemoHtml(args) {
  const { slug } = args;
  const demoDir = path.join(DEMOS_DIR, slug);

  if (!fs.existsSync(demoDir)) {
    return { error: `Demo "${slug}" not found.` };
  }

  // Look for a YAML config in the demo dir
  const configFiles = fs.readdirSync(demoDir).filter(f => f.endsWith('.yml') || f.endsWith('.yaml'));
  if (configFiles.length === 0) {
    // Try generate from session meta
    const metaPath = path.join(demoDir, '.session-meta.json');
    if (!fs.existsSync(metaPath)) {
      return { error: `No config file or session metadata found in "${slug}".` };
    }
    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    if (!meta.config_path || !fs.existsSync(meta.config_path)) {
      return { error: `Config path in session metadata does not exist.` };
    }
    try {
      await runGenerate({ config: meta.config_path });
      const outputPath = path.join(demoDir, 'output', 'demo.html');
      return {
        success: true,
        message: `HTML regenerated successfully.`,
        output_path: outputPath,
      };
    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  const configPath = path.join(demoDir, configFiles[0]);
  try {
    await runGenerate({ config: configPath });
    const outputPath = path.join(demoDir, 'output', 'demo.html');
    return {
      success: true,
      message: `HTML regenerated from ${configFiles[0]}.`,
      output_path: outputPath,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getDemoStatus(slug, demoDir) {
  const screenshotsDir = path.join(demoDir, 'screenshots');
  const clipsDir = path.join(demoDir, 'clips');
  const outputDir = path.join(demoDir, 'output');
  const outputHtml = path.join(outputDir, 'demo.html');
  const metaPath = path.join(demoDir, '.session-meta.json');

  const screenshots = fs.existsSync(screenshotsDir)
    ? fs.readdirSync(screenshotsDir).filter(f => f.endsWith('.png')).length
    : 0;

  const clips = fs.existsSync(clipsDir)
    ? fs.readdirSync(clipsDir).filter(f => f.endsWith('.mp4') || f.endsWith('.webm')).length
    : 0;

  const hasHtml = fs.existsSync(outputHtml);

  let meta = null;
  if (fs.existsSync(metaPath)) {
    try { meta = JSON.parse(fs.readFileSync(metaPath, 'utf8')); } catch {}
  }

  return {
    slug,
    path: demoDir,
    screenshots,
    clips,
    has_html: hasHtml,
    output_path: hasHtml ? outputHtml : null,
    agent_name: meta?.agent_name || null,
    created_at: meta?.created_at || null,
    last_updated: meta?.last_generated || meta?.created_at || null,
    status: hasHtml ? 'complete' : screenshots > 0 ? 'captured' : 'empty',
  };
}

function getDemoList() {
  if (!fs.existsSync(DEMOS_DIR)) return [];

  return fs.readdirSync(DEMOS_DIR)
    .filter(name => {
      const full = path.join(DEMOS_DIR, name);
      return fs.statSync(full).isDirectory();
    })
    .map(slug => getDemoStatus(slug, path.join(DEMOS_DIR, slug)));
}

// ── Server setup ──────────────────────────────────────────────────────────────

const server = new Server(
  { name: 'agentdemo', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  let result;
  switch (name) {
    case 'create_demo':       result = await handleCreateDemo(args); break;
    case 'get_demo_status':   result = await handleGetDemoStatus(args); break;
    case 'resume_demo':       result = await handleResumeDemo(args); break;
    case 'list_demos':        result = await handleListDemos(); break;
    case 'generate_demo_html': result = await handleGenerateDemoHtml(args); break;
    default:
      result = { error: `Unknown tool: ${name}` };
  }

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify(result, null, 2),
      },
    ],
    isError: !!(result.error),
  };
});

// ── Start ─────────────────────────────────────────────────────────────────────

const transport = new StdioServerTransport();
await server.connect(transport);
console.error('AgentDemo MCP server running on stdio');
