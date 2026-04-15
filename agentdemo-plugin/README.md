# AgentDemo Plugin for Claude Code

Self-contained plugin that creates Storylane-style interactive demos for Microsoft Copilot Studio agents.

## Installation

1. Upload `agentdemo-plugin.zip` to Claude Code:
   Customize > Personal plugins > Upload plugin

2. Run setup to configure credentials:
   ```
   node scripts/setup.js
   ```
   This creates `~/.agentdemo/.env` with your M365 demo account details.

   OR manually create `~/.agentdemo/.env`:
   ```
   DEMO_EMAIL=your-demo-account@company.com
   DEMO_PASSWORD=your-password
   DEMO_TENANT=your-tenant
   ANTHROPIC_API_KEY=your-key
   ```

3. Authenticate with M365:
   ```
   node scripts/auth-standalone.js
   ```
   This opens a browser to log in (supports MFA) and saves the session.

4. Restart Claude Code. The `agentdemo` MCP tools will be available.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_demo` | Create a demo from Copilot Studio + M365 Copilot URLs |
| `get_demo_status` | Check screenshots, HTML, and status for a demo |
| `resume_demo` | Resume an interrupted capture session |
| `list_demos` | List all demos with status |
| `generate_demo_html` | Regenerate HTML from existing screenshots |

## How It Works

1. Reads your agent's topics and config from Copilot Studio
2. Opens M365 Copilot and captures the agent responding to real prompts
3. Generates AI-written callout text for each slide (requires `ANTHROPIC_API_KEY`)
4. Builds a self-contained `demo.html` with browser-frame screenshots and navigation

## Where credentials are stored

The plugin looks for `.env` in this order:
1. `scripts/.env` (plugin directory)
2. Plugin root `.env`
3. `~/.agentdemo/.env` (recommended)
4. Current working directory `.env`

Browser sessions are saved to `~/.agentdemo/.browser-session/`.
