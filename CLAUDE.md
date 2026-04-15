# AgentDemo

## Installation for new users

1. Upload `agentdemo-plugin.zip` to Claude Code:
   Customize > Personal plugins > Upload plugin

2. Run setup to configure credentials:
   ```
   node scripts/setup.js
   ```
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
   This opens a browser to log in (supports MFA) and saves the session to `~/.agentdemo/.browser-session/`.

4. Start using in Claude Code:
   "Create a demo for [agent name]..."

## Do not run node agentdemo.js directly

This repo contains the full source code but the plugin zip is the correct way to install and use it.
Claude Code should never run agentdemo.js directly.

## Project structure

- `src/` — Main source code (CLI + MCP server)
- `agentdemo-plugin/` — Self-contained plugin package (zipped as `agentdemo-plugin.zip`)
- `demos/` — Output directory for captured demos
- `templates/` — HTML templates for demo generation
