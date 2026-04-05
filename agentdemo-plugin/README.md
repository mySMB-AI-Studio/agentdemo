# AgentDemo Plugin for Claude Code

Self-contained Mission Control plugin that creates Storylane-style interactive demos for Microsoft Copilot Studio agents.

## Installation

1. Install the plugin zip (`agentdemo-plugin.zip`) through Mission Control.
2. Run `npm install` in the plugin directory to install dependencies.
3. Copy `.env.example` to `.env` and fill in your credentials:

```
DEMO_EMAIL=demo@yourtenant.onmicrosoft.com
DEMO_PASSWORD=your-password
DEMO_TENANT=yourtenant
ANTHROPIC_API_KEY=sk-ant-...    # optional, for AI callout text
```

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
