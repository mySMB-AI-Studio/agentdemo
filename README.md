# AgentDemo

Automatically creates Storylane-style interactive demos for Microsoft Copilot Studio agents. Provide your agent URLs, get a polished demo HTML file — complete with AI-written callouts, platform screenshots, and a full end-to-end conversation.

**Typical run time: 6–10 minutes** for an 8-slide demo.

---

## Setup

### 1. Install dependencies

```
npm install
```

### 2. Configure credentials

Copy `.env.example` to `.env` and fill in:

```
DEMO_EMAIL=darren@yourtenant.onmicrosoft.com
DEMO_PASSWORD=your-password
ANTHROPIC_API_KEY=sk-ant-...

# Optional: named profiles for multi-account demos (e.g. recipient inbox)
PROFILE_SEAN_EMAIL=sean@yourtenant.onmicrosoft.com
PROFILE_SEAN_PASSWORD=sean-password
```

### 3. Authenticate

```
# Default account (the demo presenter / coordinator)
node agentdemo.js auth

# Named profile for a second account (e.g. email recipient)
node agentdemo.js auth --profile sean

# Check session status for all saved sessions
node agentdemo.js auth --status
```

Complete MFA if prompted. Sessions are saved to `.browser-session/` and reused on subsequent runs.

---

## Running a Demo

Demos are driven through the MCP tool interface. Use `call-mcp-tool.mjs` from the command line:

```
cd "C:\path\to\agentdemo"
node call-mcp-tool.mjs <tool_name> <json_args>
```

### Plan first (recommended)

Discovers the agent and surfaces any missing inputs before committing to a full run:

```
node call-mcp-tool.mjs plan_demo "{\"studio_url\": \"...\", \"m365_url\": \"...\"}"
```

### Create a demo

```
node call-mcp-tool.mjs create_demo "{
  \"agent_name\": \"My Agent\",
  \"studio_url\": \"https://copilotstudio.microsoft.com/...\",
  \"m365_url\": \"https://m365.cloud.microsoft/chat/?...\",
  \"platforms\": \"sharepoint,power-automate,outlook\",
  \"sharepoint_url\": \"https://...\",
  \"power_automate_url\": \"https://make.powerautomate.com/...\",
  \"outlook_url\": \"https://outlook.office.com/mail/\",
  \"outlook_recipient_url\": \"https://outlook.office.com/mail/\",
  \"outlook_recipient_profile\": \"sean\",
  \"headless\": true
}"
```

The output HTML is at `demos/{agent-slug}/output/demo.html` — open it in any browser.

```
# Open in browser (Windows)
start demos\my-agent\output\demo.html
```

---

## All MCP Tools

| Tool | Description |
|------|-------------|
| `plan_demo` | Discover agent info and surface missing inputs before running |
| `create_demo` | Full end-to-end demo capture and HTML generation |
| `get_demo_status` | Check screenshots, HTML status, and last updated time |
| `resume_demo` | Resume an interrupted capture session |
| `list_demos` | List all demos with status |
| `generate_demo_html` | Regenerate HTML from existing screenshots without re-running the browser |

### `create_demo` parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `studio_url` | ✓ | Copilot Studio agent URL |
| `m365_url` | ✓ | M365 Copilot chat URL where the agent is published |
| `agent_name` | | Override for agent name if discovery fails |
| `instructions` | | Agent context/workflow description to guide script generation |
| `platforms` | | Comma-separated list: `sharepoint`, `power-automate`, `outlook`, `teams` |
| `sharepoint_url` | | SharePoint list or Excel file URL |
| `power_automate_url` | | Power Automate cloud flows page URL |
| `outlook_url` | | Coordinator's Outlook URL (skipped if `outlook_recipient_url` set) |
| `teams_url` | | Teams channel URL |
| `outlook_recipient_url` | | Recipient's Outlook inbox URL — captures the received email |
| `outlook_recipient_profile` | | Auth profile name for recipient account (e.g. `sean`) |
| `headless` | | Run browser invisibly (default: `false`) |
| `slug` | | Custom folder name for the demo |

---

## How It Works

1. **Discover** — Reads the agent's name, topics, instructions, and connected platforms from Copilot Studio
2. **Script** — AI generates a 3–6 step conversation that tells a complete business story
3. **Capture** — Opens M365 Copilot headlessly, types each prompt, waits for the agent's response, and screenshots the result. If the agent asks follow-up questions before completing an action, AI generates contextually appropriate replies to drive the conversation to completion
4. **Platforms** — Platform screenshots (SharePoint, Power Automate, Outlook) are captured in separate tabs so the chat thread is never interrupted
5. **Recipient inbox** — A second browser context loads the recipient's saved session and waits for the email to arrive, then opens and screenshots it
6. **Narrate** — Each screenshot is sent to Claude to generate guided-tour callout text
7. **Generate** — Builds a self-contained HTML demo with browser-frame screenshots, callout bubbles, keyboard navigation, and progress dots

---

## MCP Server (Claude Code plugin)

AgentDemo runs as an MCP server so it can be driven from within a Claude Code session.

### Register in Claude Code

Add to your Claude Code `settings.json`:

```json
{
  "mcpServers": {
    "agentdemo": {
      "command": "node",
      "args": ["src/mcp-server.js"],
      "cwd": "C:\\path\\to\\agentdemo"
    }
  }
}
```

The MCP server loads credentials from `.env` automatically — no environment variables needed in the plugin config.
