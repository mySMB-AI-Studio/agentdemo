# AgentDemo

Automatically creates Storylane-style interactive demos for Microsoft Copilot Studio agents. Paste two URLs, get a polished demo HTML file.

## Getting Started

```bash
npm install
```

### 1. Log in (one time)

```bash
node agentdemo.js auth
```

Complete MFA if prompted, then press Enter. Session is saved for all future runs.

### 2. Create a demo

```bash
node agentdemo.js create
```

Paste your **Copilot Studio agent URL** and **M365 Copilot agent URL** when asked. That's it.

AgentDemo will:
- Read your agent's topics, connections, and config from Copilot Studio
- Open M365 Copilot and capture the agent responding to real prompts
- Generate AI-written callout text for each slide (requires `ANTHROPIC_API_KEY` in `.env`)
- Build a Storylane-style `demo.html` and open it in your browser

### 3. Share

The output is at `demos/{agent-name}/output/demo.html` — a single self-contained HTML file.

## Setup

Copy `.env.example` to `.env` and fill in:

```
DEMO_EMAIL=demo@yourtenant.onmicrosoft.com
DEMO_PASSWORD=your-password
DEMO_TENANT=yourtenant
ANTHROPIC_API_KEY=sk-ant-...    # optional, for AI callout text
```

## Advanced Commands

These exist for power users and debugging. You don't need them for standard use.

| Command | Description |
|---------|-------------|
| `agentdemo create` | **The main command** — does everything end to end |
| `agentdemo auth [--status]` | Manage demo account session |
| `agentdemo init --discover` | Guided interview with auto-discovery |
| `agentdemo check --config <path>` | Pre-flight verification |
| `agentdemo capture --config <path>` | Capture screenshots only |
| `agentdemo generate --config <path>` | Generate HTML from existing assets |
| `agentdemo run --config <path>` | Capture then generate |
| `agentdemo resume --config <path>` | Resume a failed capture run |

## How It Works

1. **Discover** — Reads your agent's topics, trigger phrases, and connected platforms from Copilot Studio
2. **Capture** — Opens M365 Copilot, types each prompt, waits for the agent's streaming response, and screenshots the result
3. **Narrate** — Sends each screenshot to Claude to generate guided-tour callout text
4. **Generate** — Builds a self-contained HTML demo with browser-frame screenshots, callout bubbles, keyboard navigation, and progress dots
