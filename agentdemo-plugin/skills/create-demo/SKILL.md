---
name: create-demo
description: Queues a background capture for a Copilot Studio agent demo. Returns immediately with a job_id — the capture runs in a detached process. Use when asked to create, record, or generate a demo for an agent.
argument-hint: [studio-url] [m365-url]
allowed-tools: [Bash, Read, Glob]
---

# Create Agent Demo

Use this skill when the user wants to create a demo for a Copilot Studio agent.

## What you need from the user
- Copilot Studio agent URL (required)
- M365 Copilot agent URL (required)
- Platform URLs (optional but recommended):
  - SharePoint list/site URL
  - Power Automate flows page URL
  - Teams channel URL
  - Outlook folder/email URL
  - Xero URL
  - Any custom platform URLs
- Agent name (optional — can be auto-detected)
- Agent instructions (optional — greatly improves prompt quality; always pass user-provided prompts/workflow via this field)

## How to handle platform URLs
If the user provides platform URLs in their message, extract them and pass to queue_demo as the appropriate URL parameters.

If the user does not provide platform URLs but mentions platforms, still pass the platforms list — placeholder slides will be created for those without URLs.

## Steps
1. Ask for the two required URLs if not provided
2. Call the agentdemo MCP tool: **queue_demo**
3. Pass studio_url, m365_url, and any platform URLs provided
4. It returns immediately with a **job_id** and state "queued"
5. Tell the user:
   - "Demo capture has started for [agent name]."
   - "Job ID: [job_id]"
   - "You can check progress anytime by asking me for the demo status."
   - If Teams webhook is configured: "You'll get a Teams notification when it's ready."
6. **Do NOT wait for the demo to complete. Do NOT call get_demo_status immediately after. Just confirm the job was queued and move on.**

## Checking status later
When the user asks for demo status, call **get_demo_status** with the `job_id`.
Status states: `queued` → `running` → `complete` | `failed`
The `phase` field shows progress: `discovery_complete` → `script_generated` → `capturing` → `capture_complete` → `callouts_generated` → `html_generated`

## Example
User: "Create a demo for the Performance Review Agent. Studio URL: https://copilotstudio... M365 URL: https://m365.cloud.microsoft/... The SharePoint list is at https://tenant.sharepoint.com/... Power Automate flows are at https://make.powerautomate.com/..."

Call queue_demo with:
  studio_url, m365_url,
  platforms: "sharepoint,power-automate",
  sharepoint_url: "https://tenant.sharepoint.com/...",
  power_automate_url: "https://make.powerautomate.com/..."

Result: "Demo capture has started for Performance Review Agent. Job ID: a1b2c3d4. You can check progress anytime."
