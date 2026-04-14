---
name: create-demo
description: Creates a complete interactive demo for a Copilot Studio agent published to M365 Copilot. Captures real screenshots of the agent responding and supporting platforms (SharePoint, Power Automate, Teams, Outlook, Xero) and generates a shareable demo.html file. Use when asked to create, record, or generate a demo for an agent.
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
- Agent description (optional)
- Agent instructions (optional — improves prompt quality)

## How to handle platform URLs
If the user provides platform URLs in their message, extract them and pass to create_demo as the appropriate URL parameters.

If the user does not provide platform URLs but mentions platforms, still pass the platforms list — placeholder slides will be created for those without URLs.

## Steps
1. Ask for the two required URLs if not provided
2. Call the agentdemo MCP tool: create_demo
3. Pass studio_url, m365_url, and any platform URLs provided
4. Report the result to the user including the demo path
5. If the demo has placeholder slides, show the placeholder guide instructions to the user

## Example
User: "Create a demo for the Performance Review Agent. Studio URL: https://copilotstudio... M365 URL: https://m365.cloud.microsoft/... The SharePoint list is at https://tenant.sharepoint.com/... Power Automate flows are at https://make.powerautomate.com/..."

Call create_demo with:
  studio_url, m365_url,
  platforms: "sharepoint,power-automate",
  sharepoint_url: "https://tenant.sharepoint.com/...",
  power_automate_url: "https://make.powerautomate.com/..."

Result: "Demo created at [path]. 2 platform slides + 3 M365 Copilot slides captured. 1 placeholder needs a manual screenshot."
