---
name: create-demo
description: Creates a complete interactive demo for a Copilot Studio agent published to M365 Copilot. Captures real screenshots of the agent responding and generates a shareable demo.html file. Use when asked to create, record, or generate a demo for an agent.
---

# Create Agent Demo

Use this skill when the user wants to create a demo for a Copilot Studio agent.

## What you need from the user
- Copilot Studio agent URL (the overview page URL)
- M365 Copilot agent URL (the deep link to the agent in M365)
- Agent name (optional — can be auto-detected)
- Agent description (optional)
- Agent instructions (optional — improves prompt quality)
- Connected platforms (optional — e.g. sharepoint, xero, custom)

## Steps
1. Ask for the two URLs if not provided
2. Call the agentdemo MCP tool: create_demo
3. Pass studio_url, m365_url, and any optional fields provided
4. Report the result to the user including the demo path
5. If the demo has placeholder slides, show the placeholder guide instructions to the user

## Example usage
User: "Create a demo for the Xero Invoice Agent"
→ Ask for Studio URL and M365 URL
→ Call create_demo with those URLs
→ Return: "Demo created at [path]. 5 slides captured, 1 placeholder needs a manual screenshot."
