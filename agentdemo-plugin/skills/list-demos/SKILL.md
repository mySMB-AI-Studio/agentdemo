---
name: list-demos
description: Lists all agent demos that have been created or are in progress. Use when asked what demos exist or to check demo status.
argument-hint: ""
allowed-tools: [Read, Glob]
---

# List Agent Demos

Call the agentdemo MCP tool: list_demos
Return the results in a clear table format showing:
- Agent name
- Status (completed / partial / in-progress)
- Number of slides captured
- Date created
- Path to demo.html
