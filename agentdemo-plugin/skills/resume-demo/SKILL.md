---
name: resume-demo
description: Resumes a previously interrupted demo capture from the last incomplete slide. Use when a demo capture failed or was interrupted and needs to continue.
argument-hint: [agent-name]
allowed-tools: [Bash, Read, Glob]
---

# Resume Agent Demo

## What you need from the user
- Agent name (the demo folder slug)

## Steps
1. If agent name not provided, call list_demos first to show available demos with partial status
2. Call the agentdemo MCP tool: resume_demo
3. Pass the agent_name
4. Report the result
