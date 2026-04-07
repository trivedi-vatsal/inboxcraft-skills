# InboxCraft Agent Skills

This repository contains the official Agent Skills for **InboxCraft** (`inboxcraft-outlook-rules`), built to be perfectly compatible with `npx skills`, Antigravity, Claude Code, Cursor, and any other AI assistant that follows the [vercel-labs/skills specification](https://github.com/vercel-labs/skills).

## Available Skills

### `inboxcraft-outlook-rules`
Instructs your AI Coding Agent on how to automatically write perfectly structured, idempotent PowerShell scripts to configure your Outlook Inbox Rules.
It respects our best practices for combining local Outlook COM object execution (for folder hierarchy creation) with reliable Exchange Online fallback, as well as proper configuration of Inbox rules directly on Exchange Online.

## Installation

Register this skill globally on your machine using the `npx skills` CLI:

```bash
npx skills add https://github.com/trivedi-vatsal/inboxcraft-skills --global
```

## Usage

Simply ask your AI assistant:
> "Generate an InboxCraft outlook rules script to route `alerts@github.com` into a `github` folder inside my `team` folder."

The agent will automatically use the template in `inboxcraft-outlook-rules` and synthesize a bulletproof PowerShell script that you just run locally.
