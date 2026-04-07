# InboxCraft Agent Skills

This repository contains the official Agent Skills for **InboxCraft** (`inboxcraft-outlook-rules`), built to be perfectly compatible with `npx skills`, Antigravity, Claude Code, Cursor, and any other AI assistant that follows the [vercel-labs/skills specification](https://github.com/vercel-labs/skills).

## Skills Reference

Below is a complete list of available skills, their purpose, and an example prompt to trigger them via your AI assistant.

| Skill Name | Description | Example Prompt |
| :--- | :--- | :--- |
| **`inboxcraft-outlook-rules`** | Generates idempotent PowerShell scripts to configure Outlook Inbox Rules, leveraging COM for folders and Exchange Online for server rules. | *"Generate an InboxCraft rules script to route `alerts@github.com` into a `github` folder inside my `team` folder."* |
| **`inboxcraft-outlook-show-rules`** | Fetches all existing Inbox rules via COM and outputs their names and status to the console. | *"Show me a list of all the Outlook rules I currently have."* |
| **`inboxcraft-outlook-export-rules`** | Gathers your Outlook rules and exports them to a JSON or CSV file for sharing or backup. | *"Export all my Outlook rules to a CSV file on my desktop."* |
| **`inboxcraft-outlook-show-folders`** | Recursively traverses the inbox subfolders and prints a visual tree-like hierarchy to the console. | *"Print out the full folder tree structure of my Outlook inbox."* |
| **`inboxcraft-outlook-export-folders`** | Recursively reads your Outlook folder structure and exports the paths and item counts into a CSV or JSON file. | *"Export my Outlook folder hierarchy to a JSON file."* |
| **`inboxcraft-outlook-find-large-folders`** | Iterates through your folders and reports which ones have the most items, helping you declutter. | *"Scan my Outlook and find the top 10 largest folders by item count."* |
| **`inboxcraft-outlook-export-categories`** | Reads all custom Categories/Labels along with their colors from your Outlook profile and exports them. | *"List all my custom Outlook color categories and labels."* |
| **`inboxcraft-outlook-clean-empty-folders`** | Automatically scans and safely deletes folders containing 0 items, securely cleaning up unused folder trees (includes a Dry Run mode). | *"Run a script to track down and safely delete any completely empty Outlook subfolders."* |
| **`inboxcraft-outlook-disable-all-rules`** | A "Panic Button" script that quickly unchecks (disables) all inbox routing rules without deleting them. | *"I am troubleshooting, please temporarily disable every single Outlook rule I have."* |

## Installation

To install **all skills** in the repository, register this repo globally using the `npx skills` CLI:

```bash
npx skills add https://github.com/trivedi-vatsal/inboxcraft-skills --global
```

If you only want to install an **individual skill**, you can use the `--skill` flag:

```bash
npx skills add https://github.com/trivedi-vatsal/inboxcraft-skills --skill inboxcraft-outlook-export-rules --global
```
