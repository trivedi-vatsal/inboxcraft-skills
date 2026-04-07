# InboxCraft Agent Skills

This repository contains the official Agent Skills for **InboxCraft** (`inboxcraft-outlook-rules`), built to be perfectly compatible with `npx skills`, Antigravity, Claude Code, Cursor, and any other AI assistant that follows the [vercel-labs/skills specification](https://github.com/vercel-labs/skills).

## Skills Reference

Below is a complete list of available skills, their purpose, and an example prompt to trigger them via your AI assistant.

| Skill Name | Version | Description | Example Prompt |
| :--- | :--- | :--- | :--- |
| **`inboxcraft-outlook-rules`** | `1.0.0` | Configures Outlook Inbox Rules leveraging COM for folders and Exchange Online for server rules. | *"Generate an InboxCraft rules script to route `alerts@github.com` into a `github` folder inside my `team` folder."* |
| **`inboxcraft-outlook-show-rules`** | `1.0.0` | Fetches and displays all existing Inbox rules and their status in the console. | *"Show me a list of all the Outlook rules I currently have."* |
| **`inboxcraft-outlook-export-rules`** | `1.0.0` | Gathers Outlook inbox rules and exports them to a JSON or CSV file for sharing or backup. | *"Export all my Outlook rules to a CSV file on my desktop."* |
| **`inboxcraft-outlook-show-folders`** | `1.0.0` | Recursively traverses and prints the visual tree-like hierarchy of the Outlook inbox. | *"Print out the full folder tree structure of my Outlook inbox."* |
| **`inboxcraft-outlook-export-folders`** | `1.0.0` | Recursively reads the Outlook folder structure and exports paths and item counts to a file. | *"Export my Outlook folder hierarchy to a JSON file."* |
| **`inboxcraft-outlook-find-large-folders`** | `1.0.0` | Scans Outlook to find and report the top folders consuming the most space or item count. | *"Scan my Outlook and find the top 10 largest folders by item count."* |
| **`inboxcraft-outlook-export-categories`** | `1.0.0` | Extracts all custom Outlook Categories and color labels for backup or inspection. | *"List all my custom Outlook color categories and labels."* |
| **`inboxcraft-outlook-clean-empty-folders`** | `1.0.0` | Tracks down and safely deletes completely empty Outlook subfolders (includes Dry Run mode). | *"Run a script to track down and safely delete any completely empty Outlook subfolders."* |
| **`inboxcraft-outlook-disable-all-rules`** | `1.0.0` | Safely unchecks and pauses all active Outlook inbox rules without deleting them. | *"I am troubleshooting, please temporarily disable every single Outlook rule I have."* |

## Installation

To install **all skills** in the repository, register this repo globally using the `npx skills` CLI:

```bash
npx skills add https://github.com/trivedi-vatsal/inboxcraft-skills --global
```

If you only want to install an **individual skill**, you can use the `--skill` flag:

```bash
npx skills add https://github.com/trivedi-vatsal/inboxcraft-skills --skill inboxcraft-outlook-export-rules --global
```
