# GitHub Copilot Adapter

GitHub Copilot can use this package through repository instructions and referenced Markdown files.

## Recommended Setup

Add this summary to `.github/copilot-instructions.md` or your workspace instructions:

```text
For SheetNext development, read docs/skill/sheetnext-dev/SKILL.md and match the installed version and edition. Start with references/core-api.md, then inspect the relevant references/api/ module for exact APIs. Use integration.md, events.md, enums.md, json-format.md and ai-relay.md when needed.
```

## Usage

Reference specific files in prompts when possible, for example: `Use docs/skill/sheetnext-dev/references/events.md to wire the event handler`.
