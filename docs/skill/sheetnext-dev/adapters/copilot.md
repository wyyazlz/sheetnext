# GitHub Copilot Adapter

GitHub Copilot can use this package through repository instructions and referenced Markdown files.

## Recommended Setup

Add this summary to `.github/copilot-instructions.md` or your workspace instructions:

```text
For SheetNext development, use docs/skill/sheetnext-dev/SKILL.md as the guide. The source-of-truth API docs are in docs/skill/sheetnext-dev/references/. Check core-api.md for exact public APIs before writing code. Use events.md, enums.md, json-format.md, and ai-relay.md only when the task needs those areas.
```

## Usage

Reference specific files in prompts when possible, for example: `Use docs/skill/sheetnext-dev/references/events.md to wire the event handler`.
