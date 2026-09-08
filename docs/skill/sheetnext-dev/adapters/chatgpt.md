# ChatGPT And Generic Assistant Adapter

Use this package by uploading or pasting `SKILL.md` first, then adding the reference files needed for the task.

## Minimal Context

Start with:

- `SKILL.md`
- `references/core-api.md`
- Only the relevant module files from `references/api/`, selected using that index

Add these only when needed:

- `references/events.md` for event handling.
- `references/enums.md` for constants.
- `references/json-format.md` for import/export data.
- `references/ai-relay.md` for AI workflow integration.
- `references/recipes.md` for short examples.

## Prompt

You are using the SheetNext Dev skill. Read `SKILL.md` first. Match the installed version and edition, use core-api.md to select the required references/api/ modules, and do not invent undocumented signatures. When uploading files, preserve their module names.
