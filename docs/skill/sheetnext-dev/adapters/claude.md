# Claude Adapter

Use this package as a Claude skill when the environment supports custom skills. Otherwise, provide the files as project knowledge or attachments.

## Skill-Aware Setup

Install or copy the whole `sheetnext-dev/` directory into the tool-specific skills location, then start a new session and ask the assistant to use the SheetNext Dev skill.

## Manual Setup

Provide `SKILL.md` first. Add only the reference files needed for the current task:

- `references/core-api.md` for the module index, then the relevant `references/api/*.md` file for signatures.
- `references/events.md` for event names and payloads.
- `references/enums.md` for constants and enum values.
- `references/json-format.md` for workbook JSON data.
- `references/ai-relay.md` for built-in AI relay integration.

## Prompt

Use `SKILL.md` as the operating guide. Do not invent SheetNext APIs. Search the relevant reference file for exact method names and signatures before producing code.
