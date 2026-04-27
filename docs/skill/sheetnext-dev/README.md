# SheetNext Dev Agent Skill

Generated from SheetNext API docs on 2026-04-28.

This directory is a portable Agent Skill package for SheetNext development. It is designed to work in two modes:

- Skill-aware agents can install or load the whole `sheetnext-dev/` directory.
- Other AI coding assistants can read `SKILL.md` first, then open the relevant files in `references/`.

## Package Layout

```text
sheetnext-dev/
  SKILL.md              # Portable skill entry point and operating rules
  README.md             # Human-facing usage notes
  GENERATION.md         # Regeneration notes and source inputs
  adapters/             # Optional setup notes for specific AI tools
  references/           # Source-of-truth SheetNext API references
```

## Recommended Use

1. Read `SKILL.md` before generating or modifying SheetNext code.
2. Read only the reference files needed for the task.
3. Verify exact method names and signatures in `references/core-api.md` before coding.
4. Treat `references/` as generated documentation; update SheetNext source comments or static docs, then rerun the scanner.

## Tool Adapters

- Codex: `adapters/codex.md`
- Claude: `adapters/claude.md`
- Cursor: `adapters/cursor.md`
- GitHub Copilot: `adapters/copilot.md`
- ChatGPT or generic assistants: `adapters/chatgpt.md`

The adapters are convenience instructions only. The portable contract is `SKILL.md` plus `references/`.

## Regenerate

From the `SNEditor` directory:

```powershell
node scripts\doc-scanner.js
```
