# Cursor Adapter

Cursor can use this package as repository knowledge even when it does not auto-load Agent Skills.

## Recommended Setup

Keep `docs/skill/sheetnext-dev/` in the repository. Add a short project rule that points Cursor to the skill entry point and references.

Suggested rule text:

```text
For SheetNext work, read docs/skill/sheetnext-dev/SKILL.md first. Match the package version and edition. Use references/core-api.md to choose a module in references/api/ and verify exact APIs there. Read integration.md, events.md, enums.md, json-format.md or ai-relay.md only when relevant.
```

## Usage

When requesting SheetNext work, mention the relevant area, such as events, JSON import/export, templates, formulas, or AI relay. Ask Cursor to search the matching reference before editing.
