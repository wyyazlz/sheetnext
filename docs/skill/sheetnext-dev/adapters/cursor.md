# Cursor Adapter

Cursor can use this package as repository knowledge even when it does not auto-load Agent Skills.

## Recommended Setup

Keep `docs/skill/sheetnext-dev/` in the repository. Add a short project rule that points Cursor to the skill entry point and references.

Suggested rule text:

```text
For SheetNext work, read docs/skill/sheetnext-dev/SKILL.md first. Treat docs/skill/sheetnext-dev/references/ as the source of truth. Verify public APIs in references/core-api.md before generating code, and use events.md, enums.md, json-format.md, or ai-relay.md only when relevant.
```

## Usage

When requesting SheetNext work, mention the relevant area, such as events, JSON import/export, templates, formulas, or AI relay. Ask Cursor to search the matching reference before editing.
