---
name: "sheetnext-dev"
description: "Use for SheetNext spreadsheet apps, templates, events, import/export, JSON workbook data, and AI relay integration."
---

# SheetNext Development

Generated from SheetNext API docs on 2026-04-28.

This is a portable Agent Skill package. `SKILL.md` is the entry point, `references/` contains lazy-loaded API docs, and `adapters/` contains tool-specific setup notes.

Use the bundled references as the source of truth. Do not invent APIs.

## Compatibility

- Skill-aware agents can install or load this directory directly.
- Any AI coding assistant can use this package by reading this file first, then opening the relevant files in `references/`.
- Tool-specific setup notes live in `adapters/`; they are optional and do not change the source-of-truth API docs.

## When To Use

- Building SheetNext workbook, sheet, row, column, cell, drawing, table, pivot, slicer, formula, event, import/export, or AI relay features.
- Debugging SheetNext integrations or generated spreadsheet code.
- Converting requirements into SheetNext templates, JSON workbook data, or event-driven workflows.

## Read Order

1. For normal spreadsheet development, read `references/core-api.md`.
2. For events and hooks, read `references/events.md`.
3. For constants and enum values, read `references/enums.md`.
4. For JSON import/export data shape, read `references/json-format.md`.
5. For built-in AI relay integration, read `references/ai-relay.md`.
6. For common implementation patterns, read `references/recipes.md`.

## API Rules

- Public API is whatever appears in `references/core-api.md`.
- Deprecated APIs are excluded from the generated references; do not use undocumented old names.
- Internal SheetNext source members start with `_` or `#`; avoid them unless the user is editing SheetNext internals.
- Coordinates are zero-based when using numeric objects. A1 strings such as `"A1"` are one-based Excel-style references.
- Prefer existing SheetNext core APIs over direct DOM access or ad hoc state mutation.
- For templates, prefer `sheet.insertTemplate(...)`.
- For import/export, use `SN.IO`.
- For events, use `SN.Event`.

## Workflow

1. Identify the target object: `SN`, `Sheet`, `Cell`, `IO`, `Event`, etc.
2. Search the reference docs for the exact method name and signature before coding.
3. Generate the smallest working implementation first.
4. State verification steps and any assumptions about the documented API.

## Search Tips

- Method headings in `core-api.md` start with `####` followed by the signature.
- Class headings use `## ClassName`.
- Event names are grouped in `events.md` under `## Emitted events`.
- Enum names are under `## Enum Types` in `enums.md`.
