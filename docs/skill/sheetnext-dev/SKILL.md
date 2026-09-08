---
name: "sheetnext-dev"
description: "Build and debug SheetNext spreadsheet integrations using its documented APIs, events, workbook JSON and AI relay contract. Use for applications that embed SheetNext or changes to SheetNext itself."
---

# SheetNext Development

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

Use these references for the matching SheetNext source version. First identify the installed version and Open Edition or commercial distribution; see [integration notes](references/integration.md). Do not assume the npm package has every documented capability.

Tool-specific setup notes live in `adapters/`; they are optional and do not change the API contract.

## Find The Relevant Contract

- Start with the small [API index](references/core-api.md), then open only the relevant file in `references/api/`.
- [Recipes](references/recipes.md): initialization, values, templates and filtering.
- [Events](references/events.md): subscriptions, cancellation and source locations.
- [Enums](references/enums.md): allowed constant values.
- [JSON format](references/json-format.md): persistence structure and replacement semantics.
- [Integration](references/integration.md): setup, edition limits, saving and lifecycle.
- [AI relay](references/ai-relay.md): request/stream protocol and code-execution trust boundary; read before enabling the built-in AI runtime.

## API Rules

- Search exact signatures and option fields before implementing. If a field or return type is missing, inspect matching source when available; otherwise state the gap instead of guessing.
- Deprecated APIs are excluded from the generated references; do not use undocumented old names.
- Internal SheetNext source members start with `_` or `#`; avoid them unless the user is editing SheetNext internals.
- Coordinates are zero-based when using numeric objects. A1 strings such as `"A1"` are one-based Excel-style references.
- Prefer existing SheetNext core APIs over direct DOM access or ad hoc state mutation.
- Write cell values and formulas through `cell.editVal`; read calculated values from `cell.calcVal`. There is no public `cell.value` setter.
- `SN.IO.setData` replaces the workbook, clears history on success and has no automatic rollback on failure. Pause saving and retain a recovery snapshot before replacement.
- Client read-only flags, sheet protection and event hooks do not replace server authorization.
- For templates, prefer `sheet.insertTemplate(...)`.
- For import/export, use `SN.IO`.
- For events, use `SN.Event`.

## Workflow

1. Identify the target object: `SN`, `Sheet`, `Cell`, `IO`, `Event`, etc.
2. Search the reference docs for the exact method name and signature before coding.
3. Generate the smallest working implementation first.
4. Verify the relevant behavior when requested or required by the project, and state remaining assumptions. Do not claim a smoke test passed from syntax checks alone.

## Search Tips

- Method headings in `references/api/*.md` start with `####` followed by the signature.
- Class headings use `## ClassName`.
- Event names are grouped in `events.md` under `## Emitted events`.
- Enum names are under `## Enum Types` in `enums.md`.

## Maintaining This Package

In the SheetNext source repository, edit JSDoc, `scripts/doc-examples.js`, `scripts/doc-scanner.js` or `docs/static-references/`, then run `npm run docs`. Generated references are not hand-edited. `npm run docs:check` checks regeneration drift, example syntax/imports and key scanned contracts without building the editor. API behavior still requires version-matched runtime checks.
