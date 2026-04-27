# Generation

Generated on 2026-04-28 by `scripts/doc-scanner.js`.

## Source Inputs

- Public API comments and signatures under `src/core/`.
- Enum definitions from `src/enum/index.js`.
- Static protocol references preserved from `docs/static-references/ai-relay.md` and `docs/static-references/json-format.md`.
- Built-in examples from `scripts/doc-examples.js`.

## Output Contract

- `SKILL.md` is the portable skill entry point.
- `references/` contains generated or preserved Markdown references for agent lookup.
- `adapters/` contains optional tool-specific setup notes.

Do not manually edit generated API reference files. Update the documented source or static reference input, then rerun the scanner.
