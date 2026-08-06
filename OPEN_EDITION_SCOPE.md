# Open Edition Scope

This document defines feature boundaries for SheetNext Open Edition.

## Restricted Capabilities

1. IO file import/export
   - Restricted APIs: import, importFromUrl, export, exportAllImage
   - Available APIs: getData, setData
2. Built-in AI
   - Restricted: built-in chat panel, runtime workflows, and AI action entry points
3. PivotTable
   - Restricted: PivotTable authoring, processing, and related toolbar entries

## Open Capabilities

All other SheetNext features included in this distribution are open and
available without activation, including:

- Core workbook, sheet, cell editing
- Formula calculation
- Drawing module: image/chart/shape insertion, rendering, and drawing lifecycle
- Print preview, printing, and page setup
- Undo/redo
- Tables, comments, data validation, sorting, filtering, and other non-restricted features

## Notes

- Restricted APIs return a clear "Open Edition disabled" message.
- Toolbar filtering targets only restricted IO, AI, and PivotTable entries.
- No feature outside the three areas above is disabled by the Open Edition build.
