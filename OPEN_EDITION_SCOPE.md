# Open Edition Scope

This document defines feature boundaries for SheetNext Open Edition.

## Disabled

1. IO import/export pipeline
   - Disabled: import, importFromUrl, export, exportAllImage
   - Kept as primary open API: getData
2. AI module
   - Disabled: AI chat panel, runtime workflow, and AI action entry points
3. PivotTable module
   - Disabled: PivotTable actions and related toolbar entries
4. Print module
   - Disabled: print preview, print settings, page setup APIs
5. License module
   - Replaced with open no-activation behavior

## Kept

- Core workbook, sheet, cell editing
- Formula calculation
- Drawing module: image/chart/shape insertion, rendering, and drawing lifecycle
- Undo/redo
- Data query through `IO.getData()`

## Notes

- Disabled APIs return a clear "Open Edition disabled" message.
- Toolbar is filtered to hide disabled module entries.
