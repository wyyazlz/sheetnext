# Open Edition Scope

This document defines feature boundaries for SheetNext Open Edition.

## Disabled

1. IO import/export pipeline
   - Disabled: import, importFromUrl, export, exportAllImage
   - Kept as primary open API: getData
2. AI module
   - Disabled: AI chat panel, runtime workflow, and AI action entry points
3. Drawing module
   - Disabled: image/chart/shape insertion and drawing lifecycle
4. PivotTable module
   - Disabled: PivotTable actions and related toolbar entries
5. Print module
   - Disabled: print preview, print settings, page setup APIs
6. License module
   - Replaced with open no-activation behavior

## Kept

- Core workbook, sheet, cell editing
- Formula calculation
- Undo/redo
- Data query through `IO.getData()`

## Notes

- Disabled APIs return a clear "Open Edition disabled" message.
- Toolbar is filtered to hide disabled module entries.
