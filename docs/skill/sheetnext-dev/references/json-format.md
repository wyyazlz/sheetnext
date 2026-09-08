# SheetNext JSON Format

> Maintained in `docs/static-references/json-format.md`; copied into the skill by `scripts/doc-scanner.js`. See `core-api.md` for the matching source version and fingerprint.

## Purpose

`SN.IO.getData()` exports the current workbook into a JSON object.

`SN.IO.setData(data)` restores a workbook from the same structure.

This document describes the transport format used by those two APIs.

## Versioning

- Current version: `2.0`
- `setData` requires:
  - `data.version === "2.0"`
  - `data.sheets` to be a non-empty array

If either check fails, `setData` returns `false`.

These are entry checks, not a complete JSON schema validation. The format version is independent of the npm package version; keep the package version alongside stored snapshots if you need to plan upgrades or migrations.

## Replacement And Recovery

`setData` replaces the current workbook; it does not merge a patch. On success it clears undo/redo history. It also loads the saved `readOnly` and sheet protection state. Reapply application permissions after loading, and enforce write authorization on the server.

Loading is not atomic. Once processing starts it clears the existing sheet list and mutates the workbook; a later failure can return `false` with partially replaced state. `false` does not mean the old workbook survived unchanged, and there is no automatic rollback.

Before replacing an edited workbook, save `await SN.IO.getData()` as a recovery snapshot and pause autosaving. Validate incoming data and resource limits before calling `setData`. If it fails, attempt to restore the snapshot and check that result too; rollback itself is a best-effort second load and does not restore the old undo history. Retain the snapshot outside the editor and keep saving disabled if recovery also fails. Resume only after an intentional, successful load/recovery and after restoring business permissions. Attach persistence listeners after the initial load.

See [integration notes](integration.md#save-and-restore) for save queuing and concurrency control.

## Top-Level Shape

```json
{
  "version": "2.0",
  "workbookName": "Book1",
  "activeSheet": "Sheet1",
  "calcMode": "auto",
  "readOnly": false,
  "properties": {},
  "definedNames": {},
  "styleTable": {
    "fonts": [],
    "fills": [],
    "borders": [],
    "aligns": [],
    "styles": []
  },
  "sizeTable": {
    "rowHeights": [],
    "colWidths": []
  },
  "sheets": []
}
```

## Top-Level Fields

| Field | Type | Required | Notes |
| --- | --- | --- | --- |
| `version` | string | Yes | Must currently be `"2.0"` |
| `workbookName` | string | No | Workbook display name |
| `activeSheet` | string \| null | No | Active sheet name to restore |
| `calcMode` | `"auto"` \| `"manual"` | No | Defaults to `auto` on import |
| `readOnly` | boolean | No | Workbook read-only state |
| `properties` | object | No | Workbook custom properties |
| `definedNames` | object | No | Workbook-level defined names |
| `styleTable` | object | No | Defaults to empty pools; required to preserve referenced styles |
| `sizeTable` | object | No | Defaults to empty pools; required to preserve referenced sizes |
| `sheets` | array | Yes | Workbook sheet list |
| `pivotCaches` | array | No | Pivot cache payloads |
| `slicerCaches` | array | No | Slicer cache payloads |
| `_snTrace` | any | No | Internal/export trace marker |

## Shared Tables

`styleTable` and `sizeTable` are shared lookup tables.

Instead of storing the same style or size object repeatedly on every cell/row/column, SheetNext stores the object once in a table and references it by index.

### `styleTable`

| Field | Type | Used By |
| --- | --- | --- |
| `fonts` | array | cell `fo` |
| `fills` | array | cell `fi` |
| `borders` | array | cell `b` |
| `aligns` | array | cell `a` |
| `styles` | array | row `s`, column `s`, cell `s` |

### `sizeTable`

| Field | Type | Used By |
| --- | --- | --- |
| `rowHeights` | array<number> | row `h` |
| `colWidths` | array<number> | column `w` |

## Sheet Shape

Each item in `sheets` looks like this:

```json
{
  "name": "Sheet1",
  "hidden": false,
  "showGridLines": true,
  "showRowColHeaders": true,
  "showPageBreaks": false,
  "outlinePr": {},
  "frozenCols": 0,
  "frozenRows": 0,
  "freezeStartRow": 0,
  "freezeStartCol": 0,
  "defaultColWidth": 72,
  "defaultRowHeight": 21,
  "zoom": 1,
  "activeCell": { "r": 0, "c": 0 },
  "viewStart": { "r": 0, "c": 0 },
  "printSettings": null,
  "rowCount": 200,
  "colCount": 32,
  "cols": [],
  "rows": [],
  "merges": [],
  "drawings": []
}
```

### Common Sheet Fields

| Field | Type | Notes |
| --- | --- | --- |
| `name` | string | Sheet name |
| `hidden` | boolean | Hidden sheet state |
| `showGridLines` | boolean | Gridline visibility |
| `showRowColHeaders` | boolean | Row/column header visibility |
| `showPageBreaks` | boolean | Page break visibility |
| `outlinePr` | object | Outline options |
| `frozenCols` | number | Frozen column count |
| `frozenRows` | number | Frozen row count |
| `freezeStartRow` | number | Freeze split start row |
| `freezeStartCol` | number | Freeze split start column |
| `defaultColWidth` | number | Default width in px |
| `defaultRowHeight` | number | Default height in px |
| `zoom` | number | Import is clamped to `0.1` to `4` |
| `activeCell` | `{r,c}` | Active cell |
| `viewStart` | `{r,c}` | View origin |
| `printSettings` | object \| null | Print config |
| `rowCount` | number | Sheet size hint |
| `colCount` | number | Sheet size hint |
| `sheetProtection` | object | Optional sheet protection config |

### Optional Sheet Sections

These sections are included only when data exists:

- `cols`
- `rows`
- `merges`
- `drawings`
- `sparklines`
- `comments`
- `tables`
- `autoFilters`
- `cfRules`
- `pivotTables`
- `slicers`

## Column Records

```json
{
  "cIndex": 0,
  "w": 3,
  "hidden": false,
  "ol": 1,
  "cl": false,
  "s": 2
}
```

| Field | Type | Notes |
| --- | --- | --- |
| `cIndex` | number | Zero-based column index |
| `w` | number | Index into `sizeTable.colWidths` |
| `hidden` | boolean | Hidden column |
| `ol` | number | Outline level |
| `cl` | boolean | Collapsed state |
| `s` | number | Index into `styleTable.styles` |

## Row Records

```json
{
  "rIndex": 0,
  "h": 4,
  "hidden": false,
  "ol": 1,
  "cl": false,
  "s": 5,
  "cells": []
}
```

| Field | Type | Notes |
| --- | --- | --- |
| `rIndex` | number | Zero-based row index |
| `h` | number | Index into `sizeTable.rowHeights` |
| `hidden` | boolean | Hidden row |
| `ol` | number | Outline level |
| `cl` | boolean | Collapsed state |
| `s` | number | Index into `styleTable.styles` |
| `cells` | array | Cell payload list |

## Cell Records

Cell payloads are stored inside `row.cells`.

```json
{
  "c": 0,
  "v": "Hello",
  "vt": "string",
  "s": 1
}
```

### Cell Position Fields

| Field | Type | Notes |
| --- | --- | --- |
| `c` | number | Start column index |
| `e` | number | Optional end column index for repeated identical cells |

If `e` is present, the same decoded cell template is applied from `c` through `e` inclusive.

### Cell Value Fields

| Field | Type | Notes |
| --- | --- | --- |
| `v` | any | Raw cell value |
| `vt` | string \| null | Value type marker |
| `f` | string | Formula string; importer ensures a leading `=` |
| `cv` | any | Cached formula result |
| `cvt` | string \| null | Cached result type marker |

Supported type markers:

- `date`
- `time`
- `dateTime`
- `nan`
- `inf`
- `-inf`

### Cell Style Fields

| Field | Type | Notes |
| --- | --- | --- |
| `s` | number | Index into `styleTable.styles` |
| `fo` | number | Index into `styleTable.fonts` |
| `fi` | number | Index into `styleTable.fills` |
| `b` | number | Index into `styleTable.borders` |
| `a` | number | Index into `styleTable.aligns` |
| `fmt` | string | Number format |
| `p` | object | Protection overrides |

### Cell Feature Fields

| Field | Type | Notes |
| --- | --- | --- |
| `link` | object | Hyperlink payload |
| `dv` | object | Data validation payload |
| `rt` | array | Rich text runs |

Hyperlink shape:

```json
{
  "t": "https://example.com",
  "l": null,
  "tt": "Open link"
}
```

## Merge Records

Each merge item is a normalized range object:

```json
{
  "s": { "r": 0, "c": 0 },
  "e": { "r": 1, "c": 2 }
}
```

## Drawing Records

Basic drawing shape:

```json
{
  "type": "image",
  "startCell": { "r": 1, "c": 1 },
  "offsetX": 0,
  "offsetY": 0,
  "width": 240,
  "height": 120,
  "anchorType": "twoCell",
  "rotation": 0
}
```

Additional drawing fields depend on `type`:

- `chart`: `chartOption`
- `image`: `imageBase64`
- `shape` / `connector`: `shapeType`, `shapeStyle`, `shapeText`, `isTextBox`

## Other Optional Sections

Sheet-level optional payloads:

- `sparklines`
- `comments`
- `tables`
- `autoFilters`
- `cfRules`
- `pivotTables`
- `slicers`

Workbook-level optional payloads:

- `pivotCaches`
- `slicerCaches`

These sections are round-tripped by `getData` / `setData`, but many of them are larger feature-specific payloads and are best produced by exporting an existing workbook first.

## Dynamic Value Wrapper

Some advanced cache payloads use a typed wrapper so special values survive JSON serialization:

```json
{
  "__sn_json_type": "date",
  "__sn_json_value": "2026-03-17T00:00:00.000Z"
}
```

Possible wrapper types:

- `date`
- `nan`
- `inf`
- `-inf`

## Minimal Example

```json
{
  "version": "2.0",
  "workbookName": "Demo",
  "activeSheet": "Sheet1",
  "calcMode": "auto",
  "readOnly": false,
  "properties": {},
  "definedNames": {},
  "styleTable": {
    "fonts": [],
    "fills": [],
    "borders": [],
    "aligns": [],
    "styles": []
  },
  "sizeTable": {
    "rowHeights": [],
    "colWidths": []
  },
  "sheets": [
    {
      "name": "Sheet1",
      "hidden": false,
      "showGridLines": true,
      "showRowColHeaders": true,
      "showPageBreaks": false,
      "outlinePr": {},
      "frozenCols": 0,
      "frozenRows": 0,
      "freezeStartRow": 0,
      "freezeStartCol": 0,
      "defaultColWidth": 72,
      "defaultRowHeight": 21,
      "zoom": 1,
      "activeCell": { "r": 0, "c": 0 },
      "viewStart": { "r": 0, "c": 0 },
      "printSettings": null,
      "rowCount": 10,
      "colCount": 10,
      "cols": [],
      "rows": [
        {
          "rIndex": 0,
          "cells": [
            { "c": 0, "v": "Name" },
            { "c": 1, "v": "Score" }
          ]
        },
        {
          "rIndex": 1,
          "cells": [
            { "c": 0, "v": "Alice" },
            { "c": 1, "v": 95 }
          ]
        }
      ],
      "merges": [],
      "drawings": []
    }
  ]
}
```

## Recommended Workflow

If you need to generate JSON externally:

1. create a workbook in SheetNext
2. export it with `SN.IO.getData()`
3. treat that JSON as the baseline contract
4. modify only the sections you need
5. import it back with `SN.IO.setData(data)`

This is the safest way to stay aligned with future feature additions.
