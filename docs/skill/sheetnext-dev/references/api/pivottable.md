# PivotTable API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [PivotTable](#pivottable)
- [PivotTableItem](#pivottableitem)

## PivotTable
- Source: `src/core/PivotTable/PivotTable.js`
- Perspective Table Manager<br>I'm in charge of the retrospect.

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| map | Map<string, PivotTableItem> | No | new Map() | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | number | No | - |

### Methods
#### `add(config): PivotTableItem | null`
- Add a pivot table to this sheet.<br>The source is either a worksheet range (`sourceSheet` + `sourceRangeRef`) or in-memory<br>business records (`records`, cache-first, no worksheet data is written).

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | Configure Object |
| config.sourceSheet | Sheet | No | - | Source sheet (range source) |
| config.sourceRangeRef | string \| Object | No | - | Source range (range source) |
| config.records | Array<Object> \| any[][] | No | - | Business records (cache-first source) |
| config.fields | Array<string \| {name: string, numFmtId?: number}> | No | - | Field definitions for `records`; inferred from object keys when omitted |
| config.cellRef | string \| Object | No | - | Target cell, defaults to A3 |
| config.name | string | No | - | Pivot table name, auto generated when omitted |
| config.rows | Array<string \| number> | No | - | Row fields (field name or index) |
| config.cols | Array<string \| number> | No | - | Column fields |
| config.filters | Array<string \| number> | No | - | Report filter fields |
| config.values | Array<string \| {field: string \| number, subtotal?: string, name?: string}> | No | - | Value fields |
| config.render | boolean | No | true | Calculate and render when a field layout is provided |
| config.chart | Object \| boolean | No | - | Create a pivot chart after render: {type, includeTotals, drawing} |

**Returns**
- Type: `PivotTableItem | null`

**Examples**
```js
const sheet = SN.activeSheet;
const pt = sheet.PivotTable.add({
  sourceSheet: sheet,
  sourceRangeRef: 'A1:D100',
  cellRef: 'F1',
  name: 'SalesPivot'
});

// fieldIndex is the 0-based source column index.
pt.addRowField(0);
pt.addColField(1);
pt.addDataField(3, 'sum');
pt.setFieldSort(0, 'asc');
pt.refresh();
```

#### `remove(name)`
- Remove Perceive Table

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | Vision Table Name |

#### `get(name): PivotTable | null`
- Get a visual watch

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `PivotTable | null`

#### `getAll(): Array<PivotTable>`
- Get all view tables

**Returns**
- Type: `Array<PivotTable>`

#### `getByCell(row, col): PivotTableItem | null`
- Get the pivot table containing a cell.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | Row index |
| col | number | Yes | - | Column index |

**Returns**
- Type: `PivotTableItem | null`

#### `forEach(callback)`
- I've been through all of them.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

#### `refreshAll(options = {})`
- Refresh all view tables

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

## PivotTableItem
- Source: `src/core/PivotTable/PivotTableItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| name | string | No | options.name ?? this._generateName() | - |
| cacheId | string \| number \| undefined | No | options.cacheId | - |
| uid | string \| null | No | options.uid ?? null | - |
| location | {ref: string, firstHeaderRow: number, firstDataRow: number, firstDataCol: number} | No | { | - |
| pivotFields | PivotField[] | No | [] | - |
| rowFields | number[] | No | [] | - |
| colFields | number[] | No | [] | - |
| dataFields | PivotDataField[] | No | [] | - |
| pageFields | number[] | No | [] | - |
| rowGrandTotals | boolean | No | options.rowGrandTotals !== false | - |
| colGrandTotals | boolean | No | options.colGrandTotals !== false | - |
| showRowHeaders | boolean | No | options.showRowHeaders !== false | - |
| showColHeaders | boolean | No | options.showColHeaders !== false | - |
| compact | boolean | No | options.compact !== false | - |
| compactData | boolean | No | options.compactData !== false | - |
| outline | boolean | No | options.outline !== false | - |
| outlineData | boolean | No | options.outlineData !== false | - |
| showDrill | boolean | No | options.showDrill !== false | - |
| gridDropZones | boolean | No | options.gridDropZones ?? true | - |
| multipleFieldFilters | boolean | No | options.multipleFieldFilters !== false | - |
| showPageFields | boolean | No | options.showPageFields !== false | - |
| reportFilterLayout | string | No | options.reportFilterLayout ?? 'vertical' | - |
| reportFilterFieldsPerRow | number | No | Math.max(1, Number(options.reportFilterFieldsPerRow) \|\| 1) | - |
| compactIndent | number | No | Number.isFinite(options.compactIndent) ? Math.max(0, options.compactIndent) : 1 | - |
| mergeLabelCells | boolean | No | options.mergeLabelCells ?? false | - |
| errorValueEnabled | boolean | No | options.errorValueEnabled ?? false | - |
| errorValueText | string | No | options.errorValueText ?? '' | - |
| emptyValueEnabled | boolean | No | options.emptyValueEnabled ?? false | - |
| emptyValueText | string | No | options.emptyValueText ?? '' | - |
| enableValueCellEditing | boolean | No | options.enableValueCellEditing ?? false | - |
| showContextTips | boolean | No | options.showContextTips !== false | - |
| printDrillButtons | boolean | No | options.printDrillButtons !== false | - |
| printRepeatRowLabels | boolean | No | options.printRepeatRowLabels ?? false | - |
| printTitles | boolean | No | options.printTitles ?? false | - |
| includeCalculatedItemsInTotals | boolean | No | options.includeCalculatedItemsInTotals !== false | - |
| dataCaption | string | No | - | - |
| grandTotalCaption | string | No | - | - |
| pivotTableStyleName | string | No | options.pivotTableStyleName ?? 'PivotStyleLight16' | - |
| showRowStripes | boolean | No | options.showRowStripes ?? false | - |
| showColStripes | boolean | No | options.showColStripes ?? false | - |
| formats | Array | No | [] | - |
| chartFormat | string \| number \| null | No | options.chartFormat ?? null | - |
| chartFormats | Object \| null | No | options.chartFormats ? _clonePivotXmlNode(options.chartFormats) : null | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| cache | Object \| null | get/set | No | - |

### Methods
#### `addRowField(fieldIndex, position = -1): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| position | number | No | -1 | - |

**Returns**
- Type: `boolean`

#### `removeRowField(fieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `addColField(fieldIndex, position = -1): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| position | number | No | -1 | - |

**Returns**
- Type: `boolean`

#### `removeColField(fieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `addDataField(fieldIndex, subtotal = 'sum', name = null, position = -1): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| subtotal | string | No | 'sum' | - |
| name | string \| null | No | null | - |
| position | number | No | -1 | - |

**Returns**
- Type: `boolean`

#### `removeDataField(dataFieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dataFieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `moveDataField(dataFieldIndex, position = -1): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dataFieldIndex | number | Yes | - | - |
| position | number | No | -1 | - |

**Returns**
- Type: `boolean`

#### `addPageField(fieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `removePageField(fieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `moveField(fieldIndex, fromArea, toArea, position = -1)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| fromArea | string | Yes | - | - |
| toArea | string | Yes | - | - |
| position | number | No | -1 | - |

#### `cloneToSheet(targetSheet, options = {}): PivotTableItem | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetSheet | Sheet | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `PivotTableItem | null`

#### `moveTo(targetSheet, cellRef = { r: 2, c: 0 }): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetSheet | Sheet | Yes | - | - |
| cellRef | string \| {r:number,c:number} | No | { r: 2, c: 0 } | - |

**Returns**
- Type: `boolean`

#### `changeDataSource(sourceSheet, sourceRangeRef, options = {}): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sourceSheet | Sheet | Yes | - | - |
| sourceRangeRef | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `boolean`

#### `setDataFieldSubtotal(dataFieldIndex, subtotal)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dataFieldIndex | number | Yes | - | - |
| subtotal | string | Yes | - | - |

#### `setDataFieldShowAs(dataFieldIndex, showDataAs, baseField = -1, baseItem = 0)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dataFieldIndex | number | Yes | - | - |
| showDataAs | string | Yes | - | - |
| baseField | number | No | -1 | - |
| baseItem | number | No | 0 | - |

#### `setItemHidden(fieldIndex, itemIndex, hidden): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| itemIndex | number | Yes | - | - |
| hidden | boolean | Yes | - | - |

**Returns**
- Type: `boolean`

#### `setItemsHidden(fieldIndex, hiddenMap): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| hiddenMap | Object | Yes | - | - |

**Returns**
- Type: `boolean`

#### `clearFieldFilter(fieldIndex): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `setFieldSort(fieldIndex, sortType): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fieldIndex | number | Yes | - | - |
| sortType | string | Yes | - | - |

**Returns**
- Type: `boolean`

#### `calculate()`

#### `getPivotData(dataFieldName, fieldItems = [): number | string`
- Return a pivot value using Excel GETPIVOTDATA-style field/item criteria.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dataFieldName | string | Yes | - | Pivot data field caption |
| fieldItems | Array | No | [ | ] - Pairs of field name and item value |

**Returns**
- Type: `number | string`

#### `isRowCollapsed(level, values): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| level | number | Yes | - | - |
| values | Array | Yes | - | - |

**Returns**
- Type: `boolean`

#### `toggleRowCollapsed(level, values)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| level | number | Yes | - | - |
| values | Array | Yes | - | - |

#### `getRowToggleInfo(row, col): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `Object | null`

#### `getRenderRef(): string`

**Returns**
- Type: `string`

#### `render(options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

#### `createChart(options = {}): DrawingItem | null`
- Create a PivotChart from this PivotTable.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |
| options.type | string | No | 'bar_clustered' | Chart type key. |
| options.includeTotals | boolean | No | true | Whether chart source includes grand totals. |
| options.drawing | Object | No | - | Drawing position and size options. |

**Returns**
- Type: `DrawingItem | null`

#### `static fromXml(sheet, xmlObj, cache): PivotTableItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |
| xmlObj | Object | Yes | - | - |
| cache | Object \| null | Yes | - | - |

**Returns**
- Type: `PivotTableItem`

#### `prepareForExport()`

#### `toXml()`

#### `refresh()`

#### `remove()`
