# Sheet API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Sheet](#sheet)
- [SheetProtection](#sheetprotection)

## Sheet
- Source: `src/core/Sheet/Sheet.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Workbook | No | SN | - |
| rId | string | No | meta['_$r:id'] | - |
| Utils | Utils | No | SN.Utils | - |
| Canvas | Canvas | No | SN.Canvas | - |
| merges | RangeNum[] | No | [] | - |
| vi | Object | No | null | - |
| views | Object[] | No | [{ pane: {} }] | - |
| rows | Row[] | No | [] | - |
| cols | Col[] | No | [] | - |
| initialized | boolean | No | false | - |
| showGridLines | boolean | No | true | - |
| showRowColHeaders | boolean | No | true | - |
| showPageBreaks | boolean | No | false | - |
| outlinePr | Object | No | { | - |
| printSettings | Object | No | null | - |
| protection | Object | No | new SheetProtection(this) | - |
| Comment | Comment | No | new Comment(this) | - |
| AutoFilter | AutoFilter | No | new AutoFilter(this) | - |
| Table | Table | No | new Table(this) | - |
| Drawing | null | No | null | - |
| Sparkline | null | No | null | - |
| CF | null | No | null | - |
| Slicer | null | No | null | - |
| PivotTable | null | No | null | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| activeCell | CellNum | get/set | No | - |
| activeAreas | RangeNum[] | get/set | No | - |
| zoom | number | get/set | No | - |
| viewStart | CellNum | get/set | No | - |
| defaultColWidth | number | get/set | No | - |
| defaultRowHeight | number | get/set | No | - |
| indexWidth | number | get/set | No | - |
| headHeight | number | get/set | No | - |
| frozenCols | number | get/set | No | - |
| frozenRows | number | get/set | No | - |
| hidden | boolean | get/set | No | - |
| name | string | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| rowCount | number | No | - |
| colCount | number | No | - |

### Methods
#### `zoomIn(step = 0.1): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| step | number | No | 0.1 | - |

**Returns**
- Type: `number`

#### `zoomOut(step = 0.1): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| step | number | No | 0.1 | - |

**Returns**
- Type: `number`

#### `zoomToSelection(): number`

**Returns**
- Type: `number`
- Zoom to fit the selected area.

#### `rangeStrToNum(range): RangeNum`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | RangeStr | Yes | - | - |

**Returns**
- Type: `RangeNum`

**Examples**
```js
const sheet = SN.activeSheet;
const range = sheet.rangeStrToNum('A1:C3');

console.log(range.s, range.e);
```

#### `getRow(r): Row`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | - |

**Returns**
- Type: `Row`

**Examples**
```js
const row = SN.activeSheet.getRow(0);
row.height = 32;
row.hidden = false;
```

#### `getCol(c): Col`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | number | Yes | - | - |

**Returns**
- Type: `Col`

**Examples**
```js
const col = SN.activeSheet.getCol(1);
col.width = 120;
col.hidden = false;
```

#### `getCell(r, c?): Cell`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number \| string | Yes | - | - |
| c | number | No | - | - |

**Returns**
- Type: `Cell`

**Examples**
```js
const totalCell = SN.activeSheet.getCell('B2');

totalCell.editVal = 2560;
totalCell.numFmt = '#,##0.00';
```

#### `eachCells(ranges, callback, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ranges | RangeRef \| RangeRef[] | Yes | - | - |
| callback | (r:number,c:number,area:RangeNum)=>void | Yes | - | - |
| options | {reverse?:boolean,sparse?:boolean} | No | {} | - |

**Examples**
```js
SN.activeSheet.eachCells('A1:C3', (rowIndex, colIndex) => {
  const cell = SN.activeSheet.getCell(rowIndex, colIndex);
  cell.fill = { fgColor: '#FFF4CC' };
});
```

#### `setBrush(keep = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| keep | boolean | No | false | - |

#### `cancelBrush(): boolean`

**Returns**
- Type: `boolean`

#### `applyBrush(targetArea): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetArea | RangeNum | Yes | - | - |

**Returns**
- Type: `boolean`

#### `moveArea(moveArea, targetArea): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| moveArea | Object \| string | Yes | - | - |
| targetArea | Object \| string | Yes | - | - |

**Returns**
- Type: `boolean`

#### `paddingArea(oArea = this.Canvas.lastPadding.oArea, targetArea = this.Canvas.lastPadding.targetArea, type = 'order')`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| oArea | Object \| string | No | this.Canvas.lastPadding.oArea | - |
| targetArea | Object \| string | No | this.Canvas.lastPadding.targetArea | - |
| type | string | No | 'order' | Fill area with series/copy/format. |

#### `showAllHidRows()`

#### `showAllHidCols()`

#### `areasBorder(position, options, area = this.activeAreas)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| position | string | Yes | - | - |
| options | Object \| null | Yes | - | - |
| area | Array | No | this.activeAreas | - |

#### `hyperlinkJump()`
- Jump to hyperlink target of active cell.

#### `rangeSort(sortKeys, range?, options = {}): boolean`
- Source: `src/core/Sheet/Sort.js` | Category: `Sort`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sortKeys | Array<{col:string \| number, order?:string, customOrder?:Array}> | Yes | - | - |
| range | RangeRef | No | - | - |
| options | {hasHeader?:boolean, caseSensitive?:boolean} | No | {} | - |

**Returns**
- Type: `boolean`
- Sort sheet range.

**Examples**
```js
const sheet = SN.activeSheet;

sheet.rangeSort([
  { col: 'A', order: 'asc' }
], `A2:D${sheet.rowCount}`);
```

#### `flashFill(range, options = {})`
- Source: `src/core/Sheet/FlashFill.js` | Category: `FlashFill`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `textToColumns(range, options = {})`
- Source: `src/core/Sheet/TextToColumns.js` | Category: `TextToColumns`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `groupRows(range, options = {})`
- Source: `src/core/Sheet/Outline.js` | Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `ungroupRows(range, options = {})`
- Source: `src/core/Sheet/Outline.js` | Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `groupCols(range, options = {})`
- Source: `src/core/Sheet/Outline.js` | Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `ungroupCols(range, options = {})`
- Source: `src/core/Sheet/Outline.js` | Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `subtotal(range, options = {})`
- Source: `src/core/Sheet/Subtotal.js` | Category: `Subtotal`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `consolidate(ranges, options = {})`
- Source: `src/core/Sheet/Consolidate.js` | Category: `Consolidate`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ranges | Array<string \| Object> \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `setDataValidation(range, rule)`
- Source: `src/core/Sheet/DataValidation.js` | Category: `DataValidation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |
| rule | Object | Yes | - | - |

#### `clearDataValidation(range)`
- Source: `src/core/Sheet/DataValidation.js` | Category: `DataValidation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |

#### `clearRange(range, ops?): {success:boolean, canceled?:boolean, areas:Array, options:Object, cellCount:number, commentCount?:number, rowColStyleCount?:number, mergeCount?:number}`
- Clear cells in a range. Defaults to clearing contents and formats.
- Source: `src/core/Sheet/ClearRange.js` | Category: `ClearRange`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | RangeRef \| RangeRef[] | Yes | - | - |
| ops | string \| Object | No | - | Defaults to clearing contents and formats. |

**Returns**
- Type: `{success:boolean, canceled?:boolean, areas:Array, options:Object, cellCount:number, commentCount?:number, rowColStyleCount?:number, mergeCount?:number}`

#### `getUsedRange(options = {}): {s:{r:number,c:number},e:{r:number,c:number}} | string | null`
- Get the used range of the sheet without scanning the full grid.
- Source: `src/core/Sheet/UsedRange.js` | Category: `UsedRange`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |
| options.contents | boolean | No | true | Include non-empty values. |
| options.formulas | boolean | No | true | Include formula cells. |
| options.richText | boolean | No | true | Include rich text cells. |
| options.hyperlinks | boolean | No | true | Include hyperlink cells. |
| options.dataValidations | boolean | No | true | Include data validation cells. |
| options.controls | boolean | No | true | Include cell control cells. |
| options.comments | boolean | No | true | Include comments. |
| options.merges | boolean | No | true | Include merged ranges. |
| options.tables | boolean | No | true | Include table ranges. |
| options.autoFilters | boolean | No | true | Include AutoFilter ranges. |
| options.drawings | boolean | No | false | Include drawing ranges. |
| options.sparklines | boolean | No | false | Include sparkline cells and source ranges. |
| options.conditionalFormats | boolean | No | false | Include conditional formatting ranges. |
| options.pivotTables | boolean | No | false | Include pivot table output ranges. |
| options.includeFormats | boolean | No | false | Include explicitly formatted cells. |
| options.includeRowColFormats | boolean | No | false | Include row and column format or visibility markers. |
| options.emptyAsNull | boolean | No | false | Return null when the sheet is empty. |
| options.asString | boolean | No | false | Return A1 range string instead of range object. |

**Returns**
- Type: `{s:{r:number,c:number},e:{r:number,c:number}} | string | null`

#### `getCurrentRegion(cell = this.activeCell): RangeNum | null`
- Get the region bounded by blank rows and columns, including hidden cells.
- Source: `src/core/Sheet/SelectionNavigation.js` | Category: `SelectionNavigation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | CellNum | No | this.activeCell | - |

**Returns**
- Type: `RangeNum | null`

#### `moveSelection(direction, options = {}): CellNum | null`
- Move or extend the current selection using Excel-style navigation.
- Source: `src/core/Sheet/SelectionNavigation.js` | Category: `SelectionNavigation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| direction | 'up' \| 'down' \| 'left' \| 'right' | Yes | - | - |
| options | {extend?:boolean,jump?:boolean} | No | {} | - |

**Returns**
- Type: `CellNum | null`

#### `async insertTemplate(arr, pos, options = {}): Promise<RangeNum | undefined>`
- Insert a template, yielding with progress when its target exceeds 100,000 cells.<br>Await completion before reading the inserted range or performing dependent edits.
- Source: `src/core/Sheet/InsertTable.js` | Category: `InsertTable`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| arr | (ICellConfig \| string \| number)[][] | Yes | - | Template data array. |
| pos | CellRef | Yes | - | Top-left insertion position. |
| options | {align?:string,border?:boolean,width?:number,height?:number} | No | {} | Template formatting. |

**Returns**
- Type: `Promise<RangeNum | undefined>`
- Inserted range, or undefined for empty/invalid input.<br>Await insertion; targets over 100,000 cells are processed in slices with progress.

**Examples**
```js
const sheet = SN.activeSheet;

/*
ICellConfig summary:
- v: cell value
- w: column width in px, usually set on the first row
- h: row height in px, usually set on the first column
- b: bold
- s: font size
- fg: background color
- a: text align, one of 'left' | 'right' | 'center'
- c: text color
- mr: merge cells to the right
- mb: merge cells downward

For mr or mb merges, keep the 2D array rectangular with "" placeholders.
Example: { mr: 2 }, "", ""
*/
const meetingTemplate = [
  [{ v: 'Meeting Minutes', s: 16, mr: 3, fg: '#eee', h: 45, b: true }, { w: 160 }, '', { w: 160 }],
  ['Time', '', 'Location', ''],
  ['Host', '', 'Recorder', ''],
  ['Expected', '', 'Present', ''],
  ['Absent Members', { mr: 2 }, '', ''],
  ['Topic', { mr: 2 }, '', ''],
  [{ v: 'Content', h: 280 }, { mr: 2 }, '', ''],
  [{ v: 'Remarks', h: 80 }, { mr: 2 }, '', '']
];

sheet.insertTemplate(meetingTemplate, 'A1', {
  border: true,
  align: 'center',
  height: 30,
  width: 100
});
```

#### `autoFitRows(startRow, endRow): number`
- Auto fit row heights in range.
- Source: `src/core/Sheet/AutoFit.js` | Category: `AutoFit`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| startRow | number | Yes | - | - |
| endRow | number | Yes | - | - |

**Returns**
- Type: `number`

#### `autoFitCols(startCol, endCol): number`
- Auto fit column widths in range.
- Source: `src/core/Sheet/AutoFit.js` | Category: `AutoFit`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| startCol | number | Yes | - | - |
| endCol | number | Yes | - | - |

**Returns**
- Type: `number`

#### `areaHaveMerge(area): boolean`
- Tests whether merged cells exist in the range
- Source: `src/core/Sheet/Merge.js` | Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | Area Object |

**Returns**
- Type: `boolean`
- Whether merged cells exist

#### `unMergeCells(cellAd = null, cell)`
- Unmerger, transfer to cell address or area, default take active Areas
- Source: `src/core/Sheet/Merge.js` | Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellAd | Object \| String \| null | No | null | Cell address or area |
| cell | CellRef | Yes | - | - |

#### `mergeCells(areas = null, mode = 'default', range)`
- Merge Cells to Support Multiple Modes
- Source: `src/core/Sheet/Merge.js` | Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| areas | Array \| Object \| String \| null | No | null | Section, default action Areas |
| mode | String | No | 'default' | Mode: 'default' \|center' \|content'same' |
| range | RangeRef | Yes | - | - |

#### `addRows(r, number = 1)`
- Insert Row
- Source: `src/core/Sheet/RowColHandler.js` | Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | Row Index to Insert Location |
| number | number | No | 1 | Number of rows inserted |

#### `addCols(c, number = 1)`
- Insert Columns
- Source: `src/core/Sheet/RowColHandler.js` | Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | number | Yes | - | Index of columns for inserting positions |
| number | number | No | 1 | Number of columns inserted |

#### `delRows(r, number = 1)`
- Delete Row
- Source: `src/core/Sheet/RowColHandler.js` | Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | Remove Row Index at Start Location |
| number | number | No | 1 | Number of rows deleted |

#### `delCols(c, number = 1)`
- Delete Column
- Source: `src/core/Sheet/RowColHandler.js` | Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | number | Yes | - | Delete column index for starting position |
| number | number | No | 1 | NUMBER OF ROWS REMOVED |

#### `getCellInViewInfo(rowIndex, colIndex, posMerge = true, r, c): Object`
- Fetch cell information in visible view.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | Number | Yes | - | Row index |
| colIndex | Number | Yes | - | Column index |
| posMerge | Boolean | No | true | Whether to process merging cells |
| r | number | Yes | - | - |
| c | number | Yes | - | - |

**Returns**
- Type: `Object`

#### `getAreaInviewInfo(area): Object`
- Fetch clipped area information in visible view.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | Area object |

**Returns**
- Type: `Object`

#### `getTotalHeight(): number`
- Retrieving the total height of all rows.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Returns**
- Type: `number`

#### `getTotalWidth(): number`
- Get the total width of all columns.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Returns**
- Type: `number`

#### `getScrollTop(rowIndex): number`
- Get the sum of heights before the specified row.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | Row index |

**Returns**
- Type: `number`

#### `getScrollLeft(colIndex): number`
- Sum of widths before the specified column.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | Column index |

**Returns**
- Type: `number`

#### `getRowIndexByScrollTop(scrollTop): number`
- Find row index from vertical scroll pixel position.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollTop | number | Yes | - | Scroll pixels position |

**Returns**
- Type: `number`

#### `getColIndexByScrollLeft(scrollLeft): number`
- Find column index from horizontal scroll pixel position.
- Source: `src/core/Sheet/ViewCalculator.js` | Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollLeft | number | Yes | - | Scroll pixels position |

**Returns**
- Type: `number`

#### `copy(area = null, isCut = false, areas?)`
- Copy Regional Data
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | Copy area {s: {r, c}, e: {r, c} |
| isCut | boolean | No | false | Whether to cut |
| areas | Array | No | - | - |

#### `cut(area = null, areas?)`
- Cut Area Data
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | Cut Area |
| areas | Array | No | - | - |

#### `paste(targetArea = null, options = {})`
- Paste Data to Target Area
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| targetArea | Object | No | null | Target area start location {r, c} or full area |
| options | Object | No | {} | Paste Options |
| options.mode | string | Yes | - | PasteMode |
| options.operation | string | Yes | - | Operation |
| options.skipBlanks | boolean | Yes | - | Skip empty cells |
| options.transpose | boolean | Yes | - | Convert |
| options.externalData | Object | Yes | - | External Clipboard Data (from System Clipboard) |

#### `clearClipboard()`
- Clear Clipboard
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

#### `getClipboardData(): Object`
- Fetch the current clipboard data
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

**Returns**
- Type: `Object`

#### `hasClipboardData(): boolean`
- Check for clipboard data
- Source: `src/core/Sheet/Clipboard.js` | Category: `Clipboard`

**Returns**
- Type: `boolean`

#### `setCellControl(range, control)`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |
| control | Object | Yes | - | - |

#### `clearCellControl(range)`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |

#### `toggleCheckbox(row, col): boolean`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `setCheckboxValue(range, checked)`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |
| checked | boolean | Yes | - | - |

#### `deleteCheckboxControlOrValue(range)`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |

#### `hasCheckboxControl(range): boolean`
- Source: `src/core/Sheet/CellControl.js` | Category: `CellControl`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |

**Returns**
- Type: `boolean`

## SheetProtection
- Source: `src/core/Sheet/SheetProtection.js`

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| selectLockedCells | boolean | get/set | No | - |
| selectUnlockedCells | boolean | get/set | No | - |
| formatCells | boolean | get/set | No | - |
| formatColumns | boolean | get/set | No | - |
| formatRows | boolean | get/set | No | - |
| insertColumns | boolean | get/set | No | - |
| insertRows | boolean | get/set | No | - |
| insertHyperlinks | boolean | get/set | No | - |
| deleteColumns | boolean | get/set | No | - |
| deleteRows | boolean | get/set | No | - |
| sort | boolean | get/set | No | - |
| autoFilter | boolean | get/set | No | - |
| pivotTables | boolean | get/set | No | - |
| objects | boolean | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| SN | Workbook | No | - |
| enabled | boolean | No | - |

### Methods
#### `enable(options = {})`
- ENABLE SHEET PROTECTION

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | SAVE OPTIONS |
| options.password | string | No | - | EXPRESS PASSWORD (AUTO-HASHI) |
| options.passwordHash | string | No | - | HASH'S PASSWORD |
| options.selectLockedCells | boolean | No | - | ALLOWS THE SELECTION OF LOCKED CELLS |
| options.selectUnlockedCells | boolean | No | - | ALLOW SELECTION OF UNLOCKED CELLS |
| options.formatCells | boolean | No | - | ALLOW SETTING CELL FORMATS |
| options.formatColumns | boolean | No | - | ALLOW SETTING BAR FORMATS |
| options.formatRows | boolean | No | - | ALLOW SETTING LINE FORMATS |
| options.insertColumns | boolean | No | - | ALLOW INSERT COLUMNS |
| options.insertRows | boolean | No | - | ALLOW INSERT ROWS |
| options.insertHyperlinks | boolean | No | - | ALLOW THE INSERTION OF HYPERLINKS |
| options.deleteColumns | boolean | No | - | ALLOWS THE DELETION OF COLUMNS |
| options.deleteRows | boolean | No | - | ALLOW DELETE ROWS |
| options.sort | boolean | No | - | ALLOW SORTING |
| options.autoFilter | boolean | No | - | ALLOW AUTOFILTER |
| options.pivotTables | boolean | No | - | ALLOW DATA PERCEIVE TABLES |
| options.objects | boolean | No | - | ALLOW EDITING OF OBJECTS |

#### `disable(password?): { ok: boolean, message?: string }`
- DISABLE SHEET PROTECTION

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| password | string | No | - | Password (if password protected) |

**Returns**
- Type: `{ ok: boolean, message?: string }`

#### `verifyPassword(input): { ok: boolean, message?: string }`
- Authentication password

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| input | string | Yes | - | Password entered |

**Returns**
- Type: `{ ok: boolean, message?: string }`

#### `getOptions(): Object | null`
- Get all protection options (for UI display or export)

**Returns**
- Type: `Object | null`

#### `isCellLocked(cell): boolean`
- Check if cells are locked

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | Object | Yes | - | Cell Object |

**Returns**
- Type: `boolean`

#### `isCellSelectable(cell): boolean`
- Check if cells are optional

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | Object | Yes | - | Cell Object |

**Returns**
- Type: `boolean`

#### `isAreaSelectable(area): boolean`
- Check if the area is optional

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object \| string | Yes | - | Area Object or String |

**Returns**
- Type: `boolean`

#### `areAreasSelectable(areas): boolean`
- Check if multiple areas are optional

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| areas | Array | Yes | - | Regional arrays |

**Returns**
- Type: `boolean`

#### `isSelectionVisible(): boolean`
- Whether to allow the selection

**Returns**
- Type: `boolean`

#### `setCellProtection(area, options = {})`
- Set cell protection properties

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object \| string \| Array | Yes | - | Regional |
| options | Object | No | {} | { locked?: boolean, hidden?: boolean } |

#### `buildXmlNode(): Object | null`
- SetProtec node for building xlsx export

**Returns**
- Type: `Object | null`
