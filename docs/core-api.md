# Core API

> Generated: 2026-03-27

Public callable classes detected for the generated API surface.

## SheetNext

### Constructor
- Create workbook instance.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dom | HTMLElement | Yes | - | Editor container element. |
| options | SheetNextInitOptions | No | {} | Initialization options. |
| options.licenseKey | string | No | - | License key. |
| options.locale | string | No | 'en-US' | Initial locale (e.g. 'en-US', 'zh-CN'). |
| options.locales | Object<string, Object> | No | - | Extra locale packs keyed by locale code. |
| options.menuRight | function | No | - | Callback `(defaultHTML: string) => string`. Receives the default right-menu HTML, return modified HTML. |
| options.menuList | function | No | - | Callback `(config: Array<{key: string, labelKey?: string, text?: string, groups?: Array, contextual?: boolean, contextType?: string, trigger?: 'panel'\|'action', action?: string, color?: string}>) => Array`. Receives the default toolbar panel config array, return modified array. Use `trigger: 'action'` to make a top tab run code directly without switching the toolbar panel. `color` customizes the top-tab text and accent color. |
| options.AI_URL | string | No | - | AI relay endpoint URL. |
| options.AI_TOKEN | string | No | - | Optional bearer token for AI relay endpoint. |

**Examples**
```js
import SheetNext from 'sheetnext';
import 'sheetnext.css';

const container = document.querySelector('#SNContainer');
const SN = new SheetNext(container);
```
```js
const SN = new SheetNext(document.querySelector('#SNContainer'), {
  menuList(config) {
    return [
      ...config,
      {
        key: 'githubLink',
        text: 'GitHub',
        trigger: 'action',
        color: '#0969da',
        action: "window.open('https://github.com/wyyazlz/sheetnext', '_blank', 'noopener,noreferrer')"
      }
    ];
  }
});
```

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| containerDom | HTMLElement | No | dom | - |
| namespace | string | No | this._setupGlobalNamespace() | - |
| calcMode | 'auto' \| 'manual' | No | 'auto' | - |
| sheets | Sheet[] | No | [] | - |
| properties | Object | No | {} | - |
| Event | EventEmitter | No | new EventEmitter() | - |
| Utils | Utils | No | new Utils(this) | - |
| I18n | I18n | No | this._createI18n(options) | - |
| Print | Print | No | new Print(this) | - |
| Layout | Layout | No | new Layout(this, options) | - |
| AI | AI | No | new AI(this, options, this.#license) | - |
| IO | IO | No | new IO(this, this.#license) | - |
| Formula | Formula | No | new Formula(this) | - |
| UndoRedo | UndoRedo | No | new UndoRedo(this) | - |
| Canvas | Canvas | No | new Canvas(this, this.#license) | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| activeSheet | Sheet | get/set | No | - |
| workbookName | string | get/set | No | - |
| readOnly | boolean | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| locale | string | No | - |

### Methods
#### `static registerLocale(locale, messages): typeof SheetNext`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |
| messages | Object | Yes | - | - |

**Returns**
- Type: `typeof SheetNext`

**Examples**
```js
import SheetNext from 'sheetnext';
import zhCN from 'sheetnext/locales/zh-CN.js';

SheetNext.registerLocale('zh-CN', zhCN);
```

#### `static getLocale(locale): Object | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

**Returns**
- Type: `Object | undefined`

#### `addSheet(sheetName?): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | No | - | - |

**Returns**
- Type: `Sheet | null`

**Examples**
```js
const budgetSheet = SN.addSheet('Budget 2026');

budgetSheet.getCell('A1').value = 'Month';
budgetSheet.getCell('B1').value = 'Amount';
```

#### `delSheet(name)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

#### `getSheet(sheetName): Sheet | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheetName | string | Yes | - | - |

**Returns**
- Type: `Sheet | null`

**Examples**
```js
const sheet = SN.getSheet('Budget 2026');
if (sheet) {
  sheet.getCell('B2').value = 128000;
}
```

#### `recalculate(currentSheetOnly = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| currentSheetOnly | boolean | No | false | - |

**Examples**
```js
SN.recalculate();
SN.recalculate(true);
```

#### `setLocale(locale)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

**Examples**
```js
SN.setLocale('zh-CN');
```

#### `t(key, params = {}): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| key | string | Yes | - | - |
| params | Record<string, any> | No | {} | - |

**Returns**
- Type: `string`

## Sheet

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

totalCell.value = 2560;
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
SN.activeSheet.eachCells('A1:C3', (r, c) => {
  const cell = SN.activeSheet.getCell(r, c);
  cell.fill = { fgColor: { rgb: 'FFF4CC' } };
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

#### `rangeSort(sortKeys, range?)`
- Category: `Sort`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sortKeys | Array<{col:string \| number, order?:string, customOrder?:Array}> | Yes | - | - |
| range | RangeRef | No | - | - |

**Examples**
```js
const sheet = SN.activeSheet;

sheet.rangeSort([
  { col: 'A', order: 'asc' }
], `A2:D${sheet.rowCount}`);
```

#### `flashFill(range, options = {})`
- Category: `FlashFill`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `textToColumns(range, options = {})`
- Category: `TextToColumns`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `groupRows(range, options = {})`
- Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `ungroupRows(range, options = {})`
- Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `groupCols(range, options = {})`
- Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `ungroupCols(range, options = {})`
- Category: `Outline`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `subtotal(range, options = {})`
- Category: `Subtotal`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `consolidate(ranges, options = {})`
- Category: `Consolidate`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ranges | Array<string \| Object> \| Object | Yes | - | - |
| options | Object | No | {} | - |

#### `setDataValidation(range, rule)`
- Category: `DataValidation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |
| rule | Object | Yes | - | - |

#### `clearDataValidation(range)`
- Category: `DataValidation`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| Array<string \| Object> | Yes | - | - |

#### `insertTable(arr, pos, options = {}, ops?): RangeNum`
- Insert a table from the given position
- Category: `InsertTable`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| arr | (ICellConfig \| string \| number)[][] | Yes | - | Table data array |
| pos | CellRef | Yes | - | Insert Location |
| options | Object | No | {} | Configure Options |
| ops | {align?:string,border?:boolean,width?:number,height?:number} | No | - | - |

**Returns**
- Type: `RangeNum`

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

sheet.insertTable(meetingTemplate, 'A1', {
  border: true,
  align: 'center',
  height: 30,
  width: 100
});
```

#### `autoFitRows(startRow, endRow): number`
- Auto fit row heights in range.
- Category: `AutoFit`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| startRow | number | Yes | - | - |
| endRow | number | Yes | - | - |

**Returns**
- Type: `number`

#### `autoFitCols(startCol, endCol): number`
- Auto fit column widths in range.
- Category: `AutoFit`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| startCol | number | Yes | - | - |
| endCol | number | Yes | - | - |

**Returns**
- Type: `number`

#### `areaHaveMerge(area): boolean`
- Tests whether merged cells exist in the range
- Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | Area Object |

**Returns**
- Type: `boolean`
- Whether merged cells exist

#### `unMergeCells(cellAd = null, cell)`
- Unmerger, transfer to cell address or area, default take active Areas
- Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellAd | Object \| String \| null | No | null | Cell address or area |
| cell | CellRef | Yes | - | - |

#### `mergeCells(areas = null, mode = 'default', range)`
- Merge Cells to Support Multiple Modes
- Category: `Merge`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| areas | Array \| Object \| String \| null | No | null | Section, default action Areas |
| mode | String | No | 'default' | Mode: 'default' \|center' \|content'same' |
| range | RangeRef | Yes | - | - |

#### `addRows(r, number = 1, index, count)`
- Insert Row
- Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | Number | Yes | - | Row Index to Insert Location |
| number | Number | No | 1 | Number of rows inserted |
| index | number | Yes | - | - |
| count | number | Yes | - | - |

#### `addCols(c, number = 1, index, count)`
- Insert Columns
- Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | Number | Yes | - | Index of columns for inserting positions |
| number | Number | No | 1 | Number of columns inserted |
| index | number | Yes | - | - |
| count | number | Yes | - | - |

#### `delRows(r, number = 1, start, count)`
- Delete Row
- Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | Number | Yes | - | Remove Row Index at Start Location |
| number | Number | No | 1 | Number of rows deleted |
| start | number | Yes | - | - |
| count | number | Yes | - | - |

#### `delCols(c, number = 1, start, count)`
- Delete Column
- Category: `RowColHandler`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| c | Number | Yes | - | Delete column index for starting position |
| number | Number | No | 1 | NUMBER OF ROWS REMOVED |
| start | number | Yes | - | - |
| count | number | Yes | - | - |

#### `getCellInViewInfo(rowIndex, colIndex, posMerge = true, r, c): Object`
- Fetch cell information in visible view
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | Number | Yes | - | Row index |
| colIndex | Number | Yes | - | Column Index |
| posMerge | Boolean | No | true | Whether to process merging cells |
| r | number | Yes | - | - |
| c | number | Yes | - | - |

**Returns**
- Type: `Object`
- Objects with location and size information

#### `getAreaInviewInfo(area): Object`
- Fetch area information in visible view
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | Area Object |

**Returns**
- Type: `Object`
- Objects with location and size information

#### `getTotalHeight(): number`
- Retrieving the total height of all rows (cache optimization)
- Category: `ViewCalculator`

**Returns**
- Type: `number`
- Total Height

#### `getTotalWidth(): number`
- Get the total width of all columns (cache optimization)
- Category: `ViewCalculator`

**Returns**
- Type: `number`
- Total width

#### `getScrollTop(rowIndex): number`
- Get the sum of heights for all lines before the specified line
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | Row index |

**Returns**
- Type: `number`
- High Sum

#### `getScrollLeft(colIndex): number`
- Sum of widths of all columns before the specified column
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | Column Index |

**Returns**
- Type: `number`
- Sum of width

#### `getRowIndexByScrollTop(scrollTop): number`
- Find line index from vertical scroll pixel position
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollTop | number | Yes | - | Scroll Pixels Position |

**Returns**
- Type: `number`
- Line Index

#### `getColIndexByScrollLeft(scrollLeft): number`
- Find the corresponding column index from the horizontal scroll pixel position
- Category: `ViewCalculator`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollLeft | number | Yes | - | Scroll Pixels Position |

**Returns**
- Type: `number`
- Column Index

#### `copy(area = null, isCut = false, areas?)`
- Copy Regional Data
- Category: `Clipboard`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | Copy area {s: {r, c}, e: {r, c} |
| isCut | boolean | No | false | Whether to cut |
| areas | Array | No | - | - |

#### `cut(area = null, areas?)`
- Cut Area Data
- Category: `Clipboard`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | No | null | Cut Area |
| areas | Array | No | - | - |

#### `paste(targetArea = null, options = {})`
- Paste Data to Target Area
- Category: `Clipboard`

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
- Category: `Clipboard`

#### `getClipboardData(): Object`
- Fetch the current clipboard data
- Category: `Clipboard`

**Returns**
- Type: `Object`

#### `hasClipboardData(): boolean`
- Check for clipboard data
- Category: `Clipboard`

**Returns**
- Type: `boolean`

## Row
- Row Objects

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |
| rIndex | number | No | rIndex | Row index |
| cells | Array<Cell> | No | [] | Line cell arrays |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| height | number | get/set | No | Row height (pixels) |
| hidden | boolean | get/set | No | Whether to hide |
| outlineLevel | number | get/set | No | Outline Level (0-7) |
| collapsed | boolean | get/set | No | Outline Collapse Marker |
| style | Object | get/set | No | Row Styles |
| font | Object | get/set | No | Font style |
| fill | Object | get/set | No | Fill Style |
| alignment | Object | get/set | No | Alignment Style |
| border | Object | get/set | No | Border Style |
| numFmt | string \| undefined | get/set | No | Number Format |

### Methods
#### `init(): void`
- Initialise cell in line

## Col
- Column object

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |
| cIndex | number | No | cIndex | Column Index |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| width | number | get/set | No | Column Width (px) |
| hidden | boolean | get/set | No | Whether to hide |
| outlineLevel | number | get/set | No | Outline Level (0-7) |
| collapsed | boolean | get/set | No | Outline Collapse Marker |
| style | Object | get/set | No | Column Styles |
| font | Object | get/set | No | Font style |
| fill | Object | get/set | No | Fill Style |
| alignment | Object | get/set | No | Alignment Style |
| border | Object | get/set | No | Border Style |
| numFmt | string \| undefined | get/set | No | Number Format |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| cells | Array<Cell> | No | Get all cells in a column |

## Cell
- Cell classes

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| cIndex | number | No | cIndex | Column Index |
| row | Row | No | row | Rows belonging to |
| isMerged | boolean | No | false | Whether or not to merge cells |
| master | {r:number, c:number} \| null | No | null | Merge Cell Main Cell References |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| editVal | string | get/set | No | Edit value or formula |
| richText | Array \| null | get/set | No | Rich text runs array |
| hyperlink | Object \| null | get/set | No | Hyperlink Configuration |
| dataValidation | Object \| null | get/set | No | Data Validation Configuration |
| protection | {locked: boolean, hidden: boolean} | get/set | No | Cell Protection Configuration |
| type | string | get/set | No | Cell Type |
| style | Object | get/set | No | Cell Styles |
| border | Object | get/set | No | Border Style |
| fill | Object | get/set | No | Fill Style |
| font | Object | get/set | No | Font style |
| alignment | Object | get/set | No | Alignment |
| numFmt | string \| undefined | get/set | No | Number Format |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| showVal | string | No | Show values |
| accountingData | Object \| undefined | No | Get Accounting Format Structured Data (for Rendering) |
| calcVal | any | No | Calculated value |
| spillArray | Array[] \| null | No | Get full spill array results |
| isSpillSource | boolean | No | Whether it is a spill source cell |
| isSpillRef | boolean | No | Whether it is a spill reference cell |
| horizontalAlign | string | No | Calculated Horizontal Alignment |
| verticalAlign | string | No | Calculated Vertical Alignment |
| isLocked | boolean | No | Is it locked |
| validData | boolean | No | Data validation results |
| isFormula | boolean | No | Is it a formula |
| buildXml | Object | No | Build Cell XML |

## Canvas
- Canvas Rendering and Interaction Management

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| dpr | number | No | - | - |
| ctx | CanvasRenderingContext2D \| null | No | - | - |
| ctxKZ | CanvasRenderingContext2D \| null | No | - | - |
| HDctx | CanvasRenderingContext2D \| null | No | - | - |
| SN | Object | No | SN | SheetNext Main Instance |
| Utils | Utils | No | SN.Utils | Tool Method Collection |
| dom | HTMLElement | No | SN.containerDom.querySelector('.sn-canvas') | Canvas Container |
| showLayer | HTMLCanvasElement | No | SN.containerDom.querySelector(".sn-show-layer") | Show Layer Canvas |
| handleLayer | HTMLCanvasElement | No | SN.containerDom.querySelector(".sn-handle-layer") | Operating Layer Canvas |
| buffer | HTMLCanvasElement | No | document.createElement('canvas') | Buffer Canvas |
| showLayerKZ | HTMLCanvasElement | No | document.createElement('canvas') | Snapshot Canvas |
| pCanvas | HTMLCanvasElement | No | document.createElement('canvas') | Brush preview canvas |
| pCtx | CanvasRenderingContext2D | No | this.pCanvas.getContext('2d') | Brush Preview Context |
| activeBorderInfo | Object \| null | No | null | Current Active Border Information |
| mpBorder | Object | No | {} | Multi-select border information |
| lastPadding | Object \| null | No | null | Last padding information |
| moveAreaName | string \| null | No | null | Mouse pointer is currently pointing to an area |
| disableMoveCheck | boolean | No | false | Whether to disable mobile detection |
| disableScroll | boolean | No | false | Whether to disable scrolling |
| input | HTMLElement | No | SN.containerDom.querySelector('.sn-input-editor') | Input Editor |
| formulaBar | HTMLElement | No | SN.containerDom.querySelector('.sn-formula-bar') | Formula Bar Input Box |
| inputEditing | boolean | No | false | - |
| areaInput | HTMLInputElement | No | SN.containerDom.querySelector('.sn-area-input') | Area input box |
| areaBox | HTMLElement | No | SN.containerDom.querySelector('.sn-area-box') | Area Selection Container |
| areaDropdown | HTMLElement | No | SN.containerDom.querySelector('.sn-area-dropdown') | Area Dropdown Container |
| rollX | HTMLElement | No | SN.containerDom.querySelector('.sn-roll-x') | Landscape scrollbar |
| rollXS | HTMLElement | No | SN.containerDom.querySelector('.sn-roll-x>div') | Landscape Scrollbar Slider |
| rollY | HTMLElement | No | SN.containerDom.querySelector('.sn-roll-y') | Portrait scrollbar |
| rollYS | HTMLElement | No | SN.containerDom.querySelector('.sn-roll-y>div') | Portrait Scrollbar Slider |
| drawingsCon | HTMLElement | No | SN.containerDom.querySelector('.sn-drawings-con') | Graphics Container |
| slicersCon | HTMLElement | No | SN.containerDom.querySelector('.sn-slicers-con') | Slicer Container |
| maxTop | number | No | 0 | Maximum longitudinal scrolling distance |
| maxLeft | number | No | 0 | Maximum lateral scrolling distance |
| rCount | number | No | 0 | Render Count |
| hoveredComment | Object \| null | No | null | Current hover comment |
| statSum | HTMLElement \| null | No | SN.containerDom.querySelector('[data-stat="sum"]') | Statistical Sum: Dom |
| statCount | HTMLElement \| null | No | SN.containerDom.querySelector('[data-stat="count"]') | Statistics Count: Dom |
| statAvg | HTMLElement \| null | No | SN.containerDom.querySelector('[data-stat="avg"]') | Statistical average: Dom |
| highlightColor | string \| null | No | null | Highlight color, null means closed |
| eyeProtectionMode | boolean | No | false | Whether to turn on eye protection mode |
| eyeProtectionColor | string | No | 'rgb(204,232,207)' | Eye protection color |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| inputEditing | boolean | get/set | No | Whether or not in input edit state |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| paddingType | string | No | Padding Type |
| hyperlinkJumpBtn | HTMLElement \| null | No | Hyperlink Jump Button |
| hyperlinkTip | HTMLElement \| null | No | Hyperlink Tips |
| validSelect | HTMLElement \| null | No | Data Validation Selector |
| validTip | HTMLElement \| null | No | Tips for data validation |
| scrollTip | HTMLElement \| null | No | Scroll Tips |
| width | number | No | Visible Width |
| height | number | No | Visible Height |
| viewWidth | number | No | Logical View Width |
| viewHeight | number | No | Logical View Height |
| pixelRatio | number | No | Pixel Ratio |
| halfPixel | number | No | Half Pixel Offset |
| activeSheet | Sheet | No | Current Active Sheet |

### Methods
#### `viewResize(): boolean`
- Re-fetch dom length after updating canvas zoom/length-width

**Returns**
- Type: `boolean`

#### `toLogical(value): number`
- Physical pixel to logical coordinates

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| value | number | Yes | - | Physical pixels |

**Returns**
- Type: `number`

#### `toPhysical(value): number`
- Logical Coordinates to Physical Pixels

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| value | number | Yes | - | Logical coordinates |

**Returns**
- Type: `number`

#### `px(value): number`
- Convert to value in pixel ratio

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| value | number | Yes | - | Raw Value |

**Returns**
- Type: `number`

#### `snap(value): number`
- Alignment Pixel Ratio

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| value | number | Yes | - | Raw Value |

**Returns**
- Type: `number`

#### `snapRect(x, y, w, h): {x: number, y: number, w: number, h: number}`
- Align Rectangle to Pixel Ratio

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | X |
| y | number | Yes | - | Y |
| w | number | Yes | - | Wide |
| h | number | Yes | - | High |

**Returns**
- Type: `{x: number, y: number, w: number, h: number}`

#### `getEventPosition(event): {x: number, y: number}`
- Gets the position of the event in the canvas (logical coordinates)<br>@ param {MouseEvent \| {clientX: number, clientY: number}} event - Mouse event or object with coordinates

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | - | Yes | - | - |

**Returns**
- Type: `{x: number, y: number}`

#### `getCellIndexByPosition(x, y): {c: number, r: number}`
- Get cell index by location

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | Logical coordinates X |
| y | number | Yes | - | Logical coordinates Y |

**Returns**
- Type: `{c: number, r: number}`

#### `applyZoom(): void`
- Apply zoom and synchronize scrollbar and input box positions

#### `clearBuffer(width = this.buffer.width, height = this.buffer.height): void`
- Clear Buffered Canvas (Physical Pixels)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| width | number | No | this.buffer.width | Width |
| height | number | No | this.buffer.height | Height |

#### `updateScrollBarSize(): void`
- Update scrollbar dimensions

#### `updateOverlayContainers(): void`
- Update Overlay Container Dimensions

#### `updateScrollBar(): void`
- Update scrollbar position

#### `getCellIndex(event): {r:number, c:number}`
- Get the cell index corresponding to the event<br>@ param {MouseEvent \| {clientX: number, clientY: number}} event - Mouse event or object with coordinates

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | - | Yes | - | - |

**Returns**
- Type: `{r:number, c:number}`

#### `setHighlight(event): void`
- Set highlight color and sync toolbar status

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | MouseEvent | Yes | - | Trigger event |

#### `toggleEyeProtection(): void`
- Toggle eye protection mode

#### `async r(type?, options = {}): Promise<void>`
- Render Canvas (Snapshot mode supported)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| type | string | No | - | Rendering type, 's' for snapshot rendering |
| options | Object | No | {} | Rendering Options |
| options.targetCanvas | HTMLCanvasElement | No | - | Target Canvas (Screenshot Mode) |
| options.width | number | No | - | Screenshot Width |
| options.height | number | No | - | Screenshot height |
| options.dpr | number | No | - | Specify dpr |
| options.sheet | Sheet | No | - | Assign Sheets |

**Returns**
- Type: `Promise<void>`

#### `rCell(x, y, w, h, cell)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |

#### `rCellBackground(x, y, w, h, cell, skipGridLine = false, cfFormat = null, tableStyle = null)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |
| skipGridLine | boolean | No | false | - |
| cfFormat | Object \| null | No | null | - |
| tableStyle | Object \| null | No | null | - |

#### `rCellText(x, y, w, h, cell, layoutRect = null)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |
| layoutRect | {x:number,y:number,w:number,h:number} \| null | No | null | - |

#### `rCellTextWithOverflow(r, c, x, y, w, h, cell, sheet, nonEmptyCells, cfFormat = null, tableStyle = null)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | - |
| c | number | Yes | - | - |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |
| sheet | Object | Yes | - | - |
| nonEmptyCells | Array<Object> | Yes | - | - |
| cfFormat | Object \| null | No | null | - |
| tableStyle | Object \| null | No | null | - |

#### `rMerged(sheet)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `rBorder(ciArr)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ciArr | Array<Object> | Yes | - | - |

#### `rBaseData(sheet)`
- Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `getOffset(x, y, w, h): Object`
- Category: `SheetRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |

**Returns**
- Type: `Object`

#### `async rDrawing(nowRCount, options = {}): Promise<void>`
- Category: `SheetRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| nowRCount | number | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `Promise<void>`

#### `async captureScreenshot(cellAddress, csType = 'topleft', option = {}): Promise<void>`
- Category: `ScreenshotRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellAddress | string | Yes | - | - |
| csType | string | No | 'topleft' | - |
| option | Object | No | {} | - |

**Returns**
- Type: `Promise<void>`

#### `getLight()`
- Category: `IndexRenderer`

#### `rAllBtn()`
- Category: `IndexRenderer`

#### `rRowColHeaders(sheet, lightInfo)`
- Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| lightInfo | Object | Yes | - | - |

#### `rHighlightBars(sheet, lightInfo)`
- Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| lightInfo | Object | Yes | - | - |

#### `rPageBreaks(sheet)`
- Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `rIndex(sheet, isScreenshot = false, renderOptions = {})`
- Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| isScreenshot | boolean | No | false | - |
| renderOptions | Object | No | {} | - |

#### `rDom()`
- Category: `IndexRenderer`

#### `showCellInput(clearValue = false, focusEditor = false)`
- Category: `InputManager`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| clearValue | boolean | No | false | - |
| focusEditor | boolean | No | false | - |

#### `updateInputPosition()`
- Category: `InputManager`

#### `updInputValue()`
- Category: `InputManager`

#### `addEventListeners()`
- Category: `EventManager`

#### `rSparklines(sheet)`
- Render sparklines within all views
- Category: `SparklineRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | Sheet Objects |

#### `rCommentIndicators(sheet)`
- Render comment Indicator
- Category: `CommentRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `rCommentBoxes(sheet)`
- Render comment Floating Box
- Category: `CommentRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `rConditionalFormats(sheet)`
- Render Conditional Formatting (Icons Only)<br>The data bar has been rendered in rBaseData to avoid overwriting the text
- Category: `CFRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | - | Yes | - | - |

#### `rDataBar(x, y, w, h, dataBar)`
- Render Data Bar (Export for CellRenderer)
- Category: `CFRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | - | Yes | - | - |
| y | - | Yes | - | - |
| w | - | Yes | - | - |
| h | - | Yes | - | - |
| dataBar | - | Yes | - | - |

#### `rFilterIcons(sheet)`
- Render Filter Icon
- Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Objects |

#### `clearFilterIconPositions()`
- Clear filter icon position information
- Category: `FilterRenderer`

#### `getFilterIconAtPosition(x, y): {colIndex:number, scopeId:string} | null`
- Detect if the click location is on the filter icon
- Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | Click on the X coordinates |
| y | number | Yes | - | Click on the Y position |

**Returns**
- Type: `{colIndex:number, scopeId:string} | null`

#### `drawFilterIcon(ctx, iconX, iconY, iconSize, hasFilter, sortOrder)`
- Draw a single filter icon (Excel style drop-down button)
- Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ctx | - | Yes | - | - |
| iconX | - | Yes | - | - |
| iconY | - | Yes | - | - |
| iconSize | - | Yes | - | - |
| hasFilter | - | Yes | - | - |
| sortOrder | - | Yes | - | - |

#### `rPivotTablePlaceholders(sheet)`
- Render PivotTable Placeholder
- Category: `PivotTableRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `getPivotTableAtPosition(sheet, x, y)`
- Detect if the mouse is in the pivot table placeholder area
- Category: `PivotTableRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |
| x | number | Yes | - | Mouse X coordinates |
| y | number | Yes | - | Mouse Y coordinates<br>@ returns {PivotTable \| null} Returns null if a PivotTable instance is returned within an area |

## Layout
- Layout & Toolbar Management

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | No | SN | SheetNext Main Instance |
| myModal | Object \| null | No | null | Generic Popup Instance |
| toast | Object \| null | No | null | Toast Instance |
| chatWindowOffsetX | number | No | 0 | Chat Window Drag Offset X |
| chatWindowOffsetY | number | No | 0 | Chat Window Drag Offset Y |
| menuConfig | Object | No | MenuConfig.initDefaultMenuConfig(SN, options) | Menu Configuration |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| minimalToolbarEnabled | boolean | get/set | No | - |
| minimalToolbarTextEnabled | boolean | get/set | No | - |
| showMenuBar | boolean | get/set | No | Whether to show the menu bar |
| showToolbar | boolean | get/set | No | Whether to show the toolbar |
| showAIChat | boolean | get/set | No | Whether to show the AI chat portal |
| showFormulaBar | boolean | get/set | No | Whether to show the formula bar |
| showSheetTabBar | boolean | get/set | No | Whether to show the sheet tab bar |
| showStats | boolean | get/set | No | Whether to show the statistics bar |
| showAIChatWindow | boolean | get/set | No | Whether to show the AI chat window |
| showPivotPanel | boolean | get/set | No | Whether to show the PivotTable field panel |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| isSmallWindow | boolean | No | Whether it is a small window |
| isMinimalToolbarLocked | boolean | No | - |

### Methods
#### `openPivotPanel(pt): void`
- Opens the PivotTable Fields panel (manually invoked by the user, such as a toolbar button)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | PivotTable Instance |

#### `togglePivotPanel(pt)`
- Toggle PivotTable field panel display

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | PivotTable Instance<br>@ returns {boolean} shown or not |

#### `autoOpenPivotPanel(pt): void`
- Automatically opens the PivotTable Fields panel (called when the PivotTable area is clicked)<br>Do not open automatically if the user manually closed it

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pt | PivotTable | Yes | - | PivotTable Instance |

#### `closePivotPanel(): void`
- Close the PivotTable field panel (called when you click outside the PivotTable area)<br>The manual close flag will not be set

#### `refreshActiveButtons(): void`
- Refresh button state with active getter in toolbar<br>Bidirectional binding for toggle buttons such as format brush

#### `refreshToolbar(): void`
- Unified refresh toolbar status (panel toggle, invoked after API modification)<br>Integration: Style status + checkbox + active button

#### `scrollSheetTabs(dir): void`
- Scroll Sheet Tab

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dir | number | Yes | - | Scroll Direction (1 or -1) |

#### `updateContextualToolbar(context): void`
- Update Context Toolbar (Show/Hide Context Tab Based on Selection)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| context | Object | Yes | - | Context Information {type: 'table' \| 'pivotTable' \| 'chart' \| null, data: any, autoSwitch: boolean}<br>- autoSwitch: Whether to automatically switch to the context tab (default false, only show no switch) |

#### `initResizeObserver()`
- Initialize ResizeObserver
- Category: `ResizeHandler`

#### `throttleResize()`
- Throttling resize event
- Category: `ResizeHandler`

#### `handleAllResize()`
- Handle all resize related logic
- Category: `ResizeHandler`

#### `updateCanvasSize()`
- Update Canvas Dimensions
- Category: `ResizeHandler`

#### `initDragEvents()`
- Initialize Drag Event
- Category: `DragHandler`

#### `initEventListeners()`
- Initialize all event listeners.
- Category: `EventManager`

#### `initColorComponents(scope?)`
- Initialize color components.
- Category: `EventManager`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scope | Element | No | - | Optional scope. |

#### `removeLoading()`
- Remove loading animation.
- Category: `EventManager`

## Formula
- Formula Calculator

## AutoFilter

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| ref | string \| null | get/set | No | Gets/sets the default (sheet) filter range |
| sortState | {colIndex: number, order: 'asc' \| 'desc'} \| null | get/set | No | Default (sheet) scope sort state |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} \| null | No | Default (sheet) scope |
| columns | Map<number, Object> | No | Default (sheet) scope filter |
| enabled | boolean | No | Whether the default (sheet) scope is enabled |
| headerRow | number | No | Default (sheet) scope table header row |
| hasActiveFilters | boolean | No | Whether the default (sheet) scope has a valid filter |
| activeScopeId | string | No | Current Active Scope ID |

### Methods
#### `setActiveScope(scopeId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | Yes | - | - |

**Returns**
- Type: `string`
- Set active scope id when scope exists.

#### `getTableScopeId(tableId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |

**Returns**
- Type: `string`
- Build table scope id from table id.

#### `isScopeEnabled(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope range is enabled.

#### `hasActiveFiltersInScope(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope has active column filters.

#### `hasAnyActiveFilters(): boolean`

**Returns**
- Type: `boolean`
- Check whether any scope has active filters.

#### `getScopeRange(scopeId = null): {s:{r:number,c:number},e:{r:number,c:number}} | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `{s:{r:number,c:number},e:{r:number,c:number}} | null`
- Get range of scope.

#### `getScopeHeaderRow(scopeId = null): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `number`
- Get header row index of scope.

#### `getColumnFilter(colIndex, scopeId = null): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `Object | null`
- Get filter config for column in scope.

#### `getEnabledScopes(): Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`

**Returns**
- Type: `Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`
- List enabled scopes sorted by range.

#### `registerTableScope(table, options = {}): string | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| table | Object | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `string | null`
- Register or update filter scope for table.

#### `unregisterTableScope(tableId, options = {}): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |
| options | {silent?:boolean} | No | {} | - |

**Returns**
- Unregister table scope and refresh view state.

#### `setRange(range, scopeId = this._defaultScopeId)`
- Set filter scope (default sheet scope; scopeId can be specified)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} | Yes | - | - |
| scopeId | string | No | this._defaultScopeId | - |

#### `isColumnInRange(colIndex, scopeId = null)`
- Determine if a column is within the filter range

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `isHeaderRow(rowIndex, scopeId = null)`
- Determine if a row is a table header row

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `getColumnValues(colIndex, scopeId = null): Array<{value: any, text: string, count: number, isEmpty: boolean}>`
- Get all unique values for a column (for filter panel)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

**Returns**
- Type: `Array<{value: any, text: string, count: number, isEmpty: boolean}>`

#### `setColumnFilter(colIndex, filter, scopeId = null)`
- Set column filters

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| filter | Object | Yes | - | - |
| scopeId | string | No | null | - |

#### `clearColumnFilter(colIndex, scopeId = null)`
- Clear a column filter

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `clearAllFilters(scopeId = null)`
- Clear filters (reserved range)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | null | - |

#### `clear(scopeId = this._defaultScopeId)`
- Clear filters completely (inclusive)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

#### `hasFilter(colIndex, scopeId = null)`
- Determine if columns are filtered

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `setSortState(colIndex, order, scopeId = null)`
- Set sorting status

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| order | 'asc' \| 'desc' | Yes | - | - |
| scopeId | string | No | null | - |

#### `getSortOrder(colIndex, scopeId = null)`
- Get sorting status

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `isRowFilteredHidden(rowIndex): boolean`
- Whether the row is hidden due to filtering

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `parse(xmlObj)`
- Resolve default scopes from sheet AutoFilter

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | - |

#### `parseScope(scopeId, autoFilterXml, options = {})`
- Resolves the filter state of the specified scope to memory (mainly for table import)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | Yes | - | - |
| autoFilterXml | Object | Yes | - | - |
| options | {range?: Object, restoreHiddenRowsFromSheet?: boolean, ownerType?: string, ownerId?: string} | No | {} | - |

#### `buildXml(scopeId = this._defaultScopeId): Object | null`
- Build the XML for the specified scope; export the sheet scope by default

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

**Returns**
- Type: `Object | null`

## Table

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | number | No | - |

### Methods
#### `getAll(): TableItem[]`

**Returns**
- Type: `TableItem[]`

#### `get(id): TableItem | null`
- Get table from ID

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `TableItem | null`

#### `getByName(name): TableItem | null`
- Get table by name

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `TableItem | null`

#### `getTableAt(row, col): TableItem | null`
- Get tables containing specified cells

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `TableItem | null`

#### `add(options = {}): TableItem`
- Add/create new table

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

**Returns**
- Type: `TableItem`

**Examples**
```js
const table = SN.activeSheet.Table;

table.add({
  rangeRef: 'A1:D10',
  name: 'Sales',
  showHeaderRow: true,
  showRowStripes: true
});
```

#### `createFromSelection(area, options = {}): TableItem`
- Create Table from Selection

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| area | Object | Yes | - | { s: {r, c}, e: {r, c} } |
| options | Object | No | {} | - |

**Returns**
- Type: `TableItem`

#### `remove(id): boolean`
- Remove Table

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | Form ID |

**Returns**
- Type: `boolean`

#### `removeByName(name): boolean`
- Remove table by name

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `boolean`

#### `forEach(callback)`
- Walk through all tables

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

#### `toJSON()`
- Convert to JSON array

#### `fromJSON(jsonArray)`
- Restore from JSON

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| jsonArray | Array | Yes | - | - |

## Drawing
- Drawing manager - CRUD, XML parse & export

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| map | Map<string, DrawingItem> | No | new Map() | - |

### Methods
#### `addChart(chartOption, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| chartOption | Object | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

**Examples**
```js
const sheet = SN.activeSheet;

sheet.Drawing.addChart({
  title: { text: 'Sales trend chart' },
  legend: {
    data: `${sheet.name}!B3`
  },
  xAxis: {
    type: 'category',
    data: `${sheet.name}!C2:E2`
  },
  yAxis: { type: 'value' },
  series: [{
    name: 'Sales volume',
    type: 'line',
    data: `${sheet.name}!C3:E3`
  }]
}, {
  startCell: 'B2'
});
```

#### `addImage(imageBase64, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| imageBase64 | string | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

#### `addShape(shapeType, options?): DrawingItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| shapeType | string | Yes | - | - |
| options | DrawingOptions | No | - | - |

**Returns**
- Type: `DrawingItem`

#### `remove(id)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

#### `get(id): DrawingItem | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `DrawingItem | null`

#### `getByCell(cellRef): DrawingItem[]`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | Object \| string | Yes | - | - |

**Returns**
- Type: `DrawingItem[]`

#### `getAll(): DrawingItem[]`

**Returns**
- Type: `DrawingItem[]`

## Comment
- Comment Manager<br>High-performance design: Map index (cellRef - > CommentItem) + O (1) lookup<br>Architectural Reference: Fully Replicated Drawing Design Patterns

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| map | Map<string, CommentItem> | No | new Map() | - |

### Methods
#### `parse(xmlObj)`
- Parsing comment data from xmlObj (imported from an Excel file)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | XML Object for Sheet |

#### `add(config = {}): CommentItem`
- Add Comment

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | No | {} | {cellRef, author, text, visible, ...} |

**Returns**
- Type: `CommentItem`

#### `get(cellRef): CommentItem | null`
- Get comments

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | Cell reference "A1" |

**Returns**
- Type: `CommentItem | null`

#### `remove(cellRef): boolean`
- Delete Comment

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | Cell reference "A1" |

**Returns**
- Type: `boolean`

#### `getAll(): CommentItem[]`
- Get a list of all comments

**Returns**
- Type: `CommentItem[]`

#### `clearCache()`
- Clear the location cache for all comment (called when the view changes)

#### `toXmlObject(): Object | null`
- Export as XML object (for Excel export)<br>Convert thread comment to normal comment format

**Returns**
- Type: `Object | null`

## PivotTable
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
#### `add(config): PivotTable`
- Add Perceive Table

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | Configure Object |
| config.sourceSheet | Sheet | Yes | - | Source Sheet |
| config.sourceRangeRef | string | Yes | - | Source range |
| config.cellRef | string \| Object | Yes | - | target cell<br>Is there a cycle for |

**Returns**
- Type: `PivotTable`

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

#### `forEach(callback)`
- I've been through all of them.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

#### `refreshAll()`
- Refresh all view tables

## Slicer
- Slicer Manager<br>We'll be in charge of the cutter.

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| map | Map<string, SlicerItem> | No | new Map() | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| size | number | No | - |

### Methods
#### `add(config): Slicer`
- Add Slicer

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | Configure Object |

**Returns**
- Type: `Slicer`

#### `remove(id)`
- Remove Slicer

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

#### `get(id): Slicer | null`
- Get Slice

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| id | string | Yes | - | - |

**Returns**
- Type: `Slicer | null`

#### `getAll(): Array<Slicer>`
- Get all the slicers.

**Returns**
- Type: `Array<Slicer>`

#### `forEach(callback)`
- All the slicers.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| callback | Function | Yes | - | - |

## Sparkline
- Mini Chart Manager

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |
| map | Map<string, SparklineItem> | No | new Map() | Minimap indexed by location Map: "R:C" - > SparklineItem |
| groups | Array<Object> | No | [] | Minimap List |

### Methods
#### `add(config): SparklineItem | null`
- Add Minimap

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | Yes | - | Configure Object |
| config.cellRef | string \| Object | Yes | - | Cell references "A1" or {r, c} |
| config.formula | string | Yes | - | Data Range Formula |
| config.type | string | No | 'line' | Miniature Type |
| config.colors | Object | No | {} | Colour Configuration |

**Returns**
- Type: `SparklineItem | null`

**Examples**
```js
const sparkline = SN.activeSheet.Sparkline;

sparkline.add({
  cellRef: 'D1',
  formula: 'A1:C1',
  type: 'column',
  colors: {
    series: '#376092',
    negative: '#D00000',
    high: '#00B050',
    low: '#FF6600'
  }
});
```

#### `remove(cellRef): boolean`
- Delete Minimap

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string \| Object | Yes | - | Cell references "A1" or {r, c} |

**Returns**
- Type: `boolean`

#### `get(cellRef): SparklineItem | null`
- Fetch Minimap Based on Cell Location

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string \| Object | Yes | - | Cell references "A1" or {r, c} |

**Returns**
- Type: `SparklineItem | null`

#### `getAll(): Array<SparklineItem>`
- Fetch All Miniature Examples

**Returns**
- Type: `Array<SparklineItem>`

#### `clearCache(cellRefOrRow, c?)`
- Clear the cache of specified minimaps

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRefOrRow | string \| Object \| number | Yes | - | Cell references "A1"/{r, c} or line index |
| c | number | No | - | Column Index (only when the first parameter is a line index) |

#### `clearAllCache()`
- Clear all caches

#### `parse(xmlObj)`
- Parsing MinimlObj Data

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | XML Object for Sheet |

#### `toXmlObj(): Object | null`
- Export to XML Object Structure

**Returns**
- Type: `Object | null`

## AI
- AI assistant module (tool-calling workflow)

### Methods
#### `chatInput(con)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| con | string | Yes | - | - |

#### `async conversation(p, joinChat = true): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| p | string | Yes | - | - |
| joinChat | boolean | No | true | - |

**Returns**
- Type: `Promise<void>`

#### `clearChat()`

#### `handleFileChange(event)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | Event | Yes | - | - |

#### `async screenshot(addresses = [])`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| addresses | Array | No | [] | - |

## SheetProtection

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

## CF
- CF conditional formatting Manager<br>Responsible for parsing, storing and managing Excel conditional formatting rules

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| rules | Array<CFrule | No | [] | - |

### Methods
#### `add(config = {})`
- Add conditional formatting rule

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | No | {} | Rule Configuration |
| config.rangeRef | string | Yes | - | Scope of application "A1: D10" |
| config.type | string | Yes | - | Rule type (colorScale/dataBar/iconSet/cellIs/expression/top10/aboveAverage/duplicateValues/containsText/timePeriod/containsBlanks/containsErrors)<br>@ returns {CFRule} created rule |

**Examples**
```js
const cf = SN.activeSheet.CF;

cf.add({
  rangeRef: 'A1:A20',
  type: 'cellIs',
  operator: 'greaterThan',
  formula1: 90,
  dxf: {
    font: { color: '#FF0000' },
    fill: { fgColor: '#FFCCCC' }
  }
});

cf.add({
  rangeRef: 'B1:B20',
  type: 'colorScale',
  cfvos: [{ type: 'min' }, { type: 'max' }],
  colors: ['#FFFFFF', '#FF0000']
});
```

#### `remove(index): boolean`
- Delete conditional formatting rule

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| index | number | Yes | - | Rule index |

**Returns**
- Type: `boolean`

#### `get(index): CFRule | null`
- Fetch the rule

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| index | number | Yes | - | Rule index |

**Returns**
- Type: `CFRule | null`

#### `getAll(): Array<CFRule>`
- Get all rules

**Returns**
- Type: `Array<CFRule>`

#### `parse(xmlObj): void`
- Parsing conditional formatting data from xmlObj

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | XML Object for Sheet |

#### `getFormat(r, c)`
- Get a cell's conditional formatting

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| r | number | Yes | - | Row index |
| c | number | Yes | - | Column Index<br>@ returns {Object \| null} formatted object |

#### `clearAllCache(): void`
- Clear all caches

#### `clearRangeCache(sqref): void`
- Clears the cache for the specified range

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sqref | string | Yes | - | Scope |

#### `toXmlObject(): Array`
- Export as an array of sheet XML nodes

**Returns**
- Type: `Array`

#### `getDxfId(dxf): number`
- Get dxf style index (for export)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| dxf | Object | Yes | - | Differentiated Style Objects |

**Returns**
- Type: `number`
- dxfId

#### `getDxfList(): Array`
- Get dxf style list (for exporting styles.xml)

**Returns**
- Type: `Array`

## I18n
- Lightweight i18n service with flat key lookup.

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | No | locale \|\| 'en-US' | - |
| fallbackLocale | string | No | fallbackLocale \|\| 'en-US' | - |

### Methods
#### `register(locale, pack)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |
| pack | LocalePack | Yes | - | - |

#### `setLocale(locale)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| locale | string | Yes | - | - |

#### `t(key, params = {}): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| key | string | Yes | - | - |
| params | Record<string, any> | No | {} | - |

**Returns**
- Type: `string`

## IO

### Methods
#### `async import(file): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| file | File | Yes | - | - |

**Returns**
- Type: `Promise<void>`
- Import from xlsx/csv/json file.

#### `async importFromUrl(fileUrl): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fileUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<void>`

#### `async export(type)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| type | 'XLSX' \| 'CSV' \| 'JSON' \| 'HTML' | Yes | - | - |

#### `async exportAllImage(): Promise<void>`

**Returns**
- Type: `Promise<void>`

#### `async getData(): Promise<Object>`
- Category: `JsonIO`

**Returns**
- Type: `Promise<Object>`

#### `setData(data): boolean`
- Category: `JsonIO`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| data | Object | Yes | - | - |

**Returns**
- Type: `boolean`

## Print
- Print Settings and Rendering

## UndoRedo

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| maxStack | number | No | 50 | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| canUndo | boolean | No | - |
| canRedo | boolean | No | - |

### Methods
#### `add(action): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| action | {undo: Function, redo: Function, activeCell?: {r:number,c:number}} | Yes | - | - |

**Returns**
- Add undo action into current transaction.

#### `begin(name = ''): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | No | '' | - |

**Returns**
- Begin a manual undo transaction.

#### `commit(): void`

**Returns**
- Commit current manual transaction.

#### `undo(): boolean`

**Returns**
- Type: `boolean`

#### `redo(): boolean`

**Returns**
- Type: `boolean`

#### `clear(): void`

**Returns**
- Clear undo/redo stacks and pending transaction.

## Utils

### Methods
#### `objToArr(obj): Array<any>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | any | Yes | - | - |

**Returns**
- Type: `Array<any>`
- Normalize any value to array.

#### `rgbToHex(rgb): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rgb | string | Yes | - | - |

**Returns**
- Type: `string`
- Convert rgb() string to hex color.

#### `emuToPx(emu): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| emu | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert EMU to pixels.

#### `pxToEmu(px): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| px | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert pixels to EMU.

#### `decodeEntities(encodedString): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| encodedString | string | Yes | - | - |

**Returns**
- Type: `string`
- Decode HTML entities to plain text.

#### `sortObjectInPlace(obj, keyOrder): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | Object | Yes | - | - |
| keyOrder | Array<string> | Yes | - | - |

**Returns**
- Reorder object keys in-place.

#### `debounce(func, wait): Function`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| func | Function | Yes | - | - |
| wait | number | Yes | - | - |

**Returns**
- Type: `Function`
- Create a debounced function.

#### `numToChar(num): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| num | number | Yes | - | - |

**Returns**
- Type: `string`

#### `charToNum(name): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `number`

#### `rangeNumToStr(obj, absolute = false): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | {s:{r:number,c:number},e:{r:number,c:number}} | Yes | - | - |
| absolute | boolean | No | false | - |

**Returns**
- Type: `string`

#### `cellStrToNum(cellStr): {c:number,r:number} | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellStr | string | Yes | - | - |

**Returns**
- Type: `{c:number,r:number} | undefined`

#### `cellNumToStr(cellNum): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellNum | {c:number,r:number} | Yes | - | - |

**Returns**
- Type: `string`

#### `getFillFormula(formula, originalRow, originalCol, targetRow, targetCol): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formula | string | Yes | - | - |
| originalRow | number | Yes | - | - |
| originalCol | number | Yes | - | - |
| targetRow | number | Yes | - | - |
| targetCol | number | Yes | - | - |

**Returns**
- Type: `string`
- Build fill formula by shifting relative refs only.

#### `toast(message): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| message | any | Yes | - | - |

**Returns**
- Show toast message.

#### `async downloadImageToBase64(imageUrl): Promise<string | undefined>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| imageUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<string | undefined>`
- Fetch image URL and return base64 data URL.

#### `modal(options = {}): Promise<{bodyEl: HTMLElement} | null>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

**Returns**
- Type: `Promise<{bodyEl: HTMLElement} | null>`
- Open common modal panel.
