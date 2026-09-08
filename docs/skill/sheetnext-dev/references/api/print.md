# Print API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Print](#print)

## Print
- Source: `src/core/Print/Print.js`
- Print Settings and Rendering

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| SN | Object | No | SN | SheetNext Main Instance |

### Methods
#### `createDefaultSettings(): Object`
- Create Default Printing Settings

**Returns**
- Type: `Object`

#### `ensureSettings(sheet = this.SN.activeSheet): Object | null`
- Get and ensure print settings for sheets

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `initSheetPrintSettings(sheet): Object`
- Initialising Sheet Printing Settings

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |

**Returns**
- Type: `Object`

#### `readSettingsFromXml(xmlObj, sheet = null): Object`
- Read print settings from XML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | Worksheet XML Object |
| sheet | null | No | null | - |

**Returns**
- Type: `Object`

#### `applySettingsToWorksheet(sheet, ws): void`
- Apply printing settings to worksheet XML

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| ws | Object | Yes | - | Worksheet XML Object |

#### `getSettings(sheet = this.SN.activeSheet): Object | null`
- Get Print Settings (Deep Copy)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `setSettings(partial, sheet = this.SN.activeSheet): Object | null`
- Merge Update Print Settings

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| partial | Object | Yes | - | Partial Settings |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `setPageMargins(margins, sheet = this.SN.activeSheet): Object | null`
- Set Page Margin

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| margins | Object | Yes | - | Margin Configuration |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `getPaperSizes(): Array<Object>`

**Returns**
- Type: `Array<Object>`
- Supported paper sizes in inches.

#### `getMarginPresets(): Object`

**Returns**
- Type: `Object`
- Named margin presets in inches.

#### `getPageMarginsPreset(sheet = this.SN.activeSheet): 'normal' | 'narrow' | 'wide' | 'custom'`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | - |

**Returns**
- Type: `'normal' | 'narrow' | 'wide' | 'custom'`
- Matching margin preset.

#### `setPageMarginsPreset(preset, sheet = this.SN.activeSheet)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| preset | string | Yes | - | - |
| sheet | Sheet | No | this.SN.activeSheet | - |

#### `setPageSetup(setup, sheet = this.SN.activeSheet): Object | null`
- Set Page Configuration

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| setup | Object | Yes | - | Page Settings |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `setPrintOptions(options, sheet = this.SN.activeSheet): Object | null`
- Set Print Options

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | Yes | - | Print Options |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `setPrintArea(range, sheet = this.SN.activeSheet): Object | null`
- Set Print Range

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object | Yes | - | Regional |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `Object | null`

#### `clearPrintArea(sheet = this.SN.activeSheet): void`
- Empty Print Area

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

#### `setPrintTitles(titles, sheet = this.SN.activeSheet): void`
- Set the print title row/column

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| titles | {rows?: string \| Array<number> \| null, cols?: string \| Array<number> \| null} | Yes | - | Repeated rows and columns. |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

#### `toggleGridlines(on, sheet = this.SN.activeSheet): void`
- Toggle Print Grid Line

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | Enable |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

#### `toggleHeadings(on, sheet = this.SN.activeSheet): void`
- Toggle Print Header

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | Enable |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

#### `toggleShowPageBreaks(on, sheet = this.SN.activeSheet): void`
- Show/Hide Page Break Guide

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| on | boolean | Yes | - | Whether to show |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

#### `insertPageBreak(cell = null, sheet = this.SN.activeSheet): {row:boolean, col:boolean} | null`
- Insert Manual Page Break in the current active cell

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | {r: number, c: number} \| null | No | null | Cell above/left of which to insert breaks. |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `{row:boolean, col:boolean} | null`

#### `resetPageBreaks(sheet = this.SN.activeSheet)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | Remove all manual page breaks. |

#### `getPageBreakGuides(sheet = this.SN.activeSheet): {range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>} | null`
- Get page break guide information (for canvas display)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | this.SN.activeSheet | Sheet |

**Returns**
- Type: `{range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>} | null`

#### `async preview(options = {}): Promise<void>`
- Open the print settings and preview panel.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | Preview Options |

**Returns**
- Type: `Promise<void>`

#### `async print(options = {}): Promise<void>`
- Print a job through the browser's system print dialog.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | Job options, or {job, fromPage, toPage, signal, onProgress}. |

**Returns**
- Type: `Promise<void>`

#### `createJob(options = {}): {pages: Array<Object>}`
- Create a pagination snapshot without rendering images or changing saved print areas.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | {sheet, scope: 'sheet'\|'selection'\|'workbook', ranges, ignorePrintArea}. |

**Returns**
- Type: `{pages: Array<Object>}`

#### `buildLayout(sheet, settings): Object`
- Build page break layout information

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| settings | Object | Yes | - | Print Settings |

**Returns**
- Type: `Object`

#### `getPaperDefinition(settings): Object`
- Access to paper definition

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| settings | Object | Yes | - | Print Settings |

**Returns**
- Type: `Object`

#### `getMarginPixels(margins = DEFAULT_MARGINS): {left: number, right: number, top: number, bottom: number}`
- Calculate margin pixels

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| margins | Object | No | DEFAULT_MARGINS | Margin Configuration |

**Returns**
- Type: `{left: number, right: number, top: number, bottom: number}`

#### `getPrintRange(sheet, settings): {s: {r: number, c: number}, e: {r: number, c: number}} | null`
- Get Actual Print Range

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| settings | Object | Yes | - | Print Settings |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}} | null`

#### `normalizeRange(range, sheet): {s: {r: number, c: number}, e: {r: number, c: number}} | null`
- Normative Scope Object

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | string \| Object \| null | Yes | - | Scope |
| sheet | Sheet | Yes | - | Sheet |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}} | null`

#### `getUsedRange(sheet): {s: {r: number, c: number}, e: {r: number, c: number}}`
- Get used area

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |

**Returns**
- Type: `{s: {r: number, c: number}, e: {r: number, c: number}}`

#### `buildSegments(start, end, sizeGetter, maxSize, forcedStarts = []): Array<{s: number, e: number}>`
- Build page segment

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| start | number | Yes | - | Initial Index |
| end | number | Yes | - | End Index |
| sizeGetter | Function | Yes | - | Fetch Size Method |
| maxSize | number | Yes | - | Maximum size |
| forcedStarts | Array | No | [] | - |

**Returns**
- Type: `Array<{s: number, e: number}>`

#### `getRowSize(sheet, r): number`
- Get line height (pixels)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| r | number | Yes | - | Row index |

**Returns**
- Type: `number`

#### `getColSize(sheet, c): number`
- Get column width (pixels)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| c | number | Yes | - | Column Index |

**Returns**
- Type: `number`

#### `getRangeSize(sheet, range, includeHeadings): {width: number, height: number}`
- Get range dimensions

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| range | s: {r: number, c: number}, e: {r: number, c: number} | Yes | - | range |
| includeHeadings | boolean | Yes | - | Whether to include a column header |

**Returns**
- Type: `{width: number, height: number}`

#### `buildCellAddress(sheet, cell): string`
- Build cell address with sheet name

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| cell | r: number, c: number | Yes | - | cell coordinates |

**Returns**
- Type: `string`

#### `async renderPages(sheet, settings, layout): Promise<Array<Object>>`
- Render all pages

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| settings | Object | Yes | - | Print Settings |
| layout | Object | Yes | - | Layout Information |

**Returns**
- Type: `Promise<Array<Object>>`

#### `async renderPage(page, options = {}): Promise<Object>`
- Render one job page, including repeated titles.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| page | Object | Yes | - | A page descriptor from createJob. |
| options | {signal?: AbortSignal} | No | {} | Optional cancellation signal. |

**Returns**
- Type: `Promise<Object>`

#### `async captureRange(sheet, range, width, height, settings, outputScale = 2): Promise<string>`
- Caption range specified

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet |
| range | s: {r: number, c: number}, e: {r: number, c: number} | Yes | - | range |
| width | number | Yes | - | Width |
| height | number | Yes | - | Height |
| settings | Object | Yes | - | Print Settings |
| outputScale | number | No | 2 | Raster resolution multiplier. |

**Returns**
- Type: `Promise<string>`

#### `buildPageHtml(page): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| page | Object | Yes | - | Rendered job page. |

**Returns**
- Type: `string`
- Paper HTML shared by preview and print.

#### `buildPrintHtml(pages): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| pages | Array<Object> | Yes | - | Rendered job pages. |

**Returns**
- Type: `string`
- Standalone printable document.
