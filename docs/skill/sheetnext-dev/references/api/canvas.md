# Canvas API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [Canvas](#canvas)
- [FormulaEditor](#formulaeditor)
- [HyperlinkPopover](#hyperlinkpopover)

## Canvas
- Source: `src/core/Canvas/Canvas.js`
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
| scrollSkeletonLayer | HTMLCanvasElement | No | document.createElement('canvas') | Scroll loading skeleton layer |
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
| middleLayer | - | No | SN.containerDom.querySelector('.sn-middle-layer') | - |
| slicersCon | HTMLElement | No | SN.containerDom.querySelector('.sn-slicers-con') | Slicer Container |
| maxTop | number | No | 0 | Maximum longitudinal scrolling distance |
| maxLeft | number | No | 0 | Maximum lateral scrolling distance |
| formulaEditor | FormulaEditor | No | new FormulaEditor(this) | Inline Formula Editor Controller |
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

#### `startRangePicker(options = {}): () => void`
- Start a modeless grid range-picking session.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | Picker callbacks and target context |
| options.onChange | (reference:string, detail:{area:Object, sheet:Object})=>void | Yes | - | Range change callback |
| options.targetSheetName | string | No | - | Formula target sheet used to qualify cross-sheet references |

**Returns**
- Type: `() => void`
- Session disposer

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

#### `scrollByPixels(deltaX = 0, deltaY = 0, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| deltaX | number | No | 0 | - |
| deltaY | number | No | 0 | - |
| options | Object | No | {} | - |

#### `scrollToPixels(scrollLeft, scrollTop, options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scrollLeft | - | Yes | - | - |
| scrollTop | - | Yes | - | - |
| options | Object | No | {} | - |

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
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |

#### `rCellBackground(x, y, w, h, cell, skipGridLine = false, cfFormat = null, tableStyle = null)`
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

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

#### `rCellText(x, y, w, h, cell, layoutRect = null, cfFormat = null)`
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| cell | Object | Yes | - | - |
| layoutRect | {x:number,y:number,w:number,h:number} \| null | No | null | - |
| cfFormat | Object \| null | No | null | - |

#### `rCellTextWithOverflow(r, c, x, y, w, h, cell, sheet, nonEmptyCells, cfFormat = null, tableStyle = null)`
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

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
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `rBorder(ciArr)`
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ciArr | Array<Object> | Yes | - | - |

#### `rBaseData(sheet)`
- Source: `src/core/Canvas/CellRenderer.js` | Category: `CellRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `getOffset(x, y, w, h): Object`
- Source: `src/core/Canvas/SheetRenderer.js` | Category: `SheetRenderer`

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
- Source: `src/core/Canvas/SheetRenderer.js` | Category: `SheetRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| nowRCount | number | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `Promise<void>`

#### `async captureScreenshot(cellAddress, csType = 'topleft', option = {}): Promise<string>`
- Source: `src/core/Canvas/ScreenshotRenderer.js` | Category: `ScreenshotRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellAddress | string | Yes | - | - |
| csType | string | No | 'topleft' | - |
| option | Object | No | {} | - |

**Returns**
- Type: `Promise<string>`

#### `getLight()`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

#### `rAllBtn()`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

#### `rRowColHeaders(sheet, lightInfo)`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| lightInfo | Object | Yes | - | - |

#### `rHighlightBars(sheet, lightInfo)`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| lightInfo | Object | Yes | - | - |

#### `rPageBreaks(sheet)`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |

#### `rIndex(sheet, isScreenshot = false, renderOptions = {})`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | - |
| isScreenshot | boolean | No | false | - |
| renderOptions | Object | No | {} | - |

#### `rDom()`
- Source: `src/core/Canvas/IndexRenderer.js` | Category: `IndexRenderer`

#### `showCellInput(clearValue = false, focusEditor = false)`
- Source: `src/core/Canvas/InputManager.js` | Category: `InputManager`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| clearValue | boolean | No | false | - |
| focusEditor | boolean | No | false | - |

#### `updateInputPosition()`
- Source: `src/core/Canvas/InputManager.js` | Category: `InputManager`

#### `updInputValue()`
- Source: `src/core/Canvas/InputManager.js` | Category: `InputManager`

#### `addEventListeners()`
- Source: `src/core/Canvas/EventManager.js` | Category: `EventManager`

#### `rSparklines(sheet)`
- Render sparklines within all views
- Source: `src/core/Canvas/SparklineRenderer.js` | Category: `SparklineRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Object | Yes | - | Sheet Objects |

#### `rCommentIndicators(sheet)`
- Render comment Indicator
- Source: `src/core/Canvas/CommentRenderer.js` | Category: `CommentRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `rCommentBoxes(sheet)`
- Render comment Floating Box
- Source: `src/core/Canvas/CommentRenderer.js` | Category: `CommentRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `rConditionalFormats(sheet)`
- Render Conditional Formatting (Icons Only)<br>The data bar has been rendered in rBaseData to avoid overwriting the text
- Source: `src/core/Canvas/CFRenderer.js` | Category: `CFRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | - | Yes | - | - |

#### `rDataBar(x, y, w, h, dataBar)`
- Render Data Bar (Export for CellRenderer)
- Source: `src/core/Canvas/CFRenderer.js` | Category: `CFRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | - | Yes | - | - |
| y | - | Yes | - | - |
| w | - | Yes | - | - |
| h | - | Yes | - | - |
| dataBar | - | Yes | - | - |

#### `rFilterIcons(sheet)`
- Source: `src/core/Canvas/FilterRenderer.js` | Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |

#### `clearFilterIconPositions()`
- Source: `src/core/Canvas/FilterRenderer.js` | Category: `FilterRenderer`

#### `getFilterIconAtPosition(x, y): {colIndex:number, scopeId:string} | null`
- Source: `src/core/Canvas/FilterRenderer.js` | Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| x | number | Yes | - | - |
| y | number | Yes | - | - |

**Returns**
- Type: `{colIndex:number, scopeId:string} | null`

#### `drawFilterIcon(ctx, iconX, iconY, iconSize, hasFilter, sortOrder, colors = null)`
- Source: `src/core/Canvas/FilterRenderer.js` | Category: `FilterRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| ctx | CanvasRenderingContext2D | Yes | - | - |
| iconX | number | Yes | - | - |
| iconY | number | Yes | - | - |
| iconSize | number | Yes | - | - |
| hasFilter | boolean | Yes | - | - |
| sortOrder | string \| null | Yes | - | - |
| colors | null | No | null | - |

#### `rPivotTablePlaceholders(sheet)`
- Render PivotTable Placeholder
- Source: `src/core/Canvas/PivotTableRenderer.js` | Category: `PivotTableRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |

#### `getPivotTableAtPosition(sheet, x, y)`
- Detect if the mouse is in the pivot table placeholder area
- Source: `src/core/Canvas/PivotTableRenderer.js` | Category: `PivotTableRenderer`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet Instance |
| x | number | Yes | - | Mouse X coordinates |
| y | number | Yes | - | Mouse Y coordinates<br>@ returns {PivotTable \| null} Returns null if a PivotTable instance is returned within an area |

### Potential Lifecycle Props
- Assigned outside the constructor. These are flagged for audit because they may be public state that the scanner should not silently miss.
| Name | Static | Source Line | Example Assignment |
| --- | --- | --- | --- |
| SKctx | No | 339 | this.scrollSkeletonLayer.getContext('2d') |

## FormulaEditor
- Source: `src/core/Canvas/FormulaEditor.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| canvas | Canvas | No | canvas | Canvas instance |
| SN | Object | No | canvas.SN | SheetNext Main Instance |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| active | boolean | No | Whether the formula editing mode is currently active |
| rangePicking | boolean | No | - |

### Methods
#### `beginSession(): void`
- Start a fresh editing session, called when the cell editor opens.

#### `syncFromInput(): void`
- Re-read the editor content and refresh coloring, overlays and highlights.

#### `onInput(): void`
- User typed into the cell editor.

#### `onExternalTextChange(): void`
- Editor content was mirrored from the formula bar, refresh without caret handling.

#### `handleKeyDown(event): boolean`
- Intercept keyboard events during formula editing.<br>Handles dropdown navigation, Tab-accept, Escape, F4 and arrow-key pointing.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | KeyboardEvent | Yes | - | Keydown event from the document handler |

**Returns**
- Type: `boolean`
- true when the event was consumed

#### `handleCanvasMousedown(event): boolean`
- Intercept a canvas mousedown during formula editing.<br>When the cursor is at a reference-insertable position the click/drag<br>inserts a cell/range/row/column reference instead of committing the edit.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | MouseEvent | Yes | - | Mousedown event on the handle layer |

**Returns**
- Type: `boolean`
- true when the event was consumed

#### `drawRefHighlights(sheet): void`
- Draw colored frames around referenced areas, called from the render pipeline.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | Sheet being rendered |

#### `onViewScrolled(): void`
- Reposition editor overlays after view scrolling during formula editing.

#### `deactivate(render = true): void`
- Leave formula mode and clear overlays/highlights.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| render | boolean | No | true | Whether to trigger a snapshot repaint |

## HyperlinkPopover
- Source: `src/core/Canvas/HyperlinkPopover.js`
