# Event API

> Generated: 2026-03-17

## Subscription surface

Use `SN.Event` for registration and dispatch. `SN.Event` is an `EventEmitter` instance exposed on the root workbook object.

```js
SN.Event.on('beforeCellEdit', (e) => {
  if (e.data.cell?.locked) e.cancel('Cell is locked')
})

SN.Event.once('afterSheetAdd', (e) => {
  console.log(e.data.sheet)
})
```

## EventEmitter
- Event Transmitter

### Methods
#### `on(event, handler, options = {}): EventEmitter`
- Register Event Listener

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |
| handler | Function | Yes | - | Handling function |
| options | Object | No | {} | Options |
| options.id | string | Yes | - | Unique identity (anti duplicate binding) |
| options.priority | number | Yes | - | Priority (execute as soon as possible) |
| options.once | boolean | Yes | - | Whether to execute only once |

**Returns**
- Type: `EventEmitter`
- Chain Calls Supported

#### `once(event, handler, options = {}): EventEmitter`
- Sign up for a one-time event listener

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |
| handler | Function | Yes | - | Handling function |
| options | Object | No | {} | Options |

**Returns**
- Type: `EventEmitter`

#### `off(event, idOrHandler): EventEmitter`
- Remove Event Listener

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |
| idOrHandler | string \| Function | Yes | - | id or handler reference |

**Returns**
- Type: `EventEmitter`

#### `offAll(): EventEmitter`
- Remove all event listeners

**Returns**
- Type: `EventEmitter`

#### `emit(event, data = {}): SheetEvent`
- Trigger event

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |
| data | Object | No | {} | Event Data |

**Returns**
- Type: `SheetEvent`
- Event object (canceled status can be checked)

#### `async emitAsync(event, data = {}): Promise<SheetEvent>`
- Trigger event with async hooks support.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | event name |
| data | Object | No | {} | event data |

**Returns**
- Type: `Promise<SheetEvent>`

#### `hasListeners(event): boolean`
- Check for listeners

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |

**Returns**
- Type: `boolean`

#### `listenerCount(event): number`
- Get the number of listeners

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event name (optional, return total if not passed) |

**Returns**
- Type: `number`

#### `listEvents(): Array<{event: string, count: number}>`
- Get all event names and number of monitors

**Returns**
- Type: `Array<{event: string, count: number}>`

## SheetEvent
- Event object

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| type | string | No | type | - |
| data | Object | No | data | - |
| target | Object | No | target | - |
| canceled | boolean | No | false | - |
| stopped | boolean | No | false | - |
| cancelReason | null | No | null | - |
| timestamp | number | No | Date.now() | - |

### Methods
#### `cancel(reason = '')`
- Cancel event (block default behavior)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| reason | string | No | '' | Reason for canceling (optional, used as a reminder) |

#### `stopPropagation()`
- Stop event propagation (no longer performing follow-up listener)

#### `waitUntil(promise)`
- Register async work for emitAsync to await.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| promise | Promise | Yes | - | - |

## Emitted events

- Total detected events: 114
- Dynamic event patterns: 7

### Col
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| after${eventBase} | emit | No | col, newHidden, oldHidden, row, sheet |
| afterColResize | emit | No | col, newWidth, oldWidth, sheet |
| before${eventBase} | emit | No | col, newHidden, oldHidden, row, sheet |
| beforeColResize | emit | No | col, newWidth, oldWidth, sheet |

### Sheet
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| after${eventSuffix} | emit | No | [indexKey], changedCount, limitedCount, maxLevel, sheet |
| afterClearDataValidation | emit | No | ...summary, sheet |
| afterConsolidate | emit | No | ...summary, sheet |
| afterDeleteColumns | emit | No | col, count, sheet |
| afterDeleteRows | emit | No | count, row, sheet |
| afterFlashFill | emit | No | exampleCount, filledCount, sheet, sourceRange, strategy, targetRange |
| afterInsertColumns | emit | No | col, count, sheet |
| afterInsertRows | emit | No | count, row, sheet |
| afterMerge | emit | No | areas, mode, sheet |
| afterMoveArea | emit | No | moveArea, sheet, targetArea |
| afterSelectionChange | emit | No | newAreas, newCell, oldAreas, oldCell, sheet, type |
| afterSetDataValidation | emit | No | ...summary, sheet |
| afterSheetRename | emit | No | newName, oldName, sheet |
| afterSort | emit | No | range, sheet, sortKeys, success |
| afterSubtotal | emit | No | ...summary, options, sheet |
| afterTextToColumns | emit | No | columnCount, rowCount, sheet, sourceRange, targetRange |
| afterUnmerge | emit | No | cellAd, sheet |
| afterZoomChange | emit | No | newZoom, oldZoom, sheet |
| before${eventSuffix} | emit | No | [indexKey], areas, maxLevel, sheet |
| beforeClearDataValidation | emit | No | areas, sheet |
| beforeConsolidate | emit | No | functionName, sheet, skipBlanks, sourceRanges, targetRange |
| beforeDeleteColumns | emit | No | col, count, sheet |
| beforeDeleteRows | emit | No | count, row, sheet |
| beforeFlashFill | emit | No | exampleCount, sheet, sourceRange, strategy, targetRange |
| beforeInsertColumns | emit | No | col, count, sheet |
| beforeInsertRows | emit | No | count, row, sheet |
| beforeMerge | emit | No | areas, mode, sheet |
| beforeMoveArea | emit | No | moveArea, sheet, targetArea |
| beforeSelectionChange | emit | No | newAreas, newCell, oldAreas, oldCell, sheet, type |
| beforeSetDataValidation | emit | No | areas, rule, sheet |
| beforeSheetRename | emit | No | newName, oldName, sheet |
| beforeSort | emit | No | range, sheet, sortKeys |
| beforeSubtotal | emit | No | options, range, sheet |
| beforeTextToColumns | emit | No | options, sheet, sourceRange, targetRange |
| beforeUnmerge | emit | No | cellAd, sheet |
| beforeZoomChange | emit | No | newZoom, oldZoom, sheet |

### Workbook
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterActiveSheetChange | emit | No | newSheet, oldSheet |
| afterOperation | emitAsync | No | - |
| afterOperation:${op.type} | emitAsync | No | - |
| afterSheetAdd | emit | No | name, sheet |
| afterSheetDelete | emit | No | index, name, sheet |
| afterWorkbookRename | emit | No | newName, oldName |
| beforeActiveSheetChange | emit | No | newSheet, oldSheet |
| beforeOperation | emit | No | ...data, operation |
| beforeOperation:${type} | emit | No | ...data, operation |
| beforeSheetAdd | emit | No | name |
| beforeSheetDelete | emit | No | index, name, sheet |
| beforeWorkbookRename | emit | No | newName, oldName |

### AI
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterAIRequest | emitAsync | No | durationMs, request, response, result, status, textBlock, ui |
| aiRequestChunk | emit | No | chunk, chunkIndex, request, textBlock, ui |
| aiRequestError | emitAsync | No | durationMs, error, request, response, status, textBlock, ui |
| aiRequestFinally | emitAsync | No | durationMs, error, request, response, result, status, textBlock, ui |
| aiRequestStart | emit | No | request, textBlock, ui |
| beforeAIRequest | emitAsync | No | request, textBlock, ui |

### AutoFilter
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterAutoFilterApply | emit | No | ...eventBase, allHiddenRowsAdded, allHiddenRowsRemoved, hiddenRowsAdded, hiddenRowsRemoved, range, sheet |
| afterAutoFilterClear | emit | No | ...eventBase, range, ref, sheet |
| afterAutoFilterClearAllFilters | emit | No | ...eventBase, sheet |
| afterAutoFilterClearColumnFilter | emit | No | ...eventBase, colIndex, prevFilter, sheet |
| afterAutoFilterSetColumnFilter | emit | No | ...eventBase, colIndex, filter, prevFilter, sheet |
| afterAutoFilterSetRange | emit | No | ...eventBase, range, sheet |
| beforeAutoFilter | emit | No | ...eventBase, action, colIndex, filter, range, sheet |
| beforeAutoFilterApply | emit | No | ...eventBase, columns, range, sheet |
| beforeAutoFilterClear | emit | No | ...eventBase, range, ref, sheet |
| beforeAutoFilterClearAllFilters | emit | No | ...eventBase, filters, sheet |
| beforeAutoFilterClearColumnFilter | emit | No | ...eventBase, colIndex, prevFilter, sheet |
| beforeAutoFilterSetColumnFilter | emit | No | ...eventBase, colIndex, filter, prevFilter, sheet |
| beforeAutoFilterSetRange | emit | No | ...eventBase, range, sheet |

### Cell
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterCellEdit | emit | No | cell, col, newValue, oldValue, row, sheet |
| afterCellStyleChange | emit | No | cell, col, newStyle, oldStyle, row, sheet |
| afterHyperlinkInsert | emit | No | cell, col, hyperlink, row, sheet |
| beforeCellEdit | emit | No | cell, col, newValue, oldValue, row, sheet |
| beforeCellStyleChange | emit | No | cell, col, newStyle, oldStyle, row, sheet |
| beforeHyperlinkInsert | emit | No | cell, col, hyperlink, row, sheet |

### Drawing
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterDrawingAdd | emit | No | config, drawing, sheet |
| afterDrawingRemove | emit | No | drawing, id, sheet |
| beforeDrawingAdd | emit | No | config, sheet |
| beforeDrawingRemove | emit | No | drawing, id, sheet |

### PivotTable
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterPivotTableAdd | emit | No | config, pivotTable, sheet |
| afterPivotTableChange | emit | No | action, pivotTable, sheet, sourceRangeRef, sourceSheet |
| beforePivotTableAdd | emit | No | config, sheet |
| beforePivotTableChange | emit | No | action, pivotTable, sheet, sourceRangeRef, sourceSheet |

### UndoRedo
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterRedo | emit | No | tx |
| afterUndo | emit | No | tx |
| beforeRedo | emit | No | tx |
| beforeUndo | emit | No | tx |

### Row
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterRowResize | emit | No | newHeight, oldHeight, row, sheet |
| beforeRowResize | emit | No | newHeight, oldHeight, row, sheet |

### Slicer
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterSlicerAdd | emit | No | config, sheet, slicer |
| afterSlicerRemove | emit | No | id, sheet, slicer |
| beforeSlicerAdd | emit | No | config, sheet |
| beforeSlicerRemove | emit | No | id, sheet, slicer |

### Sparkline
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterSparklineAdd | emit | No | config, sheet, sparkline |
| afterSparklineRemove | emit | No | c, r, sheet, sparkline |
| beforeSparklineAdd | emit | No | config, sheet |
| beforeSparklineRemove | emit | No | c, r, sheet, sparkline |

### Table
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| afterTableAdd | emit | No | options, sheet, table |
| afterTableAutoFilterToggle | emit | No | newValue, oldValue, sheet, table |
| afterTableRemove | emit | No | id, sheet, table |
| afterTableRename | emit | No | newName, oldName, sheet, table |
| afterTableResize | emit | No | newRef, oldRef, sheet, table |
| afterTableStyleChange | emit | No | newStyle, oldStyle, sheet, table |
| afterTableTotalsRowToggle | emit | No | newValue, oldValue, sheet, table |
| beforeTableAdd | emit | No | options, sheet |
| beforeTableAutoFilterToggle | emit | No | newValue, oldValue, sheet, table |
| beforeTableRemove | emit | No | id, sheet, table |
| beforeTableRename | emit | No | newName, oldName, sheet, table |
| beforeTableResize | emit | No | newRef, oldRef, sheet, table |
| beforeTableStyleChange | emit | No | newStyle, oldStyle, sheet, table |
| beforeTableTotalsRowToggle | emit | No | newValue, oldValue, sheet, table |

### Comment
| Event | Mode | Cancelable | Payload Keys |
| --- | --- | --- | --- |
| event | emit | No | - |
