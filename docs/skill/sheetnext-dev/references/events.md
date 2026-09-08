# Event API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

## Subscription surface

Use `SN.Event` for registration and dispatch. `SN.Event` is an `EventEmitter` instance exposed on the root workbook object.

```js
const preventNegative = (event) => {
  if (Number(event.data.newValue) < 0) event.cancel('Use a non-negative value');
};
SN.Event.on('beforeCellEdit', preventNegative);

SN.Event.once('afterSheetAdd', (event) => {
  console.log(event.data.sheet);
});
SN.Event.off('beforeCellEdit', preventNegative);
```

The last line removes this example handler; call it when the integration no longer needs the subscription. `on` returns the event emitter, not an unsubscribe function. Do not use `offAll()` for component cleanup: it also removes editor-owned listeners.

Handlers for `emit` must cancel synchronously. `emitAsync` can await handlers and `waitUntil`; this does not make every calling editor API asynchronous. Client hooks are not server authorization.

## EventEmitter
- Source: `src/core/Event/EventEmitter.js`
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
| options.id | string | No | - | Unique identity (anti duplicate binding) |
| options.priority | number | No | 0 | Priority (execute as soon as possible) |
| options.once | boolean | No | false | Whether to execute only once |

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

#### `off(event, idOrHandler?): EventEmitter`
- Remove Event Listener

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | Yes | - | Event Name |
| idOrHandler | string \| Function | No | - | Listener id or handler reference; omit to remove all listeners for this event. |

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

#### `listenerCount(event?): number`
- Get the number of listeners

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | string | No | - | Event name; omit to return the total listener count. |

**Returns**
- Type: `number`

#### `listEvents(): Array<{event: string, count: number}>`
- Get all event names and number of monitors

**Returns**
- Type: `Array<{event: string, count: number}>`

## SheetEvent
- Source: `src/core/Event/EventEmitter.js`
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

- Total detected events: 122
- Dynamic event patterns: 7
- Cancelable `Yes` means a returned event is checked at an observed call site; `Unknown` needs source review, and `No` denotes an after-event with no detected cancellation check. Conditional operation hooks still depend on caller options.
- Template names containing `${...}` are source patterns, not literal subscription names. Payload keys only list statically visible fields; spreads and variables may contain more data.
- Persistence coverage and explicit dirty tracking are described in [integration.md](integration.md).

### Col
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| after${eventBase} | emit | No | col, newHidden, oldHidden, row, sheet | src/core/Col/Col.js:230, src/core/Row/Row.js:256, src/core/Sheet/Sheet.js:425 |
| afterColResize | emit | No | col, newWidth, oldWidth, sheet | src/core/Col/Col.js:181 |
| before${eventBase} | emit | Yes | col, newHidden, oldHidden, row, sheet | src/core/Col/Col.js:202, src/core/Row/Row.js:228, src/core/Sheet/Sheet.js:415 |
| beforeColResize | emit | Yes | col, newWidth, oldWidth, sheet | src/core/Col/Col.js:153 |

### Sheet
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| after${eventSuffix} | emit | No | [indexKey], changedCount, limitedCount, maxLevel, sheet | src/core/Sheet/Outline.js:150 |
| afterClearCellControl | emit | No | ...summary, sheet | src/core/Sheet/CellControl.js:277 |
| afterClearDataValidation | emit | No | ...summary, sheet | src/core/Sheet/DataValidation.js:171 |
| afterClearRange | emit | No | ...summary, sheet | src/core/Sheet/ClearRange.js:552 |
| afterConsolidate | emit | No | ...summary, sheet | src/core/Sheet/Consolidate.js:210 |
| afterDeleteColumns | emit | No | col, count, sheet | src/core/Sheet/RowColHandler.js:912 |
| afterDeleteRows | emit | No | count, row, sheet | src/core/Sheet/RowColHandler.js:771 |
| afterFlashFill | emit | No | exampleCount, filledCount, sheet, sourceRange, strategy, targetRange | src/core/Sheet/FlashFill.js:313 |
| afterInsertColumns | emit | No | col, count, sheet | src/core/Sheet/RowColHandler.js:637 |
| afterInsertRows | emit | No | count, row, sheet | src/core/Sheet/RowColHandler.js:414 |
| afterMerge | emit | No | areas, mode, sheet | src/core/Sheet/Merge.js:207, src/core/Sheet/Merge.js:238 |
| afterMoveArea | emit | No | moveArea, sheet, targetArea | src/core/Sheet/Sheet.js:819 |
| afterSelectionChange | emit | No | newAreas, newCell, oldAreas, oldCell, sheet, type | src/core/Sheet/Sheet.js:147, src/core/Sheet/Sheet.js:182 |
| afterSetCellControl | emit | No | ...summary, control, sheet | src/core/Sheet/CellControl.js:250 |
| afterSetDataValidation | emit | No | ...summary, sheet | src/core/Sheet/DataValidation.js:139 |
| afterSheetRename | emit | No | newName, oldName, sheet | src/core/Sheet/Sheet.js:449 |
| afterSort | emit | No | options, range, sheet, sortKeys, success | src/core/Sheet/Sort.js:19 |
| afterSubtotal | emit | No | ...summary, options, sheet | src/core/Sheet/Subtotal.js:302 |
| afterTextToColumns | emit | No | columnCount, rowCount, sheet, sourceRange, targetRange | src/core/Sheet/TextToColumns.js:189 |
| afterUnmerge | emit | No | cellAd, sheet | src/core/Sheet/Merge.js:78, src/core/Sheet/Merge.js:91 |
| before${eventSuffix} | emit | Yes | [indexKey], areas, maxLevel, sheet | src/core/Sheet/Outline.js:89 |
| beforeClearCellControl | emit | Yes | areas, sheet | src/core/Sheet/CellControl.js:260 |
| beforeClearDataValidation | emit | Yes | areas, sheet | src/core/Sheet/DataValidation.js:148 |
| beforeClearRange | emit | Yes | areas, options, sheet | src/core/Sheet/ClearRange.js:500 |
| beforeConsolidate | emit | Yes | functionName, sheet, skipBlanks, sourceRanges, targetRange | src/core/Sheet/Consolidate.js:149 |
| beforeDeleteColumns | emit | Yes | col, count, sheet | src/core/Sheet/RowColHandler.js:792 |
| beforeDeleteRows | emit | Yes | count, row, sheet | src/core/Sheet/RowColHandler.js:658 |
| beforeFlashFill | emit | Yes | exampleCount, sheet, sourceRange, strategy, targetRange | src/core/Sheet/FlashFill.js:276 |
| beforeInsertColumns | emit | Yes | col, count, sheet | src/core/Sheet/RowColHandler.js:435 |
| beforeInsertRows | emit | Yes | count, row, sheet | src/core/Sheet/RowColHandler.js:254 |
| beforeMerge | emit | Yes | areas, mode, sheet | src/core/Sheet/Merge.js:158 |
| beforeMoveArea | emit | Yes | moveArea, sheet, targetArea | src/core/Sheet/Sheet.js:775 |
| beforeSelectionChange | emit | Yes | newAreas, newCell, oldAreas, oldCell, sheet, type | src/core/Sheet/Sheet.js:135, src/core/Sheet/Sheet.js:170 |
| beforeSetCellControl | emit | Yes | areas, control, sheet | src/core/Sheet/CellControl.js:228 |
| beforeSetDataValidation | emit | Yes | areas, rule, sheet | src/core/Sheet/DataValidation.js:112 |
| beforeSheetRename | emit | Yes | newName, oldName, sheet | src/core/Sheet/Sheet.js:443 |
| beforeSort | emit | Yes | options, range, sheet, sortKeys | src/core/Sheet/Sort.js:11 |
| beforeSubtotal | emit | Yes | options, range, sheet | src/core/Sheet/Subtotal.js:183 |
| beforeTextToColumns | emit | Yes | options, sheet, sourceRange, targetRange | src/core/Sheet/TextToColumns.js:146 |
| beforeUnmerge | emit | Yes | cellAd, sheet | src/core/Sheet/Merge.js:49 |

### Workbook
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterActiveSheetChange | emit | No | newSheet, oldSheet | src/core/Workbook/Workbook.js:168 |
| afterOperation | emitAsync | No | - | src/core/Workbook/Workbook.js:574 |
| afterOperation:${op.type} | emitAsync | No | - | src/core/Workbook/Workbook.js:575 |
| afterSheetAdd | emit | No | name, sheet | src/core/Workbook/Workbook.js:227 |
| afterSheetDelete | emit | No | index, name, sheet | src/core/Workbook/Workbook.js:246 |
| afterSheetMove | emit | No | fromIndex, sheet, toIndex | src/core/Workbook/Workbook.js:282 |
| afterWorkbookRename | emit | No | newName, oldName | src/core/Workbook/Workbook.js:195 |
| beforeActiveSheetChange | emit | Yes | newSheet, oldSheet | src/core/Workbook/Workbook.js:158 |
| beforeOperation | emit | Yes | ...data, operation | src/core/Workbook/Workbook.js:555 |
| beforeOperation:${type} | emit | Yes | ...data, operation | src/core/Workbook/Workbook.js:556 |
| beforeSheetAdd | emit | Yes | name | src/core/Workbook/Workbook.js:219 |
| beforeSheetDelete | emit | Yes | index, name, sheet | src/core/Workbook/Workbook.js:239 |
| beforeSheetMove | emit | Yes | fromIndex, sheet, toIndex | src/core/Workbook/Workbook.js:273 |
| beforeWorkbookRename | emit | Yes | newName, oldName | src/core/Workbook/Workbook.js:187 |

### AI
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterAIRequest | emitAsync | No | durationMs, request, response, result, status, textBlock, ui | src/core/AI/AI.js:827 |
| aiRequestChunk | emit | Unknown | chunk, chunkIndex, request, textBlock, ui | src/core/AI/AI.js:1018 |
| aiRequestError | emitAsync | Unknown | durationMs, error, request, response, status, textBlock, ui | src/core/AI/AI.js:844 |
| aiRequestFinally | emitAsync | Unknown | durationMs, error, request, response, result, status, textBlock, ui | src/core/AI/AI.js:860 |
| aiRequestStart | emit | Unknown | request, textBlock, ui | src/core/AI/AI.js:791 |
| beforeAIRequest | emitAsync | Unknown | request, textBlock, ui | src/core/AI/AI.js:778 |

### AutoFilter
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterAutoFilterApply | emit | No | ...eventBase, allHiddenRowsAdded, allHiddenRowsRemoved, hiddenRowsAdded, hiddenRowsRemoved, range, sheet | src/core/AutoFilter/AutoFilter.js:1191 |
| afterAutoFilterClear | emit | No | ...eventBase, range, ref, sheet | src/core/AutoFilter/AutoFilter.js:1133 |
| afterAutoFilterClearAllFilters | emit | No | ...eventBase, sheet | src/core/AutoFilter/AutoFilter.js:1068 |
| afterAutoFilterClearColumnFilter | emit | No | ...eventBase, colIndex, prevFilter, sheet | src/core/AutoFilter/AutoFilter.js:1018 |
| afterAutoFilterSetColumnFilter | emit | No | ...eventBase, colIndex, filter, prevFilter, sheet | src/core/AutoFilter/AutoFilter.js:963 |
| afterAutoFilterSetRange | emit | No | ...eventBase, range, sheet | src/core/AutoFilter/AutoFilter.js:777 |
| beforeAutoFilter | emit | Yes | ...eventBase, action, colIndex, filter, range, sheet | src/core/AutoFilter/AutoFilter.js:1039, src/core/AutoFilter/AutoFilter.js:1087, src/core/AutoFilter/AutoFilter.js:749, src/core/AutoFilter/AutoFilter.js:924, src/core/AutoFilter/AutoFilter.js:986 |
| beforeAutoFilterApply | emit | Yes | ...eventBase, columns, range, sheet | src/core/AutoFilter/AutoFilter.js:1204 |
| beforeAutoFilterClear | emit | Yes | ...eventBase, range, ref, sheet | src/core/AutoFilter/AutoFilter.js:1099 |
| beforeAutoFilterClearAllFilters | emit | Yes | ...eventBase, filters, sheet | src/core/AutoFilter/AutoFilter.js:1048 |
| beforeAutoFilterClearColumnFilter | emit | Yes | ...eventBase, colIndex, prevFilter, sheet | src/core/AutoFilter/AutoFilter.js:996 |
| beforeAutoFilterSetColumnFilter | emit | Yes | ...eventBase, colIndex, filter, prevFilter, sheet | src/core/AutoFilter/AutoFilter.js:935 |
| beforeAutoFilterSetRange | emit | Yes | ...eventBase, range, sheet | src/core/AutoFilter/AutoFilter.js:758 |

### Cell
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterCellEdit | emit | No | cell, col, newValue, oldValue, row, sheet | src/core/Cell/Cell.js:378, src/core/Sheet/CellControl.js:135 |
| afterCellStyleChange | emit | No | cell, col, newStyle, oldStyle, row, sheet | src/core/Cell/Cell.js:1002, src/core/Cell/Cell.js:368, src/core/Cell/Cell.js:460, src/core/Cell/Cell.js:876, src/core/Sheet/CellControl.js:111 |
| afterHyperlinkInsert | emit | No | cell, col, hyperlink, row, sheet | src/core/Cell/Cell.js:764 |
| beforeCellEdit | emit | Yes | cell, col, newValue, oldValue, row, sheet | src/core/Cell/Cell.js:189, src/core/Sheet/CellControl.js:123 |
| beforeCellStyleChange | emit | Yes | cell, col, newStyle, oldStyle, row, sheet | src/core/Cell/Cell.js:203, src/core/Cell/Cell.js:420, src/core/Cell/Cell.js:858, src/core/Cell/Cell.js:986, src/core/Sheet/CellControl.js:98 |
| beforeHyperlinkInsert | emit | Yes | cell, col, hyperlink, row, sheet | src/core/Cell/Cell.js:741 |

### Drawing
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterDrawingAdd | emit | No | config, drawing, sheet | src/core/Drawing/Drawing.js:65 |
| afterDrawingRemove | emit | No | drawing, id, sheet | src/core/Drawing/Drawing.js:85 |
| beforeDrawingAdd | emit | Yes | config, sheet | src/core/Drawing/Drawing.js:53 |
| beforeDrawingRemove | emit | Yes | drawing, id, sheet | src/core/Drawing/Drawing.js:75 |

### PivotTable
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterPivotTableAdd | emit | No | config, pivotTable, sheet | src/core/PivotTable/PivotTable.js:87 |
| afterPivotTableChange | emit | No | action, pivotTable, sheet, sourceRangeRef, sourceSheet | src/core/PivotTable/PivotTableItem.js:3917, src/core/PivotTable/PivotTableItem.js:3941, src/core/PivotTable/PivotTableItem.js:945 |
| beforePivotTableAdd | emit | Yes | config, sheet | src/core/PivotTable/PivotTable.js:55 |
| beforePivotTableChange | emit | Yes | action, pivotTable, sheet, sourceRangeRef, sourceSheet | src/core/PivotTable/PivotTableItem.js:3899, src/core/PivotTable/PivotTableItem.js:3928, src/core/PivotTable/PivotTableItem.js:772 |

### UndoRedo
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterRedo | emit | No | tx | src/core/UndoRedo/UndoRedo.js:134 |
| afterUndo | emit | No | tx | src/core/UndoRedo/UndoRedo.js:105 |
| beforeRedo | emit | Yes | tx | src/core/UndoRedo/UndoRedo.js:116 |
| beforeUndo | emit | Yes | tx | src/core/UndoRedo/UndoRedo.js:87 |

### Row
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterRowResize | emit | No | newHeight, oldHeight, row, sheet | src/core/Row/Row.js:207 |
| beforeRowResize | emit | Yes | newHeight, oldHeight, row, sheet | src/core/Row/Row.js:179 |

### Slicer
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterSlicerAdd | emit | No | config, sheet, slicer | src/core/Slicer/Slicer.js:87 |
| afterSlicerRemove | emit | No | id, sheet, slicer | src/core/Slicer/Slicer.js:132 |
| beforeSlicerAdd | emit | Yes | config, sheet | src/core/Slicer/Slicer.js:36 |
| beforeSlicerRemove | emit | Yes | id, sheet, slicer | src/core/Slicer/Slicer.js:112 |

### Sparkline
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterSparklineAdd | emit | No | config, sheet, sparkline | src/core/Sparkline/Sparkline.js:105 |
| afterSparklineRemove | emit | No | c, r, sheet, sparkline | src/core/Sparkline/Sparkline.js:138 |
| beforeSparklineAdd | emit | Yes | config, sheet | src/core/Sparkline/Sparkline.js:83 |
| beforeSparklineRemove | emit | Yes | c, r, sheet, sparkline | src/core/Sparkline/Sparkline.js:126 |

### Table
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterTableAdd | emit | No | options, sheet, table | src/core/Table/Table.js:135 |
| afterTableAutoFilterToggle | emit | No | newValue, oldValue, sheet, table | src/core/Table/TableItem.js:254 |
| afterTableRemove | emit | No | id, sheet, table | src/core/Table/Table.js:217 |
| afterTableRename | emit | No | newName, oldName, sheet, table | src/core/Table/TableItem.js:129 |
| afterTableResize | emit | No | newRef, oldRef, sheet, table | src/core/Table/TableItem.js:487 |
| afterTableStyleChange | emit | No | newStyle, oldStyle, sheet, table | src/core/Table/TableItem.js:175 |
| afterTableTotalsRowToggle | emit | No | newValue, oldValue, sheet, table | src/core/Table/TableItem.js:220 |
| beforeTableAdd | emit | Yes | options, sheet | src/core/Table/Table.js:103 |
| beforeTableAutoFilterToggle | emit | Yes | newValue, oldValue, sheet, table | src/core/Table/TableItem.js:239 |
| beforeTableRemove | emit | Yes | id, sheet, table | src/core/Table/Table.js:208 |
| beforeTableRename | emit | Yes | newName, oldName, sheet, table | src/core/Table/TableItem.js:115 |
| beforeTableResize | emit | Yes | newRef, oldRef, sheet, table | src/core/Table/TableItem.js:453 |
| beforeTableStyleChange | emit | Yes | newStyle, oldStyle, sheet, table | src/core/Table/TableItem.js:164 |
| beforeTableTotalsRowToggle | emit | Yes | newValue, oldValue, sheet, table | src/core/Table/TableItem.js:197 |

### Canvas
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| afterZoomChange | emit | No | newZoom, oldZoom, sheet | src/core/Canvas/ScrollInteraction.js:271, src/core/Sheet/Sheet.js:211 |
| beforeZoomChange | emit | Yes | newZoom, oldZoom, sheet | src/core/Canvas/ScrollInteraction.js:264, src/core/Sheet/Sheet.js:201 |

### Comment
| Event | Mode | Cancelable | Payload Keys | Source |
| --- | --- | --- | --- | --- |
| event | emit | Unknown | - | src/core/Comment/Comment.js:6, src/core/Sheet/Clipboard.js:15 |
