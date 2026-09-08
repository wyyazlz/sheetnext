# Layout API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Layout](#layout)

## Layout
- Source: `src/core/Layout/Layout.js`
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
- Source: `src/core/Layout/ResizeHandler.js` | Category: `ResizeHandler`

#### `throttleResize()`
- Throttling resize event
- Source: `src/core/Layout/ResizeHandler.js` | Category: `ResizeHandler`

#### `handleAllResize()`
- Handle all resize related logic
- Source: `src/core/Layout/ResizeHandler.js` | Category: `ResizeHandler`

#### `updateCanvasSize()`
- Update Canvas Dimensions
- Source: `src/core/Layout/ResizeHandler.js` | Category: `ResizeHandler`

#### `initDragEvents()`
- Initialize Drag Event
- Source: `src/core/Layout/DragHandler.js` | Category: `DragHandler`

#### `initEventListeners()`
- Initialize all event listeners.
- Source: `src/core/Layout/EventManager.js` | Category: `EventManager`

#### `initColorComponents(scope?)`
- Initialize color components.
- Source: `src/core/Layout/EventManager.js` | Category: `EventManager`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scope | Element | No | - | Optional scope. |

#### `removeLoading()`
- Remove loading animation.
- Source: `src/core/Layout/EventManager.js` | Category: `EventManager`
