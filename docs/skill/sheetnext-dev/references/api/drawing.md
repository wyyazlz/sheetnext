# Drawing API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [Drawing](#drawing)
- [DrawingItem](#drawingitem)

## Drawing
- Source: `src/core/Drawing/Drawing.js`
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

## DrawingItem
- Source: `src/core/Drawing/DrawingItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| id | string | No | "rId" + ++sheet._idCount | - |
| inView | boolean | No | false | - |
| anchorType | string | No | config.anchorType ?? 'twoCell' | - |
| xmlDetail | Object \| null | No | config.xmlDetail ?? null | - |
| shapeAdjustments | Object<string, number> | No | parsed.shapeAdjustments | - |
| shapeType | string | No | parsed.shapeType | - |
| isConnector | boolean | No | CONNECTOR_TYPES.includes(parsed.shapeType) | - |
| isTextBox | boolean | No | this.isConnector ? false : (parsed.isTextBox ?? this.isTextBox) | - |
| shapeStyle | Object | No | { ...this.shapeStyle, ...parsed.shapeStyle } | - |
| shapeText | string | No | parsed.shapeText | - |
| type | - | No | == 'image' && this._imageBase64) { | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| imageBase64 | string | get/set | No | - |
| updRender | boolean | get/set | No | - |
| startCell | CellNum | get/set | No | - |
| drawCache | HTMLElement \| null | get/set | No | - |
| offsetX | number | get/set | No | - |
| offsetY | number | get/set | No | - |
| rotation | number | get/set | No | - |
| width | number | get/set | No | - |
| height | number | get/set | No | - |
| chartOption | Object \| null | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| type | 'chart' \| 'image' \| 'shape' \| 'slicer' | No | - |
| imageBytes | Object | No | - |
| imageSuffix | string \| null | No | - |
| area | Object | No | - |
| renderOption | Object \| null | No | - |

### Methods
#### `updIndex(type)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| type | 'up' \| 'down' \| 'top' \| 'bottom' | Yes | - | - |
