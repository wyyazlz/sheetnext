# Slicer API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Slicer](#slicer)
- [SlicerItem](#sliceritem)

## Slicer
- Source: `src/core/Slicer/Slicer.js`
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

## SlicerItem
- Source: `src/core/Slicer/SlicerItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| id | string | No | options.id ?? `Slicer_${Date.now()}` | - |
| name | string | No | options.name ?? this.id | - |
| caption | string | No | options.caption ?? options.fieldName ?? '' | - |
| cacheId | string | No | options.cacheId | - |
| sourceType | string | No | options.sourceType ?? 'table' | - |
| sourceSheet | string | No | options.sourceSheet ?? sheet.name | - |
| sourceName | string \| null | No | options.sourceName ?? null | - |
| sourceId | string \| null | No | options.sourceId ?? null | - |
| fieldIndex | number | No | Number.isFinite(options.fieldIndex) ? options.fieldIndex : 0 | - |
| fieldName | string | No | options.fieldName ?? '' | - |
| columns | number | No | Math.max(1, Number(options.columns) \|\| 1) | - |
| buttonHeight | number | No | Math.max(18, Number(options.buttonHeight) \|\| 20) | - |
| showHeader | boolean | No | options.showHeader !== false | - |
| multiSelect | boolean | No | !!options.multiSelect | - |
| styleName | string | No | options.styleName ?? 'SlicerStyleLight1' | - |
| Drawing | DrawingItem \| null | No | options.Drawing \|\| options.drawing \|\| null | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| selectedKeys | Iterable<string> | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| cache | SlicerCache \| null | No | - |

### Methods
#### `toggleSelection(key, multi = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| key | string | Yes | - | - |
| multi | boolean | No | false | - |

#### `clearSelection(apply = false)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| apply | boolean | No | false | - |

#### `applyFilter()`

#### `toJSON(): Object`

**Returns**
- Type: `Object`

#### `render(canvas, x, y, w, h, inView = true)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| canvas | Object | Yes | - | - |
| x | number | Yes | - | - |
| y | number | Yes | - | - |
| w | number | Yes | - | - |
| h | number | Yes | - | - |
| inView | boolean | No | true | - |
