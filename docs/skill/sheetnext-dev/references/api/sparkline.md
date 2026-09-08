# Sparkline API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Sparkline](#sparkline)
- [SparklineItem](#sparklineitem)

## Sparkline
- Source: `src/core/Sparkline/Sparkline.js`
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

## SparklineItem
- Source: `src/core/Sparkline/SparklineItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| r | number | No | config.r | - |
| c | number | No | config.c | - |
| sqref | string | No | config.sqref | - |
| formula | string | No | config.formula | - |
| group | Object | No | config.group | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| type | 'line' \| 'column' \| 'stacked' | get/set | No | - |
| colors | Object | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| key | string | No | - |

### Methods
#### `getData(): Array<number | null>`

**Returns**
- Type: `Array<number | null>`
