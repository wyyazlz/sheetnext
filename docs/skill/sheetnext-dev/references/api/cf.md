# CF API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [CF](#cf)
- [CFRule](#cfrule)

## CF
- Source: `src/core/CF/CF.js`
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

## CFRule
- Source: `src/core/CF/CFRule.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| type | string | No | config.type | - |
| priority | number | No | config.priority | - |
| sqref | string | No | config.sqref | - |
| stopIfTrue | boolean | No | config.stopIfTrue \|\| false | - |
| dxf | Object \| null | No | config.dxf \|\| null | - |
| needsRangeData | boolean | No | false | - |

### Methods
#### `evaluate(cell, r, c, rangeData): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | Cell | Yes | - | - |
| r | number | Yes | - | - |
| c | number | Yes | - | - |
| rangeData | Object | Yes | - | - |

**Returns**
- Type: `boolean`

#### `getFormat(cell, r, c, rangeData): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cell | Cell | Yes | - | - |
| r | number | Yes | - | - |
| c | number | Yes | - | - |
| rangeData | Object | Yes | - | - |

**Returns**
- Type: `Object | null`

#### `getRangeData(): Object`

**Returns**
- Type: `Object`

#### `resolveFormula(formula, r, c): *`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formula | string \| number | Yes | - | - |
| r | number | Yes | - | - |
| c | number | Yes | - | - |

**Returns**
- Type: `*`
